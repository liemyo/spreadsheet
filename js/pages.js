// Page layout for printing / PDF: paper, margins, scale, title rows, print area and page breaks
// (automatic, manual, and "keep together" ranges that move a table to the next page instead of cutting it).
import { el, icon, esc, colName, rangeAddr, keyR, keyC } from './util.js';
import { dialog, field, select, checkbox, toast, segmented } from './ui/ui.js';
import { sheetHTML } from './io.js';

export const PAPER = {
  a4: { n: 'A4', mm: [210, 297] },
  f4: { n: 'F4 / Folio', mm: [215, 330] },
  letter: { n: 'Letter', mm: [215.9, 279.4] },
  legal: { n: 'Legal', mm: [215.9, 355.6] },
  a3: { n: 'A3', mm: [297, 420] },
  a5: { n: 'A5', mm: [148, 210] },
  b5: { n: 'B5', mm: [176, 250] },
  tabloid: { n: 'Tabloid', mm: [279.4, 431.8] },
};
// [id, name, [top, bottom, left, right] cm]
export const MARGINS = [
  ['normal', 'Normal', [1.91, 1.91, 1.78, 1.78]],
  ['narrow', 'Sempit', [1.91, 1.91, 0.64, 0.64]],
  ['wide', 'Lebar', [2.54, 2.54, 2.54, 2.54]],
  ['none', 'Tanpa margin', [0.5, 0.5, 0.5, 0.5]],
  ['skripsi', 'Skripsi (4-3-3-4)', [4, 3, 4, 3]],
];
const DEF = { size: 'a4', orient: 'portrait', m: { t: 1.91, b: 1.91, l: 1.78, r: 1.78 }, fit: 'width', scale: 100, hc: false, vc: false, grid: true, title: false, num: true, auto: true, rb: [], cb: [], keep: [], tr: null, area: null };
export const MARGIN_PX = 96 / 2.54;
const NUM_H = 22, TITLE_H = 34;

export function setup(sh) { const p = sh.page || {}; return { ...DEF, ...p, m: { ...DEF.m, ...(p.m || {}) }, rb: p.rb || [], cb: p.cb || [], keep: p.keep || [] }; }
export const marginPreset = m => MARGINS.find(([, , v]) => v.every((x, i) => Math.abs(x - [m.t, m.b, m.l, m.r][i]) < .005))?.[0] || '';
const fmtCm = v => String(+v.toFixed(2)).replace('.', ',');

// Change the page settings of a sheet as one undoable step.
export function patch(app, sh, label, fn) {
  app.tx(label, () => { app.wb.hist.touchMeta(sh); const p = setup(sh); fn(p); sh.page = p; sh.layoutV++; });
  app.pageCache = null;
  app.grid.render();
  app.onSelChange();
}

export function paperPx(p) {
  let [w, h] = (PAPER[p.size] || PAPER.a4).mm;
  if (p.orient === 'landscape') [w, h] = [h, w];
  return { w, h, W: w / 25.4 * 96, H: h / 25.4 * 96 };
}

// Rows/cols that get printed: the print area, or the used range grown over merges, charts excluded.
export function printRange(sh, p = setup(sh)) {
  if (p.area) return { ...p.area };
  const u = sh.usedRange();
  if (!u) return null;
  for (const m of sh.merges) if (m.r1 <= u.r2 && m.c1 <= u.c2) { u.r2 = Math.max(u.r2, m.r2); u.c2 = Math.max(u.c2, m.c2); }
  return u;
}

// Tables = blocks of consecutive rows that have content, separated by empty rows.
function autoTables(sh, g) {
  const rows = new Set();
  for (const [k, c] of sh.cells) {
    if (c.empty || (c.v == null && c.f == null)) continue;
    const r = keyR(k), cc = keyC(k);
    if (r >= g.r1 && r <= g.r2 && cc >= g.c1 && cc <= g.c2) rows.add(r);
  }
  const out = [];
  let s = -1;
  for (let r = g.r1; r <= g.r2 + 1; r++) {
    const on = rows.has(r) || (s >= 0 && sh.isRowHidden(r) && r <= g.r2);
    if (on && s < 0) s = r;
    else if (!on && s >= 0) { if (r - 1 > s) out.push({ r1: s, r2: r - 1, auto: true }); s = -1; }
  }
  return out;
}

/** Page layout of a sheet: rows of each page, columns of each page, scale. Cached per app until the next change. */
export function layout(app, sh, over = null) {
  const key = sh.id + '|' + sh.layoutV + '|' + JSON.stringify(sh.page) + '|' + JSON.stringify(over);
  app.pageCache ||= new Map();
  if (app.pageCache.has(key)) return app.pageCache.get(key);
  const p = { ...setup(sh), ...(over || {}) }, g = printRange(sh, p);
  const paper = paperPx(p);
  const cw = paper.W - (p.m.l + p.m.r) * MARGIN_PX;
  const ch = paper.H - (p.m.t + p.m.b) * MARGIN_PX - (p.num ? NUM_H : 0);
  let res = { p, g, paper, cw, ch, scale: 1, rows: [], cols: [], keeps: [], rowBreaks: [], colBreaks: [] };
  if (!g) { app.pageCache.set(key, res); return res; }
  const visC = [], visR = [];
  for (let c = g.c1; c <= g.c2; c++) if (!sh.hiddenC.has(c)) visC.push(c);
  for (let r = g.r1; r <= g.r2; r++) if (!sh.isRowHidden(r)) visR.push(r);
  const totW = visC.reduce((a, c) => a + sh.colWidth(c), 0) || 1;
  const totH = visR.reduce((a, r) => a + sh.rowHeight(r), 0) + (p.title ? TITLE_H : 0) || 1;
  let scale = p.fit === 'width' ? Math.min(1, cw / totW) : p.fit === 'page' ? Math.min(1, cw / totW, ch / totH) : (p.scale || 100) / 100;
  scale = Math.max(0.1, Math.min(4, scale));
  const W = cw / scale, H = ch / scale;

  // columns
  const cb = new Set(p.cb), cols = [];
  let cs = g.c1, acc = 0;
  for (const c of visC) {
    const w = sh.colWidth(c);
    if (c > cs && (cb.has(c) || (p.fit !== 'width' && p.fit !== 'page' && acc + w > W))) { cols.push({ c1: cs, c2: c - 1 }); cs = c; acc = 0; }
    acc += w;
  }
  cols.push({ c1: cs, c2: g.c2 });

  // rows
  const tr = p.tr && p.tr.r2 >= p.tr.r1 ? p.tr : null;
  const trH = tr ? range(tr.r1, tr.r2).reduce((a, r) => a + sh.rowHeight(r), 0) : 0;
  const repeatH = start => (tr && start > tr.r2 ? trH : 0);
  const hOf = (a, b) => { let s = 0; for (let r = a; r <= b; r++) s += sh.rowHeight(r); return s; };
  const keeps = [
    ...p.keep.filter(k => k.r2 >= g.r1 && k.r1 <= g.r2),
    ...sh.merges.filter(m => m.r2 > m.r1 && m.r1 >= g.r1 && m.r2 <= g.r2).map(m => ({ r1: m.r1, r2: m.r2, merge: true })),
    ...(p.auto ? autoTables(sh, g) : []),
  ];
  const rb = new Set(p.rb), starts = [g.r1];
  let ps = g.r1;
  acc = p.title ? TITLE_H : 0;
  for (let r = g.r1; r <= g.r2; r++) {
    const h = sh.rowHeight(r);
    if (!h && !rb.has(r)) continue;
    const forced = rb.has(r) && r > ps;
    if (forced || (acc + h > H && r > ps)) {
      let b = r;
      if (!forced) {
        let moved = true;
        while (moved) {
          moved = false;
          for (const k of keeps) if (k.r1 < b && k.r2 >= b && k.r1 > ps && hOf(k.r1, k.r2) + repeatH(k.r1) <= H) { b = k.r1; moved = true; }
        }
      }
      starts.push(b); ps = b; acc = repeatH(b); r = b - 1;
      continue;
    }
    acc += h;
  }
  const rows = starts.map((s, i) => ({ r1: s, r2: i < starts.length - 1 ? starts[i + 1] - 1 : g.r2 }));
  res = { ...res, scale, W, H, tr, trH, rows, cols, keeps, rowBreaks: starts.slice(1), colBreaks: cols.slice(1).map(x => x.c1), manualR: rb, manualC: cb };
  app.pageCache.set(key, res);
  return res;
}
const range = (a, b) => Array.from({ length: Math.max(0, b - a + 1) }, (_, i) => a + i);

// The break (if any) that cuts the selection: returns {b, page, fits, rows}.
export function cutBy(app, sh, sel) {
  const L = layout(app, sh);
  if (!L.g || sel.r2 === sel.r1) return null;
  const b = L.rowBreaks.find(x => x > sel.r1 && x <= sel.r2);
  if (b == null) return null;
  let h = 0;
  for (let r = sel.r1; r <= sel.r2; r++) h += sh.rowHeight(r);
  const page = L.rowBreaks.indexOf(b) + 1;
  return { b, page, fits: h + (L.tr && sel.r1 > L.tr.r2 ? L.trH : 0) <= L.H, L };
}

/* ---------------- commands ---------------- */
export function keepTogether(app, g) {
  const sh = app.sheet, L = layout(app, sh);
  let h = 0;
  for (let r = g.r1; r <= g.r2; r++) h += sh.rowHeight(r);
  if (h > L.H) { toast('Pilihan lebih tinggi dari satu halaman — gunakan "Ulangi baris judul" agar judul tabel muncul di tiap halaman'); return; }
  patch(app, sh, 'Jaga tetap utuh', p => { p.keep = [...p.keep.filter(k => !(k.r1 === g.r1 && k.r2 === g.r2)), { r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c2 }]; });
  showBreaks(app);
  const nb = layout(app, sh).rowBreaks.find(x => x > g.r1 && x <= g.r2);
  toast(nb == null ? `Baris ${g.r1 + 1}–${g.r2 + 1} tidak akan terpotong — dipindahkan utuh ke halaman berikutnya bila perlu` : 'Rentang ini tidak dapat dipindahkan utuh');
}
export function unkeep(app, g) {
  patch(app, app.sheet, 'Izinkan terpotong', p => { p.keep = p.keep.filter(k => k.r2 < g.r1 || k.r1 > g.r2); });
}
export function breakAt(app, axis, i, on = true) {
  patch(app, app.sheet, on ? 'Sisipkan pemisah halaman' : 'Hapus pemisah halaman', p => {
    const k = axis === 'r' ? 'rb' : 'cb';
    p[k] = on ? [...new Set([...p[k], i])].sort((a, b) => a - b) : p[k].filter(x => x !== i);
  });
  showBreaks(app);
}
export function showBreaks(app, on = true) { if (!!app.settings.breaks === on) return; app.settings.breaks = on; app.saveSettings(); app.grid.render(); }

export function layoutMenu(app) {
  const sh = app.sheet, p = setup(sh), s = app.grid.range;
  const L = () => layout(app, sh);
  const kept = p.keep.some(k => k.r1 <= s.r2 && k.r2 >= s.r1);
  const rbHere = p.rb.includes(s.r1), cbHere = p.cb.includes(s.c1);
  return [
    { label: 'Margin', icon: 'margin', sub: () => [
      ...MARGINS.map(([k, n, v]) => ({ label: `${n}  ·  ${v.map(fmtCm).join(' / ')} cm`, checked: () => marginPreset(setup(sh).m) === k, run: () => patch(app, sh, 'Margin', q => { [q.m.t, q.m.b, q.m.l, q.m.r] = v; }) })),
      { sep: true },
      { label: 'Margin kustom…', icon: 'tune', run: () => setupDialog(app) },
    ] },
    { label: 'Orientasi', icon: 'screen_rotation', sub: [
      { label: 'Potret', icon: 'crop_portrait', checked: () => setup(sh).orient !== 'landscape', run: () => patch(app, sh, 'Orientasi', q => { q.orient = 'portrait'; }) },
      { label: 'Lanskap', icon: 'crop_landscape', checked: () => setup(sh).orient === 'landscape', run: () => patch(app, sh, 'Orientasi', q => { q.orient = 'landscape'; }) },
    ] },
    { label: 'Ukuran kertas', icon: 'description', sub: () => Object.entries(PAPER).map(([k, v]) => ({ label: `${v.n}  ·  ${fmtCm(v.mm[0] / 10)} × ${fmtCm(v.mm[1] / 10)} cm`, checked: () => setup(sh).size === k, run: () => patch(app, sh, 'Ukuran kertas', q => { q.size = k; }) })) },
    { label: 'Skala', icon: 'fit_screen', sub: [
      { label: 'Normal (100%)', checked: () => setup(sh).fit === 'normal' && setup(sh).scale === 100, run: () => patch(app, sh, 'Skala', q => { q.fit = 'normal'; q.scale = 100; }) },
      { label: 'Sesuaikan dengan lebar', checked: () => setup(sh).fit === 'width', run: () => patch(app, sh, 'Skala', q => { q.fit = 'width'; }) },
      { label: 'Muat dalam 1 halaman', checked: () => setup(sh).fit === 'page', run: () => patch(app, sh, 'Skala', q => { q.fit = 'page'; }) },
      { label: 'Skala kustom…', icon: 'tune', run: () => setupDialog(app) },
    ] },
    { sep: true },
    { label: 'Tampilkan pemisah halaman', icon: 'insert_page_break', checked: () => !!app.settings.breaks, run: () => showBreaks(app, !app.settings.breaks) },
    { label: `Jangan potong ${rangeAddr(s)} (pindah ke halaman berikutnya)`, icon: 'move_down', hidden: () => kept || s.r1 === s.r2, run: () => keepTogether(app, s) },
    { label: `Izinkan ${rangeAddr(s)} terpotong`, icon: 'content_cut', hidden: () => !kept, run: () => unkeep(app, s) },
    { label: 'Hindari tabel terpotong (otomatis)', icon: 'auto_awesome', checked: () => setup(sh).auto, run: () => patch(app, sh, 'Hindari tabel terpotong', q => { q.auto = !q.auto; }) },
    { label: 'Pemisah halaman manual', icon: 'horizontal_rule', sub: () => [
      { label: `Sisipkan di atas baris ${s.r1 + 1}`, icon: 'vertical_align_top', hidden: () => rbHere || s.r1 === 0, run: () => breakAt(app, 'r', s.r1) },
      { label: `Hapus pemisah di atas baris ${s.r1 + 1}`, icon: 'close', hidden: () => !rbHere, run: () => breakAt(app, 'r', s.r1, false) },
      { label: `Sisipkan di kiri kolom ${colName(s.c1)}`, icon: 'format_indent_increase', hidden: () => cbHere || s.c1 === 0, run: () => breakAt(app, 'c', s.c1) },
      { label: `Hapus pemisah di kiri kolom ${colName(s.c1)}`, icon: 'close', hidden: () => !cbHere, run: () => breakAt(app, 'c', s.c1, false) },
      { sep: true },
      { label: 'Hapus semua pemisah & penjaga manual', icon: 'delete_sweep', disabled: () => !p.rb.length && !p.cb.length && !p.keep.length, run: () => patch(app, sh, 'Setel ulang pemisah', q => { q.rb = []; q.cb = []; q.keep = []; }) },
    ] },
    { label: 'Ulangi baris judul di tiap halaman', icon: 'title', sub: () => [
      { label: 'Tidak ada', checked: () => !setup(sh).tr, run: () => patch(app, sh, 'Baris judul', q => { q.tr = null; }) },
      { label: 'Baris 1', checked: () => setup(sh).tr?.r1 === 0 && setup(sh).tr?.r2 === 0, run: () => patch(app, sh, 'Baris judul', q => { q.tr = { r1: 0, r2: 0 }; }) },
      { label: 'Baris 1–2', checked: () => setup(sh).tr?.r1 === 0 && setup(sh).tr?.r2 === 1, run: () => patch(app, sh, 'Baris judul', q => { q.tr = { r1: 0, r2: 1 }; }) },
      { label: `Baris pilihan (${s.r1 + 1}${s.r2 > s.r1 ? '–' + (s.r2 + 1) : ''})`, run: () => patch(app, sh, 'Baris judul', q => { q.tr = { r1: s.r1, r2: Math.min(s.r2, s.r1 + 9) }; }) },
    ] },
    { label: 'Area cetak', icon: 'crop', sub: () => [
      { label: `Tetapkan dari pilihan (${rangeAddr(s)})`, run: () => patch(app, sh, 'Area cetak', q => { q.area = { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 }; }) },
      { label: 'Hapus area cetak', disabled: () => !setup(sh).area, run: () => patch(app, sh, 'Area cetak', q => { q.area = null; }) },
    ] },
    { sep: true },
    { label: 'Penyiapan halaman…', icon: 'tune', run: () => setupDialog(app) },
    { label: `Pratinjau cetak (${L().g ? L().rows.length * L().cols.length : 0} halaman)`, icon: 'preview', run: () => printPages(app, { sheets: [sh] }) },
  ];
}

/* ---------------- page setup dialog ---------------- */
export function setupDialog(app) {
  const sh = app.sheet, p = setup(sh);
  const num = v => el('input.inp', { type: 'number', min: 0, max: 10, step: 0.1, value: +(+v).toFixed(2) });
  const m = { t: num(p.m.t), b: num(p.m.b), l: num(p.m.l), r: num(p.m.r) };
  const prev = el('div.pg-prev', el('i'));
  const sizeSel = select(Object.entries(PAPER).map(([k, v]) => ({ value: k, label: `${v.n} (${fmtCm(v.mm[0] / 10)} × ${fmtCm(v.mm[1] / 10)} cm)` })), p.size, () => draw());
  let orient = p.orient;
  const orientSeg = segmented([{ value: 'portrait', label: 'Potret', icon: 'crop_portrait' }, { value: 'landscape', label: 'Lanskap', icon: 'crop_landscape' }], orient, v => { orient = v; draw(); });
  const preset = select([{ value: '', label: 'Kustom' }, ...MARGINS.map(([k, n, v]) => ({ value: k, label: `${n} (${v.map(fmtCm).join(' / ')} cm)` }))], marginPreset(p.m), v => {
    const P = MARGINS.find(x => x[0] === v)?.[2];
    if (P) { [m.t.value, m.b.value, m.l.value, m.r.value] = P; draw(); }
  });
  Object.values(m).forEach(i => i.addEventListener('input', () => { preset.value = ''; draw(); }));
  let fit = p.fit;
  const scaleIn = el('input.inp', { type: 'number', min: 10, max: 400, step: 5, value: p.scale || 100, disabled: fit !== 'normal' });
  const fitSel = select([{ value: 'normal', label: 'Skala kustom (%)' }, { value: 'width', label: 'Sesuaikan dengan lebar halaman' }, { value: 'page', label: 'Muat seluruhnya dalam 1 halaman' }], fit, v => { fit = v; scaleIn.disabled = v !== 'normal'; });
  const o = { hc: p.hc, vc: p.vc, grid: p.grid, title: p.title, num: p.num, auto: p.auto };
  const ck = (k, label) => checkbox(label, o[k], v => (o[k] = v));
  function draw() {
    let [w, h] = PAPER[sizeSel.value].mm;
    if (orient === 'landscape') [w, h] = [h, w];
    const k = 150 / Math.max(w, h);
    prev.style.width = w * k + 'px'; prev.style.height = h * k + 'px';
    const mm = x => (parseFloat(x.value) || 0) * 10 * k + 'px';
    Object.assign(prev.firstChild.style, { top: mm(m.t), bottom: mm(m.b), left: mm(m.l), right: mm(m.r) });
  }
  draw();
  const body = el('div.pgset',
    el('div.stack',
      field('Orientasi', orientSeg),
      field('Ukuran kertas', sizeSel),
      field('Preset margin', preset),
      el('div.grid4', field('Atas (cm)', m.t), field('Bawah (cm)', m.b), field('Kiri (cm)', m.l), field('Kanan (cm)', m.r)),
      el('div.grid2', field('Skala', fitSel), field('Persen', scaleIn)),
      el('div.checks',
        ck('auto', 'Hindari tabel terpotong — pindahkan tabel utuh ke halaman berikutnya'),
        ck('hc', 'Tengahkan horizontal'), ck('vc', 'Tengahkan vertikal'),
        ck('grid', 'Cetak garis kisi'), ck('num', 'Nomor halaman di bawah'), ck('title', 'Judul spreadsheet di halaman pertama'))),
    el('div.pg-prev-wrap', prev, el('small.muted', 'Pratinjau margin')));
  dialog({
    title: `Penyiapan halaman — ${sh.name}`, width: 620, body,
    buttons: [{ label: 'Pratinjau', run: () => { apply(); printPages(app, { sheets: [sh] }); } }, { spacer: true }, { label: 'Batal' }, { label: 'Terapkan', primary: true, run: () => { apply(); showBreaks(app); toast('Penyiapan halaman disimpan — garis putus-putus menandai batas halaman'); } }],
  });
  function apply() {
    const cm = x => Math.max(0, Math.min(10, parseFloat(x.value) || 0));
    patch(app, sh, 'Penyiapan halaman', q => {
      Object.assign(q, o, { size: sizeSel.value, orient, fit, scale: Math.max(10, Math.min(400, parseFloat(scaleIn.value) || 100)), m: { t: cm(m.t), b: cm(m.b), l: cm(m.l), r: cm(m.r) } });
    });
  }
}

/* ---------------- printing ---------------- */
const PRINT_CSS = `
*{box-sizing:border-box}html,body{margin:0}body{font-family:Arial,Helvetica,sans-serif;font-size:10pt;color:#000;background:#e8eaed}
table.sheet{border-collapse:collapse;table-layout:fixed}table.sheet td{padding:1px 4px;overflow:hidden;white-space:nowrap;line-height:1.25;font-size:10pt}
table.sheet.grid td{outline:1px solid #d9d9d9;outline-offset:-0.5px}table.sheet tr{break-inside:avoid}
.pg{position:relative;background:#fff;margin:18px auto;box-shadow:0 1px 4px rgba(60,64,67,.3);overflow:hidden}
.pg-in{position:relative;overflow:hidden;display:flex;flex-direction:column}
.pg-in.vc{justify-content:center}.pg-z{transform-origin:0 0}.pg-in.hc .pg-z{margin:0 auto}
.pg-num{text-align:center;font:9pt Arial,sans-serif;color:#555}
h1{font:600 13pt Arial,sans-serif;margin:0 0 4px}.meta{color:#666;font-size:8pt;margin-bottom:8px}
.bar{position:sticky;top:0;z-index:2;display:flex;gap:10px;align-items:center;justify-content:center;padding:10px;background:rgba(255,255,255,.92);backdrop-filter:blur(8px);font:500 13px Inter,Arial,sans-serif;color:#3c4043;box-shadow:0 1px 3px rgba(0,0,0,.12)}
.bar button{font:500 13px Inter,Arial;background:#1a73e8;color:#fff;border:0;border-radius:18px;padding:8px 18px;cursor:pointer}
@media print{body{background:#fff}.bar{display:none}.pg{margin:0!important;box-shadow:none;break-after:page;padding:0!important;width:auto!important;height:auto!important}.pg:last-child{break-after:auto}}
`;

function pageHTML(app, sh, L, rows, cols, idx, total, first) {
  const p = L.p;
  const tables = [];
  if (first && p.title) tables.push(`<h1>${esc(app.wb.name)}</h1><div class="meta">${esc(sh.name)} · dicetak ${new Date().toLocaleString('id-ID')}</div>`);
  if (L.tr && rows.r1 > L.tr.r2) tables.push(sheetHTML(app, sh, { r1: L.tr.r1, r2: L.tr.r2, c1: cols.c1, c2: cols.c2 }, { grid: p.grid, clip: true }));
  tables.push(sheetHTML(app, sh, { ...rows, c1: cols.c1, c2: cols.c2 }, { grid: p.grid, clip: true }));
  // zoom keeps layout size in step with the scale, so centering and the page height stay exact
  return `<div class="pg-in${p.hc ? ' hc' : ''}${p.vc ? ' vc' : ''}" style="width:${L.cw.toFixed(2)}px;height:${(L.ch - 0.5).toFixed(2)}px"><div class="pg-z" style="zoom:${L.scale.toFixed(4)}">${tables.join('')}</div></div>${p.num ? `<div class="pg-num" style="height:${NUM_H}px;line-height:${NUM_H}px">${idx} / ${total}</div>` : ''}`;
}

export function printPages(app, { sheets, area = null, auto = false }) {
  const parts = [], pageCss = new Set();
  const all = sheets.map(sh => app.withSheet(sh, () => ({ sh, L: layout(app, sh, area ? { area } : null) })));
  const total = all.reduce((a, x) => a + (x.L.g ? x.L.rows.length * x.L.cols.length : 0), 0);
  if (!total) { toast('Tidak ada isi untuk dicetak'); return; }
  let n = 0;
  for (const { sh, L } of all) {
    if (!L.g) continue;
    const p = L.p, name = 'p' + sh.id.replace(/\W/g, '');
    pageCss.add(`@page ${name}{size:${L.paper.w}mm ${L.paper.h}mm;margin:${p.m.t * 10}mm ${p.m.r * 10}mm ${p.m.b * 10}mm ${p.m.l * 10}mm}.${name}{page:${name};width:${L.paper.W.toFixed(2)}px;height:${L.paper.H.toFixed(2)}px;padding:${(p.m.t * MARGIN_PX).toFixed(2)}px ${(p.m.r * MARGIN_PX).toFixed(2)}px ${(p.m.b * MARGIN_PX).toFixed(2)}px ${(p.m.l * MARGIN_PX).toFixed(2)}px}`);
    let first = true;
    for (const cols of L.cols) for (const rows of L.rows) {   // down, then over (as Excel)
      n++;
      parts.push(`<section class="pg ${name}">${app.withSheet(sh, () => pageHTML(app, sh, L, rows, cols, n, total, first))}</section>`);
      first = false;
    }
  }
  const w = window.open('', '_blank');
  if (!w) { toast('Izinkan pop-up untuk mencetak'); return; }
  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(app.wb.name)}</title><style>${PRINT_CSS}${[...pageCss].join('')}</style></head><body><div class="bar"><span>${total} halaman</span><button onclick="print()">Cetak / Simpan PDF</button></div>${parts.join('')}</body></html>`);
  w.document.close();
  if (auto) setTimeout(() => { try { w.focus(); w.print(); } catch { /* the button is there */ } }, 500);
}

// Floating suggestion under a selection that a page break cuts in two.
export function breakChip(app) {
  const host = document.getElementById('pagechip');
  if (!host) return;
  host.classList.remove('show');
  host.replaceChildren();
  const g = app.grid, sh = app.sheet;
  if (!sh || g.edit || app.readOnly || app.settings.breakTips === false || g.sel.ranges.length !== 1) return;
  const s = g.range;
  if (s.r2 - s.r1 > 2000 || s.r1 === s.r2) return;
  const cut = cutBy(app, sh, s);
  if (!cut || app._chipSkip === rangeAddr(s)) return;
  const p = setup(sh);
  if (p.keep.some(k => k.r1 === s.r1 && k.r2 === s.r2)) return;
  host.append(
    icon('content_cut'),
    el('span', `${rangeAddr(s)} terpotong di halaman ${cut.page}–${cut.page + 1}`),
    cut.fits
      ? el('button.btn.primary.sm', { type: 'button', onclick: () => keepTogether(app, s) }, icon('move_down'), 'Pindahkan ke halaman berikutnya')
      : el('button.btn.primary.sm', { type: 'button', onclick: () => { patch(app, sh, 'Baris judul', q => { q.tr = { r1: s.r1, r2: s.r1 }; }); toast(`Baris ${s.r1 + 1} diulang di setiap halaman`); } }, icon('title'), `Ulangi baris ${s.r1 + 1} di tiap halaman`),
    el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Penyiapan halaman', onclick: () => setupDialog(app) }, icon('tune')),
    el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Biarkan terpotong', onclick: () => { app._chipSkip = rangeAddr(s); host.classList.remove('show'); } }, icon('close')),
  );
  // beside the selection at the break line when there is room, otherwise just under the line
  host.classList.add('show');
  const yb = g.Y(cut.b), hr = g.host.getBoundingClientRect(), cw = host.offsetWidth, chh = host.offsetHeight;
  const right = hr.left + g.X(s.c2 + 1) * g.z + 10;
  const beside = right + cw < hr.right - 8;
  host.style.left = (beside ? right : Math.max(hr.left + 4, Math.min(hr.right - cw - 8, hr.left + g.X(s.c1) * g.z))) + 'px';
  host.style.top = Math.max(hr.top + 30, Math.min(hr.bottom - chh - 8, hr.top + yb * g.z + (beside ? -chh / 2 : 6))) + 'px';
}
