// Cetak / PDF. Ctrl/⌘+P, tombol cetak, menu Cetak dan Download › PDF membuka dialog ini:
// - area yang dikerjakan dipilih sendiri (dan ikut terseleksi di sheet), bukan dari A1;
// - skala Otomatis mengisi lebar kertas, orientasi Otomatis memilih yang tulisannya paling terbaca;
// - tabel tidak terpotong (pindah utuh ke halaman berikutnya), tabel yang lebih panjang dari satu halaman
//   berlanjut dengan judul + header-nya diulang; − / + mengubah skala dan halaman dirapikan ulang dengan aturan yang sama.
// Pratinjau = iframe di dialog. Yang dicetak = #print-root di halaman aplikasi sendiri (CSS @media print di app.css
// menyembunyikan selebihnya dan menolkan zoom amatu.js) — juga untuk menu Print bawaan browser (beforeprint).
import { el, icon, esc, rangeAddr, colName } from './util.js';
import { dialog, checkbox, select, toast } from './ui/ui.js';
import * as PG from './pages.js';

const same = (a, b) => !!a && !!b && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
const pick = o => Object.fromEntries(PG.GLOBAL_KEYS.map(k => [k, k === 'm' ? { ...o.m } : o[k]]));
const cm = v => String(+(+v).toFixed(2)).replace('.', ',');
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const ORIENT = { portrait: 'potret', landscape: 'lanskap' };
const PREVIEW_MAX = 40;   // pages drawn in the preview (all of them are printed)
let cur = null;      // the print dialog on screen
let busy = false;    // a print of ours is running: the page's own beforeprint must stay out of it

export const isOpen = () => !!cur;
export function printNow(pdf = false) { cur?.go(pdf); }

// small segmented control whose value can also be set from outside (− / + clear the "fit" choice)
function seg(items, value, onPick) {
  const w = el('div.seg.full.prv-seg', { role: 'radiogroup' });
  const btns = items.map(([v, label, ic]) => {
    const b = el('button', { type: 'button', role: 'radio', onclick: () => { set(v); onPick(v); } }, ic ? icon(ic) : null, label);
    b._v = v;
    w.append(b);
    return b;
  });
  function set(v) { for (const b of btns) { b.classList.toggle('on', b._v === v); b.setAttribute('aria-checked', String(b._v === v)); } }
  set(value);
  w.set = set;
  return w;
}
const sec = (label, ...kids) => el('div.field.prv-sec', el('span', label), ...kids);

export function printDialog(app, { pdf = false } = {}) {
  if (cur || !app.wb) return;
  if (document.body.classList.contains('file-locked')) return;
  const sh = app.sheet;
  if (!sh) return;
  const saved = PG.setup(sh);
  const q = pick(saved);
  const auto = PG.autoArea(app, sh);
  const book = app.wb.sheets.filter(s => !s.hidden);
  // what the person had selected (whole rows / columns are cut down to the worked area)
  let sel0 = app.grid.sel.ranges.length === 1 ? { ...app.grid.range } : null;
  if (sel0 && auto) {
    if (sel0.r1 === 0 && sel0.r2 >= sh.rows - 1) { sel0.r1 = auto.r1; sel0.r2 = auto.r2; }
    if (sel0.c1 === 0 && sel0.c2 >= sh.cols - 1) { sel0.c1 = auto.c1; sel0.c2 = auto.c2; }
  }
  const multi = sel0 && (sel0.r2 > sel0.r1 || sel0.c2 > sel0.c1) && !same(sel0, auto) && !same(sel0, saved.area);
  const scopes = [];
  if (saved.area) scopes.push({ value: 'area', label: `Area cetak tersimpan · ${rangeAddr(saved.area)}` });
  if (auto) scopes.push({ value: 'auto', label: `Area kerja otomatis · ${rangeAddr(auto)}` });
  if (multi) scopes.push({ value: 'sel', label: `Sel yang dipilih · ${rangeAddr(sel0)}` });
  if (book.length > 1) scopes.push({ value: 'book', label: `Seluruh workbook · ${book.length} sheet` });
  if (!scopes.length) { toast('Tidak ada isi untuk dicetak'); return; }
  let scope = scopes[0].value;

  // settings over each sheet's own: other sheets keep their print area, breaks and keeps
  const overFor = s => {
    const o = pick(q);
    if (s !== sh) return o;
    if (scope === 'auto') o.area = null;
    else if (scope === 'sel') o.area = { r1: sel0.r1, c1: sel0.c1, r2: sel0.r2, c2: sel0.c2 };
    return o;
  };
  const sheets = () => (scope === 'book' ? book : [sh]);
  let L = null, D = null, lastG = null, big = false;

  /* ---------- controls ---------- */
  const frame = el('iframe.prv-frame', { title: 'Pratinjau cetak', tabIndex: -1 });
  const badge = el('div.prv-badge', { 'aria-live': 'polite' });
  const zoomBtn = el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Perbesar pratinjau', 'aria-label': 'Perbesar pratinjau', onclick: () => { big = !big; zoomBtn.firstChild.textContent = big ? 'zoom_out' : 'zoom_in'; zoomBtn.dataset.tip = big ? 'Lihat seluruh halaman' : 'Perbesar pratinjau'; fitPreview(); } }, icon('zoom_in'));
  const stage = el('div.prv-stage', frame, badge, el('div.prv-tools', zoomBtn));

  const scopeSel = select(scopes, scope, v => { scope = v; update(); });
  const rangeInfo = el('div.prv-hint');
  const paperSel = select(Object.entries(PG.PAPER).map(([k, v]) => ({ value: k, label: `${v.n} · ${cm(v.mm[0] / 10)} × ${cm(v.mm[1] / 10)} cm` })), q.size, v => { q.size = v; update(); });
  const orientSeg = seg([['auto', 'Otomatis', 'auto_awesome'], ['portrait', 'Potret', 'crop_portrait'], ['landscape', 'Lanskap', 'crop_landscape']], q.orient, v => { q.orient = v; update(); });
  const orientHint = el('div.prv-hint');
  const fitSeg = seg([['auto', 'Otomatis', 'auto_awesome'], ['width', 'Lebar', 'width'], ['page', '1 halaman', 'fit_page']], q.fit, v => { q.fit = v; update(); });
  const pct = el('input.inp.prv-pct', { type: 'number', min: 10, max: 400, step: 5, 'aria-label': 'Skala (%)', 'data-enter': '1' });
  const setPct = () => { const v = parseFloat(pct.value); if (v > 0) { q.fit = 'normal'; q.scale = clamp(Math.round(v), 10, 400); update(); } };
  pct.addEventListener('change', setPct);
  pct.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); setPct(); } });
  const M = app.MOD || 'Ctrl';
  const minus = el('button.icon-btn.prv-step', { type: 'button', 'data-tip': `Perkecil|${M}+−`, 'aria-label': 'Perkecil skala', onclick: () => step(-1) }, icon('remove'));
  const plus = el('button.icon-btn.prv-step', { type: 'button', 'data-tip': `Perbesar|${M}++`, 'aria-label': 'Perbesar skala', onclick: () => step(1) }, icon('add'));
  const scaleHint = el('div.prv-hint');

  const mOpts = [...PG.MARGINS.map(([k, n, v]) => ({ value: k, label: `${n} · ${v.map(cm).join(' / ')} cm` })), { value: 'custom', label: 'Kustom…' }];
  const mSel = select(mOpts, PG.marginPreset(q.m) || 'custom', v => {
    if (v === 'custom') { mBox.hidden = false; return; }
    [q.m.t, q.m.b, q.m.l, q.m.r] = PG.MARGINS.find(x => x[0] === v)[2];
    syncMargins(); update();
  });
  const mIn = k => {
    const i = el('input.inp.sm', { type: 'number', min: 0, max: 10, step: 0.1, 'data-enter': '1', 'aria-label': 'Margin (cm)' });
    const set = () => { q.m[k] = clamp(parseFloat(i.value) || 0, 0, 10); mSel.value = PG.marginPreset(q.m) || 'custom'; update(); };
    i.addEventListener('change', set);
    i.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); set(); } });
    return i;
  };
  const mInputs = { t: mIn('t'), b: mIn('b'), l: mIn('l'), r: mIn('r') };
  const syncMargins = () => { for (const k of ['t', 'b', 'l', 'r']) mInputs[k].value = +(+q.m[k]).toFixed(2); };
  syncMargins();
  const mBox = el('div.grid4.prv-m', ...[['t', 'Atas'], ['b', 'Bawah'], ['l', 'Kiri'], ['r', 'Kanan']].map(([k, n]) => el('label.field', el('span', n), mInputs[k])));
  mBox.hidden = mSel.value !== 'custom';

  const ck = (k, label) => checkbox(label, !!q[k], v => { q[k] = v; update(); });
  const gridCk = checkbox('Garis kisi', true, v => { q.grid = v; update(); });
  const notes = el('div.prv-notes');
  const side = el('div.prv-side',
    sec('Yang dicetak', scopeSel, rangeInfo),
    sec('Ukuran kertas', paperSel),
    sec('Orientasi', orientSeg, orientHint),
    sec('Skala', fitSeg, el('div.prv-scale', minus, pct, el('span.prv-unit', '%'), plus), scaleHint),
    sec('Margin', mSel, mBox),
    el('div.checks.prv-checks',
      ck('auto', 'Jangan potong tabel — pindahkan utuh ke halaman berikutnya'),
      ck('hdr', 'Ulangi judul & header tabel di halaman berikutnya'),
      ck('hc', 'Tengahkan horizontal'),
      ck('vc', 'Tengahkan vertikal'),
      gridCk,
      ck('num', 'Nomor halaman'),
      ck('title', 'Judul spreadsheet di halaman pertama')),
    notes);
  const info = el('span.prv-info');

  const dlg = dialog({
    title: pdf ? 'Download PDF' : 'Cetak', width: 'min(1180px, 100%)', cls: 'print-dlg', noBackdropClose: true,
    body: el('div.prv', stage, side),
    buttons: [
      { label: 'Batal' },
      { label: 'Simpan PDF', primary: pdf, run: () => { go(true); return false; } },
      { label: 'Cetak', primary: !pdf, run: () => { go(false); return false; } },
    ],
    onClose: () => {
      if (cur === me) cur = null;
      window.removeEventListener('keydown', keys, true);
      ro.disconnect();
      app.grid.focus();
    },
  });
  dlg.el.querySelector('.dlg-foot')?.prepend(info);
  // the first fields would take focus (and the Enter key): keep it on the dialog itself
  setTimeout(() => { if (dlg.el.contains(document.activeElement)) document.activeElement.blur(); }, 40);

  /* ---------- preview ---------- */
  const doc = frame.contentDocument;
  const fonts = [...document.querySelectorAll('link[rel=stylesheet][href*="fonts.googleapis.com"]')].filter(l => !/Material\+Symbols/.test(l.href)).map(l => `<link rel="stylesheet" href="${esc(l.href)}">`).join('');
  doc.open();
  doc.write(`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>${esc(app.wb.name)}</title>${fonts}<style id="pc"></style><style id="pv"></style></head><body></body></html>`);
  doc.close();
  const fwin = frame.contentWindow;
  fwin.addEventListener('scroll', () => badgeNow(), { passive: true });
  fwin.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); dlg.close(); } else keys(e); });
  const ro = new ResizeObserver(() => fitPreview());
  ro.observe(stage);

  function paint() {
    doc.getElementById('pc').textContent = D.css + '\nhtml,body{margin:0}@media print{html,body{background:#fff!important;zoom:1!important}}';
    doc.body.innerHTML = D.total ? D.body : '<p style="font:14px Inter,Arial,sans-serif;color:#5f6368;text-align:center;margin-top:40px">Tidak ada isi untuk dicetak</p>';
    fitPreview();
  }
  function fitPreview() {
    if (!D) return;
    let mw = 0, mh = 0;
    for (const { L: x } of D.all) if (x.g) { mw = Math.max(mw, x.paper.W); mh = Math.max(mh, x.paper.H); }
    if (!mw) { mw = 794; mh = 1123; }
    const sw = stage.clientWidth || 600, sh2 = stage.clientHeight || 500;
    const z = big ? clamp((sw - 40) / mw, 0.2, 1.6) : clamp(Math.min((sw - 40) / mw, (sh2 - 64) / mh), 0.1, 1.2);
    const bg = getComputedStyle(stage).backgroundColor || '#dfe3e8';
    // pages shorter than the stage sit in its middle
    let tall = 0, count = 0;
    for (const { L: x } of D.all) if (x.g) { tall += x.pages * x.paper.H; count += x.pages; }
    const top = Math.max(18, (sh2 - (tall * z + Math.max(0, count - 1) * 18)) / 2);
    doc.getElementById('pv').textContent = `@media screen{html{zoom:${z.toFixed(4)}}body{background:${bg};padding:${(top / z).toFixed(1)}px 0 ${(48 / z).toFixed(1)}px}.pg{margin:0 auto ${(18 / z).toFixed(1)}px!important;box-shadow:0 ${(2 / z).toFixed(1)}px ${(10 / z).toFixed(1)}px rgba(16,24,40,.22)!important}}`;
    badgeNow();
  }
  function badgeNow() {
    const pages = doc.querySelectorAll('.pg');
    badge.hidden = !pages.length;
    if (!pages.length) return;
    let i = 0;
    const mid = (fwin.innerHeight || 400) * 0.4;
    pages.forEach((p, k) => { if (p.getBoundingClientRect().top <= mid) i = k; });
    badge.textContent = `Halaman ${i + 1} dari ${D?.total || pages.length}`;
  }

  /* ---------- state → screen ---------- */
  function update() {
    L = app.withSheet(sh, () => PG.layout(app, sh, overFor(sh)));
    D = PG.printDoc(app, { sheets: sheets(), over: overFor, limit: PREVIEW_MAX });
    // what gets printed is selected on the sheet
    if (L.g && !same(L.g, lastG)) {
      lastG = { ...L.g };
      app.grid.selectRange(L.g, L.g.r1, L.g.c1);
      app.grid.scrollTo(L.g.r1, L.g.c1);
    }
    paint();
    const sc = Math.round(L.scale * 100);
    pct.value = sc;
    fitSeg.set(q.fit);
    orientSeg.set(q.orient);
    minus.disabled = sc <= 10; plus.disabled = sc >= 400;
    gridCk.querySelector('input').checked = L.grid;
    // texts
    if (scope === 'book') rangeInfo.textContent = `${D.all.filter(x => x.L.g).length} sheet berisi · area kerja tiap sheet`;
    else if (L.g) rangeInfo.textContent = `${rangeAddr(L.g)} · ${L.g.c2 - L.g.c1 + 1} kolom × ${L.g.r2 - L.g.r1 + 1} baris — terpilih di sheet`;
    else rangeInfo.textContent = 'Tidak ada isi';
    const o2 = L.other;
    orientHint.textContent = q.orient !== 'auto' ? '' : !L.g ? '' : o2
      ? `Dipilih ${ORIENT[L.orient]}: tulisan ${sc}%, ${L.pages} halaman (${ORIENT[o2.orient]}: ${Math.round(o2.scale * 100)}%, ${o2.pages} halaman)`
      : `Dipilih ${ORIENT[L.orient]}`;
    const fitTxt = {
      auto: L.scale < L.fitW - 1e-3 ? `Otomatis ${sc}% — diperkecil sedikit agar muat di ${L.pages} halaman` : `Otomatis ${sc}% — tabel memenuhi lebar kertas`,
      width: `${sc}% — selebar kertas`,
      page: `${sc}% — semua isi dalam 1 halaman`,
      normal: L.cols.length > 1 ? `${sc}% — skala manual, tabel dibagi rata ke samping` : `${sc}% — skala manual, halaman tetap dirapikan`,
    };
    scaleHint.textContent = L.g ? fitTxt[q.fit] || '' : '';
    const N = [];
    if (L.g) {
      if (!L.split) N.push(['check_circle', 'Tidak ada tabel yang terpotong']);
      else if (q.hdr && L.repeated) N.push(['view_headline', `${L.split} tabel lebih panjang dari 1 halaman — berlanjut dengan judul & header diulang`]);
      else N.push(['content_cut', `${L.split} tabel lebih panjang dari 1 halaman dan berlanjut ke halaman berikutnya`]);
      if (D.shown < D.total) N.push(['info', `Pratinjau menampilkan ${D.shown} halaman pertama dari ${D.total}; semuanya ikut dicetak.`]);
      if (L.moved) N.push(['move_down', `Pemisah halaman digeser ${L.moved}× agar tabel / baris tidak terpotong`]);
      if (L.scale < 0.5) N.push(['warning', `Tulisan kecil (${sc}%). Coba kertas lebih besar (F4/A3), margin Sempit, atau sembunyikan kolom yang tidak perlu.`, 'warn']);
      if (q.fit === 'normal' && L.cols.length > 1) N.push(['warning', `Tabel lebih lebar dari kertas, dibagi ${L.cols.length} halaman ke samping${L.lab ? ` — kolom ${colName(L.lab.c1)}${L.lab.c2 > L.lab.c1 ? '–' + colName(L.lab.c2) : ''} diulang di kiri tiap halaman` : ''}. Klik Otomatis agar tabel utuh selebar kertas.`, 'warn']);
    }
    notes.replaceChildren(...N.map(([ic, t, k]) => el('div.prv-note' + (k ? '.' + k : ''), icon(ic), el('span', t))));
    const paper = PG.PAPER[q.size]?.n || 'A4';
    info.textContent = D.total ? `${D.total} halaman · ${paper} ${scope === 'book' ? '' : ORIENT[L.orient] + ' · ' + sc + '%'}`.trim() : 'Tidak ada isi';
  }
  function step(d) {
    if (!L?.g) return;
    const c = Math.round(L.scale * 100);
    q.fit = 'normal';
    q.scale = clamp(d > 0 ? Math.floor(c / 5) * 5 + 5 : Math.ceil(c / 5) * 5 - 5, 10, 400);
    update();
  }
  function keys(e) {
    const mod = e.metaKey || e.ctrlKey;
    if (!mod || e.altKey) return;
    const k = e.key;
    if (k === '=' || k === '+') { e.preventDefault(); e.stopPropagation(); step(1); }
    else if (k === '-' || k === '_') { e.preventDefault(); e.stopPropagation(); step(-1); }
    else if (k === '0') { e.preventDefault(); e.stopPropagation(); q.fit = 'auto'; update(); }
    else if ((k === 'p' || k === 'P') && !e.shiftKey && e.currentTarget === fwin) { e.preventDefault(); go(false); }
  }
  window.addEventListener('keydown', keys, true);

  // the dialog's settings become the sheet's page setup (one undo step) once something is printed
  function save() {
    if (app.readOnly || sh.protect) return;
    if (JSON.stringify(pick(PG.setup(sh))) === JSON.stringify(pick(q))) return;
    PG.patch(app, sh, 'Penyiapan halaman', p => Object.assign(p, pick(q)));
  }
  async function go(asPdf) {
    if (!D?.total) { toast('Tidak ada isi untuk dicetak'); return; }
    save();
    if (asPdf) toast('Di jendela cetak, pilih Tujuan: "Simpan sebagai PDF"', { ms: 6000 });
    // the app page itself prints the pages (its own zoom is reset for print): same result in Chrome, Safari and iOS
    mainPrint(app, { sheets: sheets(), over: overFor }, app.wb.name || 'Spreadsheet', () => { if (cur === me) dlg.close(); });
  }

  const me = { go, close: () => dlg.close() };
  cur = me;
  update();
}

// images decoded and web fonts in, at most 2.5 s
function ready(doc, root = doc) {
  const imgs = [...root.querySelectorAll('img')].filter(i => !i.complete).map(i => new Promise(r => { i.onload = i.onerror = r; }));
  return Promise.race([Promise.all(imgs).then(() => doc.fonts?.ready), new Promise(r => setTimeout(r, 2500))]);
}

/* ---------- printing inside the app page (iPhone / iPad, and the browser's own Print menu) ---------- */
function mount(app, opts) {
  unmount();
  const d = PG.printDoc(app, { ...opts, scope: '#print-root ' });
  if (!d.total) return false;
  const style = el('style', { id: 'print-root-css' });
  style.textContent = d.css;
  const root = el('div', { id: 'print-root' });
  root.innerHTML = d.body;
  document.head.append(style);
  document.body.append(root);
  return root;
}
function unmount() { document.getElementById('print-root')?.remove(); document.getElementById('print-root-css')?.remove(); }

let busyAt = 0;
function mainPrint(app, opts, name, done) {
  if (busy && Date.now() - busyAt < 60000) return;     // a double click prints once
  const root = mount(app, opts);
  if (!root) { toast('Tidak ada isi untuk dicetak'); return; }
  const t0 = document.title;
  document.title = name;          // the name Chrome proposes for the PDF
  busy = true; busyAt = Date.now();
  let fin = false;
  const end = () => { if (fin) return; fin = true; busy = false; window.removeEventListener('afterprint', end); unmount(); document.title = t0; done?.(); };
  window.addEventListener('afterprint', end);
  ready(document, root).then(() => {
    const t = performance.now();
    window.print();
    // Chrome waits here until its print window is closed (printed or cancelled); others finish on afterprint
    if (performance.now() - t > 300) setTimeout(end, 60);
  });
}

// Ctrl/⌘+P anywhere (also while typing in a cell, a box or a dialog) opens the print dialog instead of printing
// the app screen; a print started from the browser menu still gets the worked area, fitted (beforeprint).
export function installPrint(app) {
  window.addEventListener('keydown', e => {
    if (!(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey || (e.key !== 'p' && e.key !== 'P' && e.code !== 'KeyP')) return;
    e.preventDefault(); e.stopPropagation();
    if (document.body.classList.contains('file-locked')) return;
    if (cur) { cur.go(false); return; }
    app.cmd('print');
  }, true);
  window.addEventListener('beforeprint', () => {
    if (busy || document.getElementById('print-root') || !app.wb || document.body.classList.contains('file-locked')) return;
    if (!mount(app, { sheets: [app.sheet] })) return;
    window.addEventListener('afterprint', () => { if (!busy) unmount(); }, { once: true });
  });
}
