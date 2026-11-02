// Dialogs and popups.
import { el, icon, colName, rangeAddr, addr, parseRange, esc, uid, key } from '../util.js';
import { dialog, popup, closeMenus, toast, confirmDlg, promptDlg, field, select, checkbox, segmented, colorGrid, pickColor, alertDlg } from './ui.js';
import { format, compile, NUMBER_FORMATS, MORE_FORMATS, L, LOCALES, fmtSmart } from '../format.js';
import { serial } from '../util.js';
import { DOCS, CATEGORIES, allFunctions } from '../docs.js';
import { TEMPLATES } from '../templates.js';
import * as ops from '../ops.js';
import * as io from '../io.js';
import * as store from '../store.js';
import { FErr } from '../formula/core.js';
import { THEME_FONTS, ensureFont } from './toolbar.js';

const M = () => window.amatu?.MOD || 'Ctrl';

// ---------- link & note ----------
export function linkDialog(app) {
  const g = app.grid, sh = app.sheet;
  const cell = sh.get(g.sel.r, g.sel.c);
  const text = el('input.inp', { value: app.editTextOf(sh, g.sel.r, g.sel.c).startsWith('=') ? '' : app.editTextOf(sh, g.sel.r, g.sel.c), placeholder: 'Teks yang ditampilkan' });
  const url = el('input.inp', { value: cell?.link || '', placeholder: 'Tempel link, atau ketik rentang seperti Sheet1!A1' });
  dialog({
    title: 'Sisipkan link', width: 440,
    body: el('div.stack', field('Teks', text), field('Link', url, 'Link ke situs web, alamat email (mailto:) atau rentang di spreadsheet ini.')),
    buttons: [{ label: 'Batal' }, { label: 'Terapkan', primary: true, run: () => {
      const u = url.value.trim();
      if (u && parseRange(u.replace(/^#/, ''))) { app.cmd('setLink', '#' + u.replace(/^#/, ''), text.value); return; }
      app.cmd('setLink', u, text.value);
    } }],
  });
}

export function noteDialog(app) {
  const g = app.grid, sh = app.sheet;
  const cell = sh.get(g.sel.r, g.sel.c);
  const ta = el('textarea.inp', { rows: 5, placeholder: 'Tulis catatan…' }, cell?.note || '');
  dialog({
    title: `Catatan · ${addr(g.sel.r, g.sel.c)}`, width: 380,
    body: ta,
    buttons: [cell?.note ? { label: 'Hapus catatan', danger: true, run: () => app.cmd('setNote', null) } : null, { spacer: true }, { label: 'Batal' }, { label: 'Simpan', primary: true, run: () => app.cmd('setNote', ta.value.trim()) }].filter(Boolean),
  });
}

// ---------- find / replace ----------
let findBar = null;
export function findDialog(app) {
  if (findBar) { findBar.querySelector('input').select(); return; }
  const inp = el('input.inp', { placeholder: 'Cari di sheet', 'aria-label': 'Cari' });
  const count = el('span.find-count');
  let hits = [], idx = -1;
  const run = () => {
    hits = ops.findAll(app, inp.value, {});
    app.findHits = { sh: app.sheet, list: hits };
    idx = hits.length ? 0 : -1;
    go();
  };
  const go = () => {
    count.textContent = inp.value ? (hits.length ? `${idx + 1} dari ${hits.length}` : 'Tidak ada hasil') : '';
    if (idx >= 0) { const h = hits[idx]; app.grid.select(h.r, h.c); }
    app.grid.render();
  };
  const close = () => { findBar.remove(); findBar = null; app.findHits = null; app.grid.render(); app.grid.focus(); };
  inp.addEventListener('input', run);
  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); if (hits.length) { idx = (idx + (e.shiftKey ? -1 : 1) + hits.length) % hits.length; go(); } }
    if (e.key === 'Escape') close();
  });
  findBar = el('div.findbar.glass', icon('search'), inp, count,
    el('button.icon-btn', { type: 'button', 'data-tip': 'Sebelumnya', onclick: () => { if (hits.length) { idx = (idx - 1 + hits.length) % hits.length; go(); } } }, icon('keyboard_arrow_up')),
    el('button.icon-btn', { type: 'button', 'data-tip': 'Berikutnya', onclick: () => { if (hits.length) { idx = (idx + 1) % hits.length; go(); } } }, icon('keyboard_arrow_down')),
    el('button.icon-btn', { type: 'button', 'data-tip': 'Opsi lainnya', onclick: () => { const q = inp.value; close(); replaceDialog(app, true, q); } }, icon('more_vert')),
    el('button.icon-btn', { type: 'button', 'data-tip': 'Tutup', onclick: close }, icon('close')));
  document.getElementById('grid-wrap').append(findBar);
  inp.focus();
}

export function replaceDialog(app, withReplace = true, initial = '') {
  const q = el('input.inp', { value: initial, placeholder: 'Teks yang dicari' });
  const rep = el('input.inp', { placeholder: 'Ganti dengan' });
  const scope = select([{ value: 'sheet', label: 'Sheet ini' }, { value: 'all', label: 'Semua sheet' }], 'sheet');
  const opt = { matchCase: false, whole: false, regex: false, formulas: false };
  const status = el('div.muted.small');
  let hits = [], idx = -1, lastQ = null;
  const opts = () => ({ ...opt, allSheets: scope.value === 'all' });
  const search = () => {
    hits = ops.findAll(app, q.value, opts());
    lastQ = q.value + JSON.stringify(opts());
    app.findHits = { sh: app.sheet, list: hits.filter(h => h.sh === app.sheet) };
    status.textContent = hits.length ? `${hits.length} kecocokan ditemukan` : 'Tidak ditemukan';
  };
  const next = () => {
    if (lastQ !== q.value + JSON.stringify(opts())) { search(); idx = -1; }
    if (!hits.length) return;
    idx = (idx + 1) % hits.length;
    const h = hits[idx];
    if (h.sh !== app.sheet) app.switchSheet(app.wb.sheets.indexOf(h.sh));
    app.grid.select(h.r, h.c);
    status.textContent = `${idx + 1} dari ${hits.length}`;
  };
  dialog({
    title: 'Cari dan ganti', width: 470, onClose: () => { app.findHits = null; app.grid.render(); },
    body: el('div.stack',
      field('Cari', q),
      withReplace ? field('Ganti dengan', rep) : null,
      field('Telusuri', scope),
      el('div.checks',
        checkbox('Cocokkan huruf besar/kecil', false, v => (opt.matchCase = v)),
        checkbox('Cocokkan seluruh isi sel', false, v => (opt.whole = v)),
        checkbox('Telusuri menggunakan ekspresi reguler', false, v => (opt.regex = v)),
        checkbox('Telusuri juga dalam rumus', false, v => (opt.formulas = v))),
      status),
    buttons: [
      { label: 'Cari', run: () => { next(); return false; } },
      withReplace ? { label: 'Ganti', run: () => {
        if (idx < 0 || !hits[idx]) { next(); return false; }
        const n = app.cmd('replaceAll', [hits[idx]], q.value, rep.value, opts());
        search(); idx--; next(); if (n) toast('1 kecocokan diganti'); return false;
      } } : null,
      withReplace ? { label: 'Ganti semua', run: () => { search(); const n = app.cmd('replaceAll', hits, q.value, rep.value, opts()); status.textContent = `${n} kecocokan diganti`; toast(`${n} kecocokan diganti`); search(); return false; } } : null,
      { spacer: true },
      { label: 'Selesai', primary: true },
    ].filter(Boolean),
  });
}

// ---------- format cells (Excel Ctrl+1) ----------
export function formatCellsDialog(app) {
  const g = app.grid, sh = app.sheet;
  const st0 = sh.styleOf(g.sel.r, g.sel.c);
  const patch = {};
  const st = () => ({ ...st0, ...patch });
  const sample = app.wb.eng.value(sh, g.sel.r, g.sel.c);
  const body = el('div.fc');
  const tabs = el('div.tabs-h');
  const pages = {};
  const show = id => { tabs.querySelectorAll('button').forEach(b => b.classList.toggle('on', b.dataset.id === id)); Object.entries(pages).forEach(([k, p]) => (p.style.display = k === id ? '' : 'none')); };
  const addTab = (id, label, page) => { tabs.append(el('button', { type: 'button', dataset: { id }, onclick: () => show(id) }, label)); pages[id] = page; body.append(page); };
  // Number
  const cats = [
    ['Umum', 'General'], ['Angka', '#,##0.00'], ['Mata uang', '"' + L.cur + '"#,##0.00'], ['Akuntansi', '_("' + L.cur + '"* #,##0.00_);_("' + L.cur + '"* (#,##0.00);_("' + L.cur + '"* "-"??_);_(@_)'],
    ['Tanggal', L.date], ['Waktu', 'hh:mm:ss'], ['Persentase', '0.00%'], ['Pecahan', '# ?/?'], ['Ilmiah', '0.00E+00'], ['Teks', '@'], ['Kustom', null],
  ];
  const code = el('input.inp.mono', { value: st0.fmt || 'General' });
  const prev = el('div.fc-preview');
  const upd = () => { patch.fmt = code.value === 'General' ? undefined : code.value; const v = typeof sample === 'number' || typeof sample === 'string' ? sample : 1234.5678; try { prev.textContent = format(v, code.value).t || ' '; } catch { prev.textContent = '—'; } };
  code.addEventListener('input', upd);
  const curCat = (() => { const f = st0.fmt || 'General'; const cc = compile(f); if (f === 'General') return 'Umum'; if (f === '@') return 'Teks'; if (cc?.isDate) return /[hs]/.test(f.replace(/"[^"]*"/g, '')) && !/[dy]/i.test(f) ? 'Waktu' : 'Tanggal'; if (cc?.isPct) return 'Persentase'; if (/E\+/.test(f)) return 'Ilmiah'; if (/\?\/\?/.test(f)) return 'Pecahan'; if (/_\(/.test(f)) return 'Akuntansi'; if (/Rp|\$|€|£|¥/.test(f)) return 'Mata uang'; if (/[0#]/.test(f)) return 'Angka'; return 'Kustom'; })();
  const catList = el('div.fc-cats', ...cats.map(([lb, f]) => el('button.fc-cat' + (lb === curCat ? '.on' : ''), { type: 'button', onclick: e => { catList.querySelectorAll('.on').forEach(x => x.classList.remove('on')); e.currentTarget.classList.add('on'); if (f) { code.value = f; upd(); } else code.focus(); } }, lb)));
  const presetBox = el('div.fc-presets');
  for (const [grp, list] of MORE_FORMATS) for (const f of list) presetBox.append(el('button.fc-preset', { type: 'button', onclick: () => { code.value = f; upd(); } }, el('code', f), el('span.muted', format(/[dmyhs]/i.test(f.replace(/"[^"]*"/g, '')) && !/[0#]/.test(f) ? 46000.6 : 1234.5678, f).t)));
  upd();
  addTab('num', 'Angka', el('div.fc-page.fc-num', catList, el('div.stack', el('div.muted.small', 'Contoh'), prev, field('Kode format', code), el('div.muted.small', 'Kode umum'), presetBox)));
  // Alignment
  const ha = select([{ value: '', label: 'Umum' }, { value: 'left', label: 'Kiri' }, { value: 'center', label: 'Tengah' }, { value: 'right', label: 'Kanan' }], st0.ha || '', v => (patch.ha = v || undefined));
  const va = select([{ value: 'top', label: 'Atas' }, { value: 'middle', label: 'Tengah' }, { value: 'bottom', label: 'Bawah' }], st0.va || 'bottom', v => (patch.va = v === 'bottom' ? undefined : v));
  const ind = el('input.inp', { type: 'number', min: 0, max: 15, value: st0.ind || 0 });
  ind.addEventListener('input', () => (patch.ind = +ind.value || undefined));
  const rot = el('input.inp', { type: 'number', min: -90, max: 90, value: typeof st0.rot === 'number' ? st0.rot : 0 });
  rot.addEventListener('input', () => (patch.rot = +rot.value || undefined));
  addTab('align', 'Perataan', el('div.fc-page.grid2',
    field('Horizontal', ha), field('Vertikal', va), field('Indentasi', ind), field('Orientasi (derajat)', rot),
    el('div.span2.checks',
      checkbox('Kemas teks', st0.wrap === 'wrap', v => (patch.wrap = v ? 'wrap' : undefined)),
      checkbox('Teks vertikal (tumpuk)', st0.rot === 'v', v => (patch.rot = v ? 'v' : undefined)),
      checkbox('Gabungkan sel', !!sh.mergeAt(g.sel.r, g.sel.c), v => (patch._merge = v)))));
  // Font
  const ff = select(THEME_FONTS, st0.ff || app.defFont.ff, v => { patch.ff = v; ensureFont(v, app); fprev.style.fontFamily = v; });
  const fs = el('input.inp', { type: 'number', min: 1, max: 400, value: st0.fs || app.defFont.fs });
  fs.addEventListener('input', () => { patch.fs = +fs.value || undefined; fprev.style.fontSize = (+fs.value || 10) + 'pt'; });
  const fprev = el('div.fc-fontprev', { style: { fontFamily: st0.ff || app.defFont.ff, fontSize: (st0.fs || 10) + 'pt', fontWeight: st0.b ? 700 : 400, fontStyle: st0.i ? 'italic' : 'normal', color: st0.color || '' } }, 'AaBbCcYyZz 123');
  const colorSw = el('button.swatch-btn', { type: 'button', style: { background: st0.color || '#000' }, onclick: e => pickColor(e.currentTarget.getBoundingClientRect(), c => { patch.color = c || undefined; colorSw.style.background = c || '#000'; fprev.style.color = c || ''; }) });
  addTab('font', 'Font', el('div.fc-page.grid2',
    field('Font', ff), field('Ukuran', fs),
    el('div.checks',
      checkbox('Tebal', st0.b, v => { patch.b = v || undefined; fprev.style.fontWeight = v ? 700 : 400; }),
      checkbox('Miring', st0.i, v => { patch.i = v || undefined; fprev.style.fontStyle = v ? 'italic' : 'normal'; }),
      checkbox('Garis bawah', st0.u, v => (patch.u = v || undefined)),
      checkbox('Coret', st0.s, v => (patch.s = v || undefined))),
    field('Warna', colorSw),
    el('div.span2', fprev)));
  // Border
  let bColor = '#000000', bStyle = 'solid', bW = 1, bType = null;
  const bprev = el('div.fc-bprev', el('span', 'Teks'));
  const setB = t => { bType = t; bprev.className = 'fc-bprev b-' + t; bprev.style.borderColor = bColor; };
  addTab('border', 'Batas', el('div.fc-page',
    el('div.fc-bpresets', ...[['clear', 'border_clear', 'Tidak ada'], ['outer', 'border_outer', 'Luar'], ['inner', 'border_inner', 'Dalam'], ['all', 'border_all', 'Semua'], ['top', 'border_top', 'Atas'], ['bottom', 'border_bottom', 'Bawah'], ['left', 'border_left', 'Kiri'], ['right', 'border_right', 'Kanan']].map(([t, ic, lb]) => el('button.fc-bp', { type: 'button', onclick: () => setB(t) }, icon(ic), el('span', lb)))),
    el('div.grid2',
      field('Gaya', select([{ value: 'solid:1', label: 'Tipis' }, { value: 'solid:2', label: 'Sedang' }, { value: 'solid:3', label: 'Tebal' }, { value: 'dashed:1', label: 'Putus-putus' }, { value: 'dotted:1', label: 'Titik-titik' }, { value: 'double:3', label: 'Ganda' }], 'solid:1', v => { const [s, w] = v.split(':'); bStyle = s; bW = +w; })),
      field('Warna', el('input.inp.color', { type: 'color', value: '#000000', oninput: e => { bColor = e.target.value; bprev.style.borderColor = bColor; } }))),
    bprev));
  // Fill
  const fillPrev = el('div.fc-fillprev', { style: { background: st0.bg || 'transparent' } });
  addTab('fill', 'Isian', el('div.fc-page.fc-fill', colorGrid(c => { patch.bg = c || undefined; fillPrev.style.background = c || 'transparent'; }, { resetLabel: 'Tanpa warna' }), el('div.stack', el('div.muted.small', 'Contoh'), fillPrev)));
  // Protection
  addTab('prot', 'Perlindungan', el('div.fc-page.stack', el('p.dlg-text', 'Perlindungan berlaku setelah sheet dilindungi melalui Data › Lindungi sheet. Gunakan pengecualian untuk rentang yang tetap dapat diedit.'),
    checkbox('Izinkan rentang ini diedit saat sheet dilindungi', (sh.protectExcept || []).some(x => x.r1 === g.range.r1 && x.c1 === g.range.c1 && x.r2 === g.range.r2 && x.c2 === g.range.c2), v => (patch._except = v))));
  show('num');
  dialog({
    title: 'Format sel', width: 620, cls: 'dlg-fc',
    body: el('div', tabs, body),
    buttons: [{ label: 'Batal' }, { label: 'OK', primary: true, run: () => {
      const { _merge, _except, ...p } = patch;
      app.tx('Format sel', () => {
        if (Object.keys(p).length) ops.styleRanges(app, app.ranges(), s => { const o = { ...s }; for (const [k, v] of Object.entries(p)) { if (v === undefined) delete o[k]; else o[k] = v; } return o; });
        if (bType) ops.bordersOn(app, app.ranges(), bType, { color: bColor, style: bStyle, w: bW });
        if (_merge !== undefined) ops.mergeRanges(app, app.ranges(), _merge ? 'all' : 'none');
        if (_except !== undefined) { app.wb.hist.touchMeta(sh); const list = (sh.protectExcept || []).filter(x => !(x.r1 === g.range.r1 && x.c1 === g.range.c1 && x.r2 === g.range.r2 && x.c2 === g.range.c2)); if (_except) list.push({ ...g.range }); sh.protectExcept = list; }
        app.wb.eng.recalc();
        for (const rg of app.ranges()) for (let r = rg.r1; r <= Math.min(rg.r2, sh.dataRows()); r++) ops.fitRow(app, sh, r);
      });
    } }],
  });
}

export function customFormatDialog(app, kind) {
  const g = app.grid, sh = app.sheet;
  const st = sh.styleOf(g.sel.r, g.sel.c);
  const v = app.wb.eng.value(sh, g.sel.r, g.sel.c);
  const inp = el('input.inp.mono', { value: st.fmt || (kind === 'date' ? 'd mmmm yyyy' : kind === 'currency' ? '"' + L.cur + '"#,##0.00' : '#,##0.00') });
  const prev = el('div.fc-preview');
  const sampleV = typeof v === 'number' ? v : kind === 'date' ? 46000.6 : 1234.5678;
  const upd = () => { prev.textContent = format(sampleV, inp.value).t || ' '; };
  inp.addEventListener('input', upd);
  upd();
  const group = MORE_FORMATS.find(([n]) => (kind === 'date' ? n.startsWith('Tanggal') : kind === 'currency' ? n.startsWith('Mata') : n === 'Angka'));
  const list = el('div.fc-presets.tall', ...group[1].map(f => el('button.fc-preset', { type: 'button', onclick: () => { inp.value = f; upd(); } }, el('span', format(sampleV, f).t), el('code.muted', f))));
  const help = kind === 'date' ? 'd, dd, ddd, dddd · m, mm, mmm, mmmm · yy, yyyy · h, hh · mm · ss · AM/PM' : '0 digit wajib · # digit opsional · , pemisah ribuan · . desimal · % persen · "teks" · [Red] warna · ; bagian positif;negatif;nol;teks';
  dialog({
    title: kind === 'date' ? 'Format tanggal dan waktu kustom' : kind === 'currency' ? 'Format mata uang kustom' : 'Format angka kustom',
    width: 480,
    body: el('div.stack', field('Kode format', inp, help), el('div.muted.small', 'Pratinjau'), prev, list),
    buttons: [{ label: 'Batal' }, { label: 'Terapkan', primary: true, run: () => app.cmd('numberFormat', inp.value) }],
  });
}

export function pasteSpecialDialog(app) {
  let mode = 'all', op = null, skip = false, tr = false;
  const radios = (name, opts, onSel) => el('div.radios', ...opts.map(([v, lb], i) => el('label.radio', el('input', { type: 'radio', name, checked: i === 0, onchange: () => onSel(v) }), el('span', lb))));
  dialog({
    title: 'Tempel spesial', width: 460,
    body: el('div.grid2',
      el('div', el('h4', 'Tempel'), radios('ps-mode', [['all', 'Semua'], ['formulas', 'Rumus'], ['values', 'Nilai'], ['formats', 'Format'], ['valid', 'Validasi'], ['nofmt', 'Semua kecuali format'], ['widths', 'Lebar kolom']], v => (mode = v))),
      el('div', el('h4', 'Operasi'), radios('ps-op', [[null, 'Tidak ada'], ['add', 'Tambah'], ['sub', 'Kurangi'], ['mul', 'Kalikan'], ['div', 'Bagi']], v => (op = v)),
        el('div.checks.mt', checkbox('Lewati sel kosong', false, v => (skip = v)), checkbox('Transpos', false, v => (tr = v))))),
    buttons: [{ label: 'Batal' }, { label: 'OK', primary: true, run: () => app.pasteFromMenu(tr ? 'transpose' : mode, { op, skipBlanks: skip }) }],
  });
}

// ---------- resize / insert / delete ----------
export function resizeDialog(app, kind) {
  const g = app.grid;
  const list = kind === 'col' ? g.colsFor(g.range.c1) : g.rowsFor(g.range.r1);
  const cur = kind === 'col' ? app.sheet.colWidth(list[0]) : app.sheet.rowHeight(list[0]);
  const inp = el('input.inp', { type: 'number', min: 2, max: 2000, value: cur });
  let fit = false;
  dialog({
    title: kind === 'col' ? `Ubah ukuran kolom ${colName(list[0])}${list.length > 1 ? '–' + colName(list[list.length - 1]) : ''}` : `Ubah ukuran baris ${list[0] + 1}${list.length > 1 ? '–' + (list[list.length - 1] + 1) : ''}`,
    width: 360,
    body: el('div.stack',
      el('label.radio', el('input', { type: 'radio', name: 'rz', checked: true, onchange: () => (fit = false) }), el('span', `Tentukan ${kind === 'col' ? 'lebar' : 'tinggi'} dalam piksel`)), inp,
      el('label.radio', el('input', { type: 'radio', name: 'rz', onchange: () => (fit = true) }), el('span', 'Sesuaikan dengan data'))),
    buttons: [{ label: 'Batal' }, { label: 'OK', primary: true, run: () => {
      if (fit) app.cmd(kind === 'col' ? 'autofitCols' : 'autofitRows', list);
      else app.cmd(kind === 'col' ? 'resizeCols' : 'resizeRows', list, Math.max(2, +inp.value || cur));
    } }],
  });
}

export function insertCellsDialog(app) {
  let v = 'down';
  dialog({
    title: 'Sisipkan', width: 320,
    body: el('div.stack', ...[['right', 'Geser sel ke kanan'], ['down', 'Geser sel ke bawah'], ['row', 'Seluruh baris'], ['col', 'Seluruh kolom']].map(([k, lb]) => el('label.radio', el('input', { type: 'radio', name: 'ins', checked: k === 'down', onchange: () => (v = k) }), el('span', lb)))),
    buttons: [{ label: 'Batal' }, { label: 'OK', primary: true, run: () => (v === 'row' ? app.cmd('insertRows', 'above') : v === 'col' ? app.cmd('insertCols', 'left') : app.cmd('shiftCells', v)) }],
  });
}
export function deleteCellsDialog(app) {
  let v = 'up';
  dialog({
    title: 'Hapus', width: 320,
    body: el('div.stack', ...[['left', 'Geser sel ke kiri'], ['up', 'Geser sel ke atas'], ['row', 'Seluruh baris'], ['col', 'Seluruh kolom']].map(([k, lb]) => el('label.radio', el('input', { type: 'radio', name: 'del', checked: k === 'up', onchange: () => (v = k) }), el('span', lb)))),
    buttons: [{ label: 'Batal' }, { label: 'OK', primary: true, run: () => (v === 'row' ? app.cmd('deleteRows') : v === 'col' ? app.cmd('deleteCols') : app.cmd('shiftCells', v)) }],
  });
}

// ---------- filter popup ----------
export function filterPopup(app, c) {
  const sh = app.sheet, f = sh.filter;
  if (!f) return;
  const g = app.grid;
  const crit = f.crit?.[c] || {};
  const end = Math.max(f.r2, sh.dataRows() - 1);
  const counts = new Map();
  for (let r = f.r1 + 1; r <= end; r++) {
    // values visible under the other columns' filters
    let hiddenByOther = false;
    for (const [cs, cr] of Object.entries(f.crit || {})) {
      if (+cs === c) continue;
      const d = g.display(r, +cs, sh.get(r, +cs));
      const t = d ? String(d.text ?? '') : '';
      if (cr.hidden?.includes(t) || (cr.cond && !ops.condMatch(cr.cond, app.wb.eng.value(sh, r, +cs), t, app.wb.eng, sh, r, +cs))) { hiddenByOther = true; break; }
    }
    if (hiddenByOther) continue;
    const d = g.display(r, c, sh.get(r, c));
    const t = d ? String(d.text ?? '') : '';
    counts.set(t, (counts.get(t) || 0) + 1);
  }
  const vals = [...counts.keys()].sort((a, b) => (a === '' ? 1 : b === '' ? -1 : a.localeCompare(b, undefined, { numeric: true })));
  const hidden = new Set(crit.hidden || []);
  const list = el('div.flt-list');
  const search = el('input.inp.sm', { placeholder: 'Telusuri', 'aria-label': 'Telusuri nilai' });
  const draw = () => {
    list.innerHTML = '';
    const q = search.value.toLowerCase();
    for (const v of vals) {
      if (q && !v.toLowerCase().includes(q)) continue;
      const cb = el('input', { type: 'checkbox', checked: !hidden.has(v), onchange: () => { cb.checked ? hidden.delete(v) : hidden.add(v); } });
      list.append(el('label.flt-item', cb, el('span.flt-v', v === '' ? '(Kosong)' : v), el('span.muted', String(counts.get(v)))));
    }
  };
  search.addEventListener('input', draw);
  draw();
  const condSel = select([
    { value: '', label: 'Tidak ada' }, { value: 'empty', label: 'Kosong' }, { value: 'notEmpty', label: 'Tidak kosong' },
    { value: 'contains', label: 'Teks berisi' }, { value: 'notContains', label: 'Teks tidak berisi' }, { value: 'starts', label: 'Teks diawali dengan' }, { value: 'ends', label: 'Teks diakhiri dengan' }, { value: 'exact', label: 'Teks sama persis' },
    { value: 'gt', label: 'Lebih besar dari' }, { value: 'ge', label: 'Lebih besar dari atau sama dengan' }, { value: 'lt', label: 'Kurang dari' }, { value: 'le', label: 'Kurang dari atau sama dengan' },
    { value: 'eq', label: 'Sama dengan' }, { value: 'ne', label: 'Tidak sama dengan' }, { value: 'between', label: 'Di antara' }, { value: 'notBetween', label: 'Tidak di antara' }, { value: 'formula', label: 'Rumus kustom' },
  ], crit.cond?.op || '');
  const v1 = el('input.inp.sm', { value: crit.cond?.v1 ?? '', placeholder: 'Nilai atau rumus' });
  const v2 = el('input.inp.sm', { value: crit.cond?.v2 ?? '', placeholder: 'Nilai kedua' });
  const updCond = () => { v1.style.display = ['', 'empty', 'notEmpty'].includes(condSel.value) ? 'none' : ''; v2.style.display = ['between', 'notBetween'].includes(condSel.value) ? '' : 'none'; };
  condSel.addEventListener('change', updCond);
  updCond();
  const sec = (title, content, open) => { const d = el('details.flt-sec', { open }, el('summary', title), content); return d; };
  let p;
  const box = el('div.flt',
    el('button.menu-item', { type: 'button', onclick: () => { p.close(); app.cmd('sortFilterCol', c, true); } }, el('span.mi-icon', icon('arrow_downward')), el('span.mi-label', 'Urutkan A → Z')),
    el('button.menu-item', { type: 'button', onclick: () => { p.close(); app.cmd('sortFilterCol', c, false); } }, el('span.mi-icon', icon('arrow_upward')), el('span.mi-label', 'Urutkan Z → A')),
    el('div.menu-sep'),
    sec('Filter menurut kondisi', el('div.stack.sm', condSel, v1, v2), !!crit.cond),
    sec('Filter menurut nilai', el('div.stack.sm',
      el('div.flt-tools', el('button.btn.link', { type: 'button', onclick: () => { hidden.clear(); draw(); } }, 'Pilih semua'), el('span.muted', '·'), el('button.btn.link', { type: 'button', onclick: () => { vals.forEach(v => hidden.add(v)); draw(); } }, 'Hapus'), el('span.muted.flex1.right', `Menampilkan ${vals.length} item`)),
      search, list), true),
    el('div.flt-foot', el('button.btn', { type: 'button', onclick: () => p.close() }, 'Batal'), el('button.btn.primary', { type: 'button', onclick: () => {
      const cond = condSel.value ? { op: condSel.value, v1: v1.value, v2: v2.value } : null;
      const h = [...hidden].filter(v => counts.has(v));
      app.cmd('setFilterCrit', c, h.length || cond ? { hidden: h.length ? h : undefined, cond: cond || undefined } : null);
      p.close();
    } }, 'OK')));
  const rc = g.cellRect(f.r1, c), hr = g.host.getBoundingClientRect();
  p = popup(box, new DOMRect(hr.left + (rc.x + rc.w) * g.z - 260, hr.top + rc.y * g.z, 260, rc.h * g.z), { cls: 'flt-pop' });
}

// ---------- sort / dedupe / split ----------
export function sortDialog(app) {
  const sh = app.sheet;
  let g = { ...app.grid.range };
  if (g.r1 === g.r2 && g.c1 === g.c2) g = io.currentRegion(app, g.r1, g.c1);
  const rangeInp = el('input.inp', { value: rangeAddr(g) });
  let header = typeof app.wb.eng.value(sh, g.r1, g.c1) === 'string' && typeof app.wb.eng.value(sh, g.r1 + 1, g.c1) !== 'string';
  const keysBox = el('div.stack.sm');
  const colOpts = () => { const gg = parseRange(rangeInp.value) || g; const out = []; for (let c = gg.c1; c <= Math.min(gg.c2, sh.cols - 1); c++) { const hv = app.wb.eng.value(sh, gg.r1, c); out.push({ value: c, label: header && hv != null ? `${hv}` : `Kolom ${colName(c)}` }); } return out; };
  const keys = [{ c: app.grid.sel.c, asc: true }];
  const draw = () => {
    keysBox.innerHTML = '';
    keys.forEach((k, i) => keysBox.append(el('div.sort-key',
      el('span.muted', i ? 'lalu menurut' : 'Urutkan menurut'),
      select(colOpts(), k.c, v => (k.c = +v)),
      segmented([{ value: true, label: 'A → Z' }, { value: false, label: 'Z → A' }], k.asc, v => (k.asc = v)),
      i ? el('button.icon-btn', { type: 'button', onclick: () => { keys.splice(i, 1); draw(); } }, icon('close')) : null)));
  };
  draw();
  dialog({
    title: 'Urutkan rentang', width: 520,
    body: el('div.stack', field('Rentang', rangeInp), checkbox('Data memiliki baris header', header, v => { header = v; draw(); }), keysBox,
      el('button.btn.link', { type: 'button', onclick: () => { keys.push({ c: (keys[keys.length - 1]?.c ?? g.c1) + 1, asc: true }); draw(); } }, icon('add'), 'Tambahkan kolom urutan lain')),
    buttons: [{ label: 'Batal' }, { label: 'Urutkan', primary: true, run: () => { const gg = parseRange(rangeInp.value); if (!gg) { toast('Rentang tidak valid'); return false; } app.cmd('sortRange', gg, keys, header); } }],
  });
}

export function removeDuplicatesDialog(app) {
  const sh = app.sheet;
  let g = { ...app.grid.range };
  if (g.r1 === g.r2 && g.c1 === g.c2) g = io.currentRegion(app, g.r1, g.c1);
  if (g.r2 >= sh.rows - 1) g.r2 = Math.max(g.r1, sh.dataRows() - 1);
  let header = typeof app.wb.eng.value(sh, g.r1, g.c1) === 'string';
  const sel = new Set(Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => g.c1 + i));
  const box = el('div.checks.scroll');
  const draw = () => {
    box.innerHTML = '';
    box.append(checkbox('Pilih semua', sel.size === g.c2 - g.c1 + 1, v => { for (let c = g.c1; c <= g.c2; c++) v ? sel.add(c) : sel.delete(c); draw(); }));
    for (let c = g.c1; c <= g.c2; c++) { const hv = app.wb.eng.value(sh, g.r1, c); box.append(checkbox(header && hv != null ? `Kolom ${colName(c)} – ${hv}` : `Kolom ${colName(c)}`, sel.has(c), v => (v ? sel.add(c) : sel.delete(c)))); }
  };
  draw();
  dialog({
    title: 'Hapus duplikat', width: 420,
    body: el('div.stack', el('p.muted', `Rentang data: ${rangeAddr(g)}`), checkbox('Data memiliki baris header', header, v => { header = v; draw(); }), el('div.muted.small', 'Kolom yang dianalisis'), box),
    buttons: [{ label: 'Batal' }, { label: 'Hapus duplikat', primary: true, run: () => { if (!sel.size) { toast('Pilih minimal satu kolom'); return false; } app.cmd('removeDuplicates', g, [...sel].sort((a, b) => a - b), header); } }],
  });
}

// ---------- goal seek ----------
export function goalSeekDialog(app) {
  const g = app.grid;
  const setCell = el('input.inp', { value: addr(g.sel.r, g.sel.c) });
  const toVal = el('input.inp', { type: 'text', placeholder: 'Nilai target' });
  const byCell = el('input.inp', { placeholder: 'mis. B2' });
  dialog({
    title: 'Pencarian Tujuan', width: 380,
    body: el('div.stack', field('Atur sel', setCell, 'Sel yang berisi rumus'), field('Ke nilai', toVal), field('Dengan mengubah sel', byCell, 'Sel berisi nilai (bukan rumus)')),
    buttons: [{ label: 'Batal' }, { label: 'OK', primary: true, run: () => {
      const a = parseRange(setCell.value), b = parseRange(byCell.value), v = Number(String(toVal.value).replace(',', '.'));
      if (!a || !b || !isFinite(v)) { toast('Lengkapi sel dan nilai dengan benar'); return false; }
      const sh = app.sheet;
      if (sh.get(a.r1, a.c1)?.f == null) { toast('Sel yang diatur harus berisi rumus'); return false; }
      if (sh.get(b.r1, b.c1)?.f != null) { toast('Sel yang diubah harus berisi nilai, bukan rumus'); return false; }
      const res = app.cmd('goalSeek', { r: a.r1, c: a.c1 }, v, { r: b.r1, c: b.c1 });
      if (res?.ok) alertDlg('Status Pencarian Tujuan', `Solusi ditemukan untuk sel ${addr(a.r1, a.c1)}.\nNilai target: ${format(v).t}\nNilai saat ini: ${format(+res.value.toPrecision(12)).t}\n${addr(b.r1, b.c1)} = ${format(+res.x.toPrecision(12)).t}`);
      else alertDlg('Status Pencarian Tujuan', 'Pencarian tujuan tidak menemukan solusi. Coba nilai awal lain pada sel yang diubah.');
    } }],
  });
}

// ---------- pivot ----------
export function pivotDialog(app) {
  const sh = app.sheet;
  let g = { ...app.grid.range };
  if (g.r1 === g.r2 && g.c1 === g.c2) g = io.currentRegion(app, g.r1, g.c1);
  if (g.r2 >= sh.rows - 1) g.r2 = Math.max(g.r1, sh.dataRows() - 1);
  const rangeInp = el('input.inp', { value: rangeAddr(g, sh.name) });
  let newSheet = true;
  dialog({
    title: 'Buat tabel pivot', width: 420,
    body: el('div.stack', field('Rentang data', rangeInp),
      el('div.muted.small', 'Sisipkan ke'),
      el('label.radio', el('input', { type: 'radio', name: 'pv', checked: true, onchange: () => (newSheet = true) }), el('span', 'Sheet baru')),
      el('label.radio', el('input', { type: 'radio', name: 'pv', onchange: () => (newSheet = false) }), el('span', 'Sheet yang ada (mulai dari sel aktif)'))),
    buttons: [{ label: 'Batal' }, { label: 'Buat', primary: true, run: () => {
      const r = parseRange(rangeInp.value);
      if (!r) { toast('Rentang tidak valid'); return false; }
      const def = { srcSheet: r.sheet || sh.name, range: { r1: r.r1, c1: r.c1, r2: r.r2 === Infinity ? sh.rows - 1 : r.r2, c2: r.c2 === Infinity ? sh.cols - 1 : r.c2 }, rows: [], cols: [], values: [], totals: true, at: newSheet ? { r: 0, c: 0 } : { r: app.grid.sel.r, c: app.grid.sel.c } };
      const target = app.cmd('pivotCreate', def, newSheet);
      app.sidebar.open('pivot', { sheet: target });
    } }],
  });
}

// ---------- stats ----------
export function columnStatsDialog(app) {
  const sh = app.sheet, eng = app.wb.eng;
  let c = app.grid.sel.c;
  const body = el('div.stats-dlg');
  const draw = () => {
    body.innerHTML = '';
    const vals = [];
    const end = sh.dataRows() - 1;
    let start = 0;
    const head = eng.value(sh, 0, c);
    if (typeof head === 'string' && typeof eng.value(sh, 1, c) !== 'string') start = 1;
    for (let r = start; r <= end; r++) { if (sh.isRowHidden(r)) continue; vals.push(eng.value(sh, r, c)); }
    const nonEmpty = vals.filter(v => v != null && v !== '');
    const nums = nonEmpty.filter(v => typeof v === 'number').sort((a, b) => a - b);
    const sum = nums.reduce((a, b) => a + b, 0);
    const uniq = new Map();
    for (const v of nonEmpty) { const k = v instanceof FErr ? v.code : String(v); uniq.set(k, (uniq.get(k) || 0) + 1); }
    const fmt = v => fmtSmart(v);
    const nav = el('div.stats-nav',
      el('button.icon-btn', { type: 'button', onclick: () => { if (c > 0) { c--; draw(); } } }, icon('chevron_left')),
      el('b', `Kolom ${colName(c)}${typeof head === 'string' ? ' – ' + head : ''}`),
      el('button.icon-btn', { type: 'button', onclick: () => { if (c < sh.cols - 1) { c++; draw(); } } }, icon('chevron_right')));
    const tbl = el('table.kv',
      ...[['Total baris', vals.length], ['Sel kosong', vals.length - nonEmpty.length], ['Nilai unik', uniq.size],
        ...(nums.length ? [['Jumlah', fmt(sum)], ['Rata-rata', fmt(sum / nums.length)], ['Median', fmt(nums.length % 2 ? nums[nums.length >> 1] : (nums[nums.length / 2 - 1] + nums[nums.length / 2]) / 2)], ['Minimum', fmt(nums[0])], ['Maksimum', fmt(nums[nums.length - 1])]] : [])]
        .map(([k, v]) => el('tr', el('td', k), el('td', String(v)))));
    const top = [...uniq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8);
    const max = top[0]?.[1] || 1;
    const freq = el('div.freq', ...top.map(([k, n]) => el('div.freq-row', el('span.freq-k', k), el('span.freq-bar', el('i', { style: { width: (n / max * 100) + '%' } })), el('span.muted', String(n)))));
    let hist = null;
    if (nums.length > 1) {
      const bins = Math.min(12, Math.max(4, Math.round(Math.sqrt(nums.length))));
      const lo = nums[0], hi = nums[nums.length - 1], w = (hi - lo) / bins || 1;
      const cnt = Array(bins).fill(0);
      for (const v of nums) cnt[Math.min(bins - 1, Math.floor((v - lo) / w))]++;
      const m = Math.max(...cnt);
      hist = el('div.hist', ...cnt.map((n, i) => el('div.hist-bar', { title: `${fmt(lo + i * w)} – ${fmt(lo + (i + 1) * w)}: ${n}`, style: { height: (n / m * 100) + '%' } })));
    }
    body.append(...[nav, el('div.grid2', el('div', el('h4', 'Ringkasan'), tbl), el('div', hist ? el('h4', 'Distribusi') : el('h4', 'Frekuensi'), hist || freq)), hist ? el('div', el('h4', 'Frekuensi teratas'), freq) : null].filter(Boolean));
  };
  draw();
  dialog({ title: 'Statistik kolom', width: 560, body, buttons: [{ label: 'Tutup', primary: true }] });
}

export function errorCheckDialog(app) {
  const list = [];
  for (const sh of app.wb.sheets) for (const [k, c] of sh.cells) if (c.f != null && c.sv instanceof FErr) list.push({ sh, r: Math.floor(k / 16384), c: k % 16384, e: c.sv });
  list.sort((a, b) => app.wb.sheets.indexOf(a.sh) - app.wb.sheets.indexOf(b.sh) || a.r - b.r || a.c - b.c);
  let d;
  const body = list.length ? el('div.err-list', ...list.slice(0, 300).map(x => el('button.err-item', { type: 'button', onclick: () => { d.close(); app.switchSheet(app.wb.sheets.indexOf(x.sh)); app.grid.select(x.r, x.c); } },
    el('span.err-code', x.e.code), el('span', `${x.sh.name}!${addr(x.r, x.c)}`), el('span.muted', x.e.msg || '')))) : el('p.dlg-text', 'Tidak ada rumus yang menghasilkan error. Bagus!');
  d = dialog({ title: `Pemeriksa error (${list.length})`, width: 520, body, buttons: [{ label: 'Tutup', primary: true }] });
}

// ---------- documents ----------
export async function openDialog(app, tab = 'recent') {
  const docs = await store.listDocs();
  const body = el('div.home');
  const tabs = el('div.tabs-h');
  const pages = {};
  const show = id => { tabs.querySelectorAll('button').forEach(b => b.classList.toggle('on', b.dataset.id === id)); Object.entries(pages).forEach(([k, p]) => (p.style.display = k === id ? '' : 'none')); };
  let d;
  // recent
  const search = el('input.inp', { placeholder: 'Telusuri spreadsheet', 'aria-label': 'Telusuri' });
  const list = el('div.doc-list');
  const draw = () => {
    list.innerHTML = '';
    const q = search.value.toLowerCase();
    const items = docs.filter(x => !q || x.name.toLowerCase().includes(q));
    if (!items.length) list.append(el('div.empty', icon('folder_off'), el('span', 'Belum ada spreadsheet tersimpan')));
    for (const x of items) {
      const row = el('div.doc-row' + (x.id === app.wb.id ? '.current' : ''),
        el('span.doc-ic', icon('grid_on')),
        el('div.doc-main', el('div.doc-name', x.name), el('div.muted.small', new Date(x.updated).toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' }) + ' · ' + x.sheets + ' sheet')),
        el('button.icon-btn', { type: 'button', 'data-tip': 'Hapus', onclick: async e => { e.stopPropagation(); if (await confirmDlg('Hapus spreadsheet', `Hapus "${x.name}" dari perangkat ini?`, 'Hapus', true)) { await store.deleteDoc(x.id); docs.splice(docs.indexOf(x), 1); draw(); if (x.id === app.wb.id) app.cmd('newDoc', 'blank'); } } }, icon('delete')));
      row.addEventListener('click', async () => { d.close(); if (x.id === app.wb.id) return; await app.save(); const j = await store.loadDoc(x.id); if (j) app.openWorkbook((await import('../model.js')).Workbook.fromJSON(j)); });
      list.append(row);
    }
  };
  search.addEventListener('input', draw);
  draw();
  pages.recent = el('div.stack', el('div.row', search, el('button.btn', { type: 'button', onclick: () => { d.close(); importDialog(app); } }, icon('upload_file'), 'Impor file')), list);
  // templates
  pages.templates = el('div.tpl-grid', ...TEMPLATES.filter(t => t.id !== 'welcome').map(t => el('button.tpl', { type: 'button', onclick: () => { d.close(); app.cmd('newDoc', t.id); } },
    el('div.tpl-thumb', { dataset: { kind: t.id } }, icon(t.icon)), el('div.tpl-name', t.name), el('div.muted.small', t.desc))));
  tabs.append(el('button', { type: 'button', dataset: { id: 'recent' }, onclick: () => show('recent') }, 'Terbaru'), el('button', { type: 'button', dataset: { id: 'templates' }, onclick: () => show('templates') }, 'Template'));
  body.append(tabs, pages.recent, pages.templates);
  show(tab);
  d = dialog({ title: 'Buka spreadsheet', width: 680, cls: 'dlg-home', body });
}

export function importDialog(app) {
  const file = el('input', { type: 'file', accept: '.xlsx,.xlsm,.xls,.ods,.csv,.tsv,.txt,.json', style: { display: 'none' } });
  let mode = 'new';
  const drop = el('div.dropzone', icon('upload_file'), el('div', el('b', 'Tarik file ke sini'), el('div.muted.small', 'atau klik untuk memilih · .xlsx, .xls, .ods, .csv, .tsv, .json')));
  drop.addEventListener('click', () => file.click());
  let d;
  const go = f => { d.close(); importFile(app, f, mode); };
  file.addEventListener('change', () => file.files[0] && go(file.files[0]));
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', e => { e.preventDefault(); e.stopPropagation(); const f = e.dataTransfer.files[0]; if (f) go(f); });
  d = dialog({
    title: 'Impor file', width: 460,
    body: el('div.stack', drop, file,
      el('div.muted.small', 'Lokasi impor'),
      ...[['new', 'Buat spreadsheet baru'], ['insert', 'Sisipkan sheet baru'], ['replace', 'Ganti sheet saat ini']].map(([v, lb], i) => el('label.radio', el('input', { type: 'radio', name: 'imp', checked: !i, onchange: () => (mode = v) }), el('span', lb)))),
  });
}

export async function importFile(app, f, mode = 'new') {
  try {
    toast(`Mengimpor ${f.name}…`, { ms: 1500 });
    const wb = await io.importFile(f);
    app.cmd('importSheets', wb, mode);
    toast(`${f.name} berhasil diimpor`);
  } catch (e) {
    console.error(e);
    alertDlg('Gagal mengimpor', e.message || String(e));
  }
}

export async function deleteDocDialog(app) {
  if (!(await confirmDlg('Hapus spreadsheet', `Hapus "${app.wb.name}" dari perangkat ini? Tindakan ini tidak dapat diurungkan.`, 'Hapus', true))) return;
  await store.deleteDoc(app.wb.id);
  const docs = await store.listDocs();
  if (docs.length) { const j = await store.loadDoc(docs[0].id); app.openWorkbook((await import('../model.js')).Workbook.fromJSON(j)); }
  else app.cmd('newDoc', 'blank');
  toast('Spreadsheet dihapus');
}

// Bagikan: live share links, see ../sharing.js
export { shareDialog } from '../sharing.js';

export function settingsDialog(app) {
  const wb = app.wb;
  const loc = select(Object.entries(LOCALES).map(([id, l]) => ({ value: id, label: `${l.name} (${format(1234.5, '#,##0.00').t.replace(/[\d]/g, '').length ? '' : ''}${id})` })), wb.settings.locale || 'id-ID');
  const ff = select(THEME_FONTS, app.defFont.ff);
  const fs = select([8, 9, 10, 11, 12, 14].map(String), String(app.defFont.fs));
  const theme = select(app.THEMES.map(t => ({ value: t.id, label: t.label })), app.settings.theme);
  dialog({
    title: 'Setelan spreadsheet', width: 480,
    body: el('div.stack',
      el('h4', 'Umum'),
      field('Lokal', loc, 'Memengaruhi format tanggal, angka, dan mata uang. Rumus selalu memakai koma sebagai pemisah argumen (titik koma juga diterima).'),
      el('div.grid2', field('Font default', ff), field('Ukuran', fs)),
      el('h4', 'Tampilan'),
      field('Tema', theme),
      checkbox('Efek kaca cair', app.settings.glass, v => app.setGlass(v)),
      el('h4', 'Perhitungan'),
      el('p.muted.small', `Perhitungan ulang otomatis saat ada perubahan. Fungsi volatil (NOW, TODAY, RAND) diperbarui setiap menit. Perhitungan terakhir: ${app.wb.eng.lastMs?.toFixed(1) || 0} ms.`)),
    buttons: [{ label: 'Batal' }, { label: 'Simpan', primary: true, run: () => {
      if (loc.value !== (wb.settings.locale || 'id-ID')) app.cmd('setLocale', loc.value);
      if (ff.value !== app.defFont.ff || +fs.value !== app.defFont.fs) { ensureFont(ff.value, app); app.cmd('defaultFont', ff.value, +fs.value); }
      if (theme.value !== app.settings.theme) app.applyTheme(theme.value);
    } }],
  });
}

// ---------- help ----------
export function functionListDialog(app) {
  const search = el('input.inp', { placeholder: 'Cari fungsi', 'aria-label': 'Cari fungsi' });
  const cat = select([{ value: '', label: 'Semua kategori' }, ...CATEGORIES.map(c => ({ value: c, label: c }))], '');
  const list = el('div.fn-list');
  let d;
  const all = allFunctions();
  const draw = () => {
    list.innerHTML = '';
    const q = search.value.trim().toUpperCase();
    const items = all.filter(f => (!cat.value || f.cat === cat.value) && (!q || f.name.includes(q) || f.desc.toUpperCase().includes(q)));
    for (const f of items.slice(0, 400)) list.append(el('button.fn-item', { type: 'button', onclick: () => { d.close(); app.cmd('insertFunction', f.name); } },
      el('div.fn-sig', el('b', f.name), el('span.muted', `(${f.args})`)), el('div.fn-desc', f.desc), el('span.fn-cat', f.cat)));
    count.textContent = `${items.length} fungsi`;
  };
  const count = el('span.muted.small');
  search.addEventListener('input', draw);
  cat.addEventListener('change', draw);
  draw();
  d = dialog({ title: 'Daftar fungsi', width: 640, cls: 'dlg-fn', body: el('div.stack', el('div.row', search, cat), count, list), buttons: [{ label: 'Tutup', primary: true }] });
}

export function shortcutsDialog(app) {
  const m = M();
  const groups = [
    ['Tindakan umum', [[`${m}+Z`, 'Urungkan'], [`${m}+Y`, 'Ulangi'], [`${m}+C / X / V`, 'Salin / potong / tempel'], [`${m}+Shift+V`, 'Tempel nilai saja'], [`${m}+F`, 'Cari'], [`${m}+H`, 'Cari dan ganti'], [`${m}+P`, 'Cetak'], [`${m}+S`, 'Simpan'], [`${m}+O`, 'Buka'], [`${m}+/`, 'Pintasan keyboard'], ['Alt+/', 'Telusuri menu']]],
    ['Format', [[`${m}+B / I / U`, 'Tebal / miring / garis bawah'], ['Alt+Shift+5', 'Coret'], [`${m}+Shift+E / L / R`, 'Rata tengah / kiri / kanan'], [`${m}+Shift+1`, 'Format angka'], [`${m}+Shift+3`, 'Format tanggal'], [`${m}+Shift+4`, 'Format mata uang'], [`${m}+Shift+5`, 'Format persen'], [`${m}+Shift+7`, 'Batas luar'], [`${m}+\\`, 'Hapus format'], [`${m}+1`, 'Format sel (Excel)']]],
    ['Navigasi & seleksi', [['Panah', 'Pindah sel'], [`${m}+Panah`, 'Ke tepi data'], ['Shift+Panah', 'Perluas seleksi'], [`${m}+A`, 'Pilih semua'], [`${m}+Spasi`, 'Pilih kolom'], ['Shift+Spasi', 'Pilih baris'], [`${m}+Home / End`, 'Ke awal / akhir data'], [`${m}+PageUp / PageDown`, 'Sheet sebelumnya / berikutnya'], ['Alt+PageUp / PageDown', 'Geser horizontal']]],
    ['Mengedit', [['Enter / F2', 'Edit sel'], ['Alt+Enter', 'Baris baru dalam sel'], [`${m}+Enter`, 'Isi seluruh seleksi'], [`${m}+Shift+Enter`, 'Rumus array'], ['F4', 'Ganti referensi absolut ($)'], ['Esc', 'Batalkan edit'], ['Delete', 'Hapus nilai'], [`${m}+D / R`, 'Isi ke bawah / kanan'], [`${m}+;`, 'Tanggal hari ini'], [`${m}+Shift+;`, 'Waktu sekarang'], ['Alt+=', 'Jumlah otomatis'], [`${m}+E`, 'Isi kilat']]],
    ['Baris, kolom, sheet', [[`${m}+Alt+=`, 'Sisipkan sel'], [`${m}+Alt+-`, 'Hapus sel'], [`${m}+9 / 0`, 'Sembunyikan baris / kolom'], [`${m}+Shift+9 / 0`, 'Tampilkan baris / kolom'], ['Shift+F11', 'Sheet baru'], [`${m}+Shift+F`, 'Buat / hapus filter'], ['Shift+F2', 'Catatan'], [`${m}+K`, 'Sisipkan link'], [`${m}+\``, 'Tampilkan rumus'], ['F9', 'Hitung ulang']]],
  ];
  dialog({
    title: 'Pintasan keyboard', width: 760, cls: 'dlg-keys',
    body: el('div.keys', ...groups.map(([t, rows]) => el('div.keys-grp', el('h4', t), el('table', ...rows.map(([k, d]) => el('tr', el('td', d), el('td', ...k.split(' / ').flatMap((x, i) => [i ? ' / ' : '', el('kbd', x)])))))))),
    buttons: [{ label: 'Tutup', primary: true }],
  });
}

export function aboutDialog(app) {
  dialog({
    title: 'Tentang Amatu Sheets', width: 480,
    body: el('div.about',
      el('div.about-logo', el('img', { src: 'icon.svg', alt: '', width: 56, height: 56 })),
      el('h3', 'Amatu Sheets'),
      el('p.muted', 'Versi 1.0 · Spreadsheet di browser yang memadukan alur kerja Google Sheets dengan fitur lanjutan Microsoft Excel.'),
      el('ul.about-list',
        el('li', `${Object.keys((DOCS && Object.fromEntries(DOCS)) || {}).length}+ fungsi termasuk XLOOKUP, LET, LAMBDA, FILTER, QUERY, dan array dinamis`),
        el('li', 'Format bersyarat, validasi data, filter, tabel pivot, diagram'),
        el('li', 'Impor/ekspor Excel, OpenDocument, CSV; histori versi lokal'),
        el('li', 'Data disimpan di perangkat Anda (IndexedDB), tanpa server')),
      el('p.muted.small', 'Ketuk logo kapan saja untuk membuka daftar spreadsheet.')),
    buttons: [{ label: 'Tutup', primary: true }],
  });
}
