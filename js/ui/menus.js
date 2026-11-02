// Menu bar and context menus.
import { el, icon, colName, rangeAddr, uid } from '../util.js';
import { openMenu, closeMenus, promptDlg, toast, colorGrid } from './ui.js';
import { NUMBER_FORMATS, L } from '../format.js';
import { DOCS, CATEGORIES } from '../docs.js';
import { TEMPLATES } from '../templates.js';
import * as D from './dialogs.js';
import * as io from '../io.js';
import { THEME_FONTS } from './toolbar.js';
import * as SH from '../sharing.js';
import * as PG from '../pages.js';
import * as LK from '../lock.js';

const M = () => window.amatu?.MOD || 'Ctrl';

function selInfo(app) {
  const g = app.grid.range;
  const rows = g.r2 - g.r1 + 1, cols = g.c2 - g.c1 + 1;
  return { g, rows, cols, rLabel: rows === 1 ? `${g.r1 + 1}` : `${g.r1 + 1}–${g.r2 + 1}`, cLabel: cols === 1 ? colName(g.c1) : `${colName(g.c1)}–${colName(g.c2)}` };
}

export function numberFormatItems(app) {
  return [
    ...NUMBER_FORMATS().map(f => (f.sep ? { sep: true } : { label: f.label, preview: f.sample, checked: () => (app.sheet.styleOf(app.grid.sel.r, app.grid.sel.c).fmt || 'General') === f.id, run: () => app.cmd('numberFormat', f.id) })),
    { sep: true },
    { label: 'Mata uang kustom…', icon: 'currency_exchange', run: () => D.customFormatDialog(app, 'currency') },
    { label: 'Tanggal dan waktu kustom…', icon: 'calendar_month', run: () => D.customFormatDialog(app, 'date') },
    { label: 'Format angka kustom…', icon: 'tune', run: () => D.customFormatDialog(app, 'number') },
  ];
}

export function functionItems(app) {
  const common = ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'];
  return [
    ...common.map(n => ({ label: n, run: () => app.cmd('insertFunction', n) })),
    { sep: true },
    ...CATEGORIES.map(cat => ({
      label: cat, sub: () => [...DOCS.values()].filter(d => d.cat === cat).sort((a, b) => a.name.localeCompare(b.name)).map(d => ({ label: d.name, run: () => app.cmd('insertFunction', d.name) })),
    })),
    { sep: true },
    { label: 'Daftar semua fungsi', icon: 'functions', run: () => D.functionListDialog(app) },
  ];
}

function freezeItems(app) {
  const sh = app.sheet, g = app.grid;
  return [
    { label: 'Tidak ada baris', checked: () => sh.freezeR === 0, run: () => app.cmd('freeze', 0, sh.freezeC) },
    { label: '1 baris', checked: () => sh.freezeR === 1, run: () => app.cmd('freeze', 1, sh.freezeC) },
    { label: '2 baris', checked: () => sh.freezeR === 2, run: () => app.cmd('freeze', 2, sh.freezeC) },
    { label: `Hingga baris ${g.sel.r + 1}`, checked: () => sh.freezeR === g.sel.r + 1 && g.sel.r > 1, run: () => app.cmd('freeze', g.sel.r + 1, sh.freezeC) },
    { sep: true },
    { label: 'Tidak ada kolom', checked: () => sh.freezeC === 0, run: () => app.cmd('freeze', sh.freezeR, 0) },
    { label: '1 kolom', checked: () => sh.freezeC === 1, run: () => app.cmd('freeze', sh.freezeR, 1) },
    { label: '2 kolom', checked: () => sh.freezeC === 2, run: () => app.cmd('freeze', sh.freezeR, 2) },
    { label: `Hingga kolom ${colName(g.sel.c)}`, checked: () => sh.freezeC === g.sel.c + 1 && g.sel.c > 1, run: () => app.cmd('freeze', sh.freezeR, g.sel.c + 1) },
    { sep: true },
    { label: 'Bekukan panel di sel aktif', hint: 'Excel', run: () => app.cmd('freeze', g.sel.r, g.sel.c) },
  ];
}

export function pasteSpecialItems(app) {
  return [
    { label: 'Hanya nilai', kbd: `${M()}+Shift+V`, run: () => app.pasteFromMenu('values') },
    { label: 'Hanya format', run: () => app.pasteFromMenu('formats') },
    { label: 'Hanya rumus', run: () => app.pasteFromMenu('formulas') },
    { label: 'Tanpa format', run: () => app.pasteFromMenu('nofmt') },
    { label: 'Hanya validasi data', run: () => app.pasteFromMenu('valid') },
    { label: 'Hanya lebar kolom', run: () => app.pasteFromMenu('widths') },
    { label: 'Ditransposisikan', run: () => app.pasteFromMenu('transpose') },
    { sep: true },
    { label: 'Tempel spesial…', icon: 'content_paste_go', hint: 'Excel', run: () => D.pasteSpecialDialog(app) },
  ];
}

const sysCopy = (app, cut) => {
  app.grid.focus();
  const ok = document.execCommand(cut ? 'cut' : 'copy');
  if (!ok) app.copyTo(null, cut);
};

const ROT = [
  { label: 'Tidak ada', icon: 'text_rotation_none', v: 0 },
  { label: 'Miring ke atas', icon: 'text_rotate_up', v: 45 },
  { label: 'Miring ke bawah', icon: 'text_rotation_down', v: -45 },
  { label: 'Tumpuk secara vertikal', icon: 'text_rotate_vertical', v: 'v' },
  { label: 'Putar ke atas', icon: 'text_rotation_angleup', v: 90 },
  { label: 'Putar ke bawah', icon: 'text_rotation_angledown', v: -90 },
];

export function menus(app) {
  const st = () => app.sheet.styleOf(app.grid.sel.r, app.grid.sel.c);
  const hasFilter = () => !!app.sheet.filter;
  return {
    File: [
      { label: 'Baru', icon: 'note_add', sub: [
        { label: 'Spreadsheet kosong', icon: 'grid_on', run: () => app.cmd('newDoc', 'blank') },
        { sep: true },
        ...TEMPLATES.filter(t => t.id !== 'blank' && t.id !== 'welcome').map(t => ({ label: t.name, icon: t.icon, run: () => app.cmd('newDoc', t.id) })),
        { sep: true },
        { label: 'Galeri template', icon: 'dashboard', run: () => D.openDialog(app, 'templates') },
      ] },
      { label: 'Buka', icon: 'folder_open', kbd: `${M()}+O`, run: () => D.openDialog(app) },
      { label: 'Impor', icon: 'upload_file', run: () => D.importDialog(app) },
      { label: 'Buat salinan', icon: 'file_copy', run: () => app.cmd('makeCopy') },
      { sep: true },
      { label: 'Bagikan', icon: 'person_add', run: () => SH.shareDialog(app) },
      { label: 'Download', icon: 'download', sub: [
        { label: 'Microsoft Excel (.xlsx)', run: () => io.exportXLSX(app) },
        { label: 'OpenDocument (.ods)', run: () => io.exportXLSX(app, 'ods') },
        { label: 'Dokumen PDF (.pdf)', hint: 'cetak', run: () => io.printView(app) },
        { label: 'Halaman web (.html)', run: () => io.exportHTML(app) },
        { label: 'Nilai yang dipisahkan koma (.csv)', run: () => io.exportCSV(app, ',') },
        { label: 'Nilai yang dipisahkan tab (.tsv)', run: () => io.exportCSV(app, '\t') },
        { label: 'Amatu Sheets (.json)', run: () => io.exportJSON(app) },
      ] },
      { sep: true },
      { label: 'Ganti nama', icon: 'drive_file_rename_outline', run: () => { const t = document.getElementById('doc-title'); t.focus(); t.select(); } },
      { label: 'Histori versi', icon: 'history', sub: [
        { label: 'Beri nama versi saat ini', run: async () => { const n = await promptDlg('Beri nama versi', 'Nama versi', ''); if (n) { await (await import('../store.js')).saveVersion(app.wb, n); toast('Versi disimpan'); } } },
        { label: 'Lihat histori versi', run: () => app.sidebar.open('versions') },
      ] },
      { label: 'Hapus', icon: 'delete', run: () => D.deleteDocDialog(app) },
      { sep: true },
      { label: 'Keamanan file', icon: 'lock', sub: () => LK.menuItems(app) },
      { label: 'Setelan', icon: 'settings', run: () => D.settingsDialog(app) },
      { sep: true },
      { label: 'Penyiapan halaman (kertas & margin)…', icon: 'description', run: () => PG.setupDialog(app) },
      { label: 'Pratinjau cetak', icon: 'preview', run: () => PG.printPages(app, { sheets: [app.sheet] }) },
      { label: 'Cetak', icon: 'print', kbd: `${M()}+P`, run: () => io.printView(app) },
    ],
    Edit: [
      { label: 'Urungkan', icon: 'undo', kbd: `${M()}+Z`, disabled: () => !app.wb.hist.undoS.length, run: () => app.undo() },
      { label: 'Ulangi', icon: 'redo', kbd: `${M()}+Y`, disabled: () => !app.wb.hist.redoS.length, run: () => app.redo() },
      { sep: true },
      { label: 'Potong', icon: 'content_cut', kbd: `${M()}+X`, run: () => sysCopy(app, true) },
      { label: 'Salin', icon: 'content_copy', kbd: `${M()}+C`, run: () => sysCopy(app, false) },
      { label: 'Tempel', icon: 'content_paste', kbd: `${M()}+V`, run: () => app.pasteFromMenu('all') },
      { label: 'Tempel spesial', icon: 'content_paste_go', sub: () => pasteSpecialItems(app) },
      { sep: true },
      { label: 'Hapus', icon: 'backspace', sub: [
        { label: 'Nilai', kbd: 'Delete', run: () => app.cmd('clear', 'contents') },
        { label: 'Format', kbd: `${M()}+\\`, run: () => app.cmd('clear', 'formats') },
        { label: 'Catatan', run: () => app.cmd('clear', 'notes') },
        { label: 'Semua', run: () => app.cmd('clear', 'all') },
      ] },
      { label: () => '', hidden: true },
      { label: `Hapus baris`, icon: 'table_rows', run: () => app.cmd('deleteRows') },
      { label: `Hapus kolom`, icon: 'view_column', run: () => app.cmd('deleteCols') },
      { sep: true },
      { label: 'Cari dan ganti', icon: 'find_replace', kbd: `${M()}+H`, run: () => D.replaceDialog(app, true) },
    ],
    Lihat: [
      { label: 'Tampilkan', icon: 'visibility', sub: [
        { label: 'Garis kisi', checked: () => app.sheet.grid, run: () => app.cmd('toggleGridlines') },
        { label: 'Rumus', kbd: `${M()}+\``, checked: () => app.showFormulas, run: () => app.cmd('toggleFormulas') },
        { label: 'Sel yang bisa diedit pelihat', checked: () => !!app.settings.showVE, run: () => app.cmd('toggleShowVE') },
        { label: 'Pemisah halaman', checked: () => !!app.settings.breaks, run: () => PG.showBreaks(app, !app.settings.breaks) },
        { label: 'Saran saat tabel terpotong halaman', checked: () => app.settings.breakTips !== false, run: () => { app.settings.breakTips = app.settings.breakTips === false; app.saveSettings(); app.onSelChange(); } },
        { label: 'Bilah formula', checked: () => !document.body.classList.contains('no-fx'), run: () => { document.body.classList.toggle('no-fx'); app.grid.resize(); } },
        { label: 'Bilah alat', checked: () => !document.body.classList.contains('no-toolbar'), run: () => { document.body.classList.toggle('no-toolbar'); app.grid.resize(); } },
      ] },
      { label: 'Bekukan', icon: 'ac_unit', sub: () => freezeItems(app) },
      { sep: true },
      { label: 'Zoom', icon: 'zoom_in', sub: [50, 75, 90, 100, 125, 150, 200].map(z => ({ label: z + '%', checked: () => Math.round(app.grid.z * 100) === z, run: () => app.setZoom(z / 100) })) },
      { label: 'Tema', icon: 'palette', sub: () => app.themeMenu() },
      { label: 'Efek kaca cair', icon: 'blur_on', checked: () => app.settings.glass, run: () => app.setGlass(!app.settings.glass) },
      { sep: true },
      { label: 'Layar penuh', icon: 'fullscreen', run: () => (document.fullscreenElement ? document.exitFullscreen() : document.documentElement.requestFullscreen?.()) },
    ],
    Sisipkan: [
      { label: 'Sel', icon: 'grid_view', sub: [
        { label: 'Sisipkan sel dan geser ke kanan', run: () => app.cmd('shiftCells', 'right') },
        { label: 'Sisipkan sel dan geser ke bawah', run: () => app.cmd('shiftCells', 'down') },
      ] },
      { label: 'Baris', icon: 'table_rows', sub: () => { const s = selInfo(app); return [
        { label: `Sisipkan ${s.rows} baris di atas`, run: () => app.cmd('insertRows', 'above') },
        { label: `Sisipkan ${s.rows} baris di bawah`, run: () => app.cmd('insertRows', 'below') },
      ]; } },
      { label: 'Kolom', icon: 'view_column', sub: () => { const s = selInfo(app); return [
        { label: `Sisipkan ${s.cols} kolom ke kiri`, run: () => app.cmd('insertCols', 'left') },
        { label: `Sisipkan ${s.cols} kolom ke kanan`, run: () => app.cmd('insertCols', 'right') },
      ]; } },
      { label: 'Sheet', icon: 'add_box', kbd: 'Shift+F11', run: () => app.cmd('addSheet') },
      { sep: true },
      { label: 'Diagram', icon: 'insert_chart', run: () => app.sidebar.open('chart', { create: true }) },
      { label: 'Tabel pivot', icon: 'pivot_table_chart', run: () => D.pivotDialog(app) },
      { label: 'Gambar dalam sel', icon: 'image', run: async () => { const u = await promptDlg('Sisipkan gambar dalam sel', 'URL gambar', 'https://'); if (u && u !== 'https://') { const g = app.grid; app.tx('Gambar', () => app.sheet.put(g.sel.r, g.sel.c, { f: `IMAGE("${u.replace(/"/g, '""')}")`, v: null })); } } },
      { label: 'Fungsi', icon: 'functions', sub: () => functionItems(app) },
      { sep: true },
      { label: 'Link', icon: 'link', kbd: `${M()}+K`, run: () => D.linkDialog(app) },
      { label: 'Kotak centang', icon: 'check_box', run: () => app.cmd('insertCheckbox') },
      { label: 'Dropdown', icon: 'arrow_drop_down_circle', run: () => app.cmd('insertDropdown') },
      { label: 'Catatan', icon: 'sticky_note_2', kbd: 'Shift+F2', run: () => D.noteDialog(app) },
      { sep: true },
      { label: 'Tanggal hari ini', icon: 'today', kbd: `${M()}+;`, run: () => app.cmd('insertNow', 'date') },
      { label: 'Waktu sekarang', icon: 'schedule', kbd: `${M()}+Shift+;`, run: () => app.cmd('insertNow', 'time') },
    ],
    Format: [
      { label: 'Angka', icon: 'tag', sub: () => numberFormatItems(app) },
      { label: 'Teks', icon: 'format_bold', sub: [
        { label: 'Tebal', icon: 'format_bold', kbd: `${M()}+B`, checked: () => st().b, run: () => app.cmd('toggleStyle', 'b') },
        { label: 'Miring', icon: 'format_italic', kbd: `${M()}+I`, checked: () => st().i, run: () => app.cmd('toggleStyle', 'i') },
        { label: 'Garis bawah', icon: 'format_underlined', kbd: `${M()}+U`, checked: () => st().u, run: () => app.cmd('toggleStyle', 'u') },
        { label: 'Coret', icon: 'strikethrough_s', kbd: 'Alt+Shift+5', checked: () => st().s, run: () => app.cmd('toggleStyle', 's') },
      ] },
      { label: 'Perataan', icon: 'format_align_left', sub: [
        { label: 'Kiri', icon: 'format_align_left', kbd: `${M()}+Shift+L`, run: () => app.cmd('align', 'left') },
        { label: 'Tengah', icon: 'format_align_center', kbd: `${M()}+Shift+E`, run: () => app.cmd('align', 'center') },
        { label: 'Kanan', icon: 'format_align_right', kbd: `${M()}+Shift+R`, run: () => app.cmd('align', 'right') },
        { sep: true },
        { label: 'Atas', icon: 'vertical_align_top', run: () => app.cmd('valign', 'top') },
        { label: 'Tengah', icon: 'vertical_align_center', run: () => app.cmd('valign', 'middle') },
        { label: 'Bawah', icon: 'vertical_align_bottom', run: () => app.cmd('valign', 'bottom') },
      ] },
      { label: 'Pengemasan', icon: 'wrap_text', sub: [
        { label: 'Luapan', icon: 'format_overline', checked: () => !st().wrap, run: () => app.cmd('wrap', 'overflow') },
        { label: 'Kemas', icon: 'wrap_text', checked: () => st().wrap === 'wrap', run: () => app.cmd('wrap', 'wrap') },
        { label: 'Potong', icon: 'content_cut', checked: () => st().wrap === 'clip', run: () => app.cmd('wrap', 'clip') },
      ] },
      { label: 'Rotasi', icon: 'text_rotate_up', sub: ROT.map(r => ({ label: r.label, icon: r.icon, checked: () => (st().rot || 0) === r.v, run: () => app.cmd('rotate', r.v) })) },
      { label: 'Ukuran font', icon: 'format_size', sub: [6, 7, 8, 9, 10, 11, 12, 14, 18, 24, 36].map(s => ({ label: String(s), checked: () => (st().fs || app.defFont.fs) === s, run: () => app.cmd('fontSize', s) })) },
      { label: 'Gabungkan sel', icon: 'cell_merge', sub: [
        { label: 'Gabungkan semua', run: () => app.cmd('merge', 'all') },
        { label: 'Gabungkan secara horizontal', run: () => app.cmd('merge', 'h') },
        { label: 'Gabungkan secara vertikal', run: () => app.cmd('merge', 'v') },
        { label: 'Pisahkan', run: () => app.cmd('merge', 'none') },
      ] },
      { sep: true },
      { label: 'Format sel…', icon: 'tune', kbd: `${M()}+1`, run: () => D.formatCellsDialog(app) },
      { label: 'Format bersyarat', icon: 'format_color_fill', run: () => app.sidebar.open('cf') },
      { label: 'Warna bergantian', icon: 'format_list_bulleted', run: () => app.sidebar.open('banding') },
      { label: 'Font default', icon: 'font_download', sub: () => THEME_FONTS.map(f => ({ label: f, checked: () => app.defFont.ff === f, run: () => app.cmd('defaultFont', f, app.defFont.fs) })) },
      { sep: true },
      { label: 'Ubah huruf', icon: 'match_case', sub: [
        { label: 'HURUF BESAR', run: () => app.cmd('changeCase', 'upper') },
        { label: 'huruf kecil', run: () => app.cmd('changeCase', 'lower') },
        { label: 'Huruf Awal Kapital', run: () => app.cmd('changeCase', 'proper') },
      ] },
      { label: 'Hapus format', icon: 'format_clear', kbd: `${M()}+\\`, run: () => app.cmd('clear', 'formats') },
    ],
    'Tata letak': () => PG.layoutMenu(app),
    Data: [
      { label: 'Urutkan sheet', icon: 'sort', sub: () => [
        { label: `Urutkan sheet menurut kolom ${colName(app.grid.sel.c)} (A–Z)`, run: () => app.cmd('sortSheet', true) },
        { label: `Urutkan sheet menurut kolom ${colName(app.grid.sel.c)} (Z–A)`, run: () => app.cmd('sortSheet', false) },
      ] },
      { label: 'Urutkan rentang', icon: 'sort_by_alpha', sub: () => [
        { label: `Urutkan rentang menurut kolom ${colName(app.grid.sel.c)} (A–Z)`, run: () => app.cmd('sortRange', app.grid.range, [{ c: app.grid.sel.c, asc: true }], false) },
        { label: `Urutkan rentang menurut kolom ${colName(app.grid.sel.c)} (Z–A)`, run: () => app.cmd('sortRange', app.grid.range, [{ c: app.grid.sel.c, asc: false }], false) },
        { label: 'Opsi pengurutan lanjutan', icon: 'tune', run: () => D.sortDialog(app) },
      ] },
      { sep: true },
      { label: () => '', hidden: true },
      { label: 'Buat filter', icon: 'filter_alt', kbd: `${M()}+Shift+F`, hidden: hasFilter, run: () => app.cmd('toggleFilter') },
      { label: 'Hapus filter', icon: 'filter_alt_off', kbd: `${M()}+Shift+F`, hidden: () => !hasFilter(), run: () => app.cmd('toggleFilter') },
      { sep: true },
      { label: 'Validasi data', icon: 'rule', run: () => app.sidebar.open('dv') },
      { label: 'Statistik kolom', icon: 'query_stats', run: () => D.columnStatsDialog(app) },
      { label: 'Rentang bernama', icon: 'label', run: () => app.sidebar.open('names') },
      { label: 'Proteksi & izin', icon: 'lock', sub: () => protectItems(app) },
      { sep: true },
      { label: 'Pembersihan data', icon: 'cleaning_services', sub: [
        { label: 'Hapus duplikat…', run: () => D.removeDuplicatesDialog(app) },
        { label: 'Pangkas spasi kosong', run: () => app.cmd('trim') },
      ] },
      { label: 'Pisahkan teks menjadi kolom', icon: 'call_split', sub: [
        { label: 'Deteksi otomatis', run: () => app.cmd('splitText', 'auto') },
        { label: 'Koma', run: () => app.cmd('splitText', ',') },
        { label: 'Titik koma', run: () => app.cmd('splitText', ';') },
        { label: 'Titik', run: () => app.cmd('splitText', '.') },
        { label: 'Spasi', run: () => app.cmd('splitText', ' ') },
        { label: 'Kustom…', run: async () => { const d = await promptDlg('Pemisah kustom', 'Karakter pemisah', '|'); if (d) app.cmd('splitText', d); } },
      ] },
      { label: 'Isi kilat', icon: 'bolt', kbd: `${M()}+E`, hint: 'Excel', run: () => app.cmd('flashFill') },
      { label: 'Analisis bagaimana-jika', icon: 'psychology_alt', sub: [
        { label: 'Pencarian Tujuan (Goal Seek)…', run: () => D.goalSeekDialog(app) },
      ] },
      { label: 'Tabel pivot', icon: 'pivot_table_chart', run: () => D.pivotDialog(app) },
      { label: 'Perbarui tabel pivot', icon: 'refresh', hidden: () => !app.sheet.pivot, run: () => app.cmd('pivotRefresh') },
      { label: 'Acak rentang', icon: 'shuffle', run: () => app.cmd('randomize') },
      { label: 'Konversi rumus ke nilai', icon: 'pin', run: () => app.cmd('toValues') },
    ],
    Alat: [
      { label: 'Isi otomatis teks', icon: 'auto_awesome', checked: () => app.settings.autocomplete !== false, run: () => { app.settings.autocomplete = app.settings.autocomplete === false; app.saveSettings(); } },
      { label: 'Hitung ulang', icon: 'calculate', kbd: 'F9', run: () => { app.after({ save: false }); toast(`Dihitung ulang dalam ${app.wb.eng.lastMs.toFixed(1)} ms`); } },
      { label: 'Pemeriksa rumus error', icon: 'bug_report', run: () => D.errorCheckDialog(app) },
      { sep: true },
      { label: 'Setelan spreadsheet', icon: 'settings', run: () => D.settingsDialog(app) },
    ],
    Bantuan: [
      { label: 'Daftar fungsi', icon: 'functions', run: () => D.functionListDialog(app) },
      { label: 'Pintasan keyboard', icon: 'keyboard', kbd: `${M()}+/`, run: () => D.shortcutsDialog(app) },
      { sep: true },
      { label: 'Tentang Amatu Sheets', icon: 'info', run: () => D.aboutDialog(app) },
    ],
  };
}

export function buildMenubar(app, host) {
  const defs = menus(app);
  let active = null;
  for (const name of Object.keys(defs)) {
    const b = el('button.mb-item', { type: 'button' }, name);
    const open = () => {
      document.querySelectorAll('.mb-item.open').forEach(x => x.classList.remove('open'));
      b.classList.add('open');
      active = b;
      const m = openMenu(menus(app)[name], b.getBoundingClientRect());
      const obs = new MutationObserver(() => { if (!m.isConnected) { b.classList.remove('open'); if (active === b) active = null; obs.disconnect(); } });
      obs.observe(document.getElementById('popups'), { childList: true });
    };
    b.addEventListener('click', () => { if (b.classList.contains('open')) { closeMenus(); b.classList.remove('open'); active = null; } else open(); });
    b.addEventListener('mouseenter', () => { if (active && active !== b) open(); });
    host.append(b);
  }
  // menu search (like "Help > Search the menus")
  const search = el('input.mb-search', { placeholder: 'Telusuri menu (Alt+/)', 'aria-label': 'Telusuri menu' });
  search.addEventListener('input', () => {
    const q = search.value.trim().toLowerCase();
    if (!q) { closeMenus(); return; }
    const found = [];
    const walk = (items, path) => {
      for (const it of (typeof items === 'function' ? items() : items)) {
        if (!it || it.sep || it.head || it.hidden === true || typeof it.label !== 'string') continue;
        if (it.sub) walk(it.sub, [...path, it.label]);
        else if (it.label.toLowerCase().includes(q)) found.push({ ...it, hint: path.join(' › ') });
      }
    };
    for (const [n, items] of Object.entries(menus(app))) walk(items, [n]);
    openMenu(found.slice(0, 14).map(f => ({ ...f, run: () => { search.value = ''; f.run(); } })), search.getBoundingClientRect());
  });
  search.addEventListener('keydown', e => { if (e.key === 'Escape') { search.value = ''; closeMenus(); app.grid.focus(); } });
  document.addEventListener('keydown', e => { if (e.altKey && (e.key === '/' || e.code === 'Slash')) { e.preventDefault(); search.focus(); } });
  host.append(el('span.flex1'), search);
}

// Data › Proteksi & izin: password lock of the sheet (with cells left open) and viewer-editable cells of share links.
export function protectItems(app) {
  const sh = app.sheet;
  return [
    { head: `Sheet "${sh.name}"` },
    { label: 'Kunci sheet dengan kata sandi…', icon: 'lock', hidden: () => sh.protect, run: () => SH.lockDialog(app, sh) },
    { label: 'Buka kunci sheet…', icon: 'lock_open', hidden: () => !sh.protect, run: () => SH.unlockDialog(app, sh) },
    { label: 'Sel yang tetap terbuka saat dikunci…', icon: 'lock_open_right', disabled: () => !sh.protect, run: () => SH.lockedCellsDialog(app, sh) },
    { sep: true },
    { head: 'Tautan "Lihat saja"' },
    { label: 'Sel yang bisa diedit pelihat…', icon: 'edit_note', run: () => SH.viewerCellsDialog(app, sh) },
    { label: 'Izinkan pelihat mengedit pilihan ini', icon: 'add_task', disabled: () => app.readOnly, run: () => {
      const add = app.grid.sel.ranges.map(g => ({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }));
      app.tx('Izin pelihat', () => { app.wb.hist.touchMeta(sh); sh.ve = [...(sh.ve || []), ...add.filter(g => !(sh.ve || []).some(x => x.r1 === g.r1 && x.c1 === g.c1 && x.r2 === g.r2 && x.c2 === g.c2))]; });
      if (!app.settings.showVE) app.cmd('toggleShowVE');
      toast(`Pelihat dapat mengedit ${add.map(g => rangeAddr(g)).join(', ')}`);
    } },
    { label: 'Tandai sel yang bisa diedit pelihat', icon: 'border_style', checked: () => !!app.settings.showVE, run: () => app.cmd('toggleShowVE') },
    { sep: true },
    { label: 'Bagikan…', icon: 'person_add', run: () => SH.shareDialog(app) },
  ];
}

export function cellContextMenu(app) {
  const s = selInfo(app);
  return [
    { label: 'Potong', icon: 'content_cut', kbd: `${M()}+X`, run: () => sysCopy(app, true) },
    { label: 'Salin', icon: 'content_copy', kbd: `${M()}+C`, run: () => sysCopy(app, false) },
    { label: 'Tempel', icon: 'content_paste', kbd: `${M()}+V`, run: () => app.pasteFromMenu('all') },
    { label: 'Tempel spesial', icon: 'content_paste_go', sub: () => pasteSpecialItems(app) },
    { sep: true },
    { label: `Sisipkan ${s.rows} baris di atas`, icon: 'add', run: () => app.cmd('insertRows', 'above') },
    { label: `Sisipkan ${s.cols} kolom di kiri`, icon: 'add', run: () => app.cmd('insertCols', 'left') },
    { label: 'Sisipkan sel', icon: 'add', sub: [
      { label: 'Geser ke kanan', run: () => app.cmd('shiftCells', 'right') },
      { label: 'Geser ke bawah', run: () => app.cmd('shiftCells', 'down') },
    ] },
    { sep: true },
    { label: `Hapus baris ${s.rLabel}`, icon: 'delete', run: () => app.cmd('deleteRows') },
    { label: `Hapus kolom ${s.cLabel}`, icon: 'delete', run: () => app.cmd('deleteCols') },
    { label: 'Hapus sel', icon: 'delete', sub: [
      { label: 'Geser ke kiri', run: () => app.cmd('shiftCells', 'left') },
      { label: 'Geser ke atas', run: () => app.cmd('shiftCells', 'up') },
    ] },
    { sep: true },
    { label: 'Buat filter', icon: 'filter_alt', hidden: () => !!app.sheet.filter, run: () => app.cmd('toggleFilter') },
    { label: 'Hapus filter', icon: 'filter_alt_off', hidden: () => !app.sheet.filter, run: () => app.cmd('toggleFilter') },
    { label: 'Format bersyarat', icon: 'format_color_fill', run: () => app.sidebar.open('cf') },
    { label: 'Validasi data', icon: 'rule', run: () => app.sidebar.open('dv') },
    { sep: true },
    { label: 'Sisipkan link', icon: 'link', kbd: `${M()}+K`, run: () => D.linkDialog(app) },
    { label: 'Sisipkan catatan', icon: 'sticky_note_2', kbd: 'Shift+F2', run: () => D.noteDialog(app) },
    { label: 'Dropdown', icon: 'arrow_drop_down_circle', run: () => app.cmd('insertDropdown') },
    { label: 'Diagram', icon: 'insert_chart', run: () => app.sidebar.open('chart', { create: true }) },
    { sep: true },
    { label: 'Halaman cetak', icon: 'insert_page_break', sub: () => PG.layoutMenu(app).filter(x => x.sep || ['move_down', 'content_cut', 'horizontal_rule', 'title', 'insert_page_break', 'tune'].includes(x.icon)) },
    { label: 'Lihat tindakan sel lainnya', icon: 'more_horiz', sub: [
      { label: 'Format sel…', icon: 'tune', kbd: `${M()}+1`, run: () => D.formatCellsDialog(app) },
      { label: 'Tentukan rentang bernama', icon: 'label', run: () => app.sidebar.open('names', { define: true }) },
      { label: 'Konversi rumus ke nilai', icon: 'pin', run: () => app.cmd('toValues') },
      { label: 'Acak rentang', icon: 'shuffle', run: () => app.cmd('randomize') },
      { label: 'Urutkan rentang A–Z', icon: 'sort_by_alpha', run: () => app.cmd('sortRange', app.grid.range, [{ c: app.grid.sel.c, asc: true }], false) },
      { label: 'Urutkan rentang Z–A', icon: 'sort_by_alpha', run: () => app.cmd('sortRange', app.grid.range, [{ c: app.grid.sel.c, asc: false }], false) },
      { label: 'Hapus format', icon: 'format_clear', run: () => app.cmd('clear', 'formats') },
    ] },
  ];
}

export function colContextMenu(app) {
  const s = selInfo(app), sh = app.sheet;
  const hiddenNear = [...sh.hiddenC].some(c => c >= s.g.c1 - 1 && c <= s.g.c2 + 1);
  return [
    { label: 'Potong', icon: 'content_cut', run: () => sysCopy(app, true) },
    { label: 'Salin', icon: 'content_copy', run: () => sysCopy(app, false) },
    { label: 'Tempel', icon: 'content_paste', run: () => app.pasteFromMenu('all') },
    { label: 'Tempel spesial', icon: 'content_paste_go', sub: () => pasteSpecialItems(app) },
    { sep: true },
    { label: `Sisipkan ${s.cols} kolom di kiri`, icon: 'add', run: () => app.cmd('insertCols', 'left') },
    { label: `Sisipkan ${s.cols} kolom di kanan`, icon: 'add', run: () => app.cmd('insertCols', 'right') },
    { label: `Hapus kolom ${s.cLabel}`, icon: 'delete', run: () => app.cmd('deleteCols') },
    { label: `Kosongkan kolom ${s.cLabel}`, icon: 'backspace', run: () => app.cmd('clear', 'contents') },
    { label: `Sembunyikan kolom ${s.cLabel}`, icon: 'visibility_off', kbd: `${M()}+0`, run: () => app.cmd('hideCols') },
    { label: 'Tampilkan kolom', icon: 'visibility', hidden: () => !hiddenNear && !sh.hiddenC.size, run: () => app.cmd('unhideCols') },
    { label: 'Ubah ukuran kolom', icon: 'width', run: () => D.resizeDialog(app, 'col') },
    { label: 'Sesuaikan dengan data', icon: 'fit_width', run: () => app.cmd('autofitCols', app.grid.colsFor(s.g.c1)) },
    { sep: true },
    { label: 'Urutkan sheet A–Z', icon: 'sort_by_alpha', run: () => app.cmd('sortSheet', true) },
    { label: 'Urutkan sheet Z–A', icon: 'sort_by_alpha', run: () => app.cmd('sortSheet', false) },
    { sep: true },
    { label: 'Buat filter', icon: 'filter_alt', hidden: () => !!sh.filter, run: () => app.cmd('toggleFilter') },
    { label: 'Format bersyarat', icon: 'format_color_fill', run: () => app.sidebar.open('cf') },
    { label: 'Validasi data', icon: 'rule', run: () => app.sidebar.open('dv') },
    { label: 'Statistik kolom', icon: 'query_stats', run: () => D.columnStatsDialog(app) },
  ];
}

export function rowContextMenu(app) {
  const s = selInfo(app), sh = app.sheet;
  const hiddenNear = [...sh.hiddenR].some(r => r >= s.g.r1 - 1 && r <= s.g.r2 + 1);
  return [
    { label: 'Potong', icon: 'content_cut', run: () => sysCopy(app, true) },
    { label: 'Salin', icon: 'content_copy', run: () => sysCopy(app, false) },
    { label: 'Tempel', icon: 'content_paste', run: () => app.pasteFromMenu('all') },
    { label: 'Tempel spesial', icon: 'content_paste_go', sub: () => pasteSpecialItems(app) },
    { sep: true },
    { label: `Sisipkan ${s.rows} baris di atas`, icon: 'add', run: () => app.cmd('insertRows', 'above') },
    { label: `Sisipkan ${s.rows} baris di bawah`, icon: 'add', run: () => app.cmd('insertRows', 'below') },
    { label: `Hapus baris ${s.rLabel}`, icon: 'delete', run: () => app.cmd('deleteRows') },
    { label: `Kosongkan baris ${s.rLabel}`, icon: 'backspace', run: () => app.cmd('clear', 'contents') },
    { label: `Sembunyikan baris ${s.rLabel}`, icon: 'visibility_off', kbd: `${M()}+9`, run: () => app.cmd('hideRows') },
    { label: 'Tampilkan baris', icon: 'visibility', hidden: () => !hiddenNear && !sh.hiddenR.size, run: () => app.cmd('unhideRows') },
    { label: 'Ubah ukuran baris', icon: 'height', run: () => D.resizeDialog(app, 'row') },
    { label: 'Sesuaikan dengan data', icon: 'fit_screen', run: () => app.cmd('autofitRows', app.grid.rowsFor(s.g.r1)) },
    { sep: true },
    { label: 'Format bersyarat', icon: 'format_color_fill', run: () => app.sidebar.open('cf') },
    { label: 'Validasi data', icon: 'rule', run: () => app.sidebar.open('dv') },
  ];
}

export function tabContextMenu(app, sh) {
  const i = app.wb.sheets.indexOf(sh);
  return [
    { label: 'Hapus', icon: 'delete', run: () => app.cmd('deleteSheet', sh) },
    { label: 'Duplikat', icon: 'content_copy', run: () => app.cmd('duplicateSheet') },
    { label: 'Ganti nama', icon: 'drive_file_rename_outline', run: () => { const tab = document.querySelectorAll('#tabs .tab')[app.wb.sheets.filter((s, j) => !s.hidden && j <= i).length - 1]; if (tab) app.renameSheetInline(tab, sh); } },
    { label: 'Ubah warna', icon: 'palette', sub: [{ node: close => colorGrid(c => { close(); app.cmd('tabColor', sh, c); }, { resetLabel: 'Setel ulang' }) }] },
    { label: sh.protect ? 'Buka kunci sheet…' : 'Kunci sheet dengan kata sandi…', icon: sh.protect ? 'lock_open' : 'lock', run: () => app.cmd('protectSheet', sh) },
    { label: 'Sel yang bisa diedit pelihat…', icon: 'edit_note', run: () => SH.viewerCellsDialog(app, sh) },
    { label: 'Sembunyikan sheet', icon: 'visibility_off', run: () => app.cmd('hideSheet', sh) },
    { sep: true },
    { label: 'Pindahkan ke kanan', icon: 'arrow_forward', disabled: () => i >= app.wb.sheets.length - 1, run: () => app.cmd('moveSheet', i, i + 1) },
    { label: 'Pindahkan ke kiri', icon: 'arrow_back', disabled: () => i <= 0, run: () => app.cmd('moveSheet', i, i - 1) },
  ];
}
