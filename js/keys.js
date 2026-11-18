// Keyboard shortcuts of Amatu Sheets in one list. The key handler (app.onKey), the key hints in the
// menus and tooltips, and the "Pintasan keyboard" dialog all read it, so what is shown is what works.
//
// Keys are written like 'Mod+Shift+V':
//   Mod   = ⌘ Command on Mac, Ctrl elsewhere      Alt   = ⌥ Option on Mac
//   Ctrl  = the Control key itself (⌃ on Mac)     Shift = ⇧
// Several entries in keys[] are alternatives; 'Mod+Alt+V N' (space) = press one, then the other.
// mac / win replace keys[] on that platform. `run` lives in app.js (SHORTCUT_RUN), keyed by id.
export const isMac = /Mac|iPhone|iPad/.test(navigator.platform);

const MAC_MOD = { Ctrl: '⌃', Alt: '⌥', Shift: '⇧', Mod: '⌘' };
const NAMES_MAC = { Up: '↑', Down: '↓', Left: '←', Right: '→', Space: 'Spasi', Esc: 'Esc', Enter: 'Enter', Tab: 'Tab', Backspace: '⌫', Delete: '⌦', PageUp: 'PgUp', PageDown: 'PgDn', Home: 'Home', End: 'End', '-': '−' };
const NAMES_WIN = { Up: '↑', Down: '↓', Left: '←', Right: '→', Space: 'Spasi', Esc: 'Esc', Enter: 'Enter', Tab: 'Tab', Backspace: 'Backspace', Delete: 'Delete', PageUp: 'PgUp', PageDown: 'PgDn', Home: 'Home', End: 'End' };
// key of a spec → KeyboardEvent.code (layout-free, and Option on a Mac changes e.key into a symbol)
const CODES = { '=': ['Equal', 'NumpadAdd'], '-': ['Minus', 'NumpadSubtract'], ',': ['Comma'], '.': ['Period'], '/': ['Slash', 'NumpadDivide'], '\\': ['Backslash', 'IntlBackslash'], ';': ['Semicolon'], '`': ['Backquote'], "'": ['Quote'], '[': ['BracketLeft'], ']': ['BracketRight'] };
const KEYS = { Up: 'ArrowUp', Down: 'ArrowDown', Left: 'ArrowLeft', Right: 'ArrowRight', Space: ' ', Esc: 'Escape', Enter: 'Enter', Tab: 'Tab', Backspace: 'Backspace', Delete: 'Delete', PageUp: 'PageUp', PageDown: 'PageDown', Home: 'Home', End: 'End' };

// The menu bar: ⌃⌥ + letter on a Mac, Alt + letter elsewhere (like Google Sheets).
export const MENU_KEYS = { File: 'F', Edit: 'E', Lihat: 'L', Sisipkan: 'S', Format: 'O', 'Tata letak': 'T', Data: 'D', Alat: 'A', Bantuan: 'B' };
const menuKey = l => ({ mac: [`Ctrl+Alt+${l}`], win: [`Alt+${l}`] });

export const GROUPS = ['Umum', 'Menu', 'Edit & tempel', 'Tempel spesial', 'Format teks & sel', 'Format angka', 'Batas', 'Baris, kolom & sel', 'Sisipkan', 'Data & alat', 'Navigasi & pilihan', 'Sheet', 'Saat mengedit sel', 'Tampilan'];

const S = (id, group, label, keys, o = {}) => ({ id, group, label, keys, ...o });
export const SHORTCUTS = [
  // ---------- umum ----------
  S('undo', 'Umum', 'Urungkan', ['Mod+Z']),
  S('redo', 'Umum', 'Ulangi', ['Mod+Y', 'Mod+Shift+Z']),
  S('repeat', 'Umum', 'Ulangi tindakan terakhir di sel lain', ['F4']),
  S('save', 'Umum', 'Simpan sekarang (juga tersimpan otomatis)', ['Mod+S']),
  S('open', 'Umum', 'Buka spreadsheet lain', ['Mod+O']),
  S('print', 'Umum', 'Cetak / simpan sebagai PDF', ['Mod+P']),
  S('find', 'Umum', 'Cari di sheet', ['Mod+F']),
  S('replace', 'Umum', 'Cari dan ganti', ['Mod+H', 'Mod+Shift+H'], { mac: ['Mod+Shift+H'] }),
  S('versions', 'Umum', 'Histori versi', ['Mod+Alt+Shift+H']),
  S('shortcuts', 'Umum', 'Daftar pintasan keyboard (jendela ini)', ['Mod+/']),
  S('menuSearch', 'Umum', 'Telusuri semua menu', ['Alt+/']),
  S('escape', 'Umum', 'Tutup menu / dialog, batalkan pilihan salin', ['Esc']),

  // ---------- menu ----------
  ...Object.entries(MENU_KEYS).map(([name, l]) => S('menu' + l, 'Menu', `Buka menu ${name}`, [], { ...menuKey(l), menu: name })),
  S('menuLetter', 'Menu', 'Di dalam menu: tekan huruf yang bergaris bawah', ['Huruf']),
  S('menuArrows', 'Menu', 'Di dalam menu: pindah, buka submenu, jalankan', ['Up', 'Right', 'Enter']),

  // ---------- edit & tempel ----------
  S('cut', 'Edit & tempel', 'Potong', ['Mod+X']),
  S('copy', 'Edit & tempel', 'Salin', ['Mod+C']),
  S('paste', 'Edit & tempel', 'Tempel', ['Mod+V']),
  S('pasteValues', 'Edit & tempel', 'Tempel nilai saja (tanpa rumus & format)', ['Mod+Shift+V']),
  S('pasteSpecial', 'Edit & tempel', 'Tempel spesial… (pilih jenis tempel)', ['Mod+Alt+V'], { mac: ['Mod+Alt+V', 'Ctrl+Mod+V'] }),
  S('clear', 'Edit & tempel', 'Hapus isi sel', ['Delete', 'Backspace'], { mac: ['Backspace', 'Delete'] }),
  S('clearFormats', 'Edit & tempel', 'Hapus format', ['Mod+\\']),
  S('clearAll', 'Edit & tempel', 'Hapus isi dan format', ['Mod+Alt+\\']),
  S('fillDown', 'Edit & tempel', 'Isi ke bawah (salin sel teratas)', ['Mod+D']),
  S('fillRight', 'Edit & tempel', 'Isi ke kanan (salin sel terkiri)', ['Mod+R']),
  S('flashFill', 'Edit & tempel', 'Isi kilat (ikuti pola contoh)', ['Mod+E']),

  // ---------- tempel spesial (in the Tempel spesial dialog) ----------
  S('psValues', 'Tempel spesial', 'Tempel nilai saja', ['Mod+Alt+V N']),
  S('psFormats', 'Tempel spesial', 'Tempel format saja', ['Mod+Alt+V F']),
  S('psFormulas', 'Tempel spesial', 'Tempel rumus saja', ['Mod+Alt+V R']),
  S('psNoFmt', 'Tempel spesial', 'Tempel semua kecuali format', ['Mod+Alt+V K']),
  S('psValid', 'Tempel spesial', 'Tempel validasi data saja', ['Mod+Alt+V V']),
  S('psWidths', 'Tempel spesial', 'Tempel lebar kolom saja', ['Mod+Alt+V L']),
  S('psTranspose', 'Tempel spesial', 'Tempel ditransposisikan (baris ↔ kolom)', ['Mod+Alt+V T']),
  S('psBlanks', 'Tempel spesial', 'Lewati sel kosong', ['Mod+Alt+V W']),
  S('psOk', 'Tempel spesial', 'Tempel dengan pilihan itu', ['Enter', 'Mod+V']),

  // ---------- format teks & sel ----------
  S('bold', 'Format teks & sel', 'Tebal', ['Mod+B']),
  S('italic', 'Format teks & sel', 'Miring', ['Mod+I']),
  S('underline', 'Format teks & sel', 'Garis bawah', ['Mod+U']),
  S('strike', 'Format teks & sel', 'Coret', ['Mod+Shift+X', 'Alt+Shift+5', 'Mod+5']),
  S('alignLeft', 'Format teks & sel', 'Rata kiri', ['Mod+Shift+L']),
  S('alignCenter', 'Format teks & sel', 'Rata tengah', ['Mod+Shift+E']),
  S('alignRight', 'Format teks & sel', 'Rata kanan', ['Mod+Shift+R']),
  S('valignMenu', 'Format teks & sel', 'Perataan vertikal (atas / tengah / bawah)', ['Alt+Shift+A']),
  S('wrap', 'Format teks & sel', 'Kemas teks (aktif / nonaktif)', ['Alt+Shift+W']),
  S('rotateMenu', 'Format teks & sel', 'Rotasi teks', ['Alt+Shift+R']),
  S('fontMenu', 'Format teks & sel', 'Pilih font', ['Alt+Shift+F']),
  S('fontBigger', 'Format teks & sel', 'Perbesar ukuran font', ['Mod+Shift+.']),
  S('fontSmaller', 'Format teks & sel', 'Perkecil ukuran font', ['Mod+Shift+,']),
  S('textColor', 'Format teks & sel', 'Warna teks', ['Alt+Shift+C']),
  S('fillColor', 'Format teks & sel', 'Warna isian sel', ['Alt+Shift+I']),
  S('mergeToggle', 'Format teks & sel', 'Gabungkan sel / pisahkan kembali', ['Alt+Shift+M']),
  S('unmerge', 'Format teks & sel', 'Pisahkan sel yang digabung', ['Alt+Shift+U']),
  S('paintFormat', 'Format teks & sel', 'Format cat (salin format ke sel lain, lalu Enter)', ['Mod+Alt+C']),
  S('formatCells', 'Format teks & sel', 'Format sel… (dialog lengkap)', ['Mod+1']),

  // ---------- format angka ----------
  S('numMenu', 'Format angka', 'Pilih format angka', ['Alt+Shift+N']),
  S('fmtNumber', 'Format angka', 'Angka 1.234,56', ['Mod+Shift+1']),
  S('fmtTime', 'Format angka', 'Waktu', ['Mod+Shift+2']),
  S('fmtDate', 'Format angka', 'Tanggal', ['Mod+Shift+3']),
  S('fmtCurrency', 'Format angka', 'Mata uang (Rp)', ['Mod+Shift+4']),
  S('fmtPercent', 'Format angka', 'Persen', ['Mod+Shift+5']),
  S('fmtSci', 'Format angka', 'Ilmiah (eksponen)', ['Mod+Shift+6']),
  S('decMore', 'Format angka', 'Tambah tempat desimal', ['Mod+Alt+.']),
  S('decLess', 'Format angka', 'Kurangi tempat desimal', ['Mod+Alt+,']),

  // ---------- batas ----------
  S('bordersMenu', 'Batas', 'Pilih batas, warna & gaya garis', ['Alt+Shift+G']),
  S('borderTop', 'Batas', 'Batas atas', ['Alt+Shift+1']),
  S('borderRight', 'Batas', 'Batas kanan', ['Alt+Shift+2']),
  S('borderBottom', 'Batas', 'Batas bawah', ['Alt+Shift+3']),
  S('borderLeft', 'Batas', 'Batas kiri', ['Alt+Shift+4']),
  S('borderClear', 'Batas', 'Hapus batas', ['Alt+Shift+6', 'Mod+Shift+-']),
  S('borderOuter', 'Batas', 'Batas luar', ['Alt+Shift+7', 'Mod+Shift+7']),
  S('borderAll', 'Batas', 'Semua batas', ['Alt+Shift+8']),

  // ---------- baris, kolom & sel ----------
  S('insRowAbove', 'Baris, kolom & sel', 'Sisipkan baris di atas', ['Alt+Shift+Up']),
  S('insRowBelow', 'Baris, kolom & sel', 'Sisipkan baris di bawah', ['Alt+Shift+Down']),
  S('insColLeft', 'Baris, kolom & sel', 'Sisipkan kolom di kiri', ['Alt+Shift+Left']),
  S('insColRight', 'Baris, kolom & sel', 'Sisipkan kolom di kanan', ['Alt+Shift+Right']),
  S('insert', 'Baris, kolom & sel', 'Sisipkan… (baris/kolom terpilih langsung, sel: pilih geser)', ['Mod+Alt+=', 'Mod+Shift+=']),
  S('insRowSeq', 'Baris, kolom & sel', 'Sisipkan seluruh baris (dari dialog Sisipkan)', ['Mod+Alt+= B']),
  S('insColSeq', 'Baris, kolom & sel', 'Sisipkan seluruh kolom (dari dialog Sisipkan)', ['Mod+Alt+= K']),
  S('delete', 'Baris, kolom & sel', 'Hapus… (baris/kolom terpilih langsung, sel: pilih geser)', ['Mod+Alt+-']),
  S('delRowSeq', 'Baris, kolom & sel', 'Hapus seluruh baris', ['Mod+Alt+- B', 'Shift+Space Mod+Alt+-']),
  S('delColSeq', 'Baris, kolom & sel', 'Hapus seluruh kolom', ['Mod+Alt+- K', 'Ctrl+Space Mod+Alt+-']),
  S('hideRows', 'Baris, kolom & sel', 'Sembunyikan baris', ['Mod+9', 'Mod+Alt+9']),
  S('hideCols', 'Baris, kolom & sel', 'Sembunyikan kolom', ['Mod+0', 'Mod+Alt+0']),
  S('unhideRows', 'Baris, kolom & sel', 'Tampilkan baris tersembunyi', ['Mod+Shift+9']),
  S('unhideCols', 'Baris, kolom & sel', 'Tampilkan kolom tersembunyi', ['Mod+Shift+0']),
  S('autofit', 'Baris, kolom & sel', 'Sesuaikan lebar kolom dengan isi', ['Alt+Shift+O']),

  // ---------- sisipkan ----------
  S('link', 'Sisipkan', 'Sisipkan link', ['Mod+K']),
  S('note', 'Sisipkan', 'Sisipkan / edit catatan', ['Shift+F2', 'Mod+Alt+M']),
  S('date', 'Sisipkan', 'Tanggal hari ini', ['Mod+;']),
  S('time', 'Sisipkan', 'Waktu sekarang', ['Mod+Shift+;']),
  S('datetime', 'Sisipkan', 'Tanggal dan waktu', ['Mod+Alt+Shift+;']),
  S('autosum', 'Sisipkan', 'Jumlah otomatis (SUM)', ['Alt+=']),
  S('function', 'Sisipkan', 'Sisipkan fungsi (daftar fungsi)', ['Shift+F3']),
  S('chart', 'Sisipkan', 'Sisipkan diagram', ['Alt+Shift+D', 'Alt+F1']),
  S('signature', 'Sisipkan', 'Sisipkan tanda tangan', ['Alt+Shift+S']),
  S('checkbox', 'Sisipkan', 'Sisipkan kotak centang', ['Alt+Shift+X']),
  S('dropdown', 'Sisipkan', 'Sisipkan dropdown', ['Alt+Shift+L']),

  // ---------- data & alat ----------
  S('filter', 'Data & alat', 'Buat / hapus filter', ['Mod+Shift+F']),
  S('banding', 'Data & alat', 'Warna bergantian / format sebagai tabel', ['Alt+Shift+T']),
  S('cf', 'Data & alat', 'Format bersyarat', ['Alt+Shift+B']),
  S('dv', 'Data & alat', 'Validasi data', ['Alt+Shift+V']),
  S('pivot', 'Data & alat', 'Tabel pivot', ['Alt+Shift+P']),
  S('names', 'Data & alat', 'Rentang bernama', ['Mod+F3']),
  S('recalc', 'Data & alat', 'Hitung ulang semua rumus', ['F9']),

  // ---------- navigasi ----------
  S('arrows', 'Navigasi & pilihan', 'Pindah satu sel', ['Up', 'Down', 'Left', 'Right']),
  S('edge', 'Navigasi & pilihan', 'Lompat ke tepi data', ['Mod+Up', 'Mod+Down']),
  S('extend', 'Navigasi & pilihan', 'Perluas pilihan', ['Shift+Up', 'Shift+Down']),
  S('extendEdge', 'Navigasi & pilihan', 'Perluas pilihan sampai tepi data', ['Mod+Shift+Down', 'Mod+Shift+Right']),
  S('nextCell', 'Navigasi & pilihan', 'Sel kanan / kiri', ['Tab', 'Shift+Tab']),
  S('home', 'Navigasi & pilihan', 'Ke awal baris', ['Home']),
  S('ctrlHome', 'Navigasi & pilihan', 'Ke awal sheet (A1)', ['Mod+Home']),
  S('ctrlEnd', 'Navigasi & pilihan', 'Ke sel terakhir yang berisi data', ['Mod+End']),
  S('page', 'Navigasi & pilihan', 'Gulir satu layar ke bawah / atas', ['PageDown', 'PageUp']),
  S('pageSide', 'Navigasi & pilihan', 'Gulir satu layar ke kanan / kiri', ['Alt+PageDown', 'Alt+PageUp']),
  S('goto', 'Navigasi & pilihan', 'Buka ke sel / rentang (kotak nama)', ['Mod+G', 'F5']),
  S('selRow', 'Navigasi & pilihan', 'Pilih seluruh baris', ['Shift+Space']),
  S('selCol', 'Navigasi & pilihan', 'Pilih seluruh kolom', ['Ctrl+Space']),
  S('selAll', 'Navigasi & pilihan', 'Pilih semua', ['Mod+A', 'Mod+Shift+Space']),
  S('dropdownOpen', 'Navigasi & pilihan', 'Buka pilihan dropdown di sel', ['Alt+Down']),

  // ---------- sheet ----------
  S('addSheet', 'Sheet', 'Sheet baru', ['Shift+F11']),
  S('nextSheet', 'Sheet', 'Sheet berikutnya', ['Mod+Shift+PageDown'], { mac: ['Mod+PageDown', 'Mod+Shift+PageDown'] }),
  S('prevSheet', 'Sheet', 'Sheet sebelumnya', ['Mod+Shift+PageUp'], { mac: ['Mod+PageUp', 'Mod+Shift+PageUp'] }),
  S('sheetList', 'Sheet', 'Daftar semua sheet', ['Alt+Shift+K']),

  // ---------- saat mengedit ----------
  S('edit', 'Saat mengedit sel', 'Mulai edit sel (atau langsung ketik)', ['Enter', 'F2']),
  S('newline', 'Saat mengedit sel', 'Baris baru di dalam sel', ['Alt+Enter']),
  S('commitDown', 'Saat mengedit sel', 'Selesai, pindah ke bawah / atas', ['Enter', 'Shift+Enter']),
  S('commitRight', 'Saat mengedit sel', 'Selesai, pindah ke kanan / kiri', ['Tab', 'Shift+Tab']),
  S('fillSel', 'Saat mengedit sel', 'Isi semua sel terpilih dengan isian ini', ['Mod+Enter']),
  S('arrayF', 'Saat mengedit sel', 'Masukkan sebagai rumus array', ['Mod+Shift+Enter']),
  S('absRef', 'Saat mengedit sel', 'Ganti referensi A1 → $A$1 → A$1 → $A1', ['F4']),
  S('cancelEdit', 'Saat mengedit sel', 'Batalkan edit', ['Esc']),

  // ---------- tampilan ----------
  S('zoomIn', 'Tampilan', 'Perbesar zoom', ['Mod+=']),
  S('zoomOut', 'Tampilan', 'Perkecil zoom', ['Mod+-']),
  S('formulas', 'Tampilan', 'Tampilkan rumus di sel', ['Ctrl+`']),
  S('compact', 'Tampilan', 'Sembunyikan / tampilkan bilah alat', ['Alt+Shift+H']),
];
const BY_ID = new Map(SHORTCUTS.map(s => [s.id, s]));

// ---------- parsing / formatting ----------
function parseChord(spec) {
  const parts = spec.split('+');
  // a trailing '+' would split into ['', ''] — not used; '-' / '=' are plain keys
  const key = parts.pop();
  const m = new Set(parts);
  return { mod: m.has('Mod'), ctrl: m.has('Ctrl'), alt: m.has('Alt'), shift: m.has('Shift'), key };
}
export const keysFor = (s, mac) => (s ? (mac ? s.mac : s.win) || s.keys : []);
export const keysOf = s => keysFor(s, isMac);

function keyName(k, mac) {
  const n = (mac ? NAMES_MAC : NAMES_WIN)[k];
  if (n) return n;
  return k.length === 1 ? k.toUpperCase() : k;
}
// one chord → ['⌥', '⇧', 'M'] (Mac: Apple order ⌃⌥⇧⌘) or ['Ctrl', 'Alt', 'M']
export function chordParts(spec, mac = isMac) {
  const c = parseChord(spec);
  const out = [];
  if (mac) {
    if (c.ctrl) out.push('⌃');
    if (c.alt) out.push('⌥');
    if (c.shift) out.push('⇧');
    if (c.mod) out.push('⌘');
  } else {
    if (c.mod || c.ctrl) out.push('Ctrl');
    if (c.alt) out.push('Alt');
    if (c.shift) out.push('Shift');
  }
  out.push(keyName(c.key, mac));
  return out;
}
export function fmtChord(spec, mac = isMac) {
  const p = chordParts(spec, mac);
  return mac ? (p.length > 1 && p[p.length - 1].length > 1 ? p.slice(0, -1).join('') + ' ' + p[p.length - 1] : p.join('')) : p.join('+');
}
// 'Mod+Alt+V N' → '⌥⌘V lalu N'
export const fmtKeys = (spec, mac = isMac) => spec.split(' ').map(x => fmtChord(x, mac)).join(' lalu ');
// Key hint for a menu row / tooltip: the first alternative (or the first two, joined)
export function K(id, n = 1) {
  const ks = keysOf(BY_ID.get(id));
  return ks.length ? ks.slice(0, n).map(k => fmtKeys(k)).join(' / ') : '';
}
// 'Tooltip label|keys'
export const tip = (label, id, n) => { const k = K(id, n); return k ? `${label}|${k}` : label; };
export const shortcut = id => BY_ID.get(id);

// ---------- matching ----------
// legacy keyCodes, for keyboards / remote desktops / automation that leave e.code empty
const KEYCODES = { '=': [187, 61, 107], '-': [189, 173, 109], ',': [188], '.': [190], '/': [191, 111], '\\': [220], ';': [186, 59], '`': [192], "'": [222], '[': [219], ']': [221] };
function keyMatches(e, key) {
  const noCode = !e.code;
  if (CODES[key]) return CODES[key].includes(e.code) || (noCode && (e.key === key || KEYCODES[key]?.includes(e.keyCode)));
  if (KEYS[key]) return e.key === KEYS[key] || (key === 'Space' && e.code === 'Space');
  if (/^F\d+$/.test(key)) return e.key === key;
  // letters: the printed key unless Option/Alt is down (on a Mac it turns e.key into a symbol)
  if (/^[A-Z]$/.test(key)) {
    if (!e.altKey && /^[a-z]$/i.test(e.key)) return e.key.toUpperCase() === key;
    return e.code === 'Key' + key || (noCode && (e.key.toUpperCase() === key || e.keyCode === key.charCodeAt(0)));
  }
  if (/^\d$/.test(key)) return e.code === 'Digit' + key || e.code === 'Numpad' + key || (noCode && (e.key === key || e.keyCode === 48 + +key));
  return false;
}
function chordMatches(e, spec) {
  if (spec.includes(' ')) return false;   // sequences are shown, not matched
  const c = parseChord(spec);
  if (isMac) { if (!!e.metaKey !== c.mod || !!e.ctrlKey !== c.ctrl) return false; }
  else if (e.metaKey || !!e.ctrlKey !== (c.mod || c.ctrl)) return false;
  if (!!e.altKey !== c.alt || !!e.shiftKey !== c.shift) return false;
  return keyMatches(e, c.key);
}
// First shortcut with a handler whose keys match the event
export function matchShortcut(e, handlers) {
  for (const s of SHORTCUTS) {
    if (!handlers[s.id]) continue;
    for (const k of keysOf(s)) if (chordMatches(e, k)) return s;
  }
  return null;
}
// The menu name for ⌃⌥ + letter (Mac) / Alt + letter (elsewhere), or null
export function menuFor(e) {
  const ok = isMac ? e.ctrlKey && e.altKey && !e.metaKey && !e.shiftKey : e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey;
  if (!ok) return null;
  const l = /^Key[A-Z]$/.test(e.code) ? e.code.slice(3) : !e.code && e.keyCode >= 65 && e.keyCode <= 90 ? String.fromCharCode(e.keyCode) : !e.code && /^[a-z]$/i.test(e.key) ? e.key.toUpperCase() : null;
  if (!l) return null;
  return Object.keys(MENU_KEYS).find(n => MENU_KEYS[n] === l) || null;
}
