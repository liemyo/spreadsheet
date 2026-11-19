// Starter templates.
import { newWorkbook, setInput } from './model.js';
import { parseRange, uid, serial, jsDateToSerial } from './util.js';

function make(name, fn) {
  const wb = newWorkbook(name);
  wb.hist.paused++;
  fn(wb, wb.sheets[0]);
  wb.hist.paused--;
  wb.dirty = false;
  return wb;
}
function at(a) { const g = parseRange(a); return [g.r1, g.c1]; }
function tall(sh, r, st) {
  if (st && st.fs > 11) sh.rowH.set(r, Math.max(sh.rowH.get(r) || 21, Math.ceil(st.fs * 4 / 3 * 1.3 + 8)));
}
function put(wb, sh, a, v, st) {
  const [r, c] = at(a);
  tall(sh, r, st);
  if (v !== null && v !== undefined) setInput(wb, sh, r, c, String(v));
  if (st) sh.setStyle(r, c, { ...sh.styleOf(r, c), ...st });
}
function row(wb, sh, a, vals, st) {
  const [r, c] = at(a);
  tall(sh, r, st);
  vals.forEach((v, j) => { if (v !== null && v !== undefined) setInput(wb, sh, r, c + j, String(v)); if (st) sh.setStyle(r, c + j, { ...sh.styleOf(r, c + j), ...st }); });
}
function styleRange(sh, a, st) {
  const g = parseRange(a);
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) sh.setStyle(r, c, { ...sh.styleOf(r, c), ...st });
}
const widths = (sh, list) => list.forEach((w, i) => w && sh.colW.set(i, w));
const RP = '"Rp"#,##0';
const today = () => Math.floor(jsDateToSerial(new Date()));
const H = { b: true, color: '#ffffff', bg: '#1f6f54', va: 'middle' };
const thin = { w: 1, style: 'solid', color: '#d0d7de' };

function budget() {
  return make('Anggaran Bulanan', (wb, sh) => {
    sh.name = 'Anggaran';
    widths(sh, [200, 130, 130, 130, 120, 30]);
    put(wb, sh, 'A1', 'Anggaran Bulanan', { b: true, fs: 18, color: '#1f6f54' });
    put(wb, sh, 'A2', '=TEXT(TODAY(),"mmmm yyyy")', { color: '#5f6368' });
    row(wb, sh, 'A4', ['Kategori', 'Anggaran', 'Aktual', 'Selisih', '% Terpakai'], H);
    styleRange(sh, 'B4:E4', { ha: 'right' });
    const data = [['Rumah & sewa', 3500000, 3500000], ['Makan & belanja', 2500000, 2730000], ['Transportasi', 900000, 760000], ['Listrik & air', 650000, 702000], ['Internet & pulsa', 400000, 400000], ['Kesehatan', 500000, 150000], ['Hiburan', 600000, 845000], ['Tabungan & investasi', 2000000, 2000000]];
    data.forEach((d, i) => {
      const r = 5 + i;
      row(wb, sh, 'A' + r, [d[0], d[1], d[2], `=B${r}-C${r}`, `=IFERROR(C${r}/B${r},0)`]);
      styleRange(sh, `B${r}:D${r}`, { fmt: RP });
      put(wb, sh, 'E' + r, null, { fmt: '0%' });
    });
    row(wb, sh, 'A13', ['Total', '=SUM(B5:B12)', '=SUM(C5:C12)', '=B13-C13', '=IFERROR(C13/B13,0)'], { b: true, bt: { w: 2, style: 'solid', color: '#1f6f54' } });
    styleRange(sh, 'B13:D13', { fmt: RP }); put(wb, sh, 'E13', null, { fmt: '0%' });
    put(wb, sh, 'A16', 'Ringkasan', { b: true, fs: 12 });
    row(wb, sh, 'A17', ['Pemasukan bulan ini', 12500000]); put(wb, sh, 'B17', null, { fmt: RP });
    row(wb, sh, 'A18', ['Total pengeluaran', '=C13']); put(wb, sh, 'B18', null, { fmt: RP });
    row(wb, sh, 'A19', ['Sisa dana', '=B17-B18'], { b: true }); put(wb, sh, 'B19', null, { fmt: RP });
    row(wb, sh, 'A20', ['Rasio tabungan', '=C12/B17']); put(wb, sh, 'B20', null, { fmt: '0.0%' });
    styleRange(sh, 'A17:B20', { bb: thin });
    sh.banding.push({ r1: 3, c1: 0, r2: 11, c2: 4, h: '#1f6f54', a: '#ffffff', b: '#eef6f2', header: true });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: 4, c1: 4, r2: 11, c2: 4 }], type: 'bar', bar: { color: '#34a853' } });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: 4, c1: 3, r2: 12, c2: 3 }], type: 'num', op: 'lt', v1: '0', style: { color: '#c5221f', b: true } });
    sh.charts.push({ id: uid('ch'), type: 'column', range: 'A4:C12', anchor: { r: 3, c: 6 }, dx: 0, dy: 0, w: 520, h: 320, opts: { title: 'Anggaran vs Aktual', legend: 'bottom', headers: 'on', labels: 'on', colors: ['#9ad0b5', '#1f6f54'] } });
    sh.freezeR = 0;
  });
}

function todo() {
  return make('Daftar Tugas', (wb, sh) => {
    sh.name = 'Tugas';
    widths(sh, [70, 300, 150, 110, 130, 110, 240]);
    row(wb, sh, 'A1', ['Selesai', 'Tugas', 'Penanggung jawab', 'Prioritas', 'Status', 'Tenggat', 'Catatan'], { b: true, bg: '#f1f3f4', bb: { w: 2, style: 'solid', color: '#5f6368' }, va: 'middle' });
    sh.rowH.set(0, 30);
    const t = today();
    const tasks = [
      [true, 'Susun rencana kerja kuartal IV', 'Rina', 'Tinggi', 'Selesai', t - 6, ''],
      [false, 'Presentasi hasil survei pelanggan', 'Budi', 'Tinggi', 'Dikerjakan', t + 2, 'Draf di folder tim'],
      [false, 'Perbarui daftar harga pemasok', 'Sari', 'Sedang', 'Dikerjakan', t + 5, ''],
      [false, 'Rapat koordinasi mingguan', 'Andi', 'Sedang', 'Belum mulai', t + 1, 'Ruang rapat 2'],
      [false, 'Audit stok gudang', 'Dewi', 'Tinggi', 'Belum mulai', t - 1, 'Terlambat'],
      [true, 'Kirim laporan bulanan', 'Rina', 'Rendah', 'Selesai', t - 3, ''],
      [false, 'Evaluasi kampanye media sosial', 'Budi', 'Rendah', 'Belum mulai', t + 10, ''],
      [false, 'Siapkan materi pelatihan', 'Sari', 'Sedang', 'Dikerjakan', t + 7, ''],
    ];
    tasks.forEach((x, i) => {
      const r = i + 1;
      sh.put(r, 0, { v: x[0] });
      [1, 2, 3, 4].forEach(c => setInput(wb, sh, r, c, x[c]));
      sh.put(r, 5, { v: x[5] });
      if (x[6]) setInput(wb, sh, r, 6, x[6]);
    });
    for (let r = tasks.length + 1; r < 30; r++) sh.put(r, 0, { v: false });
    styleRange(sh, 'F2:F30', { fmt: 'd mmm yyyy' });
    sh.dv.push({ id: uid('dv'), type: 'checkbox', ranges: [{ r1: 1, c1: 0, r2: 29, c2: 0 }] });
    sh.dv.push({ id: uid('dv'), type: 'list', ranges: [{ r1: 1, c1: 3, r2: 29, c2: 3 }], items: ['Tinggi', 'Sedang', 'Rendah'], colors: ['#fad2cf', '#fce8b2', '#ceead6'], chip: 'chip', strict: true });
    sh.dv.push({ id: uid('dv'), type: 'list', ranges: [{ r1: 1, c1: 4, r2: 29, c2: 4 }], items: ['Belum mulai', 'Dikerjakan', 'Selesai'], colors: ['#e8eaed', '#c6dafc', '#ceead6'], chip: 'chip', strict: true });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: 1, c1: 1, r2: 29, c2: 6 }], type: 'formula', v1: '=$A2=TRUE', style: { s: true, color: '#9aa0a6' } });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: 1, c1: 5, r2: 29, c2: 5 }], type: 'formula', v1: '=AND($F2<TODAY(),$A2=FALSE,$F2<>"")', style: { color: '#c5221f', b: true } });
    sh.freezeR = 1;
    sh.filter = { r1: 0, c1: 0, r2: 29, c2: 6, crit: {} };
    styleRange(sh, 'A1:G1', { ha: 'left' });
  });
}

function invoice() {
  return make('Faktur', (wb, sh) => {
    sh.name = 'Faktur';
    sh.grid = false;
    widths(sh, [40, 60, 300, 80, 130, 150, 40]);
    put(wb, sh, 'B2', 'FAKTUR', { b: true, fs: 24, color: '#174ea6' });
    put(wb, sh, 'B3', 'PT Contoh Nusantara', { b: true });
    put(wb, sh, 'B4', 'Jl. Merdeka No. 10, Jakarta 10110', { color: '#5f6368' });
    put(wb, sh, 'B5', 'finance@contoh.co.id · (021) 555-0100', { color: '#5f6368' });
    row(wb, sh, 'E2', ['No. faktur', 'INV-2026-0142'], { ha: 'right' }); put(wb, sh, 'E2', null, { color: '#5f6368' });
    row(wb, sh, 'E3', ['Tanggal', '=TODAY()'], { ha: 'right' }); put(wb, sh, 'E3', null, { color: '#5f6368' }); put(wb, sh, 'F3', null, { fmt: 'd mmmm yyyy' });
    row(wb, sh, 'E4', ['Jatuh tempo', '=F3+14'], { ha: 'right' }); put(wb, sh, 'E4', null, { color: '#5f6368' }); put(wb, sh, 'F4', null, { fmt: 'd mmmm yyyy' });
    put(wb, sh, 'B7', 'DITAGIHKAN KEPADA', { b: true, fs: 8, color: '#5f6368' });
    put(wb, sh, 'B8', 'CV Maju Bersama', { b: true });
    put(wb, sh, 'B9', 'Jl. Diponegoro 21, Bandung');
    row(wb, sh, 'B11', ['No', 'Deskripsi', 'Qty', 'Harga satuan', 'Jumlah'], { b: true, bg: '#174ea6', color: '#ffffff', va: 'middle' });
    styleRange(sh, 'D11:F11', { ha: 'right' });
    sh.rowH.set(10, 28);
    const items = [['Jasa desain antarmuka aplikasi', 1, 8500000], ['Pengembangan modul laporan', 2, 4250000], ['Lisensi tahunan server', 1, 3600000], ['Pelatihan pengguna (per sesi)', 3, 1250000]];
    for (let i = 0; i < 8; i++) {
      const r = 12 + i, it = items[i];
      row(wb, sh, 'B' + r, [it ? i + 1 : null, it ? it[0] : null, it ? it[1] : null, it ? it[2] : null, `=IF(D${r}="","",D${r}*E${r})`], { bb: thin });
      styleRange(sh, `E${r}:F${r}`, { fmt: RP });
      put(wb, sh, 'B' + r, null, { ha: 'center', color: '#5f6368' });
    }
    row(wb, sh, 'E21', ['Subtotal', '=SUM(F12:F19)']);
    row(wb, sh, 'E22', ['Diskon', '=F21*0.05']);
    row(wb, sh, 'E23', ['PPN 11%', '=(F21-F22)*0.11']);
    row(wb, sh, 'E24', ['TOTAL', '=F21-F22+F23'], { b: true, fs: 12, color: '#174ea6', bt: { w: 2, style: 'solid', color: '#174ea6' } });
    styleRange(sh, 'F21:F24', { fmt: RP }); styleRange(sh, 'E21:E24', { ha: 'right' });
    put(wb, sh, 'B26', 'Terbilang:', { i: true, color: '#5f6368' });
    put(wb, sh, 'C26', '=PROPER(TERBILANG(ROUND(F24,0)))&" Rupiah"', { i: true });
    put(wb, sh, 'B28', 'Pembayaran ditransfer ke BCA 123-456-7890 a.n. PT Contoh Nusantara.', { color: '#5f6368', fs: 9 });
    put(wb, sh, 'B29', 'Terima kasih atas kerja sama Anda.', { color: '#5f6368', fs: 9 });
  });
}

function sales() {
  return make('Laporan Penjualan', (wb, sh) => {
    sh.name = 'Data';
    widths(sh, [110, 110, 140, 110, 70, 120, 140]);
    row(wb, sh, 'A1', ['Tanggal', 'Wilayah', 'Produk', 'Sales', 'Unit', 'Harga', 'Total'], { b: true, bg: '#e8f0fe', color: '#174ea6', bb: { w: 1, style: 'solid', color: '#174ea6' } });
    const regions = ['Jawa', 'Sumatra', 'Kalimantan', 'Sulawesi', 'Bali & NT'];
    const products = [['Laptop Pro 14', 14500000], ['Monitor 27"', 3900000], ['Keyboard Nirkabel', 650000], ['Headset Rapat', 1200000], ['Webcam HD', 850000]];
    const people = ['Rina', 'Budi', 'Sari', 'Andi', 'Dewi', 'Fajar'];
    let seed = 7;
    const rnd = n => { seed = (seed * 9301 + 49297) % 233280; return Math.floor(seed / 233280 * n); };
    const start = serial(2026, 1, 5);
    for (let i = 0; i < 60; i++) {
      const r = i + 2;
      const p = products[rnd(products.length)];
      sh.put(r - 1, 0, { v: start + i * 4 + rnd(3) });
      row(wb, sh, 'B' + r, [regions[rnd(regions.length)], p[0], people[rnd(people.length)], 1 + rnd(12), p[1], `=E${r}*F${r}`]);
    }
    styleRange(sh, 'A2:A61', { fmt: 'dd/mm/yyyy' });
    styleRange(sh, 'F2:G61', { fmt: RP });
    sh.freezeR = 1;
    sh.filter = { r1: 0, c1: 0, r2: 60, c2: 6, crit: {} };
    const s2 = wb.addSheet('Ringkasan', 1);
    widths(s2, [160, 160, 120, 30, 160, 160]);
    put(wb, s2, 'A1', 'Ringkasan Penjualan 2026', { b: true, fs: 16 });
    put(wb, s2, 'A2', '=COUNTA(Data!B2:B61)&" transaksi · total "&TEXT(SUM(Data!G2:G61),"""Rp""#,##0")', { color: '#5f6368' });
    row(wb, s2, 'A4', ['Wilayah', 'Total penjualan', 'Porsi'], { b: true, bb: { w: 2, style: 'solid', color: '#202124' } });
    put(wb, s2, 'A5', '=SORT(UNIQUE(Data!B2:B61))');
    put(wb, s2, 'B5', '=SUMIF(Data!B2:B61,A5#,Data!G2:G61)');
    put(wb, s2, 'C5', '=B5#/SUM(Data!G2:G61)');
    styleRange(s2, 'B5:B9', { fmt: RP }); styleRange(s2, 'C5:C9', { fmt: '0.0%' });
    s2.cf.push({ id: uid('cf'), ranges: [{ r1: 4, c1: 2, r2: 8, c2: 2 }], type: 'bar', bar: { color: '#4285f4' } });
    row(wb, s2, 'E4', ['Produk terlaris', 'Unit terjual'], { b: true, bb: { w: 2, style: 'solid', color: '#202124' } });
    put(wb, s2, 'E5', '=QUERY(Data!C2:E61,"select C, sum(E) group by C order by sum(E) desc label sum(E) \'\'",0)');
    put(wb, s2, 'A12', 'Penjualan per sales (XLOOKUP)', { b: true });
    row(wb, s2, 'A13', ['Nama sales', 'Rina'], {}); put(wb, s2, 'B13', null, { bg: '#fef7e0' });
    row(wb, s2, 'A14', ['Total', '=SUMIFS(Data!G2:G61,Data!D2:D61,B13)']); put(wb, s2, 'B14', null, { fmt: RP, b: true });
    row(wb, s2, 'A15', ['Transaksi terakhir', '=XLOOKUP(B13,Data!D2:D61,Data!A2:A61,"-",0,-1)']); put(wb, s2, 'B15', null, { fmt: 'd mmm yyyy' });
    s2.dv.push({ id: uid('dv'), type: 'list', ranges: [{ r1: 12, c1: 1, r2: 12, c2: 1 }], items: people, colors: people.map(() => '#e8f0fe'), chip: 'chip', strict: true });
    s2.charts.push({ id: uid('ch'), type: 'doughnut', range: 'A4:B9', anchor: { r: 16, c: 0 }, dx: 0, dy: 0, w: 440, h: 300, opts: { title: 'Penjualan per wilayah', legend: 'right', headers: 'on', labels: 'on' } });
    s2.charts.push({ id: uid('ch'), type: 'bar', range: 'E4:F9', anchor: { r: 16, c: 4 }, dx: 0, dy: 0, w: 440, h: 300, opts: { title: 'Unit terjual per produk', legend: 'none', headers: 'on', labels: 'on', colors: ['#34a853'] } });
    wb.active = 1;
  });
}

function gantt() {
  return make('Jadwal Proyek', (wb, sh) => {
    sh.name = 'Gantt';
    widths(sh, [220, 100, 70, 100, 80]);
    put(wb, sh, 'A1', 'Jadwal Proyek', { b: true, fs: 16 });
    row(wb, sh, 'A2', ['Mulai proyek', '=TODAY()-WEEKDAY(TODAY(),3)']); put(wb, sh, 'B2', null, { fmt: 'd mmm yyyy', b: true });
    row(wb, sh, 'A4', ['Tugas', 'Mulai', 'Durasi', 'Selesai', 'Progres'], { b: true, bg: '#f1f3f4', va: 'middle' });
    const tasks = [['Kick-off & kebutuhan', 0, 4, 1], ['Desain sistem', 3, 6, 0.9], ['Pengembangan backend', 8, 12, 0.55], ['Pengembangan frontend', 10, 12, 0.4], ['Integrasi & pengujian', 20, 7, 0.1], ['Pelatihan pengguna', 26, 3, 0], ['Go-live', 29, 1, 0]];
    tasks.forEach((t, i) => {
      const r = 5 + i;
      row(wb, sh, 'A' + r, [t[0], `=$B$2+${t[1]}`, t[2], `=B${r}+C${r}-1`, t[3]]);
      put(wb, sh, 'B' + r, null, { fmt: 'd mmm' }); put(wb, sh, 'D' + r, null, { fmt: 'd mmm' }); put(wb, sh, 'E' + r, null, { fmt: '0%' });
    });
    const days = 35;
    for (let j = 0; j < days; j++) {
      const c = 5 + j;
      sh.colW.set(c, 26);
      setInput(wb, sh, 3, c, j === 0 ? '=$B$2' : `=${colLetter(c - 1)}4+1`);
      sh.setStyle(3, c, { fmt: 'd', ha: 'center', b: true, bg: '#f1f3f4', fs: 8, va: 'middle' });
      setInput(wb, sh, 2, c, `=IF(OR(${colLetter(c)}4=$B$2,DAY(${colLetter(c)}4)=1),TEXT(${colLetter(c)}4,"mmm"),"")`);
      sh.setStyle(2, c, { fs: 8, color: '#5f6368' });
    }
    sh.rowH.set(3, 26);
    const g = { r1: 4, c1: 5, r2: 4 + tasks.length - 1, c2: 5 + days - 1 };
    sh.cf.push({ id: uid('cf'), ranges: [g], type: 'formula', v1: '=AND(F$4>=$B5,F$4<$B5+$C5*$E5)', style: { bg: '#1a73e8' } });
    sh.cf.push({ id: uid('cf'), ranges: [g], type: 'formula', v1: '=AND(F$4>=$B5,F$4<=$D5)', style: { bg: '#aecbfa' } });
    sh.cf.push({ id: uid('cf'), ranges: [{ ...g, r1: 3 }], type: 'formula', v1: '=F$4=TODAY()', style: { bg: '#fce8b2' } });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: 4, c1: 4, r2: 4 + tasks.length - 1, c2: 4 }], type: 'bar', bar: { color: '#34a853' } });
    sh.freezeC = 5; sh.freezeR = 4;
    sh.grid = true;
  });
}
function colLetter(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }

function loan() {
  return make('Simulasi Kredit', (wb, sh) => {
    sh.name = 'Kredit';
    widths(sh, [190, 160, 150, 150, 160, 30]);
    put(wb, sh, 'A1', 'Simulasi Kredit', { b: true, fs: 16 });
    put(wb, sh, 'A2', 'Ubah sel kuning; tabel angsuran dihitung otomatis dengan array dinamis.', { color: '#5f6368' });
    const inp = { bg: '#fef7e0', bb: { w: 1, style: 'solid', color: '#f9ab00' } };
    row(wb, sh, 'A4', ['Pokok pinjaman', 250000000]); put(wb, sh, 'B4', null, { ...inp, fmt: RP });
    row(wb, sh, 'A5', ['Bunga per tahun', '9.5%']); put(wb, sh, 'B5', null, { ...inp, fmt: '0.00%' });
    row(wb, sh, 'A6', ['Tenor (bulan)', 36]); put(wb, sh, 'B6', null, inp);
    row(wb, sh, 'A8', ['Angsuran per bulan', '=PMT(B5/12,B6,-B4)'], { b: true }); put(wb, sh, 'B8', null, { fmt: RP, color: '#174ea6', fs: 12 });
    row(wb, sh, 'A9', ['Total bunga', '=B8*B6-B4']); put(wb, sh, 'B9', null, { fmt: RP });
    row(wb, sh, 'A10', ['Total pembayaran', '=B8*B6']); put(wb, sh, 'B10', null, { fmt: RP });
    row(wb, sh, 'A12', ['Bulan ke', 'Angsuran', 'Bunga', 'Pokok', 'Sisa pinjaman'], { b: true, bg: '#e8f0fe', color: '#174ea6' });
    put(wb, sh, 'A13', '=SEQUENCE(B6)');
    put(wb, sh, 'B13', '=PMT($B$5/12,$B$6,-$B$4)+0*A13#');
    put(wb, sh, 'C13', '=IPMT($B$5/12,A13#,$B$6,-$B$4)');
    put(wb, sh, 'D13', '=PPMT($B$5/12,A13#,$B$6,-$B$4)');
    put(wb, sh, 'E13', '=$B$4-SCAN(0,D13#,LAMBDA(a,x,a+x))');
    styleRange(sh, 'B13:E372', { fmt: '#,##0' });
    sh.freezeR = 12;
  });
}

/* ---------- Timeline (Harian / Mingguan / Bulanan) ---------- */
// Skala dipilih dari dropdown B4; baris 5–6 (tersembunyi) berisi tanggal akhir/awal tiap kolom periode, header bulan /
// tahun, nomor periode dan bilah jadwal (format bersyarat) mengikuti otomatis. Bilah tua = progres, muda = rencana.
const TL_N = 40, TL_ROWS = 40;
const monday = () => { const t = today(); return t - ((new Date().getDay() + 6) % 7); };
function timelineSheet(wb, sh, { title, sub, scale, start, pic, tasks, picHead = 'Penanggung jawab', colors = {} }) {
  const P = { bar: '#f6b26b', done: '#e8710a', head: '#6d28d9', head2: '#8b5cf6', label: '#fff59d', ...colors };
  const first = 8, last = first + TL_N - 1, L = colLetter, r0 = 8, r1 = r0 + TL_ROWS - 1;   // columns I.., task rows 9..
  widths(sh, [34, 250, 140, 88, 62, 88, 64, 104]);
  for (let c = first; c <= last; c++) sh.colW.set(c, 32);
  put(wb, sh, 'A1', title, { b: true, fs: 18, color: P.head });
  put(wb, sh, 'D1', sub, { color: '#5f6368', i: true });
  row(wb, sh, 'A2', ['', 'Nama proyek', 'Proyek baru']); put(wb, sh, 'B2', null, { b: true }); put(wb, sh, 'C2', null, { bg: '#fef7e0' });
  row(wb, sh, 'A3', ['', 'Tanggal mulai', String(start)]); put(wb, sh, 'B3', null, { b: true }); put(wb, sh, 'C3', null, { fmt: 'd mmm yyyy', bg: '#fef7e0', ha: 'left' });
  row(wb, sh, 'A4', ['', 'Skala timeline', scale]); put(wb, sh, 'B4', null, { b: true }); put(wb, sh, 'C4', null, { bg: '#fef7e0', b: true, color: P.head });
  put(wb, sh, 'D4', 'pilih Harian · Mingguan · Bulanan — kolom jadwal menyesuaikan otomatis', { color: '#5f6368', fs: 9 });
  sh.dv.push({ id: uid('dv'), type: 'list', ranges: [{ r1: 3, c1: 2, r2: 3, c2: 2 }], items: ['Harian', 'Mingguan', 'Bulanan'], colors: ['#ede9fe', '#ddd6fe', '#c4b5fd'], chip: 'arrow', strict: true });
  // hidden helper rows: 5 = period end, 6 = period start
  for (let c = first; c <= last; c++) {
    const C = L(c), prev = L(c - 1);
    setInput(wb, sh, 5, c, c === first ? '=IF($C$4="Harian",$C$3,IF($C$4="Mingguan",$C$3-WEEKDAY($C$3,3),DATE(YEAR($C$3),MONTH($C$3),1)))'
      : `=IF($C$4="Harian",${prev}6+1,IF($C$4="Mingguan",${prev}6+7,EDATE(${prev}6,1)))`);
    setInput(wb, sh, 4, c, `=IF($C$4="Harian",${C}6,IF($C$4="Mingguan",${C}6+6,EDATE(${C}6,1)-1))`);
    const f = 'IF($C$4="Bulanan","yyyy","mmm")';
    setInput(wb, sh, 6, c, c === first ? `=TEXT(${C}6,${f})` : `=IF(TEXT(${C}6,${f})<>TEXT(${prev}6,${f}),TEXT(${C}6,${f}),"")`);
    setInput(wb, sh, 7, c, `=IF($C$4="Harian",DAY(${C}6),IF($C$4="Mingguan",COLUMN()-${first},TEXT(${C}6,"mmm")))`);
    sh.setStyle(4, c, { fmt: 'dd/mm/yy' }); sh.setStyle(5, c, { fmt: 'dd/mm/yy' });
    sh.setStyle(6, c, { b: true, color: '#ffffff', bg: P.head, fs: 9, va: 'middle' });
    sh.setStyle(7, c, { b: true, color: '#ffffff', bg: P.head2, fs: 8, ha: 'center', va: 'middle' });
  }
  sh.hiddenR.add(4); sh.hiddenR.add(5);
  styleRange(sh, `A7:H7`, { bg: P.head, color: '#ffffff', b: true, fs: 11, va: 'middle' });
  put(wb, sh, 'B7', 'Jadwal kegiatan', null);
  row(wb, sh, 'A8', ['No', 'Kegiatan', picHead, 'Mulai', 'Durasi (hari)', 'Selesai', 'Progres', 'Status'], { b: true, bg: P.label, va: 'middle', wrap: 'wrap' });
  styleRange(sh, 'D8:G8', { ha: 'center' });
  sh.rowH.set(6, 24); sh.rowH.set(7, 34);
  // tasks: [kegiatan, pic, hari ke- (dari tanggal mulai), durasi, progres]
  for (let i = 0; i < TL_ROWS; i++) {
    const r = r0 + i, R = r + 1, t = tasks[i];
    setInput(wb, sh, r, 0, `=IF(B${R}="","",COUNTA($B$${r0 + 1}:B${R}))`);
    setInput(wb, sh, r, 5, `=IF(OR(D${R}="",E${R}=""),"",D${R}+E${R}-1)`);
    if (t) {
      row(wb, sh, 'B' + R, [t[0], t[1], `=$C$3+${t[2]}`, t[3]]);
      setInput(wb, sh, r, 6, String(t[4] * 100) + '%');
      setInput(wb, sh, r, 7, t[4] >= 1 ? 'Selesai' : t[4] > 0 ? 'Berjalan' : 'Belum mulai');
    }
  }
  styleRange(sh, `A${r0 + 1}:A${r1 + 1}`, { ha: 'center', color: '#5f6368' });
  styleRange(sh, `D${r0 + 1}:D${r1 + 1}`, { fmt: 'dd/mm/yy', ha: 'center' });
  styleRange(sh, `E${r0 + 1}:E${r1 + 1}`, { ha: 'center' });
  styleRange(sh, `F${r0 + 1}:F${r1 + 1}`, { fmt: 'dd/mm/yy', ha: 'center' });
  styleRange(sh, `G${r0 + 1}:G${r1 + 1}`, { fmt: '0%', ha: 'center' });
  styleRange(sh, `A${r0 + 1}:H${r1 + 1}`, { bb: thin });
  sh.dv.push({ id: uid('dv'), type: 'list', ranges: [{ r1: r0, c1: 2, r2: r1, c2: 2 }], items: pic, colors: pic.map(() => '#f3e8ff'), chip: 'arrow', strict: false });
  sh.dv.push({ id: uid('dv'), type: 'list', ranges: [{ r1: r0, c1: 7, r2: r1, c2: 7 }], items: ['Belum mulai', 'Berjalan', 'Selesai', 'Tertunda'], colors: ['#e8eaed', '#c6dafc', '#ceead6', '#fad2cf'], chip: 'chip', strict: true });
  const g = { r1: r0, c1: first, r2: r1, c2: last }, I = L(first);
  sh.cf.push({ id: uid('cf'), ranges: [g], type: 'formula', v1: `=AND(ISNUMBER($D${r0 + 1}),$G${r0 + 1}>0,$D${r0 + 1}<=${I}$5,$D${r0 + 1}+$E${r0 + 1}*$G${r0 + 1}-1>=${I}$6)`, style: { bg: P.done } });
  sh.cf.push({ id: uid('cf'), ranges: [g], type: 'formula', v1: `=AND(ISNUMBER($D${r0 + 1}),$D${r0 + 1}<=${I}$5,$F${r0 + 1}>=${I}$6)`, style: { bg: P.bar } });
  sh.cf.push({ id: uid('cf'), ranges: [g], type: 'formula', v1: `=AND($C$4="Harian",WEEKDAY(${I}$6,2)>5)`, style: { bg: '#f1f3f4' } });
  sh.cf.push({ id: uid('cf'), ranges: [{ r1: 6, c1: first, r2: 7, c2: last }], type: 'formula', v1: `=AND(TODAY()>=${I}$6,TODAY()<=${I}$5)`, style: { bg: '#fbbc04', color: '#202124' } });
  sh.cf.push({ id: uid('cf'), ranges: [g], type: 'formula', v1: `=AND(TODAY()>=${I}$6,TODAY()<=${I}$5)`, style: { bg: '#fef7e0' } });
  sh.cf.push({ id: uid('cf'), ranges: [{ r1: r0, c1: 5, r2: r1, c2: 5 }], type: 'formula', v1: `=AND(ISNUMBER($F${r0 + 1}),$F${r0 + 1}<TODAY(),$G${r0 + 1}<1)`, style: { color: '#c5221f', b: true } });
  sh.cf.push({ id: uid('cf'), ranges: [{ r1: r0, c1: 6, r2: r1, c2: 6 }], type: 'bar', bar: { color: '#34a853' } });
  // vertical lines between periods like a paper schedule
  for (let r = 6; r <= r1; r++) for (let c = first; c <= last; c++) sh.setStyle(r, c, { ...sh.styleOf(r, c), br: { w: 1, style: 'solid', color: r >= r0 ? '#e0e0e0' : '#ffffff' } });
  sh.freezeR = 8; sh.freezeC = 2;
  sh.grid = true;
  sh.page = { v: 2, orient: 'landscape', fit: 'width', title: false, num: true, hdr: true, auto: true, hc: false };
}
function timeline() {
  return make('Timeline Proyek', (wb, sh) => {
    sh.name = 'Timeline';
    const pic = ['Manajer Proyek', 'Tim Teknik', 'Admin & Legal', 'Pengadaan', 'Kontraktor', 'Quality Control', 'Klien / Owner', 'Keuangan'];
    timelineSheet(wb, sh, {
      title: 'Timeline Proyek', sub: 'Isi kegiatan, tanggal mulai & durasi — bilah jadwal terisi otomatis', scale: 'Mingguan', start: monday(), pic,
      tasks: [
        ['Inisiasi & kick-off meeting', 'Manajer Proyek', 0, 3, 1],
        ['Survei lapangan & pengukuran', 'Tim Teknik', 2, 6, 1],
        ['Perencanaan & desain teknis', 'Tim Teknik', 7, 14, 0.6],
        ['Perizinan & administrasi', 'Admin & Legal', 7, 21, 0.4],
        ['Rencana anggaran biaya (RAB)', 'Keuangan', 14, 7, 0.2],
        ['Pengadaan material & alat', 'Pengadaan', 21, 14, 0],
        ['Mobilisasi alat & tenaga kerja', 'Kontraktor', 32, 4, 0],
        ['Pelaksanaan pekerjaan tahap 1', 'Kontraktor', 35, 28, 0],
        ['Pelaksanaan pekerjaan tahap 2', 'Kontraktor', 63, 28, 0],
        ['Pengawasan & quality control', 'Quality Control', 35, 56, 0],
        ['Uji coba / commissioning', 'Tim Teknik', 91, 7, 0],
        ['Serah terima & dokumentasi', 'Manajer Proyek', 98, 5, 0],
      ],
    });
  });
}
function timelineHR() {
  return make('Timeline Rekrutmen', (wb, sh) => {
    sh.name = 'Timeline';
    timelineSheet(wb, sh, {
      title: 'Timeline Proses Rekrutmen', sub: 'Jabatan: ____________ · Perkiraan karyawan mulai: ____________', scale: 'Mingguan', start: monday(), picHead: 'Tanggung jawab',
      pic: ['HR', 'Hiring Manager', 'HR/Hiring Manager', 'Klien'],
      tasks: [
        ['Konsultasi awal', 'HR/Hiring Manager', 0, 5, 1],
        ['Pembuatan kontrak', 'HR/Hiring Manager', 0, 5, 1],
        ['Kriteria seleksi & pengumuman lowongan', 'HR/Hiring Manager', 7, 14, 0.5],
        ['Riset kampanye pemasaran / iklan', 'HR/Hiring Manager', 7, 14, 0.3],
        ['Pemasaran / outreach / sourcing', 'HR', 21, 28, 0],
        ['Penyaringan pelamar terbaik', 'HR', 28, 28, 0],
        ['Kandidat teratas diajukan ke klien', 'HR/Hiring Manager', 56, 7, 0],
        ['Jadwal & koordinasi wawancara', 'HR', 63, 14, 0],
        ['Pertanyaan wawancara & seleksi akhir', 'HR/Hiring Manager', 49, 28, 0],
        ['Wawancara tahap pertama', 'HR/Hiring Manager', 77, 7, 0],
        ['Cek latar belakang / referensi finalis', 'HR', 84, 14, 0],
        ['Wawancara akhir', 'HR/Hiring Manager', 98, 7, 0],
        ['Penawaran kerja', 'Hiring Manager', 105, 7, 0],
        ['Karyawan baru mulai bekerja', 'HR', 133, 5, 0],
      ],
      colors: { head: '#7c3aed', head2: '#8b5cf6', bar: '#f9a65a', done: '#e8710a' },
    });
  });
}

/* ---------- Kurva S ---------- */
// Bobot = biaya / total biaya; bobot tiap pekerjaan dibagi rata ke periode jalannya (Mulai, Durasi). Baris bawah:
// rencana per periode, kumulatif rencana, realisasi (isi), kumulatif realisasi, deviasi; tabel bantu + diagram garis Kurva S.
const KS_N = 24;
function kurvaS() {
  return make('Kurva S Proyek', (wb, sh) => {
    sh.name = 'Kurva S';
    const items = [
      ['Pekerjaan persiapan (mobilisasi, direksi keet, pengukuran)', 25000000, 1, 2],
      ['Pekerjaan tanah (galian & urugan)', 40000000, 2, 3],
      ['Pekerjaan pondasi', 120000000, 3, 4],
      ['Pekerjaan beton struktur (sloof, kolom, balok, plat)', 350000000, 5, 8],
      ['Pekerjaan dinding & plesteran', 150000000, 9, 6],
      ['Pekerjaan mekanikal & elektrikal', 90000000, 10, 10],
      ['Pekerjaan sanitasi & plumbing', 55000000, 11, 8],
      ['Pekerjaan rangka & penutup atap', 110000000, 12, 4],
      ['Pekerjaan kusen, pintu & jendela', 80000000, 13, 4],
      ['Pekerjaan lantai & keramik', 95000000, 15, 4],
      ['Pekerjaan plafon', 60000000, 16, 3],
      ['Pekerjaan pengecatan', 70000000, 18, 3],
      ['Pekerjaan lain-lain & pembersihan akhir', 15000000, 20, 1],
    ];
    const T = items.length, first = 6, last = first + KS_N - 1, L = colLetter;
    const r0 = 8, rl = r0 + T - 1;               // task rows (0-based)
    const rTot = rl + 1, rPlan = rl + 2, rCum = rl + 3, rReal = rl + 4, rRCum = rl + 5, rDev = rl + 6;
    const X = r => r + 1;                         // 1-based row number in formulas
    widths(sh, [34, 330, 120, 70, 62, 62]);
    for (let c = first; c <= last; c++) sh.colW.set(c, 58);
    const H2 = { b: true, color: '#ffffff', bg: '#0b5394', va: 'middle', ha: 'center', wrap: 'wrap' };
    put(wb, sh, 'A1', 'KURVA S — JADWAL PELAKSANAAN PEKERJAAN', { b: true, fs: 16, color: '#0b5394' });
    row(wb, sh, 'A2', ['', 'Nama proyek', 'Pembangunan Gedung']); put(wb, sh, 'B2', null, { b: true }); put(wb, sh, 'C2', null, { bg: '#fef7e0' });
    row(wb, sh, 'A3', ['', 'Lokasi', '—']); put(wb, sh, 'B3', null, { b: true }); put(wb, sh, 'C3', null, { bg: '#fef7e0' });
    row(wb, sh, 'A4', ['', 'Tanggal mulai', String(monday())]); put(wb, sh, 'B4', null, { b: true }); put(wb, sh, 'C4', null, { fmt: 'd mmmm yyyy', bg: '#fef7e0', ha: 'left' });
    row(wb, sh, 'A5', ['', 'Periode', 'Mingguan']); put(wb, sh, 'B5', null, { b: true }); put(wb, sh, 'C5', null, { bg: '#fef7e0', b: true, color: '#0b5394' });
    put(wb, sh, 'D5', 'pilih Mingguan / Bulanan · Mulai & Durasi dalam periode · isi baris Realisasi tiap periode', { color: '#5f6368', fs: 9 });
    sh.dv.push({ id: uid('dv'), type: 'list', ranges: [{ r1: 4, c1: 2, r2: 4, c2: 2 }], items: ['Mingguan', 'Bulanan'], colors: ['#cfe2f3', '#9fc5e8'], chip: 'arrow', strict: true });
    // header rows 7 (label) and 8 (start date of the period)
    row(wb, sh, 'A7', ['No', 'Uraian pekerjaan', 'Biaya (Rp)', 'Bobot', 'Mulai', 'Durasi'], H2);
    row(wb, sh, 'A8', ['', '', '', '(%)', '(ke-)', '(periode)'], { ...H2, fs: 8, b: false });
    for (let i = 1; i <= KS_N; i++) {
      const c = first + i - 1;
      setInput(wb, sh, 6, c, `=IF($C$5="Bulanan","Bulan ","Minggu ")&${i}`);
      setInput(wb, sh, 7, c, `=IF($C$5="Bulanan",EDATE($C$4,${i - 1}),$C$4+7*${i - 1})`);
      sh.setStyle(6, c, { ...H2, fs: 9 }); sh.setStyle(7, c, { ...H2, fs: 8, b: false, fmt: 'dd/mm/yy' });
    }
    sh.rowH.set(6, 30); sh.rowH.set(7, 22);
    items.forEach((it, k) => {
      const r = r0 + k, R = X(r);
      row(wb, sh, 'A' + R, [k + 1, it[0], it[1], `=IFERROR(C${R}/$C$${X(rTot)},0)`, it[2], it[3]]);
      for (let i = 1; i <= KS_N; i++) setInput(wb, sh, r, first + i - 1, `=IF(AND(${i}>=$E${R},${i}<=$E${R}+$F${R}-1),$D${R}/$F${R},"")`);
    });
    styleRange(sh, `A${X(r0)}:A${X(rl)}`, { ha: 'center' });
    styleRange(sh, `B${X(r0)}:B${X(rl)}`, { wrap: 'wrap' });
    styleRange(sh, `C${X(r0)}:C${X(rTot)}`, { fmt: '#,##0' });
    styleRange(sh, `D${X(r0)}:D${X(rTot)}`, { fmt: '0.00%', ha: 'right' });
    styleRange(sh, `E${X(r0)}:F${X(rl)}`, { ha: 'center', bg: '#fef7e0' });
    styleRange(sh, `${L(first)}${X(r0)}:${L(last)}${X(rl)}`, { fmt: '0.00%', fs: 8, ha: 'center', color: '#073763' });
    for (let k = 0; k < T; k++) sh.rowH.set(r0 + k, 36);
    // total + S-curve rows
    row(wb, sh, 'A' + X(rTot), ['', 'JUMLAH', `=SUM(C${X(r0)}:C${X(rl)})`, `=SUM(D${X(r0)}:D${X(rl)})`], { b: true, bt: { w: 2, style: 'solid', color: '#0b5394' } });
    const lab = { b: true, ha: 'right' };
    put(wb, sh, 'B' + X(rPlan), 'Rencana per periode', lab);
    put(wb, sh, 'B' + X(rCum), 'Kumulatif rencana', { ...lab, color: '#0b5394' });
    put(wb, sh, 'B' + X(rReal), 'Realisasi per periode (isi)', { ...lab, color: '#b45f06' });
    put(wb, sh, 'B' + X(rRCum), 'Kumulatif realisasi', { ...lab, color: '#b45f06' });
    put(wb, sh, 'B' + X(rDev), 'Deviasi (realisasi − rencana)', lab);
    const real = ['0,8%', '1,8%', '3%', '3,2%', '5%', '5,2%'];   // id-ID input: comma decimals
    for (let i = 1; i <= KS_N; i++) {
      const c = first + i - 1, C = L(c), P = L(c - 1);
      setInput(wb, sh, rPlan, c, `=SUM(${C}${X(r0)}:${C}${X(rl)})`);
      setInput(wb, sh, rCum, c, i === 1 ? `=${C}${X(rPlan)}` : `=${P}${X(rCum)}+${C}${X(rPlan)}`);
      if (real[i - 1] != null) setInput(wb, sh, rReal, c, real[i - 1]);
      setInput(wb, sh, rRCum, c, `=IF(${C}${X(rReal)}="","",SUM($${L(first)}${X(rReal)}:${C}${X(rReal)}))`);
      setInput(wb, sh, rDev, c, `=IF(${C}${X(rRCum)}="","",${C}${X(rRCum)}-${C}${X(rCum)})`);
    }
    styleRange(sh, `${L(first)}${X(rPlan)}:${L(last)}${X(rDev)}`, { fmt: '0.00%', fs: 9, ha: 'center' });
    styleRange(sh, `${L(first)}${X(rCum)}:${L(last)}${X(rCum)}`, { b: true, color: '#0b5394', bg: '#eaf1fb' });
    styleRange(sh, `${L(first)}${X(rReal)}:${L(last)}${X(rReal)}`, { bg: '#fef7e0' });
    styleRange(sh, `${L(first)}${X(rRCum)}:${L(last)}${X(rRCum)}`, { b: true, color: '#b45f06', bg: '#fdf1e6' });
    styleRange(sh, `A${X(r0)}:${L(last)}${X(rl)}`, { bb: { w: 1, style: 'solid', color: '#cfd8e3' } });
    styleRange(sh, `A${X(rPlan)}:${L(last)}${X(rDev)}`, { bb: thin });
    // bars of the schedule + deviation colours
    const g = { r1: r0, c1: first, r2: rl, c2: last };
    sh.cf.push({ id: uid('cf'), ranges: [g], type: 'formula', v1: `=ISNUMBER(${L(first)}${X(r0)})`, style: { bg: '#9fc5e8' } });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: rDev, c1: first, r2: rDev, c2: last }], type: 'num', op: 'lt', v1: '0', style: { color: '#c5221f', b: true, bg: '#fce8e6' } });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: rDev, c1: first, r2: rDev, c2: last }], type: 'num', op: 'ge', v1: '0', style: { color: '#188038', b: true } });
    sh.cf.push({ id: uid('cf'), ranges: [{ r1: 6, c1: first, r2: 7, c2: last }], type: 'formula', v1: `=AND(TODAY()>=${L(first)}$8,TODAY()<IF($C$5="Bulanan",EDATE(${L(first)}$8,1),${L(first)}$8+7))`, style: { bg: '#fbbc04', color: '#202124' } });
    // helper table (vertical) for the chart
    const h0 = rDev + 3;
    put(wb, sh, 'B' + X(h0 - 1), 'Data Kurva S (otomatis)', { b: true, color: '#0b5394' });
    row(wb, sh, 'B' + X(h0), ['Periode', 'Rencana (%)', 'Realisasi (%)'], { b: true, bg: '#cfe2f3', ha: 'center' });
    for (let i = 1; i <= KS_N; i++) {
      const C = L(first + i - 1), r = h0 + i;
      row(wb, sh, 'B' + X(r), [`=${C}7`, `=ROUND(${C}${X(rCum)}*100,2)`, `=IF(${C}${X(rRCum)}="","",ROUND(${C}${X(rRCum)}*100,2))`]);
    }
    styleRange(sh, `C${X(h0 + 1)}:D${X(h0 + KS_N)}`, { fmt: '0.00', ha: 'center' });
    styleRange(sh, `B${X(h0 + 1)}:D${X(h0 + KS_N)}`, { bb: thin });
    sh.charts.push({ id: uid('ch'), type: 'line', range: `B${X(h0)}:D${X(h0 + KS_N)}`, anchor: { r: h0 - 1, c: 5 }, dx: 0, dy: 0, w: 760, h: 380, opts: { title: 'Kurva S — kumulatif rencana vs realisasi (%)', legend: 'bottom', headers: 'on', labels: 'on', colors: ['#0b5394', '#e69138'] } });
    sh.freezeR = 8; sh.freezeC = 2;
    sh.page = { v: 2, orient: 'landscape', fit: 'width', title: false, num: true, hdr: true, auto: true, hc: false, area: { r1: 0, c1: 0, r2: rDev, c2: last } };
  });
}

export const TEMPLATES = [
  { id: 'blank', name: 'Spreadsheet kosong', icon: 'add', desc: 'Mulai dari awal', build: () => newWorkbook() },
  { id: 'welcome', name: 'Spreadsheet kosong', icon: 'add', desc: '', build: () => newWorkbook() },
  { id: 'budget', name: 'Anggaran bulanan', icon: 'account_balance_wallet', desc: 'Anggaran vs aktual, bilah data, diagram', build: budget },
  { id: 'todo', name: 'Daftar tugas', icon: 'checklist', desc: 'Kotak centang, chip status, tenggat', build: todo },
  { id: 'invoice', name: 'Faktur', icon: 'receipt_long', desc: 'Faktur dengan PPN dan terbilang', build: invoice },
  { id: 'sales', name: 'Laporan penjualan', icon: 'monitoring', desc: 'UNIQUE, SUMIF, QUERY, XLOOKUP, diagram', build: sales },
  { id: 'gantt', name: 'Jadwal proyek', icon: 'view_timeline', desc: 'Gantt dengan format bersyarat', build: gantt },
  { id: 'loan', name: 'Simulasi kredit', icon: 'calculate', desc: 'PMT, IPMT, SEQUENCE, SCAN, LAMBDA', build: loan },
  { id: 'timeline', name: 'Timeline proyek', icon: 'date_range', desc: 'Harian / mingguan / bulanan dari dropdown, bilah otomatis', build: timeline, plan: true },
  { id: 'timeline-hr', name: 'Timeline rekrutmen', icon: 'event_note', desc: 'Kegiatan + tanggung jawab per minggu', build: timelineHR, plan: true },
  { id: 'kurvas', name: 'Kurva S proyek', icon: 'stacked_line_chart', desc: 'Bobot, rencana vs realisasi, deviasi, diagram', build: kurvaS, plan: true },
];
