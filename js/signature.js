// signature.js — Tanda tangan untuk Amatu Sheets dan Amatu Docs (file yang sama di kedua aplikasi).
// Buat dengan menggambar (goresan halus mengikuti kecepatan / tekanan pena), mengetik nama dengan huruf
// tangan, atau mengunggah foto / scan tanda tangan yang sudah jadi: latarnya dihapus otomatis di perangkat
// (gambar tidak dikirim ke mana pun). Hasil = PNG transparan yang sudah dipangkas: {url, w, h}.
// Tanda tangan yang disimpan ada di IndexedDB 'amatu-ttd', dipakai bersama oleh Sheets dan Docs di domain yang sama.

const clamp = (v, a = 0, b = 255) => (v < a ? a : v > b ? b : v);
const MAX_OUT = 720;   // sisi terpanjang PNG di dokumen: ±360 dpi bila dicetak selebar 5 cm
const MAX_IN = 1600;   // foto unggahan diperkecil dulu agar pemrosesan cepat
const INK = [['#111827', 'Hitam'], ['#1d3fae', 'Biru'], ['#b91c1c', 'Merah']];
const FONTS = ['Dancing Script', 'Great Vibes', 'Sacramento', 'Allura', 'Alex Brush', 'Mr Dafoe', 'Yellowtail', 'Satisfy', 'Kaushan Script', 'Caveat'];

/* ---------------- hapus latar (sama dengan Amatu PDF) ---------------- */
function otsu(hist, n) {
  let sum = 0; for (let i = 0; i < 256; i++) sum += i * hist[i];
  let sB = 0, wB = 0, best = 0, th = 40;
  for (let i = 0; i < 256; i++) {
    wB += hist[i]; if (!wB) continue;
    const wF = n - wB; if (!wF) break;
    sB += i * hist[i];
    const mB = sB / wB, mF = (sum - sB) / wF, v = wB * wF * (mB - mF) * (mB - mF);
    if (v > best) { best = v; th = i; }
  }
  return th;
}

// Peta warna kertas: tiap blok mengambil warna kertas (terang bila latar terang, gelap bila latar gelap), lalu dihaluskan.
// Bayangan tangan / kertas yang tidak rata ikut terhapus karena tiap bagian dibandingkan dengan kertas di sekitarnya.
function backgroundMap(d, w, h, light) {
  const bs = Math.max(6, Math.round(Math.max(w, h) / 40)), gw = Math.ceil(w / bs), gh = Math.ceil(h / bs), g = new Float32Array(gw * gh * 3);
  for (let j = 0; j < gh; j++) for (let i = 0; i < gw; i++) {
    const vals = [];
    const x1 = Math.min(w, (i + 1) * bs), y1 = Math.min(h, (j + 1) * bs);
    for (let y = j * bs; y < y1; y += 2) for (let x = i * bs; x < x1; x += 2) { const o = (y * w + x) * 4; vals.push([d[o] + d[o + 1] + d[o + 2], o]); }
    vals.sort((a, b) => a[0] - b[0]);
    const k = light ? Math.floor(vals.length * .8) : Math.floor(vals.length * .2), o = vals[Math.min(vals.length - 1, k)][1];
    for (let c = 0; c < 3; c++) g[(j * gw + i) * 3 + c] = d[o + c];
  }
  // haluskan beberapa kali, jaga agar blok yang tertutup tinta tak menodai latar
  for (let it = 0; it < 3; it++) {
    const t = g.slice();
    for (let j = 0; j < gh; j++) for (let i = 0; i < gw; i++) for (let c = 0; c < 3; c++) {
      let s = 0, n = 0, m = light ? -1 : 1e9;
      for (let y = Math.max(0, j - 1); y <= Math.min(gh - 1, j + 1); y++) for (let x = Math.max(0, i - 1); x <= Math.min(gw - 1, i + 1); x++) {
        const v = t[(y * gw + x) * 3 + c]; s += v; n++;
        m = light ? Math.max(m, v) : Math.min(m, v);
      }
      g[(j * gw + i) * 3 + c] = it === 0 ? m : s / n;
    }
  }
  return (x, y, out) => {
    const gx = clamp(x / bs - .5, 0, gw - 1.001), gy = clamp(y / bs - .5, 0, gh - 1.001), i = gx | 0, j = gy | 0, fx = gx - i, fy = gy - j;
    const i1 = Math.min(gw - 1, i + 1), j1 = Math.min(gh - 1, j + 1);
    for (let c = 0; c < 3; c++) {
      const a = g[(j * gw + i) * 3 + c], b = g[(j * gw + i1) * 3 + c], e = g[(j1 * gw + i) * 3 + c], f = g[(j1 * gw + i1) * 3 + c];
      out[c] = (a * (1 - fx) + b * fx) * (1 - fy) + (e * (1 - fx) + f * fx) * fy;
    }
  };
}

/**
 * Foto tanda tangan / stempel → PNG transparan. Hanya menghitung dari piksel, berjalan di perangkat.
 * opt: {strength: -40..40, specks: 0..100, color: null|'#rrggbb'}
 * → {canvas, inkShare, hasAlpha} sudah dipangkas.
 */
export function removeBackground(src, opt = {}) {
  const w = src.width, h = src.height, sx = src.getContext('2d', { willReadFrequently: true }), im = sx.getImageData(0, 0, w, h), d = im.data;
  let transp = 0;
  for (let i = 3; i < d.length; i += 28) if (d[i] < 200) transp++;
  const hasAlpha = transp > d.length / 28 * .02;
  const out = new ImageData(w, h), o = out.data, A = new Uint8ClampedArray(w * h);
  if (hasAlpha) {
    o.set(d);
    for (let i = 0; i < w * h; i++) A[i] = d[i * 4 + 3];
  } else {
    // latar terang atau gelap? lihat tepi gambar
    let edge = 0, en = 0;
    for (let x = 0; x < w; x += 3) for (const y of [0, h - 1]) { const q = (y * w + x) * 4; edge += d[q] + d[q + 1] + d[q + 2]; en++; }
    for (let y = 0; y < h; y += 3) for (const x of [0, w - 1]) { const q = (y * w + x) * 4; edge += d[q] + d[q + 1] + d[q + 2]; en++; }
    const light = edge / en / 3 > 110;
    const bg = backgroundMap(d, w, h, light), b = [0, 0, 0], raw = new Float32Array(w * h), hist = new Uint32Array(256);
    for (let y = 0, i = 0; y < h; y++) for (let x = 0; x < w; x++, i++) {
      bg(x, y, b);
      const q = i * 4;
      // seberapa jauh dari kertas: kanal paling berbeda (tinta biru/merah menonjol di kanal lawannya)
      let m = 0;
      for (let c = 0; c < 3; c++) {
        const v = light ? (b[c] - d[q + c]) / Math.max(30, b[c]) : (d[q + c] - b[c]) / Math.max(30, 255 - b[c]);
        if (v > m) m = v;
      }
      raw[i] = m; hist[Math.round(clamp(m * 255))]++;
    }
    const th = otsu(hist, w * h) / 255, adj = (opt.strength || 0) / 100;
    // lo = batas derau kertas, hi = sudah tinta penuh
    const lo = clamp(th * .55 - adj * .5, .035, .6), hi = clamp(th * 1.35 - adj * .3, lo + .06, .95);
    for (let y = 0, i = 0; y < h; y++) for (let x = 0; x < w; x++, i++) {
      let a = clamp((raw[i] - lo) / (hi - lo), 0, 1); a = a * a * (3 - 2 * a);
      A[i] = a * 255;
      if (a > 0) {
        bg(x, y, b);
        const q = i * 4, inv = 1 / Math.max(.25, a);
        // warna tinta murni: buang campuran kertas dari tepi yang setengah transparan
        for (let c = 0; c < 3; c++) o[q + c] = clamp(b[c] + (d[q + c] - b[c]) * inv);
      }
    }
  }
  // buang bintik & bayangan tepi: komponen terhubung
  const lab = new Int32Array(w * h), comps = [], st = [];
  for (let s = 0; s < w * h; s++) {
    if (lab[s] || A[s] < 64) continue;
    const id = comps.length + 1, c = { id, n: 0, x0: w, y0: h, x1: 0, y1: 0, edge: 0 };
    lab[s] = id; st.push(s);
    while (st.length) {
      const i = st.pop(), x = i % w, y = i / w | 0;
      c.n++;
      if (x < c.x0) c.x0 = x; if (x > c.x1) c.x1 = x; if (y < c.y0) c.y0 = y; if (y > c.y1) c.y1 = y;
      if (!x || !y || x === w - 1 || y === h - 1) c.edge++;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const nx = x + dx, ny = y + dy;
        if (nx < 0 || ny < 0 || nx >= w || ny >= h) continue;
        const k = ny * w + nx;
        if (!lab[k] && A[k] >= 64) { lab[k] = id; st.push(k); }
      }
    }
    comps.push(c);
  }
  const total = comps.reduce((s, c) => s + c.n, 0) || 1, sp = opt.specks == null ? 50 : opt.specks;
  const minArea = Math.max(2, total * .0012 * sp / 50), keep = new Uint8Array(comps.length + 1);
  let maxN = 0; for (const c of comps) if (c.n > maxN) maxN = c.n;
  for (const c of comps) {
    const bw = c.x1 - c.x0 + 1, bh = c.y1 - c.y0 + 1, fill = c.n / (bw * bh);
    const shadow = c.edge > 8 && (c.edge > (bw + bh) * .3 || fill > .55) && c.n !== maxN;     // bayangan/pinggir kertas menempel tepi foto
    keep[c.id] = (c.n >= minArea && !shadow) ? 1 : 0;
  }
  let x0 = w, y0 = h, x1 = -1, y1 = -1, ink = 0;
  const fix = opt.color ? [parseInt(opt.color.slice(1, 3), 16), parseInt(opt.color.slice(3, 5), 16), parseInt(opt.color.slice(5, 7), 16)] : null;
  for (let y = 0, i = 0; y < h; y++) for (let x = 0; x < w; x++, i++) {
    let a = A[i];
    if (a) {
      // tepi halus (alpha < 64) ikut komponen tetangganya
      let on = lab[i] ? keep[lab[i]] : 0;
      if (!lab[i]) for (let dy = -2; dy <= 2 && !on; dy++) for (let dx = -2; dx <= 2 && !on; dx++) {
        const nx = x + dx, ny = y + dy;
        if (nx >= 0 && ny >= 0 && nx < w && ny < h && lab[ny * w + nx] && keep[lab[ny * w + nx]]) on = 1;
      }
      if (!on) a = 0;
    }
    const q = i * 4;
    o[q + 3] = a;
    if (a) {
      if (fix) { o[q] = fix[0]; o[q + 1] = fix[1]; o[q + 2] = fix[2]; }
      if (a > 30) { ink++; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    }
  }
  if (x1 < 0) { x0 = y0 = 0; x1 = w - 1; y1 = h - 1; }
  const pad = Math.round(Math.max(w, h) * .012) + 1;
  x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(w - 1, x1 + pad); y1 = Math.min(h - 1, y1 + pad);
  const c = document.createElement('canvas'); c.width = x1 - x0 + 1; c.height = y1 - y0 + 1;
  c.getContext('2d').putImageData(out, -x0, -y0);
  return { canvas: c, inkShare: ink / (w * h), hasAlpha };
}

/* ---------------- kanvas ---------------- */
const loadImage = src => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('img')); i.src = src; });

// <img> menerapkan orientasi EXIF foto HP, jadi tanda tangan tidak miring/terbalik.
export async function fileToCanvas(file, max = MAX_IN) {
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url);
    const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(img.naturalWidth * k)); c.height = Math.max(1, Math.round(img.naturalHeight * k));
    const x = c.getContext('2d', { willReadFrequently: true }); x.imageSmoothingQuality = 'high'; x.drawImage(img, 0, 0, c.width, c.height);
    return c;
  } finally { URL.revokeObjectURL(url); }
}

export function trimCanvas(src, pad = 6) {
  const w = src.width, h = src.height, d = src.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, w, h).data;
  let x0 = w, y0 = h, x1 = -1, y1 = -1;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (d[(y * w + x) * 4 + 3] > 10) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  if (x1 < 0) return null;
  x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(w - 1, x1 + pad); y1 = Math.min(h - 1, y1 + pad);
  const c = document.createElement('canvas'); c.width = x1 - x0 + 1; c.height = y1 - y0 + 1;
  c.getContext('2d').drawImage(src, -x0, -y0);
  return c;
}

// Pangkas + perkecil bertahap (separuh demi separuh, tetap tajam) → {url, w, h}
function finish(src) {
  let c = trimCanvas(src, 4) || src;
  const target = Math.min(1, MAX_OUT / Math.max(c.width, c.height));
  if (target < 1) {
    let w = c.width, h = c.height;
    const tw = Math.max(1, Math.round(c.width * target)), th = Math.max(1, Math.round(c.height * target));
    while (w > tw || h > th) {
      w = Math.max(tw, Math.round(w / 2)); h = Math.max(th, Math.round(h / 2));
      const t = document.createElement('canvas'); t.width = w; t.height = h;
      const x = t.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(c, 0, 0, w, h);
      c = t;
    }
  }
  return { url: c.toDataURL('image/png'), w: c.width, h: c.height };
}

/* ---------------- papan gambar ---------------- */
// Goresan alami: lebar mengikuti kecepatan (atau tekanan stylus), dihaluskan kurva.
// Koordinat dalam piksel tata letak kanvas, jadi tetap tepat di dalam dialog yang dianimasikan / halaman yang di-zoom.
export class Pad {
  constructor(canvas, onChange) {
    this.c = canvas; this.x = canvas.getContext('2d'); this.color = INK[0][0]; this.size = 3; this.strokes = []; this.onChange = onChange || (() => { });
    canvas.style.touchAction = 'none';
    canvas.addEventListener('pointerdown', e => this.down(e));
    canvas.addEventListener('pointermove', e => this.move(e));
    canvas.addEventListener('pointerup', e => this.up(e));
    canvas.addEventListener('pointercancel', e => this.up(e));
  }
  resize() {
    const w = this.c.offsetWidth, h = this.c.offsetHeight, k = Math.max(2, devicePixelRatio || 1);
    if (!w || !h) return;
    if (this.k === k && this.c.width === Math.round(w * k) && this.c.height === Math.round(h * k)) return;
    this.k = k; this.c.width = Math.round(w * k); this.c.height = Math.round(h * k);
    this.redraw();
  }
  pos(e) {
    const r = this.c.getBoundingClientRect(), sx = r.width ? this.c.offsetWidth / r.width : 1, sy = r.height ? this.c.offsetHeight / r.height : 1;
    return { x: (e.clientX - r.left) * sx, y: (e.clientY - r.top) * sy, t: e.timeStamp, p: e.pointerType === 'pen' ? e.pressure : 0 };
  }
  down(e) {
    if (e.button > 0) return;
    e.preventDefault();
    try { this.c.setPointerCapture(e.pointerId); } catch { /* sudah dilepas */ }
    this.resize();
    const p = this.pos(e); p.w = this.size;
    this.cur = { color: this.color, size: this.size, pts: [p] }; this.strokes.push(this.cur); this.dot(this.cur, p);
  }
  move(e) {
    if (!this.cur) return;
    const list = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
    for (const ev of (list.length ? list : [e])) {
      const s = this.cur, p = this.pos(ev), a = s.pts[s.pts.length - 1], dist = Math.hypot(p.x - a.x, p.y - a.y);
      if (dist < 1) continue;
      const v = dist / Math.max(1, p.t - a.t), target = p.p ? s.size * (.4 + p.p * 1.2) : s.size * clamp(1.5 - v * .55, .45, 1.5);
      p.w = a.w + (target - a.w) * .35;
      s.pts.push(p); this.seg(s, s.pts.length - 1);
    }
  }
  up() { if (this.cur) { this.cur = null; this.onChange(); } }
  dot(s, p) { const x = this.x, k = this.k || 2; x.fillStyle = s.color; x.beginPath(); x.arc(p.x * k, p.y * k, p.w * k / 2, 0, 7); x.fill(); }
  seg(s, i) {
    const x = this.x, k = this.k || 2, P = s.pts, a = P[Math.max(0, i - 2)], b = P[i - 1], c = P[i];
    const m1 = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }, m2 = { x: (b.x + c.x) / 2, y: (b.y + c.y) / 2 };
    x.strokeStyle = s.color; x.lineCap = 'round'; x.lineJoin = 'round'; x.lineWidth = (b.w + c.w) / 2 * k;
    x.beginPath(); x.moveTo(m1.x * k, m1.y * k); x.quadraticCurveTo(b.x * k, b.y * k, m2.x * k, m2.y * k); x.stroke();
  }
  redraw() {
    this.x.clearRect(0, 0, this.c.width, this.c.height);
    for (const s of this.strokes) { this.dot(s, s.pts[0]); for (let i = 1; i < s.pts.length; i++) this.seg(s, i); }
  }
  undo() { this.strokes.pop(); this.redraw(); this.onChange(); }
  clear() { this.strokes = []; this.redraw(); this.onChange(); }
  recolor(c) { this.color = c; for (const s of this.strokes) s.color = c; this.redraw(); this.onChange(); }
  get empty() { return !this.strokes.length; }
  output() { return trimCanvas(this.c, 8); }
}

/* ---------------- simpanan (IndexedDB) ---------------- */
const DBN = 'amatu-ttd', ST = 'sigs', KEEP_MAX = 24;
let dbp = null;
function idb() {
  if (!dbp) {
    dbp = new Promise((res, rej) => {
      const r = indexedDB.open(DBN, 1);
      r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains(ST)) r.result.createObjectStore(ST, { keyPath: 'id' }); };
      r.onsuccess = () => { const db = r.result; db.onversionchange = () => { db.close(); dbp = null; }; res(db); };
      r.onerror = () => rej(r.error);
      r.onblocked = () => rej(new Error('blocked'));
    });
    dbp.catch(() => { dbp = null; });
  }
  return dbp;
}
async function run(mode, fn) {
  const db = await idb();
  return new Promise((res, rej) => {
    const t = db.transaction(ST, mode), q = fn(t.objectStore(ST));
    t.oncomplete = () => res(q && 'result' in q ? q.result : undefined);
    t.onerror = () => rej(t.error); t.onabort = () => rej(t.error);
  });
}
export const sigStore = {
  async all() { return ((await run('readonly', s => s.getAll())) || []).filter(r => r && r.url).sort((a, b) => b.at - a.at); },
  async add(sig) {
    const rows = await this.all();
    if (rows.some(r => r.url === sig.url)) return;
    await run('readwrite', s => {
      s.put({ id: 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), url: sig.url, w: sig.w, h: sig.h, at: Date.now() });
      for (const r of rows.slice(KEEP_MAX - 1)) s.delete(r.id);
    });
  },
  del(id) { return run('readwrite', s => { s.delete(id); }); },
};

/* ---------------- tampilan ---------------- */
function h(tag, cls, attrs, ...kids) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k === 'style') Object.assign(e.style, v);
    else if (k in e && k !== 'list' && typeof v !== 'string') e[k] = v;
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const k of kids.flat()) if (k != null && k !== false) e.append(k);
  return e;
}
const ms = name => h('span', 'ms', { 'aria-hidden': 'true' }, name);

let fontsLoaded = false;
function loadFonts() {
  if (fontsLoaded) return; fontsLoaded = true;
  const l = document.createElement('link'); l.rel = 'stylesheet';
  l.href = 'https://fonts.googleapis.com/css2?' + FONTS.map(f => 'family=' + f.replace(/ /g, '+')).join('&') + '&display=swap';
  document.head.append(l);
}

const CSS = `
.sgp{display:flex;flex-direction:column;gap:12px;min-width:0;--sg-line:var(--line-2,var(--line2,rgba(0,0,0,.16)));--sg-hov:var(--hover,var(--hov,rgba(0,0,0,.05)));--sg-tx2:var(--text-2,var(--tx2,#5f6368));--sg-acc:var(--accent,#1a73e8)}
.sgp .ms{line-height:1}
.sgp [hidden]{display:none!important}
.sgp-res{display:flex;flex-direction:column;gap:10px}
.sgp-lbl{font-size:12px;font-weight:600;color:var(--sg-tx2);margin:0 0 6px}
.sgp-saved{display:flex;gap:8px;overflow-x:auto;padding:2px 2px 6px;scrollbar-width:thin}
.sgp-sv{position:relative;flex:none;width:132px;height:62px;border:1px solid var(--sg-line);border-radius:10px;background:#fff;display:grid;place-items:center;padding:6px 10px;cursor:pointer;outline:none;transition:border-color .12s,box-shadow .12s}
.sgp-sv:hover,.sgp-sv:focus-visible{border-color:var(--sg-acc);box-shadow:0 0 0 2px color-mix(in srgb,var(--sg-acc) 25%,transparent)}
.sgp-sv img{max-width:100%;max-height:100%;display:block;pointer-events:none}
.sgp-del{position:absolute;top:3px;right:3px;width:22px;height:22px;border:0;border-radius:6px;background:rgba(16,24,40,.07);color:#475467;display:grid;place-items:center;cursor:pointer;opacity:0;padding:0}
.sgp-sv:hover .sgp-del,.sgp-del:focus-visible{opacity:1}
.sgp-del:hover{background:#fee4e2;color:#d92d20}
.sgp-del .ms{font-size:15px}
.sgp-tabs{display:inline-flex;align-self:flex-start;gap:2px;padding:3px;border-radius:11px;background:var(--sg-hov);max-width:100%;overflow-x:auto}
.sgp-tabs button{border:0;background:transparent;height:32px;padding:0 12px;border-radius:8px;color:var(--sg-tx2);font:inherit;font-size:13px;font-weight:500;display:inline-flex;align-items:center;gap:6px;cursor:pointer;white-space:nowrap}
.sgp-tabs button .ms{font-size:18px}
.sgp-tabs button.on{background:var(--glass-strong,var(--glass3,#fff));color:inherit;box-shadow:0 1px 3px rgba(0,0,0,.14)}
.sgp-body{display:flex;flex-direction:column;gap:10px;min-height:268px}
.sgp-padwrap{position:relative;border:1px solid var(--sg-line);border-radius:12px;background:#fff;overflow:hidden}
.sgp-pad{position:relative;z-index:1;display:block;width:100%;height:220px;touch-action:none;cursor:crosshair}
.sgp-line{position:absolute;left:34px;right:28px;bottom:58px;border-bottom:1.5px dashed #c7ccd4;pointer-events:none}
.sgp-x{position:absolute;left:14px;bottom:56px;font:600 17px/1 Arial,sans-serif;color:#a3a9b3;pointer-events:none}
.sgp-hint{position:absolute;left:0;right:0;top:34%;text-align:center;color:#a3a9b3;font-size:14px;pointer-events:none}
.sgp-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.sgp-sp{flex:1}
.sgp-mut{font-size:12px;color:var(--sg-tx2)}
.sgp-ink{width:24px;height:24px;border-radius:50%;border:2px solid #fff;box-shadow:0 0 0 1px var(--sg-line);cursor:pointer;padding:0;flex:none}
.sgp-ink.on{box-shadow:0 0 0 2px var(--sg-acc)}
.sgp-btn{height:30px;padding:0 11px;border:1px solid var(--sg-line);border-radius:8px;background:transparent;color:inherit;font:inherit;font-size:13px;display:inline-flex;align-items:center;gap:6px;cursor:pointer;white-space:nowrap}
.sgp-btn:hover{background:var(--sg-hov)}
.sgp-btn .ms{font-size:17px}
.sgp-chip{height:28px;padding:0 11px;border:1px solid var(--sg-line);border-radius:14px;background:transparent;color:inherit;font:inherit;font-size:12px;cursor:pointer}
.sgp-chip.on{border-color:var(--sg-acc);background:color-mix(in srgb,var(--sg-acc) 12%,transparent);color:var(--sg-acc)}
.sgp-size{width:110px;accent-color:var(--sg-acc)}
.sgp-name{width:100%;height:40px;padding:0 12px;border:1px solid var(--sg-line);border-radius:10px;background:var(--inp-bg,var(--field,#fff));color:inherit;font:inherit;font-size:15px;box-sizing:border-box}
.sgp-name:focus{outline:2px solid var(--sg-acc);outline-offset:-1px}
.sgp-fonts{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;max-height:230px;overflow:auto;padding:2px}
.sgp-fonts button{height:64px;border:1px solid var(--sg-line);border-radius:10px;background:#fff;font-size:30px;line-height:1.15;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;padding:0 10px;cursor:pointer}
.sgp-fonts button.on{border-color:var(--sg-acc);box-shadow:0 0 0 2px color-mix(in srgb,var(--sg-acc) 30%,transparent)}
.sgp-drop{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;min-height:220px;box-sizing:border-box;text-align:center;padding:24px 16px;border:1.5px dashed var(--sg-line);border-radius:12px;cursor:pointer;color:var(--sg-tx2);outline:none;font-size:13px;line-height:1.45}
.sgp-drop:hover,.sgp-drop:focus-visible,.sgp-drop.drag{border-color:var(--sg-acc);background:color-mix(in srgb,var(--sg-acc) 6%,transparent)}
.sgp-drop .ms{font-size:36px;color:var(--sg-acc)}
.sgp-drop b{font-size:14px;font-weight:600;color:var(--text,var(--tx,inherit))}
.sgp-prev{position:relative;height:220px;border:1px solid var(--sg-line);border-radius:12px;display:grid;place-items:center;padding:12px;overflow:hidden;box-sizing:border-box;background-color:#fff;background-image:linear-gradient(45deg,#e8eaee 25%,transparent 25%,transparent 75%,#e8eaee 75%),linear-gradient(45deg,#e8eaee 25%,transparent 25%,transparent 75%,#e8eaee 75%);background-size:16px 16px;background-position:0 0,8px 8px}
.sgp-prev canvas{max-width:100%;max-height:194px;display:block}
.sgp-before{position:absolute;left:8px;bottom:8px;width:84px;height:56px;object-fit:cover;border-radius:6px;border:2px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.25)}
.sgp-cap{position:absolute;right:8px;bottom:8px;font-size:11px;background:rgba(255,255,255,.88);color:#475467;padding:3px 8px;border-radius:10px}
.sgp-busy{position:absolute;inset:0;display:grid;place-items:center;background:rgba(255,255,255,.6);color:#475467;font-size:13px}
.sgp-slider{display:flex;align-items:center;gap:8px;flex:1 1 170px;min-width:150px}
.sgp-slider input{flex:1;min-width:0;accent-color:var(--sg-acc)}
.sgp-chk{display:inline-flex;align-items:center;gap:8px;font-size:13px;cursor:pointer;user-select:none}
.sgp-chk input{accent-color:var(--sg-acc);width:16px;height:16px;margin:0}
.dlg:has(.sgp){min-width:0}
.dlg:has(.sgp) .btn:disabled{opacity:.45;cursor:default;box-shadow:none;transform:none;filter:none}
@media (pointer:coarse){.sgp-del{opacity:1}}
@media (max-width:560px){.sgp-pad,.sgp-drop{height:180px;min-height:180px}.sgp-prev{height:180px}.sgp-prev canvas{max-height:154px}.sgp-fonts{grid-template-columns:1fr}.sgp-body{min-height:228px}}
`;
function injectCss() {
  if (document.getElementById('sgp-css')) return;
  document.head.append(h('style', null, { id: 'sgp-css' }, CSS));
}

/**
 * Isi dialog tanda tangan. Aplikasi membungkusnya dengan dialognya sendiri.
 * onReady(ok) — tombol Sisipkan aktif/nonaktif; onPick(sig) — tanda tangan tersimpan diklik (pakai langsung).
 * → {el, result(): Promise<{url,w,h}|null>, focus(), destroy()}
 */
export function signaturePanel({ toast = () => { }, onReady = () => { }, onPick = null, tab: tab0 = 'draw' } = {}) {
  injectCss();
  let tab = null, getCanvas = () => null, ready = false, stopTab = () => { };
  const setReady = ok => { ready = !!ok; onReady(ready); };
  const savedEl = h('div');
  const tabsEl = h('div', 'sgp-tabs', { role: 'tablist' });
  const body = h('div', 'sgp-body');
  const keep = h('input', null, { type: 'checkbox', checked: true });
  const root = h('div', 'sgp', null, savedEl, tabsEl, body,
    h('div', 'sgp-row', null, h('label', 'sgp-chk', null, keep, 'Simpan untuk dipakai lagi'), h('span', 'sgp-sp'), h('span', 'sgp-mut', null, 'Diproses di perangkat ini')));
  const TABS = [['draw', 'draw', 'Gambar'], ['type', 'keyboard', 'Ketik'], ['upload', 'add_photo_alternate', 'Unggah foto']];
  for (const [id, ic, label] of TABS) tabsEl.append(h('button', null, { type: 'button', role: 'tab', 'data-tab': id, onclick: () => go(id) }, ms(ic), label));

  async function renderSaved() {
    let rows = [];
    try { rows = await sigStore.all(); } catch { /* IndexedDB tidak tersedia */ }
    savedEl.replaceChildren();
    if (!rows.length || !onPick) return;
    const list = h('div', 'sgp-saved');
    for (const r of rows) {
      const del = h('button', 'sgp-del', { type: 'button', title: 'Hapus dari simpanan', 'aria-label': 'Hapus dari simpanan' }, ms('close'));
      const tile = h('div', 'sgp-sv', { role: 'button', tabIndex: 0, title: 'Pakai tanda tangan ini' }, h('img', null, { src: r.url, alt: 'Tanda tangan tersimpan' }), del);
      const use = () => onPick({ url: r.url, w: r.w, h: r.h, saved: true });
      tile.addEventListener('click', e => { if (!e.target.closest('.sgp-del')) use(); });
      tile.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); use(); } });
      del.addEventListener('click', async e => { e.stopPropagation(); try { await sigStore.del(r.id); } catch { /* biarkan */ } renderSaved(); toast('Dihapus dari simpanan. Yang sudah ditempel tetap ada.'); });
      list.append(tile);
    }
    savedEl.append(h('div', 'sgp-lbl', null, 'Tersimpan — klik untuk langsung memakai'), list);
  }

  function go(t) {
    if (t === tab) return;
    stopTab(); stopTab = () => { };
    tab = t;
    for (const b of tabsEl.children) { const on = b.dataset.tab === t; b.classList.toggle('on', on); b.setAttribute('aria-selected', on); }
    body.replaceChildren(); setReady(false); getCanvas = () => null;
    if (t === 'draw') tabDraw(); else if (t === 'type') tabType(); else tabUpload();
  }

  function inkRow(cur, pick) {
    return INK.map(([c, n]) => h('button', 'sgp-ink' + (c === cur ? ' on' : ''), { type: 'button', title: n, 'aria-label': 'Tinta ' + n.toLowerCase(), style: { background: c }, 'data-ink': c, onclick: e => { for (const b of e.currentTarget.parentNode.querySelectorAll('.sgp-ink')) b.classList.toggle('on', b === e.currentTarget); pick(c); } }));
  }

  function tabDraw() {
    const cv = h('canvas', 'sgp-pad', { 'aria-label': 'Area tanda tangan' });
    const hint = h('div', 'sgp-hint', null, 'Tanda tangan di sini dengan mouse, jari, atau pena');
    const pad = new Pad(cv, () => { hint.hidden = !pad.empty; setReady(!pad.empty); });
    const size = h('input', 'sgp-size', { type: 'range', min: 1, max: 7, step: .5, value: 3, 'aria-label': 'Tebal goresan', oninput: e => { pad.size = +e.target.value; } });
    body.append(h('div', 'sgp-padwrap', null, h('div', 'sgp-line'), h('div', 'sgp-x', null, '×'), hint, cv),
      h('div', 'sgp-row', null, ...inkRow(pad.color, c => pad.recolor(c)), h('span', 'sgp-mut', { style: { marginLeft: '6px' } }, 'Tebal'), size, h('span', 'sgp-sp'),
        h('button', 'sgp-btn', { type: 'button', onclick: () => pad.undo() }, ms('undo'), 'Urungkan'),
        h('button', 'sgp-btn', { type: 'button', onclick: () => pad.clear() }, ms('ink_eraser'), 'Bersihkan')));
    const ro = new ResizeObserver(() => pad.resize());
    ro.observe(cv);
    stopTab = () => ro.disconnect();
    getCanvas = () => pad.output();
  }

  function tabType() {
    loadFonts();
    let font = FONTS[0], color = INK[0][0];
    const name = h('input', 'sgp-name', { type: 'text', placeholder: 'Ketik nama Anda', autocomplete: 'name', spellcheck: false, 'aria-label': 'Nama' });
    const grid = h('div', 'sgp-fonts');
    const paint = () => {
      const t = name.value.trim() || 'Nama Anda';
      grid.replaceChildren(...FONTS.map(f => h('button', f === font ? 'on' : '', { type: 'button', title: f, style: { fontFamily: `'${f}', cursive`, color }, onclick: () => { font = f; paint(); } }, t)));
      setReady(!!name.value.trim());
    };
    name.addEventListener('input', paint);
    body.append(name, h('div', 'sgp-row', null, h('span', 'sgp-mut', null, 'Warna tinta'), ...inkRow(color, c => { color = c; paint(); })), grid);
    paint();
    setTimeout(() => name.focus(), 40);
    getCanvas = async () => {
      const t = name.value.trim(), px = 150;
      if (!t) return null;
      await document.fonts.load(`${px}px "${font}"`, t).catch(() => { });
      const c = document.createElement('canvas'), m = c.getContext('2d');
      m.font = `${px}px "${font}", cursive`;
      c.width = Math.ceil(m.measureText(t).width + px * 1.2); c.height = Math.ceil(px * 2);
      const x = c.getContext('2d'); x.font = `${px}px "${font}", cursive`; x.fillStyle = color; x.textBaseline = 'middle'; x.fillText(t, px * .6, c.height / 2);
      return trimCanvas(c, 6);
    };
  }

  function tabUpload() {
    let src = null, out = null, srcUrl = '', strength = 0, specks = 50, color = null, removeBg = true, timer = 0, seq = 0;
    const file = h('input', null, { type: 'file', accept: 'image/*', hidden: true, onchange: () => { if (file.files[0]) load(file.files[0]); file.value = ''; } });
    const drop = h('div', 'sgp-drop', { tabIndex: 0, role: 'button', onclick: () => file.click(), onkeydown: e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); file.click(); } } },
      ms('add_photo_alternate'), h('b', null, null, 'Pilih foto atau scan tanda tangan'),
      h('span', null, null, 'Seret & lepas gambar ke sini, atau tempel dengan Ctrl/⌘+V.'),
      h('span', null, null, 'Paling bagus: tinta gelap di kertas polos. Latar dihapus otomatis.'));
    const prev = h('div', 'sgp-prev');
    const res = h('div', 'sgp-res', { hidden: true });
    const chips = h('span', 'sgp-row', null, ...[['', 'Asli'], ...INK].map(([c, n]) => h('button', 'sgp-chip' + (c === '' ? ' on' : ''), { type: 'button', onclick: e => { color = c || null; for (const b of chips.children) b.classList.toggle('on', b === e.currentTarget); later(0); } }, n)));
    const bgChk = h('input', null, { type: 'checkbox', checked: true, onchange: () => { removeBg = bgChk.checked; for (const x of [tune, chips]) { x.style.opacity = removeBg ? '' : '.45'; x.style.pointerEvents = removeBg ? '' : 'none'; } later(0); } });
    const tune = h('div', 'sgp-row', null,
      h('label', 'sgp-slider', null, h('span', 'sgp-mut', null, 'Kekuatan'), h('input', null, { type: 'range', min: -40, max: 40, value: 0, 'aria-label': 'Kekuatan hapus latar', oninput: e => { strength = +e.target.value; later(); } })),
      h('label', 'sgp-slider', null, h('span', 'sgp-mut', null, 'Bintik'), h('input', null, { type: 'range', min: 0, max: 100, value: 50, 'aria-label': 'Buang bintik', oninput: e => { specks = +e.target.value; later(); } })));
    res.append(prev,
      h('div', 'sgp-row', null, h('span', 'sgp-mut', null, 'Warna'), chips, h('span', 'sgp-sp'), h('button', 'sgp-btn', { type: 'button', onclick: () => file.click() }, ms('photo_library'), 'Ganti foto')),
      h('div', 'sgp-row', null, h('label', 'sgp-chk', null, bgChk, 'Hapus latar otomatis')), tune);
    body.append(file, drop, res);

    function later(ms_ = 140) { clearTimeout(timer); timer = setTimeout(process, ms_); }
    function process() {
      if (!src) return;
      const my = ++seq;
      const busy = h('div', 'sgp-busy', null, 'Menghapus latar…');
      prev.append(busy);
      setTimeout(() => {
        if (my !== seq) return;
        let r;
        try {
          r = removeBg ? removeBackground(src, { strength, specks, color }) : { canvas: src, inkShare: 1 };
        } catch (e) { console.error(e); busy.remove(); toast('Gambar tidak dapat diproses'); return; }
        out = r.canvas;
        const before = h('img', 'sgp-before', { src: srcUrl, alt: 'Foto asli', title: 'Foto asli' });
        prev.replaceChildren(out, before, h('span', 'sgp-cap', null, removeBg ? 'Latar dihapus' : 'Tanpa hapus latar'));
        setReady(r.inkShare > 0);
        if (!r.inkShare) toast('Tinta tidak terdeteksi. Geser Kekuatan ke kanan.');
      }, 16);
    }
    async function load(f) {
      if (!f || !/^image\//.test(f.type || 'image/')) { toast('Pilih file gambar (JPG, PNG, HEIC dari HP)'); return; }
      try {
        src = await fileToCanvas(f, MAX_IN);
        if (srcUrl) URL.revokeObjectURL(srcUrl);
        srcUrl = URL.createObjectURL(f);
        drop.hidden = true; res.hidden = false;
        later(0);
      } catch { toast('Gambar tidak dapat dibaca. Coba simpan sebagai JPG atau PNG.'); }
    }
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('drag'); const f = [...(e.dataTransfer?.files || [])].find(x => /^image\//.test(x.type)); if (f) load(f); });
    root._load = load;
    setTimeout(() => drop.focus({ preventScroll: true }), 40);
    getCanvas = () => out;
    stopTab = () => { clearTimeout(timer); seq++; if (srcUrl) URL.revokeObjectURL(srcUrl); root._load = null; };
  }

  // Gambar yang ditempel (Ctrl/⌘+V) di mana pun dalam dialog → langsung ke tab Unggah.
  root.addEventListener('paste', e => {
    const f = [...(e.clipboardData?.files || [])].find(x => /^image\//.test(x.type));
    if (!f) return;
    e.preventDefault(); e.stopPropagation();
    go('upload');
    root._load?.(f);
  });
  root.addEventListener('dragover', e => { if ([...(e.dataTransfer?.types || [])].includes('Files')) e.preventDefault(); });
  root.addEventListener('drop', e => {
    const f = [...(e.dataTransfer?.files || [])].find(x => /^image\//.test(x.type));
    if (!f) return;
    e.preventDefault();
    go('upload'); root._load?.(f);
  });

  renderSaved();
  go(TABS.some(t => t[0] === tab0) ? tab0 : 'draw');

  return {
    el: root,
    async result() {
      if (!ready) return null;
      const c = await getCanvas();
      if (!c) return null;
      const sig = finish(c);
      if (keep.checked) await sigStore.add(sig).catch(() => { });
      return sig;
    },
    focus() { (tabsEl.querySelector('.on') || tabsEl.firstChild)?.focus({ preventScroll: true }); },
    destroy() { stopTab(); stopTab = () => { }; },
  };
}

// Ukuran tampil awal (px) untuk tanda tangan setinggi ±3 baris, tidak terlalu lebar.
export function defaultSize(sig, height = 64, maxW = 230) {
  const ar = sig.w && sig.h ? sig.w / sig.h : 3;
  let w = height * ar, hh = height;
  if (w > maxW) { w = maxW; hh = w / ar; }
  return { w: Math.round(w), h: Math.round(hh) };
}
