// Bagian murni (tanpa impor) dari gambar mengambang: dipakai cetak (io.js / pages.js) tanpa siklus modul.
// ---------- cetak / HTML ----------
// Gambar yang terlihat pada rentang halaman g, ditempel ke sel <td> tempat sudut kiri-atasnya berada
// (baris tabel cetak bisa lebih tinggi dari di layar, jadi posisi dihitung dari sel itu, bukan dari atas tabel).
// → Map key(r,c) → [{src, x, y, w, h, z}] dengan x,y relatif ke sudut kiri-atas sel; merged/skip dari sheetHTML.
export function imagesForPrint(sh, g, merged, skip, key) {
  const out = new Map();
  const list = sh.images || [];
  if (!list.length) return out;
  const visR = [], visC = [];
  for (let r = g.r1; r <= g.r2; r++) if (!sh.isRowHidden(r)) visR.push(r);
  for (let c = g.c1; c <= g.c2; c++) if (!sh.hiddenC.has(c)) visC.push(c);
  if (!visR.length || !visC.length) return out;
  const H = (a, b) => { let s = 0; for (let r = a; r < b; r++) s += sh.rowHeight(r); return s; };   // tinggi baris a..b-1
  const W = (a, b) => { let s = 0; for (let c = a; c < b; c++) s += sh.colWidth(c); return s; };
  const pageH = H(g.r1, g.r2 + 1), pageW = W(g.c1, g.c2 + 1);
  list.forEach((d, z) => {
    const ar = Math.min(d.anchor.r, sh.rows - 1), ac = Math.min(d.anchor.c, sh.cols - 1);
    // posisi relatif ke sudut kiri-atas halaman
    const top = (ar >= g.r1 ? H(g.r1, ar) : -H(ar, g.r1)) + d.dy, left = (ac >= g.c1 ? W(g.c1, ac) : -W(ac, g.c1)) + d.dx;
    if (top >= pageH || top + d.h <= 0 || left >= pageW || left + d.w <= 0) return;
    // sel tampil pertama di/atau setelah jangkar (atau awal halaman)
    let r = visR.find(x => x >= Math.max(ar, g.r1)), c = visC.find(x => x >= Math.max(ac, g.c1));
    if (r == null) r = visR[visR.length - 1];
    if (c == null) c = visC[visC.length - 1];
    let y = top - H(g.r1, r), x = left - W(g.c1, c);
    if (skip.has(key(r, c))) {
      for (const m of merged.values()) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) { y += H(m.r1, r); x += W(m.c1, c); r = m.r1; c = m.c1; break; }
    }
    const k = key(r, c);
    if (!out.has(k)) out.set(k, []);
    out.get(k).push({ src: d.src, x, y, w: d.w, h: d.h, z });
  });
  return out;
}

// Baris paling bawah / kolom paling kanan yang ditutupi gambar (untuk rentang cetak & "jaga tabel tetap utuh").
export function imageExtent(sh, d) {
  const r0 = Math.min(d.anchor.r, sh.rows - 1), c0 = Math.min(d.anchor.c, sh.cols - 1);
  let r = r0, rem = d.dy + d.h;
  while (r < sh.rows - 1 && rem > sh.rowHeight(r)) { rem -= sh.rowHeight(r); r++; }
  let c = c0, remC = d.dx + d.w;
  while (c < sh.cols - 1 && remC > sh.colWidth(c)) { remC -= sh.colWidth(c); c++; }
  return { r1: r0, c1: c0, r2: r, c2: c };
}
