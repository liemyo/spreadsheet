// Page layout for printing / PDF: paper, margins, scale, title rows, print area and page breaks
// (automatic, manual, and "keep together" ranges that move a table to the next page instead of cutting it).
// Automatic printing (settings v2): the printed range is the worked area itself (not from A1), the scale fills
// the paper, the orientation that reads best is chosen, a table is never cut when it fits on a page, and a table
// longer than a page continues on the next one under its own title + header rows.
import { el, icon, esc, colName, rangeAddr, key, keyR, keyC } from './util.js';
import { dialog, field, select, checkbox, toast, segmented } from './ui/ui.js';
import { sheetHTML, cellDisplay } from './io.js';
import { imageExtent } from './imgcore.js';

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
// orient: auto | portrait | landscape · fit: auto | width | page | normal (+ scale %) · grid: auto | true | false
// auto: keep tables whole · hdr: repeat a cut table's title + header rows on the next page
const DEF = { v: 2, size: 'a4', orient: 'auto', m: { t: 1.91, b: 1.91, l: 1.78, r: 1.78 }, fit: 'auto', scale: 100, hc: true, vc: false, grid: 'auto', title: false, num: true, auto: true, hdr: true, rb: [], cb: [], keep: [], tr: null, area: null };
// what the print dialog edits (and shares between sheets); breaks, keeps, title rows and the print area stay per sheet
export const GLOBAL_KEYS = ['size', 'orient', 'm', 'fit', 'scale', 'hc', 'vc', 'grid', 'title', 'num', 'auto', 'hdr'];
export const MARGIN_PX = 96 / 2.54;
const NUM_H = 22, TITLE_H = 40;
const AUTO_MAX = 2;      // automatic fitting enlarges a small table at most to 200 %
const SLACK = 0.985;     // printed rows can come out a hair taller than on screen (fonts, thick borders)

export function setup(sh) {
  const p = sh.page || {};
  const q = { ...DEF, ...p, m: { ...DEF.m, ...(p.m || {}) }, rb: p.rb || [], cb: p.cb || [], keep: p.keep || [] };
  if (!p.v) {
    // saved before automatic printing: the old defaults (portrait, shrink to width, gridlines, left) become automatic
    if (!p.orient || p.orient === 'portrait') q.orient = 'auto';
    if (!p.fit || p.fit === 'width') q.fit = 'auto';
    if (p.grid == null || p.grid === true) q.grid = 'auto';
    if (!p.hc) q.hc = true;
    q.v = 2;
  }
  return q;
}
export const marginPreset = m => MARGINS.find(([, , v]) => v.every((x, i) => Math.abs(x - [m.t, m.b, m.l, m.r][i]) < .005))?.[0] || '';
const fmtCm = v => String(+(+v).toFixed(2)).replace('.', ',');

// Change the page settings of a sheet as one undoable step.
export function patch(app, sh, label, fn) {
  app.tx(label, () => { app.wb.hist.touchMeta(sh); const p = setup(sh); fn(p); sh.page = p; sh.layoutV++; });
  app.pageCache = null;
  app.grid.render();
  app.onSelChange();
}

export function paperPx(p, orient = p.orient) {
  let [w, h] = (PAPER[p.size] || PAPER.a4).mm;
  if (orient === 'landscape') [w, h] = [h, w];
  return { w, h, W: w / 25.4 * 96, H: h / 25.4 * 96 };
}

/* ---------------- what is on the sheet ---------------- */
const WHITE = /^(#fff(fff)?|white|transparent|rgba?\(\s*255\s*,\s*255\s*,\s*255\s*(,\s*1(\.0+)?\s*)?\))$/i;
const fillOf = st => (st && st.bg && !WHITE.test(String(st.bg).trim()) ? st.bg : null);
const borderOf = st => !!(st && (st.bt || st.bb || st.bl || st.br));
const inked = st => !!(fillOf(st) || borderOf(st));
const hasVal = cell => cell.f != null || (cell.v != null && cell.v !== '');
const hasValueAt = (sh, r, c) => { const cell = sh.get(r, c); return cell ? hasVal(cell) : sh.spillMap.has(key(r, c)); };
const valueAt = (sh, r, c) => { const cell = sh.get(r, c); return cell ? (cell.f != null ? cell.sv : cell.v) : sh.spillMap.get(key(r, c))?.v ?? null; };
const isText = v => typeof v === 'string' && v.trim() !== '';
const isNum = v => typeof v === 'number';
const fontCss = (st, def) => `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}${((st.fs || def.fs) * 4) / 3}px "${st.ff || def.ff}", Arial, Helvetica, sans-serif`;

/** The worked area of a sheet: every cell with a value, a fill or a border, spilled arrays, signatures and merged
 *  blocks, grown over text that runs on into empty cells (a long title). Empty leading rows/columns are left out. */
export function autoArea(app, sh) {
  const styles = sh.wb.styles;
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  const add = (a, b, c, d) => { if (a < r1) r1 = a; if (b < c1) c1 = b; if (c > r2) r2 = c; if (d > c2) c2 = d; };
  const texts = [];
  for (const [k, cell] of sh.cells) {
    const val = hasVal(cell);
    if (!val && !(cell.s && inked(styles[cell.s]))) continue;
    const r = keyR(k), c = keyC(k);
    if (sh.isRowHidden(r) || sh.hiddenC.has(c)) continue;
    add(r, c, r, c);
    if (val) texts.push([r, c, cell]);
  }
  for (const [k, s] of sh.spillOrigins) add(keyR(k), keyC(k), keyR(k) + s.rows - 1, keyC(k) + s.cols - 1);
  for (const d of sh.images || []) { const x = imageExtent(sh, d); add(x.r1, x.c1, x.r2, x.c2); }
  if (r2 < 0) return null;
  // merged blocks reaching into the area are printed whole
  for (let changed = true, n = 0; changed && n < 8; n++) {
    changed = false;
    for (const m of sh.merges) {
      if (m.r1 > r2 || m.r2 < r1 || m.c1 > c2 || m.c2 < c1) continue;
      if (m.r1 < r1 || m.c1 < c1 || m.r2 > r2 || m.c2 > c2) { add(m.r1, m.c1, m.r2, m.c2); changed = true; }
    }
  }
  // text that runs on into empty cells is printed whole: the area grows over those cells
  const def = app?.defFont || { ff: 'Arial', fs: 10 }, g = app?.grid;
  if (g?.measure && texts.length <= 60000) {
    for (const [r, c, cell] of texts) {
      const st = styles[cell.s] || {};
      if (st.wrap === 'wrap' || st.wrap === 'clip' || st.rot || (sh.merges.length && sh.mergeAt(r, c))) continue;
      if (typeof (cell.f != null ? cell.sv : cell.v) !== 'string') continue;   // numbers show ### instead
      const ha = st.ha || 'left';
      if (ha === 'left' && c < c2 - 12) continue;    // far from the right edge: cannot reach past it
      const text = cellDisplay(app, sh, r, c);
      if (!text || text.includes('\n')) continue;
      const need = g.measure(fontCss(st, def), text) + 6 + (st.ind || 0) * 10 - sh.colWidth(c);
      if (need <= 0) continue;
      const run = dir => {
        let rest = ha === 'center' ? need / 2 : need, cc = c;
        while (rest > 0) {
          const nc = cc + dir;
          if (nc < 0 || nc >= sh.cols || hasValueAt(sh, r, nc) || (sh.merges.length && sh.mergeAt(r, nc))) break;
          rest -= sh.colWidth(nc); cc = nc;
        }
        return cc;
      };
      if (ha !== 'right') c2 = Math.max(c2, run(1));
      if (ha !== 'left') c1 = Math.min(c1, run(-1));
    }
  }
  return { r1, c1, r2, c2 };
}

// Rows/cols that get printed: the print area if one is set, otherwise the worked area.
export function printRange(sh, p = setup(sh), app = null) {
  if (p.area) return { ...p.area };
  return autoArea(app, sh);
}

/* ---------------- analysis of the printed range ---------------- */
// Tables = blocks of consecutive rows that carry something (value, fill, border, signature), separated by empty rows.
// Each gets `rep` = its title + header rows (repeated when the table continues on a next page) and `lead` = a short
// heading block that belongs with the table under it.
function analyze(app, sh, g, p) {
  const styles = sh.wb.styles;
  const ink = new Set(), inkC = new Set(), vals = new Map();
  let borders = false;
  const mark = (r, c, val) => { ink.add(r); inkC.add(c); if (val) { let a = vals.get(r); if (!a) vals.set(r, a = []); a.push(c); } };
  for (const [k, cell] of sh.cells) {
    const r = keyR(k), c = keyC(k);
    if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2 || sh.hiddenC.has(c)) continue;
    const st = cell.s ? styles[cell.s] : null;
    if (!borders && borderOf(st)) borders = true;
    const val = hasVal(cell);
    if (val || inked(st)) mark(r, c, val);
  }
  for (const [k] of sh.spillMap) {
    const r = keyR(k), c = keyC(k);
    if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2 && !sh.get(r, c)) mark(r, c, true);
  }
  const extra = [];
  for (const d of sh.images || []) {
    const x = imageExtent(sh, d);
    if (x.r2 < g.r1 || x.r1 > g.r2 || x.c2 < g.c1 || x.c1 > g.c2) continue;
    for (let r = Math.max(x.r1, g.r1); r <= Math.min(x.r2, g.r2); r++) ink.add(r);
    for (let c = Math.max(x.c1, g.c1); c <= Math.min(x.c2, g.c2); c++) inkC.add(c);
    if (x.r2 > x.r1) extra.push({ r1: Math.max(x.r1, g.r1), r2: Math.min(x.r2, g.r2), image: true });   // a signature is never cut
  }
  const hm = [];
  for (const m of sh.merges) {
    if (m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2) continue;
    const tl = sh.get(m.r1, m.c1);
    if (tl && (hasVal(tl) || inked(styles[tl.s]))) {
      for (let r = Math.max(m.r1, g.r1); r <= Math.min(m.r2, g.r2); r++) ink.add(r);
      for (let c = Math.max(m.c1, g.c1); c <= Math.min(m.c2, g.c2); c++) inkC.add(c);
    }
    if (m.r2 > m.r1 && m.r1 >= g.r1 && m.r2 <= g.r2) extra.push({ r1: m.r1, r2: m.r2, merge: true });
    if (m.c2 > m.c1) hm.push(m);
  }
  // heights / widths as prefix sums
  const N = g.r2 - g.r1 + 1, pre = new Float64Array(N + 1);
  for (let i = 0; i < N; i++) pre[i + 1] = pre[i] + sh.rowHeight(g.r1 + i);
  const hOf = (a, b) => { a = Math.max(a, g.r1); b = Math.min(b, g.r2); return b < a ? 0 : pre[b - g.r1 + 1] - pre[a - g.r1]; };
  const NC = g.c2 - g.c1 + 1, preC = new Float64Array(NC + 1);
  for (let i = 0; i < NC; i++) preC[i + 1] = preC[i] + sh.colWidth(g.c1 + i);
  const wOf = (a, b) => { a = Math.max(a, g.c1); b = Math.min(b, g.c2); return b < a ? 0 : preC[b - g.c1 + 1] - preC[a - g.c1]; };

  // tables
  const tables = [];
  for (let r = g.r1, s = -1; r <= g.r2 + 1; r++) {
    const on = r <= g.r2 && (ink.has(r) || (s >= 0 && sh.isRowHidden(r)));
    if (on && s < 0) s = r;
    else if (!on && s >= 0) { tables.push({ r1: s, r2: r - 1 }); s = -1; }
  }
  const tableOf = new Int32Array(N).fill(-1);
  tables.forEach((t, i) => { t.i = i; for (let r = t.r1; r <= t.r2; r++) tableOf[r - g.r1] = i; });

  const n = r => vals.get(r)?.length || 0;
  const look = r => {
    const cs = vals.get(r) || [];
    let bold = 0, fill = null;
    for (const c of cs) { const st = sh.styleOf(r, c); if (st.b) bold++; fill ||= fillOf(st) || (sh.banding.length ? app.bandingFor(sh, r, c) : null); }
    return { cs, bold: cs.length > 0 && bold === cs.length, anyBold: bold > 0, fill };
  };
  const widest = (a, b) => { let w = 0; for (let r = a; r <= b && w < 2; r++) w = Math.max(w, n(r)); return w; };
  const groupHead = r => sh.merges.some(m => m.r1 === r && m.r2 === r && m.c2 > m.c1 && m.c1 >= g.c1 && m.c2 <= g.c2);
  // the header row(s) starting at h, or -1
  const headEnd = (h, t) => {
    if (h >= t.r2 || n(h) < 2) return -1;
    const L0 = look(h), cs = L0.cs;
    if (cs.filter(c => isText(valueAt(sh, h, c))).length < Math.max(2, cs.length * 0.6)) return -1;   // a header is (mostly) words
    const next = look(h + 1);
    let nums = false;
    for (let r = h + 1; r <= Math.min(t.r2, h + 4) && !nums; r++) for (const c of cs) if (isNum(valueAt(sh, r, c)) && !isNum(valueAt(sh, h, c))) { nums = true; break; }
    const styled = (L0.anyBold && !next.bold) || (!!L0.fill && L0.fill !== next.fill);
    const frozen = sh.freezeR > h && sh.freezeR <= h + 3;
    const filt = sh.filter && sh.filter.r1 === h;
    const band = sh.banding.some(b => b.header && b.r1 === h);
    if (!(nums || styled || frozen || filt || band)) return -1;
    let end = h;
    // a second / third header line (sub-headers under group headers) looks like the first one
    while (end - h < 2 && end + 2 <= t.r2) {
      const nx = look(end + 1);
      if (!nx.cs.length || !nx.cs.every(c => isText(valueAt(sh, end + 1, c)))) break;
      if (styled ? !(nx.anyBold === L0.anyBold && nx.fill === L0.fill) : !groupHead(end)) break;
      end++;
    }
    if (frozen) end = Math.max(end, Math.min(sh.freezeR - 1, t.r2 - 1));
    return end;
  };
  const titleRows = new Set();   // headings: shown on the first page of columns only, not cut into the label columns
  for (const t of tables) {
    const rows = t.r2 - t.r1 + 1;
    t.lead = rows <= 3 && (() => { for (let r = t.r1; r <= t.r2; r++) { if (n(r) > 1) return false; for (const c of vals.get(r) || []) if (!isText(valueAt(sh, r, c))) return false; } return n(t.r1) === 1; })();
    if (t.lead) for (let r = t.r1; r <= t.r2; r++) titleRows.add(r);
    if (rows < 3) continue;
    let h = t.r1;
    // title rows: one cell (often merged over the table) above a wider row, at most two
    while (h - t.r1 < 2 && h < t.r2 - 1 && n(h) === 1 && widest(h + 1, t.r2) >= 2) h++;
    for (let r = t.r1; r < h; r++) titleRows.add(r);
    const e = headEnd(h, t);
    if (e >= 0) { t.rep = { r1: t.r1, r2: e }; t.head = h; }
    else if (h > t.r1 && t.r2 - h >= 2) t.rep = { r1: t.r1, r2: h - 1 };
  }
  // label columns: the first columns of the widest table that name its rows (No, Nama, Proyek, Tanggal …);
  // repeated on the left of the next pages when the table is wider than the paper
  let lab = null;
  {
    let best = null, bw = 0;
    for (const t of tables) {
      let a = Infinity, b = -1;
      for (let r = t.r1; r <= t.r2; r++) for (const c of vals.get(r) || []) { if (c < a) a = c; if (c > b) b = c; }
      if (b > a && wOf(a, b) > bw) { bw = wOf(a, b); best = { t, a, b }; }
    }
    if (best) {
      const { t, a, b } = best, d1 = t.rep ? t.rep.r2 + 1 : t.r1 + 1, cols = [];
      for (let c = a; c < b && cols.length < 3; c++) {
        let txt = 0, num = 0, seq = true, prev = null;
        for (let r = d1; r <= t.r2; r++) {
          const v = valueAt(sh, r, c);
          if (isText(v)) txt++;
          else if (isNum(v)) { num++; if (prev != null && v !== prev + 1) seq = false; prev = v; }
        }
        const all = txt + num;
        if (!all || !(txt >= all * 0.6 || (!cols.length && seq && num >= all * 0.6))) break;
        cols.push(c);
      }
      if (cols.length) lab = { c1: cols[0], c2: cols[cols.length - 1] };
    }
  }
  const keeps = [
    ...p.keep.filter(k => k.r2 >= g.r1 && k.r1 <= g.r2).map(k => ({ r1: Math.max(k.r1, g.r1), r2: Math.min(k.r2, g.r2), manual: true })),
    ...extra,
    ...(p.auto ? tables.filter(t => t.r2 > t.r1) : []),
  ];
  const tr = p.tr && p.tr.r2 >= p.tr.r1 ? p.tr : null;
  return { sh, g, ink, inkC, tables, tableOf, keeps, hm, hOf, wOf, borders, tr, lab, titleRows, totW: wOf(g.c1, g.c2), totH: hOf(g.r1, g.r2), rb: new Set(p.rb), cb: new Set(p.cb), rbs: [...p.rb].sort((a, b) => a - b) };
}

/* ---------------- pagination ---------------- */
// Rows of each page for an available height H (sheet px at the chosen scale).
function rowPages(A, H, p) {
  const { g, hOf, rb, tables, tableOf } = A;
  const tableAt = r => (r >= g.r1 && r <= g.r2 && tableOf[r - g.r1] >= 0 ? tables[tableOf[r - g.r1]] : null);
  const rep = start => {
    if (A.tr) return start > A.tr.r2 ? A.tr : null;
    if (!p.hdr) return null;
    const t = tableAt(start);
    return t && t.rep && start > t.rep.r2 ? t.rep : null;
  };
  const repH = r => { const x = rep(r); return x ? hOf(x.r1, x.r2) : 0; };
  const pages = [];
  let start = g.r1, moved = 0;
  while (start <= g.r2) {
    const head = rep(start);
    const avail = H - (head ? hOf(head.r1, head.r2) : 0) - (!pages.length && p.title ? TITLE_H : 0);
    // the furthest row that still fits (a page holds at least one row)
    let e = start;
    if (hOf(start, start) <= avail) {
      let lo = start, hi = g.r2;
      while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (hOf(start, mid) <= avail + 1e-6) lo = mid; else hi = mid - 1; }
      e = lo;
    }
    let b = e + 1, forced = false;
    const mb = A.rbs.find(x => x > start);
    if (mb != null && mb <= b) { b = mb; forced = true; }
    if (b <= g.r2 && !forced) { const nb = tidyRows(A, b, start, H, tableAt, repH); if (nb !== b) moved++; b = nb; }
    pages.push({ r1: start, r2: Math.min(b - 1, g.r2), rep: head });
    start = b;
    while (start <= g.r2 && !A.ink.has(start) && !rb.has(start)) start++;   // a page never starts with empty rows
  }
  return { pages, moved };
}

// Move a page break up so nothing is cut untidily.
function tidyRows(A, b0, start, H, tableAt, repH) {
  const { g, hOf, keeps, tableOf, tables } = A;
  let b = b0;
  for (let guard = 0; guard < 60; guard++) {
    let nb = b;
    // 1. ranges kept together (tables, merged cells, signatures, manual) go to the next page whole when they fit on one
    for (const k of keeps) if (k.r1 < nb && k.r2 >= nb && k.r1 > start && hOf(k.r1, k.r2) + repH(k.r1) <= H) nb = k.r1;
    // 2. a table longer than a page is split, but not right under its header and not leaving its last row alone
    const t = tableAt(nb - 1);
    if (t && nb > t.r1 && nb <= t.r2) {
      const hEnd = t.rep ? t.rep.r2 : t.r1 - 1;
      if (nb - 1 <= hEnd + 1 && t.r1 > start) nb = t.r1;
      else if (nb === t.r2 && nb - 1 > Math.max(start, hEnd + 2)) nb--;
    }
    // 3. a heading just above a table goes to the next page with it
    let r = nb - 1;
    while (r > start && tableOf[r - g.r1] < 0) r--;
    const P = r > start && tableOf[r - g.r1] >= 0 ? tables[tableOf[r - g.r1]] : null;
    if (P && P.r2 < nb && P.lead && P.r1 > start) {
      let q = nb;
      while (q <= g.r2 && tableOf[q - g.r1] < 0) q++;
      const N = q <= g.r2 ? tables[tableOf[q - g.r1]] : null;
      if (N && N.r1 >= nb && N.r1 - P.r2 - 1 <= 2) nb = P.r1;
    }
    if (nb === b) break;
    b = nb;
  }
  return b > start ? b : b0;
}

// Columns of each page for an available width W; cuts between tables (an empty column) rather than through one.
// When a table is wider than the paper its label columns are repeated on the next pages and the columns are
// spread evenly (no page with a lone last column).
function colPages(A, W) {
  const { g, cb, inkC, wOf, hm, sh } = A;
  if (!isFinite(W) && !cb.size) return [{ c1: g.c1, c2: g.c2 }];
  const lab = A.lab && wOf(A.lab.c1, A.lab.c2) <= W * 0.35 ? A.lab : null, labW = lab ? wOf(lab.c1, lab.c2) : 0;
  const run = Wl => {
    const pages = [];
    let start = g.c1;
    while (start <= g.c2) {
      const avail = lab && start > lab.c2 ? Wl - labW : Wl;
      let e = start, acc = 0, forced = false;
      for (let c = start; c <= g.c2; c++) {
        if (c > start && cb.has(c)) { forced = true; break; }
        const w = sh.colWidth(c);
        if (c > start && acc + w > avail + 1e-6) break;
        acc += w; e = c;
      }
      let b = e + 1;
      if (b <= g.c2 && !forced) {
        let nb = b;
        for (const m of hm) if (m.c1 < nb && m.c2 >= nb && m.c1 > start && wOf(m.c1, m.c2) <= avail) nb = m.c1;
        if (inkC.has(nb - 1) && inkC.has(nb)) {
          for (let c = nb - 1; c > start; c--) if (!inkC.has(c) && sh.colWidth(c) > 0) { if (wOf(start, c) >= avail * 0.4) nb = c + 1; break; }
        }
        if (nb > start) b = nb;
      }
      pages.push({ c1: start, c2: b - 1 });
      start = b;
      while (start <= g.c2 && !inkC.has(start) && !cb.has(start)) start++;
    }
    return pages;
  };
  let pages = run(W);
  if (pages.length > 1 && isFinite(W)) {
    // the narrowest page width that needs no extra page spreads the columns evenly
    let lo = W * 0.4, hi = W;
    for (let i = 0; i < 14; i++) { const mid = (lo + hi) / 2; if (run(mid).length <= pages.length) hi = mid; else lo = mid; }
    pages = run(hi);
  }
  if (lab) for (const pg of pages) if (pg.c1 > lab.c2) pg.lab = lab;
  return pages;
}

// Scale and pages for one orientation.
function solve(A, p, orient) {
  const paper = paperPx(p, orient);
  const cw = paper.W - (p.m.l + p.m.r) * MARGIN_PX;
  const ch = paper.H - (p.m.t + p.m.b) * MARGIN_PX - (p.num ? NUM_H : 0);
  // +2: the outer border of a table sticks out half its width on each side
  const fitW = cw / ((A.totW || 1) + 2), fitP = Math.min(fitW, ch / ((A.totH || 1) + 2 + (p.title ? TITLE_H : 0)));
  const at = s => rowPages(A, (ch / s) * SLACK, p);
  let s;
  if (p.fit === 'normal') s = (p.scale || 100) / 100;
  else if (p.fit === 'page') s = Math.min(AUTO_MAX, fitP);
  else if (p.fit === 'width') s = Math.min(AUTO_MAX, fitW);
  else {
    // automatic: fill the paper width; a short document shrinks a little more when that saves a page
    // (a table that just spills over onto an almost empty last page)
    s = Math.min(AUTO_MAX, fitW);
    const n = at(s).pages.length;
    if (n > 1 && n <= 4) {
      const lo = s * (s >= 0.9 ? 0.75 : s >= 0.7 ? 0.8 : 0.9), nLo = at(lo).pages.length;
      if (nLo < n) {
        let a = lo, b = s;
        for (let i = 0; i < 10; i++) { const mid = (a + b) / 2; if (at(mid).pages.length <= nLo) a = mid; else b = mid; }
        s = a;
      }
    }
  }
  s = Math.max(0.1, Math.min(4, s));
  const R = at(s);
  const cols = colPages(A, p.fit === 'normal' ? cw / s : Infinity);
  return { orient, paper, cw, ch, scale: s, fitW, W: cw / s, H: (ch / s) * SLACK, rows: R.pages, moved: R.moved, cols, pages: R.pages.length * cols.length };
}

/** Page layout of a sheet: rows of each page, columns of each page, scale, orientation. Cached per app until the next change. */
export function layout(app, sh, over = null) {
  const ck = sh.id + '|' + sh.layoutV + '|' + JSON.stringify(sh.page) + '|' + JSON.stringify(over);
  app.pageCache ||= new Map();
  if (app.pageCache.has(ck)) return app.pageCache.get(ck);
  const p = { ...setup(sh), ...(over || {}) };
  const g = printRange(sh, p, app);
  const orient0 = p.orient === 'landscape' ? 'landscape' : 'portrait';
  let res = { p, g, orient: orient0, paper: paperPx(p, orient0), cw: 0, ch: 0, scale: 1, W: 0, H: 0, rows: [], cols: [], pages: 0, keeps: [], tables: [], tr: null, trH: 0, grid: p.grid !== false, rowBreaks: [], colBreaks: [], manualR: new Set(p.rb), manualC: new Set(p.cb), moved: 0, split: 0, repeated: 0, other: null };
  if (g) {
    const A = analyze(app, sh, g, p);
    const cands = (p.orient === 'auto' ? ['portrait', 'landscape'] : [orient0]).map(o => solve(A, p, o));
    // fewest pages, but tiny print counts too (every 10 % under 85 % weighs like one more page) and a table cut
    // sideways weighs half a page more; a tie goes to the bigger print, then to the paper the table fits best
    const cost = c => c.pages + Math.max(0, 0.85 - c.scale) * 10 + (c.cols.length - 1) * 0.5;
    let best = cands[0];
    for (const c of cands.slice(1)) {
      const d = cost(c) - cost(best);
      if (d < -1e-9 || (Math.abs(d) <= 1e-9 && (c.scale > best.scale + 1e-9 || (Math.abs(c.scale - best.scale) <= 1e-9 && c.fitW > best.fitW + 1e-9)))) best = c;
    }
    res = {
      ...res, ...best,
      keeps: A.keeps, tables: A.tables, tr: A.tr, trH: A.tr ? A.hOf(A.tr.r1, A.tr.r2) : 0, titleRows: A.titleRows, lab: best.cols.some(c => c.lab) ? A.lab : null,
      grid: p.grid === 'auto' ? !A.borders : !!p.grid,
      rowBreaks: best.rows.slice(1).map(x => x.r1), colBreaks: best.cols.slice(1).map(x => x.c1),
      split: A.tables.filter(t => best.rows.some(x => x.r1 > t.r1 && x.r1 <= t.r2)).length,
      repeated: best.rows.filter(x => x.rep).length,
      other: cands.find(c => c !== best) || null,
    };
  }
  app.pageCache.set(ck, res);
  return res;
}

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
  if (h > L.H) { toast('Pilihan lebih tinggi dari satu halaman — judul & header tabelnya diulang otomatis di halaman berikutnya'); return; }
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
  const pct = () => Math.round(L().scale * 100);
  return [
    { label: 'Margin', icon: 'margin', sub: () => [
      ...MARGINS.map(([k, n, v]) => ({ label: `${n}  ·  ${v.map(fmtCm).join(' / ')} cm`, checked: () => marginPreset(setup(sh).m) === k, run: () => patch(app, sh, 'Margin', q => { [q.m.t, q.m.b, q.m.l, q.m.r] = v; }) })),
      { sep: true },
      { label: 'Margin kustom…', icon: 'tune', run: () => setupDialog(app) },
    ] },
    { label: 'Orientasi', icon: 'screen_rotation', sub: [
      { label: 'Otomatis (yang paling terbaca)', icon: 'auto_awesome', checked: () => setup(sh).orient === 'auto', run: () => patch(app, sh, 'Orientasi', q => { q.orient = 'auto'; }) },
      { label: 'Potret', icon: 'crop_portrait', checked: () => setup(sh).orient === 'portrait', run: () => patch(app, sh, 'Orientasi', q => { q.orient = 'portrait'; }) },
      { label: 'Lanskap', icon: 'crop_landscape', checked: () => setup(sh).orient === 'landscape', run: () => patch(app, sh, 'Orientasi', q => { q.orient = 'landscape'; }) },
    ] },
    { label: 'Ukuran kertas', icon: 'description', sub: () => Object.entries(PAPER).map(([k, v]) => ({ label: `${v.n}  ·  ${fmtCm(v.mm[0] / 10)} × ${fmtCm(v.mm[1] / 10)} cm`, checked: () => setup(sh).size === k, run: () => patch(app, sh, 'Ukuran kertas', q => { q.size = k; }) })) },
    { label: 'Skala', icon: 'fit_screen', sub: () => [
      { label: `Otomatis — penuhi kertas (${pct()}%)`, checked: () => setup(sh).fit === 'auto', run: () => patch(app, sh, 'Skala', q => { q.fit = 'auto'; }) },
      { label: 'Sesuaikan dengan lebar', checked: () => setup(sh).fit === 'width', run: () => patch(app, sh, 'Skala', q => { q.fit = 'width'; }) },
      { label: 'Muat dalam 1 halaman', checked: () => setup(sh).fit === 'page', run: () => patch(app, sh, 'Skala', q => { q.fit = 'page'; }) },
      { label: 'Normal (100%)', checked: () => setup(sh).fit === 'normal' && setup(sh).scale === 100, run: () => patch(app, sh, 'Skala', q => { q.fit = 'normal'; q.scale = 100; }) },
      { label: 'Skala kustom…', icon: 'tune', run: () => setupDialog(app) },
    ] },
    { sep: true },
    { label: 'Tampilkan pemisah halaman', icon: 'insert_page_break', checked: () => !!app.settings.breaks, run: () => showBreaks(app, !app.settings.breaks) },
    { label: `Jangan potong ${rangeAddr(s)} (pindah ke halaman berikutnya)`, icon: 'move_down', hidden: () => kept || s.r1 === s.r2, run: () => keepTogether(app, s) },
    { label: `Izinkan ${rangeAddr(s)} terpotong`, icon: 'content_cut', hidden: () => !kept, run: () => unkeep(app, s) },
    { label: 'Hindari tabel terpotong (otomatis)', icon: 'auto_awesome', checked: () => setup(sh).auto, run: () => patch(app, sh, 'Hindari tabel terpotong', q => { q.auto = !q.auto; }) },
    { label: 'Ulangi judul & header tabel di halaman berikutnya', icon: 'view_headline', checked: () => setup(sh).hdr, run: () => patch(app, sh, 'Ulangi header tabel', q => { q.hdr = !q.hdr; }) },
    { label: 'Pemisah halaman manual', icon: 'horizontal_rule', sub: () => [
      { label: `Sisipkan di atas baris ${s.r1 + 1}`, icon: 'vertical_align_top', hidden: () => rbHere || s.r1 === 0, run: () => breakAt(app, 'r', s.r1) },
      { label: `Hapus pemisah di atas baris ${s.r1 + 1}`, icon: 'close', hidden: () => !rbHere, run: () => breakAt(app, 'r', s.r1, false) },
      { label: `Sisipkan di kiri kolom ${colName(s.c1)}`, icon: 'format_indent_increase', hidden: () => cbHere || s.c1 === 0, run: () => breakAt(app, 'c', s.c1) },
      { label: `Hapus pemisah di kiri kolom ${colName(s.c1)}`, icon: 'close', hidden: () => !cbHere, run: () => breakAt(app, 'c', s.c1, false) },
      { sep: true },
      { label: 'Hapus semua pemisah & penjaga manual', icon: 'delete_sweep', disabled: () => !p.rb.length && !p.cb.length && !p.keep.length, run: () => patch(app, sh, 'Setel ulang pemisah', q => { q.rb = []; q.cb = []; q.keep = []; }) },
    ] },
    { label: 'Ulangi baris judul di tiap halaman', icon: 'title', sub: () => [
      { label: 'Otomatis (judul & header tiap tabel)', checked: () => !setup(sh).tr && setup(sh).hdr, run: () => patch(app, sh, 'Baris judul', q => { q.tr = null; q.hdr = true; }) },
      { label: 'Tidak ada', checked: () => !setup(sh).tr && !setup(sh).hdr, run: () => patch(app, sh, 'Baris judul', q => { q.tr = null; q.hdr = false; }) },
      { label: 'Baris 1', checked: () => setup(sh).tr?.r1 === 0 && setup(sh).tr?.r2 === 0, run: () => patch(app, sh, 'Baris judul', q => { q.tr = { r1: 0, r2: 0 }; }) },
      { label: 'Baris 1–2', checked: () => setup(sh).tr?.r1 === 0 && setup(sh).tr?.r2 === 1, run: () => patch(app, sh, 'Baris judul', q => { q.tr = { r1: 0, r2: 1 }; }) },
      { label: `Baris pilihan (${s.r1 + 1}${s.r2 > s.r1 ? '–' + (s.r2 + 1) : ''})`, run: () => patch(app, sh, 'Baris judul', q => { q.tr = { r1: s.r1, r2: Math.min(s.r2, s.r1 + 9) }; }) },
    ] },
    { label: 'Area cetak', icon: 'crop', sub: () => [
      { label: `Tetapkan dari pilihan (${rangeAddr(s)})`, run: () => patch(app, sh, 'Area cetak', q => { q.area = { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 }; }) },
      { label: 'Hapus area cetak (kembali otomatis)', disabled: () => !setup(sh).area, run: () => patch(app, sh, 'Area cetak', q => { q.area = null; }) },
    ] },
    { sep: true },
    { label: 'Penyiapan halaman…', icon: 'tune', run: () => setupDialog(app) },
    { label: `Pratinjau cetak (${L().g ? L().pages : 0} halaman)`, icon: 'preview', run: () => app.cmd('print') },
  ];
}

/* ---------------- page setup dialog ---------------- */
export function setupDialog(app) {
  const sh = app.sheet, p = setup(sh), L0 = layout(app, sh);
  const num = v => el('input.inp', { type: 'number', min: 0, max: 10, step: 0.1, value: +(+v).toFixed(2) });
  const m = { t: num(p.m.t), b: num(p.m.b), l: num(p.m.l), r: num(p.m.r) };
  const prev = el('div.pg-prev', el('i'));
  const sizeSel = select(Object.entries(PAPER).map(([k, v]) => ({ value: k, label: `${v.n} (${fmtCm(v.mm[0] / 10)} × ${fmtCm(v.mm[1] / 10)} cm)` })), p.size, () => draw());
  let orient = p.orient;
  const orientSeg = segmented([{ value: 'auto', label: 'Otomatis', icon: 'auto_awesome' }, { value: 'portrait', label: 'Potret', icon: 'crop_portrait' }, { value: 'landscape', label: 'Lanskap', icon: 'crop_landscape' }], orient, v => { orient = v; draw(); });
  const preset = select([{ value: '', label: 'Kustom' }, ...MARGINS.map(([k, n, v]) => ({ value: k, label: `${n} (${v.map(fmtCm).join(' / ')} cm)` }))], marginPreset(p.m), v => {
    const P = MARGINS.find(x => x[0] === v)?.[2];
    if (P) { [m.t.value, m.b.value, m.l.value, m.r.value] = P; draw(); }
  });
  Object.values(m).forEach(i => i.addEventListener('input', () => { preset.value = ''; draw(); }));
  let fit = p.fit;
  const scaleIn = el('input.inp', { type: 'number', min: 10, max: 400, step: 5, value: p.fit === 'normal' ? p.scale || 100 : Math.round(L0.scale * 100), disabled: fit !== 'normal' });
  const fitSel = select([{ value: 'auto', label: 'Otomatis — penuhi kertas, tabel utuh' }, { value: 'normal', label: 'Skala kustom (%)' }, { value: 'width', label: 'Sesuaikan dengan lebar halaman' }, { value: 'page', label: 'Muat seluruhnya dalam 1 halaman' }], fit, v => { fit = v; scaleIn.disabled = v !== 'normal'; });
  const o = { hc: p.hc, vc: p.vc, grid: p.grid, title: p.title, num: p.num, auto: p.auto, hdr: p.hdr };
  const ck = (k, label) => checkbox(label, o[k], v => (o[k] = v));
  function draw() {
    let [w, h] = PAPER[sizeSel.value].mm;
    if ((orient === 'auto' ? L0.orient : orient) === 'landscape') [w, h] = [h, w];
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
        ck('hdr', 'Ulangi judul & header tabel di halaman berikutnya'),
        ck('hc', 'Tengahkan horizontal'), ck('vc', 'Tengahkan vertikal'),
        checkbox('Cetak garis kisi', L0.grid, v => (o.grid = v)), ck('num', 'Nomor halaman di bawah'), ck('title', 'Judul spreadsheet di halaman pertama'))),
    el('div.pg-prev-wrap', prev, el('small.muted', 'Pratinjau margin')));
  dialog({
    title: `Penyiapan halaman — ${sh.name}`, width: 640, body,
    buttons: [{ label: 'Pratinjau', run: () => { apply(); app.cmd('print'); } }, { spacer: true }, { label: 'Batal' }, { label: 'Terapkan', primary: true, run: () => { apply(); showBreaks(app); toast('Penyiapan halaman disimpan — garis putus-putus menandai batas halaman'); } }],
  });
  function apply() {
    const cm = x => Math.max(0, Math.min(10, parseFloat(x.value) || 0));
    patch(app, sh, 'Penyiapan halaman', q => {
      Object.assign(q, o, { size: sizeSel.value, orient, fit, scale: Math.max(10, Math.min(400, parseFloat(scaleIn.value) || 100)), m: { t: cm(m.t), b: cm(m.b), l: cm(m.l), r: cm(m.r) } });
    });
  }
}

/* ---------------- printing ---------------- */
// CSS of the printed pages. S scopes it (`#print-root ` when the pages are printed inside the app page itself).
export function printCSS(S = '', f = { ff: 'Arial', fs: 10 }) {
  const ff = `"${String(f.ff || 'Arial').replace(/["\\]/g, '')}", Arial, Helvetica, sans-serif`;
  return `${S}.pg{position:relative;box-sizing:border-box;background:#fff;color:#000;margin:18px auto;box-shadow:0 1px 4px rgba(60,64,67,.3);overflow:hidden;isolation:isolate;font:${f.fs}pt/normal ${ff};text-align:left;letter-spacing:normal;text-transform:none}
${S}.pg *{box-sizing:border-box}
${S}.pg table.sheet{border-collapse:collapse;table-layout:fixed;border-spacing:0;margin:0}
${S}.pg table.sheet td{padding:1px 4px;overflow:hidden;white-space:nowrap;line-height:var(--lh,1.25);font-size:${f.fs}pt}
${S}.pg table.sheet.grid td{outline:1px solid #d9d9d9;outline-offset:-0.5px}
${S}.pg table.sheet tr{break-inside:avoid}
${S}.pg a{color:#1155cc}
${S}.pg-in{position:relative;overflow:hidden;display:flex;flex-direction:column}
${S}.pg-in.vc{justify-content:center}
${S}.pg-z{transform-origin:0 0}
${S}.pg-in.hc .pg-z{margin:0 auto}
${S}.pg-rep,${S}.pg-t{position:relative;overflow:hidden;width:max-content;max-width:none}
${S}.pg-rep+.pg-t{margin-top:-1px}
${S}.pg-band{display:flex;align-items:flex-start}
${S}.pg-lab{flex:none;overflow:hidden;margin-right:-1px}
${S}.pg-title{height:${TITLE_H}px;overflow:hidden}
${S}.pg-title h1{font:600 13pt/20px Arial,Helvetica,sans-serif;margin:0 0 2px;color:#000}
${S}.pg-title .meta{color:#666;font:8pt/14px Arial,Helvetica,sans-serif}
${S}.pg-num{text-align:center;font:9pt Arial,Helvetica,sans-serif;color:#555}
@media print{${S}.pg{margin:0!important;box-shadow:none!important;break-after:page;padding:0!important;width:auto!important;height:auto!important}${S}.pg:last-child{break-after:auto}}`;
}

function pageHTML(app, sh, L, rows, cols, idx, total, first) {
  const p = L.p, parts = [], opt = { grid: L.grid, clip: true };
  if (first && p.title) parts.push(`<div class="pg-title"><h1>${esc(app.wb.name)}</h1><div class="meta">${esc(sh.name)} · dicetak ${esc(new Date().toLocaleString('id-ID'))}</div></div>`);
  // the label columns of a table wider than the paper stand on the left of its next pages (headings left out there)
  const band = (r1, r2, o) => {
    const t = sheetHTML(app, sh, { r1, r2, c1: cols.c1, c2: cols.c2 }, o);
    if (!cols.lab) return t;
    return `<div class="pg-band"><div class="pg-lab">${sheetHTML(app, sh, { r1, r2, c1: cols.lab.c1, c2: cols.lab.c2 }, { ...o, images: false, blank: L.titleRows })}</div>${t}</div>`;
  };
  // a table continued from the previous page starts with its title + header rows again
  if (rows.rep) parts.push(`<div class="pg-rep">${band(rows.rep.r1, rows.rep.r2, { ...opt, images: false })}</div>`);
  // the wrapper cuts a signature that crosses the page edge; the next page shows the rest
  parts.push(`<div class="pg-t">${band(rows.r1, rows.r2, opt)}</div>`);
  // zoom keeps layout size in step with the scale, so centering and the page height stay exact
  return `<div class="pg-in${p.hc ? ' hc' : ''}${p.vc ? ' vc' : ''}" style="width:${L.cw.toFixed(2)}px;height:${(L.ch - 0.5).toFixed(2)}px"><div class="pg-z" style="zoom:${L.scale.toFixed(4)}">${parts.join('')}</div></div>${p.num ? `<div class="pg-num" style="height:${NUM_H}px;line-height:${NUM_H}px">${idx} / ${total}</div>` : ''}`;
}

/** Printable pages of some sheets: {css, body, total, all:[{sh, L}]}. over = settings over each sheet's own (object or sh => object). */
export function printDoc(app, { sheets, over = null, scope = '', limit = Infinity }) {
  const all = sheets.map(sh => ({ sh, L: app.withSheet(sh, () => layout(app, sh, typeof over === 'function' ? over(sh) : over)) }));
  const total = all.reduce((a, x) => a + (x.L.g ? x.L.pages : 0), 0);
  const css = [printCSS(scope, app.defFont)], parts = [];
  let n = 0;
  for (const { sh, L } of all) {
    if (!L.g) continue;
    const p = L.p, name = 'pg' + sh.id.replace(/\W/g, '');
    css.push(`@page ${name}{size:${L.paper.w}mm ${L.paper.h}mm;margin:${p.m.t * 10}mm ${p.m.r * 10}mm ${p.m.b * 10}mm ${p.m.l * 10}mm}${scope}.${name}{page:${name};width:${L.paper.W.toFixed(2)}px;height:${L.paper.H.toFixed(2)}px;padding:${(p.m.t * MARGIN_PX).toFixed(2)}px ${(p.m.r * MARGIN_PX).toFixed(2)}px ${(p.m.b * MARGIN_PX).toFixed(2)}px ${(p.m.l * MARGIN_PX).toFixed(2)}px}`);
    let first = true;
    for (const cols of L.cols) for (const rows of L.rows) {   // down, then over (as Excel)
      if (++n > limit) break;
      parts.push(`<section class="pg ${name}">${app.withSheet(sh, () => pageHTML(app, sh, L, rows, cols, n, total, first))}</section>`);
      first = false;
    }
  }
  return { css: css.join('\n'), body: parts.join(''), total, shown: Math.min(total, limit), all };
}

// Floating suggestion under a selection that a page break cuts in two.
export function breakChip(app) {
  const host = document.getElementById('pagechip');
  if (!host) return;
  host.classList.remove('show');
  host.replaceChildren();
  const g = app.grid, sh = app.sheet;
  if (!sh || g.edit || app.readOnly || app.settings.breakTips === false || g.sel.ranges.length !== 1) return;
  if (document.querySelector('.dlg-back')) return;
  const s = g.range;
  if (s.r2 - s.r1 > 2000 || s.r1 === s.r2) return;
  const cut = cutBy(app, sh, s);
  if (!cut || app._chipSkip === rangeAddr(s)) return;
  // the whole printed range selected, or a long table that already continues under its own header: nothing to fix
  if (s.r1 <= cut.L.g.r1 && s.r2 >= cut.L.g.r2) return;
  if (cut.L.rows[cut.page]?.rep && !cut.fits) return;
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
