// Editing operations on the model. Callers wrap these in a history transaction (app.tx).
import { key, keyR, keyC, colName, rangeAddr, uid, norm, inRange, intersects, parseRange } from './util.js';
import { insertDelete, rangeCells, setInput, COL_W, ROW_H } from './model.js';
import { shiftFormula, mapRefs } from './formula/parse.js';
import { compare, FErr } from './formula/core.js';
import { format, editText, parseInput, L, compile } from './format.js';
import { fontOf } from './grid.js';
import { dvAt, dvValid } from './cf.js';

const clampG = (sh, g) => ({ r1: g.r1, c1: g.c1, r2: Math.min(g.r2, sh.rows - 1), c2: Math.min(g.c2, sh.cols - 1) });

// ---------- content / styles ----------
export function clearRanges(app, ranges, what = 'contents') {
  const sh = app.sheet;
  for (const g0 of ranges) {
    const g = clampG(sh, g0);
    rangeCells(sh, g, (r, c, cell) => {
      if (!cell) return;
      if (what === 'contents') { if (cell.v != null || cell.f != null) sh.put(r, c, { v: null, f: null, link: null }); }
      else if (what === 'formats') { if (cell.s) sh.put(r, c, { s: 0 }); }
      else if (what === 'notes') { if (cell.note) sh.put(r, c, { note: null }); }
      else sh.put(r, c, { v: null, f: null, s: 0, note: null, link: null });
    });
    if (what === 'formats' || what === 'all') {
      const cf = sh.cf.filter(rule => !rule.ranges.every(x => x.r1 >= g.r1 && x.r2 <= g.r2 && x.c1 >= g.c1 && x.c2 <= g.c2));
      if (cf.length !== sh.cf.length) { app.wb.hist.touchMeta(sh); sh.cf = cf; sh.cfVer = (sh.cfVer || 0) + 1; }
      const bd = sh.banding.filter(b => !intersects(b, g));
      if (bd.length !== sh.banding.length) { app.wb.hist.touchMeta(sh); sh.banding = bd; }
    }
  }
}

export function styleRanges(app, ranges, patch) {
  const sh = app.sheet, wb = app.wb;
  const fn = typeof patch === 'function' ? patch : st => ({ ...st, ...patch });
  for (const g0 of ranges) {
    const g = clampG(sh, g0);
    const area = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
    if (area > 200000) {
      // whole-sheet style: only touch existing cells plus the used area
      const u = sh.usedRange() || { r2: 0, c2: 0 };
      const gg = { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, Math.max(u.r2, 200)), c2: Math.min(g.c2, Math.max(u.c2, 25)) };
      for (let r = gg.r1; r <= gg.r2; r++) for (let c = gg.c1; c <= gg.c2; c++) sh.setStyle(r, c, fn(sh.styleOf(r, c), r, c));
      continue;
    }
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) sh.setStyle(r, c, fn(sh.styleOf(r, c), r, c));
  }
  wb.dirty = true;
}

export function bordersOn(app, ranges, type, b) {
  const sh = app.sheet;
  const B = b ? { w: b.w || 1, style: b.style || 'solid', color: b.color || '#000000' } : null;
  for (const g0 of ranges) {
    const g = clampG(sh, g0);
    styleRanges(app, [g], (st, r, c) => {
      const o = { ...st };
      const top = r === g.r1, bot = r === g.r2, left = c === g.c1, right = c === g.c2;
      const set = (k, v) => { if (v) o[k] = B; else if (v === false) delete o[k]; };
      switch (type) {
        case 'all': set('bt', true); set('bb', true); set('bl', true); set('br', true); break;
        case 'outer': if (top) set('bt', true); if (bot) set('bb', true); if (left) set('bl', true); if (right) set('br', true); break;
        case 'inner': if (!top) set('bt', true); if (!bot) set('bb', true); if (!left) set('bl', true); if (!right) set('br', true); break;
        case 'horizontal': if (!top) set('bt', true); if (!bot) set('bb', true); break;
        case 'vertical': if (!left) set('bl', true); if (!right) set('br', true); break;
        case 'top': if (top) set('bt', true); break;
        case 'bottom': if (bot) set('bb', true); break;
        case 'left': if (left) set('bl', true); break;
        case 'right': if (right) set('br', true); break;
        case 'clear': delete o.bt; delete o.bb; delete o.bl; delete o.br; break;
      }
      return o;
    });
    if (type === 'clear' || type === 'outer') {
      // also clear neighbour edges pointing into the range for a clean result
      if (type === 'clear') {
        if (g.r1 > 0) for (let c = g.c1; c <= g.c2; c++) { const st = sh.styleOf(g.r1 - 1, c); if (st.bb) { const o = { ...st }; delete o.bb; sh.setStyle(g.r1 - 1, c, o); } }
        if (g.c1 > 0) for (let r = g.r1; r <= g.r2; r++) { const st = sh.styleOf(r, g.c1 - 1); if (st.br) { const o = { ...st }; delete o.br; sh.setStyle(r, g.c1 - 1, o); } }
      }
    }
  }
}

export function mergeRanges(app, ranges, mode) {
  const sh = app.sheet;
  let merges = sh.merges.slice();
  for (const g0 of ranges) {
    const g = clampG(sh, g0);
    merges = merges.filter(m => !intersects(m, g));
    if (mode === 'none') continue;
    const parts = [];
    if (mode === 'all') parts.push(g);
    else if (mode === 'h') for (let r = g.r1; r <= g.r2; r++) parts.push({ r1: r, c1: g.c1, r2: r, c2: g.c2 });
    else if (mode === 'v') for (let c = g.c1; c <= g.c2; c++) parts.push({ r1: g.r1, c1: c, r2: g.r2, c2: c });
    for (const p of parts) {
      if (p.r1 === p.r2 && p.c1 === p.c2) continue;
      // keep top-left value only
      rangeCells(sh, p, (r, c, cell) => { if (cell && (r !== p.r1 || c !== p.c1) && (cell.v != null || cell.f != null)) sh.put(r, c, { v: null, f: null }); });
      merges.push(p);
    }
  }
  sh.setMerges(merges);
}

export function canMerge(app, ranges) {
  const sh = app.sheet;
  for (const g of ranges) {
    let n = 0;
    rangeCells(sh, clampG(sh, g), (r, c, cell) => { if (cell && (cell.v != null || cell.f != null)) n++; });
    if (n > 1) return false;
  }
  return true;
}

// ---------- sizes ----------
export function resizeCols(app, cols, w) {
  const sh = app.sheet;
  app.wb.hist.touchMeta(sh);
  for (const c of cols) { if (w === COL_W) sh.colW.delete(c); else sh.colW.set(c, w); sh.hiddenC.delete(c); }
  sh.layoutV++;
}
export function resizeRows(app, rows, h) {
  const sh = app.sheet;
  app.wb.hist.touchMeta(sh);
  for (const r of rows) { if (h === ROW_H) sh.rowH.delete(r); else sh.rowH.set(r, h); sh.manualH.add(r); sh.hiddenR.delete(r); }
  sh.layoutV++;
}
export function autofitCols(app, cols) {
  const sh = app.sheet, grid = app.grid;
  app.wb.hist.touchMeta(sh);
  const R = Math.min(sh.dataRows(), 20000);
  for (const c of cols) {
    let w = 0;
    for (let r = 0; r < R; r++) {
      if (sh.isRowHidden(r) || sh.mergeAt(r, c)) continue;
      const cell = sh.get(r, c);
      if (!cell && !sh.spillMap.has(key(r, c))) continue;
      const d = grid.display(r, c, cell);
      if (!d || !d.text) continue;
      const font = fontOf(d.st, app.defFont);
      for (const ln of String(d.text).split('\n')) w = Math.max(w, grid.measure(font, ln));
      if (sh.filter && r === sh.filter.r1) w += 18;
    }
    if (w) sh.colW.set(c, Math.ceil(w + 10)); else sh.colW.delete(c);
  }
  sh.layoutV++;
}
export function autofitRows(app, rows) {
  const sh = app.sheet;
  app.wb.hist.touchMeta(sh);
  for (const r of rows) { sh.manualH.delete(r); fitRow(app, sh, r, true); }
  sh.layoutV++;
}
export function fitRow(app, sh, r, force = false) {
  if (!force && sh.manualH.has(r)) return;
  const grid = app.grid;
  let h = ROW_H;
  for (let c = 0; c <= Math.min(sh.maxC, sh.cols - 1); c++) {
    const cell = sh.get(r, c);
    if (!cell) continue;
    // merged cells: one spanning several rows never stretches a row (like Excel / Sheets), one across
    // columns wraps in its full width, the cells it covers do not count
    const m = sh.merges.length ? sh.mergeAt(r, c) : null;
    if (m && (m.r1 !== m.r2 || m.c1 !== c)) continue;
    const st = app.wb.styles[cell.s] || {};
    const px = ((st.fs || app.defFont.fs) * 4) / 3;
    const d = grid.display(r, c, cell);
    let lines = 1;
    if (d && d.text && (st.wrap === 'wrap' || String(d.text).includes('\n'))) {
      let w = -6;
      for (let cc = c; cc <= (m ? m.c2 : c); cc++) w += sh.colWidth(cc);
      lines = st.wrap === 'wrap' ? grid.wrapLines(fontOf(st, app.defFont), d.text, Math.max(10, w)).length : String(d.text).split('\n').length;
    }
    h = Math.max(h, Math.ceil(lines * px * 1.25 + 6));
    if (st.rot === 'v' && d && d.text) h = Math.max(h, Math.ceil(String(d.text).length * px * 1.05 + 8));
    else if (st.rot && d && d.text) h = Math.max(h, Math.ceil(Math.abs(Math.sin(st.rot * Math.PI / 180)) * grid.measure(fontOf(st, app.defFont), d.text) + px + 6));
  }
  const cur = sh.rowH.get(r) ?? ROW_H;
  if (cur !== h) { app.wb.hist.touchMeta(sh); if (h === ROW_H) sh.rowH.delete(r); else sh.rowH.set(r, h); sh.layoutV++; }
}

export function hideRows(app, rows, hide = true) {
  const sh = app.sheet; app.wb.hist.touchMeta(sh);
  for (const r of rows) hide ? sh.hiddenR.add(r) : sh.hiddenR.delete(r);
  sh.layoutV++;
}
export function hideCols(app, cols, hide = true) {
  const sh = app.sheet; app.wb.hist.touchMeta(sh);
  for (const c of cols) hide ? sh.hiddenC.add(c) : sh.hiddenC.delete(c);
  sh.layoutV++;
}

// ---------- structure ----------
export function insertRows(app, at, n) { insertDelete(app.wb, app.sheet, 'r', at, n); }
export function deleteRows(app, at, n) { insertDelete(app.wb, app.sheet, 'r', at, -n); }
export function insertCols(app, at, n) { insertDelete(app.wb, app.sheet, 'c', at, n); }
export function deleteCols(app, at, n) { insertDelete(app.wb, app.sheet, 'c', at, -n); }

function snapRange(sh, g) {
  const out = [];
  for (let r = g.r1; r <= g.r2; r++) {
    const row = [];
    for (let c = g.c1; c <= g.c2; c++) { const cell = sh.get(r, c); row.push(cell ? { v: cell.v, f: cell.f, s: cell.s, note: cell.note, link: cell.link } : null); }
    out.push(row);
  }
  return out;
}
function writeSnap(sh, r0, c0, snap, shiftF = null) {
  snap.forEach((row, i) => row.forEach((s, j) => {
    const r = r0 + i, c = c0 + j;
    if (!s) { const cur = sh.get(r, c); if (cur) sh.put(r, c, { v: null, f: null, s: 0, note: null, link: null }); return; }
    let f = s.f;
    if (f != null && shiftF) f = shiftFormula(f, shiftF[0], shiftF[1]);
    sh.put(r, c, { v: f != null ? null : s.v, f, s: s.s, note: s.note, link: s.link });
  }));
}

// shift cells right/down (insert) or left/up (delete) inside a range
export function shiftCells(app, g, dir) {
  const sh = app.sheet;
  if (dir === 'down' || dir === 'up') {
    const rows = g.r2 - g.r1 + 1;
    const bottom = Math.max(sh.dataRows(), g.r2 + 1) + rows;
    if (dir === 'down') {
      const block = snapRange(sh, { r1: g.r1, c1: g.c1, r2: bottom, c2: g.c2 });
      writeSnap(sh, g.r1, g.c1, block.map(r => r.map(() => null)));
      writeSnap(sh, g.r1 + rows, g.c1, block.slice(0, block.length - rows));
      sh.growTo(bottom + rows + 1, sh.cols);
    } else {
      const block = snapRange(sh, { r1: g.r2 + 1, c1: g.c1, r2: bottom, c2: g.c2 });
      writeSnap(sh, g.r1, g.c1, [...block, ...Array.from({ length: rows }, () => Array(g.c2 - g.c1 + 1).fill(null))]);
    }
  } else {
    const cols = g.c2 - g.c1 + 1;
    const right = Math.max(sh.dataCols(), g.c2 + 1) + cols;
    if (dir === 'right') {
      const block = snapRange(sh, { r1: g.r1, c1: g.c1, r2: g.r2, c2: right });
      writeSnap(sh, g.r1, g.c1, block.map(r => r.map(() => null)));
      writeSnap(sh, g.r1, g.c1 + cols, block.map(r => r.slice(0, r.length - cols)));
      sh.growTo(sh.rows, right + cols + 1);
    } else {
      const block = snapRange(sh, { r1: g.r1, c1: g.c2 + 1, r2: g.r2, c2: right });
      writeSnap(sh, g.r1, g.c1, block.map(r => [...r, ...Array(cols).fill(null)]));
    }
  }
}

export function freeze(app, r, c) {
  const sh = app.sheet; app.wb.hist.touchMeta(sh);
  sh.freezeR = Math.max(0, Math.min(r, sh.rows - 1));
  sh.freezeC = Math.max(0, Math.min(c, sh.cols - 1));
  sh.layoutV++;
}

// ---------- autofill ----------
const SERIES = [
  ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'],
  ['Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab', 'Min'],
  ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'],
  ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['Q1', 'Q2', 'Q3', 'Q4'],
];
function seriesOf(v) {
  if (typeof v !== 'string') return null;
  const low = v.toLowerCase();
  for (const list of SERIES) { const i = list.findIndex(x => x.toLowerCase() === low); if (i >= 0) return { list, i, upper: v === v.toUpperCase() && v !== v.toLowerCase() }; }
  return null;
}

function buildLine(sh, cells, wb) {
  // cells: [{r,c,snap}] source sequence; returns fn(pos, tr, tc) -> snap
  const n = cells.length;
  const vals = cells.map(x => x.snap && x.snap.f == null ? x.snap.v : undefined);
  const allNum = n >= 1 && vals.every(v => typeof v === 'number');
  const isDate = allNum && cells.every(x => compile(wb.styles[x.snap?.s || 0]?.fmt)?.isDate);
  const copy = (pos, tr, tc) => {
    const i = ((pos % n) + n) % n;
    const s = cells[i].snap;
    if (!s) return null;
    if (s.f != null) return { ...s, f: shiftFormula(s.f, tr - cells[i].r, tc - cells[i].c) };
    return { ...s };
  };
  if (allNum && (n >= 2 || isDate)) {
    let step = n >= 2 ? (vals[n - 1] - vals[0]) / (n - 1) : 1;
    // monthly step for dates on the same day-of-month
    return (pos, tr, tc) => {
      const base = copy(pos, tr, tc);
      let v = vals[0] + step * pos;
      if (!Number.isInteger(step)) v = +v.toPrecision(15);
      return { ...base, v, f: null };
    };
  }
  const pat = vals.map(v => (typeof v === 'string' ? /^(.*?)(\d+)(\D*)$/.exec(v) : null));
  if (n >= 1 && pat.every(Boolean) && pat.every(p => p[1] === pat[0][1] && p[3] === pat[0][3])) {
    const nums = pat.map(p => +p[2]);
    const step = n >= 2 ? (nums[n - 1] - nums[0]) / (n - 1) : 1;
    const width = pat[0][2].length;
    return (pos, tr, tc) => {
      const base = copy(pos, tr, tc);
      const k = Math.round(nums[0] + step * pos);
      const ks = k < 0 ? String(k) : String(k).padStart(pat[0][2].startsWith('0') ? width : 1, '0');
      return { ...base, v: pat[0][1] + ks + pat[0][3] };
    };
  }
  const ser = vals.map(seriesOf);
  if (n >= 1 && ser.every(s => s && s.list === ser[0].list)) {
    const list = ser[0].list;
    const step = n >= 2 ? ((ser[n - 1].i - ser[0].i) / (n - 1)) || 1 : 1;
    return (pos, tr, tc) => {
      const base = copy(pos, tr, tc);
      const i = (((ser[0].i + Math.round(step * pos)) % list.length) + list.length) % list.length;
      return { ...base, v: ser[0].upper ? list[i].toUpperCase() : list[i] };
    };
  }
  return copy;
}

export function autofill(app, src, target, mode = 'series') {
  const sh = app.sheet, wb = app.wb;
  const vertical = target.r1 !== src.r1 || target.r2 !== src.r2;
  const forward = vertical ? target.r2 > src.r2 : target.c2 > src.c2;
  const lines = vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
  for (let l = 0; l < lines; l++) {
    const cells = [];
    if (vertical) for (let r = src.r1; r <= src.r2; r++) cells.push({ r, c: src.c1 + l, snap: snapRange(sh, { r1: r, c1: src.c1 + l, r2: r, c2: src.c1 + l })[0][0] });
    else for (let c = src.c1; c <= src.c2; c++) cells.push({ r: src.r1 + l, c, snap: snapRange(sh, { r1: src.r1 + l, c1: c, r2: src.r1 + l, c2: c })[0][0] });
    const fn = mode === 'copy' ? (pos, tr, tc) => { const n = cells.length, i = ((pos % n) + n) % n, s = cells[i].snap; return s ? { ...s, f: s.f != null ? shiftFormula(s.f, tr - cells[i].r, tc - cells[i].c) : null } : null; } : buildLine(sh, cells, wb);
    const n = cells.length;
    if (vertical) {
      const c = src.c1 + l;
      if (forward) for (let r = src.r2 + 1, pos = n; r <= target.r2; r++, pos++) put(sh, r, c, fn(pos, r, c));
      else for (let r = src.r1 - 1, pos = -1; r >= target.r1; r--, pos--) put(sh, r, c, fn(pos, r, c));
    } else {
      const r = src.r1 + l;
      if (forward) for (let c = src.c2 + 1, pos = n; c <= target.c2; c++, pos++) put(sh, r, c, fn(pos, r, c));
      else for (let c = src.c1 - 1, pos = -1; c >= target.c1; c--, pos--) put(sh, r, c, fn(pos, r, c));
    }
  }
  function put(s, r, c, snap) {
    if (!snap) { if (s.get(r, c)) s.put(r, c, { v: null, f: null, s: 0 }); return; }
    s.put(r, c, { v: snap.f != null ? null : snap.v, f: snap.f ?? null, s: snap.s });
  }
}

// Ctrl+D / Ctrl+R
export function fillFrom(app, ranges, dir) {
  const sh = app.sheet;
  for (const g of ranges) {
    if (dir === 'down') {
      if (g.r1 === g.r2) { if (g.r1 === 0) continue; autofill(app, { ...g, r1: g.r1 - 1, r2: g.r1 - 1 }, { ...g, r1: g.r1 - 1 }, 'copy'); }
      else autofill(app, { ...g, r2: g.r1 }, g, 'copy');
    } else {
      if (g.c1 === g.c2) { if (g.c1 === 0) continue; autofill(app, { ...g, c1: g.c1 - 1, c2: g.c1 - 1 }, { ...g, c1: g.c1 - 1 }, 'copy'); }
      else autofill(app, { ...g, c2: g.c1 }, g, 'copy');
    }
  }
}

// double-click on fill handle: extend down to the length of the adjacent column
export function fillDownDouble(app) {
  const sh = app.sheet, g = app.grid.range;
  const has = (r, c) => app.grid.hasValue(r, c);
  let end = -1;
  for (const c of [g.c1 - 1, g.c2 + 1]) {
    if (c < 0 || c >= sh.cols) continue;
    if (!has(g.r2 + 1, c) && !has(g.r2, c)) continue;
    let r = g.r2;
    while (r + 1 < sh.rows && has(r + 1, c)) r++;
    end = Math.max(end, r);
  }
  if (end <= g.r2) return null;
  const target = { ...g, r2: end };
  autofill(app, g, target);
  return target;
}

// ---------- move / clipboard ----------
export function moveRange(app, src, dr, dc, copy = false, dstSheet = app.sheet) {
  const sh = app.sheet, wb = app.wb;
  const snap = snapRange(sh, src);
  const dst = { r1: src.r1 + dr, c1: src.c1 + dc, r2: src.r2 + dr, c2: src.c2 + dc };
  if (!copy) {
    rangeCells(sh, src, (r, c, cell) => { if (cell) sh.put(r, c, { v: null, f: null, s: 0, note: null, link: null }); });
    // references pointing into the moved block follow it
    wb.rewriteFormulas((a, sheetName, fsh) => {
      const tsh = sheetName != null ? wb.sheetByName(sheetName) : fsh;
      if (tsh !== sh || a.kind === 'col' || a.kind === 'row') return null;
      const r2 = a.r2 === Infinity ? a.r1 : a.r2, c2 = a.c2;
      if (a.r1 >= src.r1 && r2 <= src.r2 && a.c1 >= src.c1 && c2 <= src.c2) {
        const o = { ...a, r1: a.r1 + dr, c1: a.c1 + dc, r2: a.r2 === Infinity ? Infinity : a.r2 + dr, c2: a.c2 + dc };
        if (dstSheet !== sh) o.sheet = dstSheet.name;
        else if (fsh !== sh && sheetName == null) o.sheet = sh.name;
        return o;
      }
      return null;
    });
  }
  writeSnap(dstSheet, dst.r1, dst.c1, snap, copy ? [dr, dc] : null);
  return dst;
}

export function cellsForClipboard(app, g) {
  const sh = app.sheet, wb = app.wb;
  const rows = [];
  for (let r = g.r1; r <= g.r2; r++) {
    if (sh.filteredR.has(r)) continue;
    const row = [];
    for (let c = g.c1; c <= g.c2; c++) {
      const cell = sh.get(r, c);
      const d = app.grid.display(r, c, cell);
      row.push({
        v: cell ? cell.v : (sh.spillMap.get(key(r, c))?.v ?? null), f: cell ? cell.f : null, st: wb.styles[cell ? cell.s : 0] || {},
        note: cell?.note || null, link: cell?.link || null, text: d ? (d.text ?? '') : '', val: cell && cell.f != null ? cell.sv : cell ? cell.v : sh.spillMap.get(key(r, c))?.v ?? null, r, c,
        dv: dvAt(sh, r, c),
      });
    }
    rows.push(row);
  }
  return { rows, g, sheetId: sh.id, colW: Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => sh.colW.get(g.c1 + i)) };
}

export function toTSV(clip) {
  return clip.rows.map(row => row.map(x => { const t = String(x.text ?? ''); return /[\t\n"]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; }).join('\t')).join('\n');
}

const cssBorder = b => (b ? `${b.w || 1}px ${b.style || 'solid'} ${b.color || '#000'}` : '');
export function toHTML(clip, stamp) {
  const td = x => {
    const st = x.st, css = [];
    if (st.b) css.push('font-weight:bold'); if (st.i) css.push('font-style:italic');
    if (st.u || st.s) css.push('text-decoration:' + [st.u ? 'underline' : '', st.s ? 'line-through' : ''].join(' '));
    if (st.color) css.push('color:' + st.color); if (st.bg) css.push('background-color:' + st.bg);
    if (st.ff) css.push(`font-family:${st.ff}`); if (st.fs) css.push(`font-size:${st.fs}pt`);
    if (st.ha) css.push('text-align:' + st.ha); if (st.va) css.push('vertical-align:' + st.va);
    if (st.bt) css.push('border-top:' + cssBorder(st.bt)); if (st.bb) css.push('border-bottom:' + cssBorder(st.bb));
    if (st.bl) css.push('border-left:' + cssBorder(st.bl)); if (st.br) css.push('border-right:' + cssBorder(st.br));
    const t = String(x.text ?? '').replace(/[&<>]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch])).replace(/\n/g, '<br>');
    return `<td style="${css.join(';')}">${t}</td>`;
  };
  return `<meta charset="utf-8"><google-sheets-html-origin><table data-amatu="${stamp}" style="border-collapse:collapse">` + clip.rows.map(r => '<tr>' + r.map(td).join('') + '</tr>').join('') + '</table>';
}

export function parseTSV(text) {
  const rows = [];
  let row = [], cur = '', q = false;
  const t = text.replace(/\r\n?/g, '\n');
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) {
      if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
      continue;
    }
    if (ch === '"' && cur === '') { q = true; continue; }
    if (ch === '\t') { row.push(cur); cur = ''; continue; }
    if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; continue; }
    cur += ch;
  }
  row.push(cur);
  if (!(row.length === 1 && row[0] === '' && rows.length)) rows.push(row);
  return rows;
}

export function parseHTMLTable(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const out = [];
  for (const tr of table.querySelectorAll('tr')) {
    const row = [];
    for (const td of tr.querySelectorAll('td,th')) {
      const s = td.getAttribute('style') || '';
      const st = {};
      const get = p => (new RegExp('(?:^|;)\\s*' + p + '\\s*:\\s*([^;]+)', 'i').exec(s) || [])[1]?.trim();
      const fw = get('font-weight');
      if (fw && (fw === 'bold' || +fw >= 600) || td.querySelector('b,strong') || td.tagName === 'TH') st.b = true;
      if (get('font-style') === 'italic' || td.querySelector('i,em')) st.i = true;
      const td2 = get('text-decoration') || '';
      if (/underline/.test(td2) || td.querySelector('u')) st.u = true;
      if (/line-through/.test(td2) || td.querySelector('s,strike')) st.s = true;
      const color = get('color'); if (color && !/^(#000000|black|rgb\(0, ?0, ?0\)|windowtext)$/i.test(color)) st.color = toHex(color);
      const bg = get('background-color') || get('background'); if (bg && !/^(transparent|#ffffff|white|rgb\(255, ?255, ?255\)|initial)$/i.test(bg)) st.bg = toHex(bg);
      const ta = get('text-align'); if (ta && /left|center|right/.test(ta)) st.ha = ta;
      const text = td.innerText !== undefined ? td.innerText : td.textContent;
      const span = +(td.getAttribute('colspan') || 1);
      row.push({ text: (text ?? td.textContent).replace(/ /g, ' ').replace(/\n$/, ''), st });
      for (let k = 1; k < span; k++) row.push({ text: '', st: {} });
    }
    out.push(row);
  }
  return out.length ? out : null;
}
function toHex(c) {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(c);
  if (m) return '#' + [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join('');
  return c;
}

// mode: 'all' | 'values' | 'formats' | 'formulas' | 'transpose' | 'nofmt' | 'widths' | 'cond' | 'valid'; op: add/sub/mul/div
export function pasteClip(app, clip, dest, mode = 'all', opts = {}) {
  const sh = app.sheet, wb = app.wb;
  let rows = clip.rows;
  if (mode === 'transpose' || opts.transpose) rows = rows[0].map((_, j) => rows.map(r => r[j]));
  const R = rows.length, C = rows[0].length;
  // tile into a larger selection when it is an exact multiple
  let TR = R, TC = C;
  const dR = dest.r2 - dest.r1 + 1, dC = dest.c2 - dest.c1 + 1;
  if (dR % R === 0 && dC % C === 0 && (dR > R || dC > C) && !opts.noTile) { TR = dR; TC = dC; }
  for (let i = 0; i < TR; i++) {
    for (let j = 0; j < TC; j++) {
      const x = rows[i % R][j % C];
      const r = dest.r1 + i, c = dest.c1 + j;
      if (!x) continue;
      if (mode === 'widths') continue;
      const cur = sh.get(r, c);
      const curSt = wb.styles[cur ? cur.s : 0] || {};
      const sIdx = wb.intern(x.st || {});
      if (mode === 'formats') { sh.put(r, c, { s: sIdx }); continue; }
      if (mode === 'valid') continue;
      if (mode === 'cond') continue;
      let v = x.v, f = x.f;
      if (x.raw !== undefined) {
        const p = parseInput(x.raw);
        v = p.v ?? null; f = p.f ?? null;
        if (p.fmt && !x.st?.fmt) x.st = { ...x.st, fmt: p.fmt };
      }
      if (f != null && clip.sheetId !== undefined && !clip.cut) f = shiftFormula(f, r - x.r, c - x.c);
      if (mode === 'values') { f = null; v = x.val instanceof FErr ? x.val.code : x.val ?? x.v; if (v && typeof v === 'object') v = null; }
      if (mode === 'formulas' && x.f == null) { /* plain values pass through */ }
      if (opts.op && typeof (x.val ?? x.v) === 'number') {
        const a = typeof wb.eng.value(sh, r, c) === 'number' ? wb.eng.value(sh, r, c) : 0, b = x.val ?? x.v;
        v = opts.op === 'add' ? a + b : opts.op === 'sub' ? a - b : opts.op === 'mul' ? a * b : b ? a / b : a;
        f = null;
      }
      if (opts.skipBlanks && v == null && f == null) continue;
      const patch = { v: f != null ? null : v, f: f ?? null };
      if (mode === 'all' || mode === 'transpose') { patch.s = wb.intern({ ...(x.st || {}) }); if (x.note !== undefined) patch.note = x.note; if (x.link !== undefined) patch.link = x.link; }
      else if (mode === 'nofmt') { /* keep destination formatting */ }
      else if (mode === 'values' || mode === 'formulas') {
        if (x.st && x.st.fmt && !curSt.fmt) patch.s = wb.intern({ ...curSt, fmt: x.st.fmt });
      }
      sh.put(r, c, patch);
    }
  }
  if ((mode === 'all' || mode === 'widths') && clip.colW && !opts.transpose && mode !== 'transpose') {
    if (mode === 'widths' || clip.wholeCols) {
      wb.hist.touchMeta(sh);
      clip.colW.forEach((w, j) => { if (w) sh.colW.set(dest.c1 + j, w); else sh.colW.delete(dest.c1 + j); });
      sh.layoutV++;
    }
  }
  if ((mode === 'all' || mode === 'valid') && clip.sheetId === sh.id) {
    // carry data validation of the source cells
    const seen = new Set();
    rows.flat().forEach(x => { if (x && x.dv && !seen.has(x.dv.id)) seen.add(x.dv.id); });
    if (seen.size) {
      wb.hist.touchMeta(sh);
      for (const id of seen) {
        const rule = sh.dv.find(d => d.id === id);
        if (rule) rule.ranges.push({ r1: dest.r1, c1: dest.c1, r2: dest.r1 + TR - 1, c2: dest.c1 + TC - 1 });
      }
    }
  }
  return { r1: dest.r1, c1: dest.c1, r2: dest.r1 + TR - 1, c2: dest.c1 + TC - 1 };
}

// ---------- sorting & filtering ----------
export function sortRange(app, g, keys, header = false) {
  const sh = app.sheet, eng = app.wb.eng;
  const r1 = header ? g.r1 + 1 : g.r1;
  const rows = [];
  for (let r = r1; r <= g.r2; r++) {
    if (sh.filteredR.has(r)) continue;
    rows.push({ r, vals: keys.map(k => eng.value(sh, r, k.c)), snap: snapRange(sh, { r1: r, c1: g.c1, r2: r, c2: g.c2 })[0] });
  }
  const targets = rows.map(x => x.r);
  rows.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      let x = a.vals[i], y = b.vals[i];
      const bx = x == null || x === '', by = y == null || y === '';
      if (bx && by) continue;
      if (bx) return 1;
      if (by) return -1;
      if (x instanceof FErr) x = x.code; if (y instanceof FErr) y = y.code;
      const c = compare(x, y);
      if (c) return keys[i].asc === false ? -c : c;
    }
    return a.r - b.r;
  });
  rows.forEach((x, i) => writeSnap(sh, targets[i], g.c1, [x.snap], [targets[i] - x.r, 0]));
}

export function filterRows(app) {
  const sh = app.sheet, f = sh.filter, eng = app.wb.eng;
  const old = sh.filteredR;
  sh.filteredR = new Set();
  if (f) {
    const end = Math.max(f.r2, sh.dataRows() - 1);
    for (let r = f.r1 + 1; r <= end; r++) {
      for (const [cs, crit] of Object.entries(f.crit || {})) {
        const c = +cs;
        const cell = sh.get(r, c);
        const d = app.grid.display(r, c, cell);
        const text = d ? String(d.text ?? '') : '';
        const v = eng.value(sh, r, c);
        if (crit.hidden && crit.hidden.includes(text)) { sh.filteredR.add(r); break; }
        if (crit.cond && !condMatch(crit.cond, v, text, eng, sh, r, c)) { sh.filteredR.add(r); break; }
      }
    }
  }
  if (old.size !== sh.filteredR.size || [...old].some(r => !sh.filteredR.has(r))) { app.wb.hist.touchMeta(sh); sh.layoutV++; }
}

export function condMatch(cond, v, text, eng, sh, r, c) {
  const t = text.toLowerCase(), q = String(cond.v1 ?? '').toLowerCase();
  const n = Number(String(cond.v1).replace(',', '.')), n2 = Number(String(cond.v2).replace(',', '.'));
  switch (cond.op) {
    case 'empty': return text === '';
    case 'notEmpty': return text !== '';
    case 'contains': return t.includes(q);
    case 'notContains': return !t.includes(q);
    case 'starts': return t.startsWith(q);
    case 'ends': return t.endsWith(q);
    case 'exact': return t === q;
    case 'gt': return typeof v === 'number' && v > n;
    case 'ge': return typeof v === 'number' && v >= n;
    case 'lt': return typeof v === 'number' && v < n;
    case 'le': return typeof v === 'number' && v <= n;
    case 'eq': return typeof v === 'number' ? v === n : t === q;
    case 'ne': return typeof v === 'number' ? v !== n : t !== q;
    case 'between': return typeof v === 'number' && v >= Math.min(n, n2) && v <= Math.max(n, n2);
    case 'notBetween': return typeof v === 'number' && (v < Math.min(n, n2) || v > Math.max(n, n2));
    case 'formula': { const res = eng.evalText(cond.v1 || '', sh, r, c); return res === true || (typeof res === 'number' && res !== 0); }
  }
  return true;
}

export function removeDuplicates(app, g, cols, header) {
  const sh = app.sheet, eng = app.wb.eng;
  const r1 = header ? g.r1 + 1 : g.r1;
  const seen = new Set(), keep = [];
  let removed = 0;
  for (let r = r1; r <= g.r2; r++) {
    const k = cols.map(c => { const v = eng.value(sh, r, c); return v == null ? '' : typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + v; }).join('\u0001');
    if (seen.has(k)) { removed++; continue; }
    seen.add(k);
    keep.push(snapRange(sh, { r1: r, c1: g.c1, r2: r, c2: g.c2 })[0]);
  }
  const W = g.c2 - g.c1 + 1;
  const all = [...keep, ...Array.from({ length: g.r2 - r1 + 1 - keep.length }, () => Array(W).fill(null))];
  writeSnap(sh, r1, g.c1, all);
  return { removed, left: keep.length };
}

export function trimRange(app, ranges) {
  const sh = app.sheet;
  let n = 0;
  for (const g of ranges) rangeCells(sh, clampG(sh, g), (r, c, cell) => {
    if (cell && typeof cell.v === 'string' && cell.f == null) {
      const t = cell.v.replace(/[ \t ]+/g, ' ').trim();
      if (t !== cell.v) { setInput(app.wb, sh, r, c, t); n++; }
    }
  });
  return n;
}

export function splitToColumns(app, g, delim) {
  const sh = app.sheet;
  const c = g.c1;
  for (let r = g.r1; r <= g.r2; r++) {
    const cell = sh.get(r, c);
    if (!cell || typeof cell.v !== 'string') continue;
    let parts;
    if (delim === 'auto') {
      const t = cell.v;
      const d = [',', ';', '\t', '|', ' '].find(x => t.includes(x)) || ',';
      parts = t.split(d);
    } else parts = cell.v.split(delim);
    parts.forEach((p, j) => setInput(app.wb, sh, r, c + j, p.trim()));
  }
}

export function changeCase(app, ranges, mode) {
  const sh = app.sheet;
  for (const g of ranges) rangeCells(sh, clampG(sh, g), (r, c, cell) => {
    if (!cell || typeof cell.v !== 'string' || cell.f != null) return;
    const t = mode === 'upper' ? cell.v.toUpperCase() : mode === 'lower' ? cell.v.toLowerCase() : cell.v.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
    if (t !== cell.v) sh.put(r, c, { v: t });
  });
}

export function toValues(app, ranges) {
  const sh = app.sheet;
  for (const g of ranges) {
    const gg = clampG(sh, g);
    const vals = [];
    for (let r = gg.r1; r <= gg.r2; r++) for (let c = gg.c1; c <= gg.c2; c++) {
      const cell = sh.get(r, c);
      const sp = sh.spillMap.get(key(r, c));
      if (cell && cell.f != null) vals.push([r, c, cell.sv instanceof FErr ? cell.sv.code : (cell.sv && typeof cell.sv === 'object' ? null : cell.sv), cell.af]);
      else if (sp) vals.push([r, c, sp.v instanceof FErr ? sp.v.code : sp.v, null]);
    }
    for (const [r, c, v, af] of vals) {
      sh.put(r, c, { f: null, v });
      if (af) { const st = sh.styleOf(r, c); if (!st.fmt) sh.setStyle(r, c, { ...st, fmt: af }); }
    }
  }
}

// ---------- find & replace ----------
export function findAll(app, q, opts) {
  const out = [];
  if (!q) return out;
  let re;
  try { re = opts.regex ? new RegExp(q, opts.matchCase ? 'g' : 'gi') : new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), opts.matchCase ? 'g' : 'gi'); } catch { return out; }
  const sheets = opts.allSheets ? app.wb.sheets : [app.sheet];
  for (const sh of sheets) {
    const keys = [...sh.cells.keys()].sort((a, b) => a - b);
    for (const k of keys) {
      const cell = sh.cells.get(k);
      const r = keyR(k), c = keyC(k);
      let text;
      if (opts.formulas && cell.f != null) text = '=' + cell.f;
      else {
        const v = cell.f != null ? cell.sv : cell.v;
        if (v == null) continue;
        const st = app.wb.styles[cell.s] || {};
        text = v instanceof FErr ? v.code : format(v, st.fmt || cell.af).t;
      }
      re.lastIndex = 0;
      const ok = opts.whole ? new RegExp('^(?:' + re.source + ')$', re.flags.replace('g', '')).test(text) : re.test(text);
      if (ok) out.push({ sh, r, c });
    }
  }
  return out;
}

export function replaceIn(app, hits, q, rep, opts) {
  const flags = opts.matchCase ? 'g' : 'gi';
  const re = opts.regex ? new RegExp(q, flags) : new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
  let n = 0;
  for (const h of hits) {
    const cell = h.sh.get(h.r, h.c);
    if (!cell) continue;
    if (cell.f != null) {
      if (!opts.formulas) continue;
      const t = '=' + cell.f;
      const nt = opts.whole ? (new RegExp('^(?:' + re.source + ')$', flags.replace('g', '')).test(t) ? rep : t) : t.replace(re, rep);
      if (nt !== t) { setInput(app.wb, h.sh, h.r, h.c, nt); n++; }
      continue;
    }
    const st = app.wb.styles[cell.s] || {};
    const t = editText(cell.v, st.fmt);
    const nt = opts.whole ? rep : t.replace(re, rep);
    if (nt !== t) { setInput(app.wb, h.sh, h.r, h.c, nt); n++; }
  }
  return n;
}

// ---------- goal seek ----------
export function goalSeek(app, sh, target, goal, changing) {
  const eng = app.wb.eng;
  const get = () => { eng.recalc(); const v = eng.value(sh, target.r, target.c); return typeof v === 'number' ? v : NaN; };
  const setX = x => sh.put(changing.r, changing.c, { v: x, f: null });
  const cur = eng.value(sh, changing.r, changing.c);
  let x0 = typeof cur === 'number' ? cur : 0;
  let x1 = x0 === 0 ? 1 : x0 * 1.1;
  setX(x0); let f0 = get() - goal;
  setX(x1); let f1 = get() - goal;
  for (let i = 0; i < 100; i++) {
    if (!isFinite(f1)) break;
    if (Math.abs(f1) < 1e-9) return { ok: true, x: x1, value: f1 + goal, iter: i };
    const d = f1 - f0;
    if (d === 0) { x1 += 1e-3 * (Math.abs(x1) || 1); setX(x1); f1 = get() - goal; continue; }
    const x2 = x1 - f1 * (x1 - x0) / d;
    x0 = x1; f0 = f1; x1 = x2;
    setX(x1); f1 = get() - goal;
  }
  return { ok: Math.abs(f1) < 1e-6, x: x1, value: f1 + goal };
}

// ---------- pivot ----------
export function buildPivot(app, def) {
  const wb = app.wb;
  const src = wb.sheetByName(def.srcSheet) || app.sheet;
  const g = def.range;
  const eng = wb.eng;
  const head = [];
  for (let c = g.c1; c <= g.c2; c++) { const v = eng.value(src, g.r1, c); head.push(v == null || v === '' ? 'Kolom ' + colName(c) : String(v)); }
  const end = Math.min(g.r2, src.dataRows() - 1);
  const data = [];
  for (let r = g.r1 + 1; r <= end; r++) {
    const row = [];
    let any = false;
    for (let c = g.c1; c <= g.c2; c++) { const v = eng.value(src, r, c); if (v != null && v !== '') any = true; row.push(v instanceof FErr ? null : v); }
    if (any) data.push(row);
  }
  const keyOf = v => (v == null || v === '' ? '(kosong)' : typeof v === 'number' ? v : String(v));
  const rowF = def.rows || [], colF = def.cols || [], vals = def.values?.length ? def.values : [{ col: 0, agg: 'COUNTA' }];
  const filt = def.filters || [];
  const rowsData = data.filter(row => filt.every(f => !f.hidden || !f.hidden.includes(String(keyOf(row[f.col])))));
  const agg = (fn, arr) => {
    const nums = arr.filter(v => typeof v === 'number');
    switch (fn) {
      case 'SUM': return nums.reduce((a, b) => a + b, 0);
      case 'COUNT': return nums.length;
      case 'COUNTA': return arr.filter(v => v != null && v !== '').length;
      case 'AVERAGE': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
      case 'MAX': return nums.length ? Math.max(...nums) : null;
      case 'MIN': return nums.length ? Math.min(...nums) : null;
      case 'COUNTUNIQUE': return new Set(arr.filter(v => v != null && v !== '').map(String)).size;
      case 'MEDIAN': { const s = nums.slice().sort((a, b) => a - b); const m = s.length >> 1; return s.length ? (s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2) : null; }
    }
    return null;
  };
  const sortKeys = (a, b) => compare(a[0], b[0]);
  const rowKeys = new Map(), colKeys = new Map();
  for (const row of rowsData) {
    const rk = rowF.map(f => keyOf(row[f])), ck = colF.map(f => keyOf(row[f]));
    rowKeys.set(JSON.stringify(rk), rk); colKeys.set(JSON.stringify(ck), ck);
  }
  const rk = [...rowKeys.values()].sort((a, b) => { for (let i = 0; i < a.length; i++) { const c = compare(a[i], b[i]); if (c) return def.rowDesc ? -c : c; } return 0; });
  const ck = [...colKeys.values()].sort((a, b) => { for (let i = 0; i < a.length; i++) { const c = compare(a[i], b[i]); if (c) return c; } return 0; });
  const groupVals = (pred, v) => rowsData.filter(pred).map(row => row[v.col]);
  const matchR = k => row => rowF.every((f, i) => keyOf(row[f]) === k[i]);
  const matchC = k => row => colF.every((f, i) => keyOf(row[f]) === k[i]);
  const valLabel = v => `${{ SUM: 'SUM', COUNT: 'COUNT', COUNTA: 'COUNTA', AVERAGE: 'AVERAGE', MAX: 'MAX', MIN: 'MIN', COUNTUNIQUE: 'COUNTUNIQUE', MEDIAN: 'MEDIAN' }[v.agg]} dari ${head[v.col]}`;
  const out = [];
  const hdr = rowF.map(f => head[f]);
  const showTotals = def.totals !== false;
  if (colF.length) {
    const top = [...rowF.map(() => ''), ...ck.flatMap(k => vals.map(() => k.join(' · '))), ...(showTotals ? vals.map(() => 'Total Keseluruhan') : [])];
    top[0] = colF.map(f => head[f]).join(' / ');
    out.push(top);
  }
  out.push([...hdr, ...(colF.length ? ck.flatMap(() => vals.map(valLabel)) : vals.map(valLabel)), ...(colF.length && showTotals ? vals.map(valLabel) : [])]);
  for (const k of rk) {
    const pr = matchR(k);
    const line = [...k];
    if (colF.length) {
      for (const c of ck) for (const v of vals) line.push(agg(v.agg, groupVals(row => pr(row) && matchC(c)(row), v)));
      if (showTotals) for (const v of vals) line.push(agg(v.agg, groupVals(pr, v)));
    } else for (const v of vals) line.push(agg(v.agg, groupVals(pr, v)));
    out.push(line);
  }
  if (showTotals) {
    const line = rowF.length ? ['Total Keseluruhan', ...rowF.slice(1).map(() => '')] : [];
    if (colF.length) {
      for (const c of ck) for (const v of vals) line.push(agg(v.agg, groupVals(matchC(c), v)));
      for (const v of vals) line.push(agg(v.agg, groupVals(() => true, v)));
    } else for (const v of vals) line.push(agg(v.agg, groupVals(() => true, v)));
    out.push(line);
  }
  return { table: out, headRows: colF.length ? 2 : 1, rowFields: rowF.length };
}

export function writePivot(app, sh, def) {
  const { table, headRows, rowFields } = buildPivot(app, def);
  const wbS = app.wb;
  const wb = app.wb;
  // clear previous output
  for (const [k, c] of [...sh.cells]) sh.put(keyR(k), keyC(k), { v: null, f: null, s: 0 });
  const allInt = table.every(row => row.every(v => typeof v !== 'number' || Number.isInteger(v)));
  const nf = allInt ? '#,##0' : '#,##0.00';
  const hs = wb.intern({ b: true, bg: '#e8f0fe' }), ts = wb.intern({ b: true, bg: '#f1f3f4' }), ns = wb.intern({ fmt: nf }), tns = wb.intern({ b: true, bg: '#f1f3f4', fmt: nf });
  const r0 = def.at?.r || 0, c0 = def.at?.c || 0;
  table.forEach((row, i) => row.forEach((v, j) => {
    const isHead = i < headRows, isTot = def.totals !== false && i === table.length - 1;
    const s = isHead ? hs : isTot ? (typeof v === 'number' ? tns : ts) : typeof v === 'number' ? ns : j < rowFields ? 0 : 0;
    sh.put(r0 + i, c0 + j, { v: v ?? null, f: null, s });
  }));
  app.wb.hist.touchMeta(sh);
  sh.pivot = def;
  sh.growTo(r0 + table.length + 2, Math.max(sh.cols, c0 + table[0].length + 1));
  return { rows: table.length, cols: table[0].length };
}

// ---------- validation helpers ----------
export function validateCell(app, sh, r, c, v) {
  const rule = dvAt(sh, r, c);
  if (!rule) return { ok: true };
  const ok = dvValid(sh, rule, r, c, v);
  return { ok, rule };
}

export function markInvalid(app, sh) {
  for (const cell of sh.cells.values()) cell.bad = null;
  if (!sh.dv.length) return;
  const eng = app.wb.eng;
  for (const rule of sh.dv) {
    if (rule.type === 'checkbox') continue;
    for (const g of rule.ranges) {
      const r2 = Math.min(g.r2, sh.dataRows() - 1), c2 = Math.min(g.c2, sh.cols - 1);
      for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) {
        const cell = sh.get(r, c);
        if (!cell) continue;
        const v = eng.value(sh, r, c);
        if (v == null || v === '') continue;
        if (!dvValid(sh, rule, r, c, v)) cell.bad = rule.help || 'Nilai tidak memenuhi aturan validasi data.';
      }
    }
  }
}

export { snapRange, writeSnap, clampG };
