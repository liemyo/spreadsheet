// Workbook / sheet data model, undo history, serialization and structural edits.
import { Engine } from './formula/engine.js';
import { mapRefs } from './formula/parse.js';
import { parseInput, compile } from './format.js';
import { key, keyR, keyC, uid, MAXC, intersects } from './util.js';

export const COL_W = 100;
export const ROW_H = 21;

export class Cell {
  constructor() {
    this.v = null; this.f = null; this.s = 0; this.note = null; this.link = null;
    this.ast = undefined; this.sv = null; this.val = null; this.cid = 0; this.af = null; this.hl = null;
  }
  get empty() { return this.v == null && this.f == null && !this.s && !this.note && !this.link; }
}

const snapCell = c => (c ? { v: c.v, f: c.f, s: c.s, note: c.note, link: c.link } : null);
const clone = o => (o == null ? o : JSON.parse(JSON.stringify(o)));

export class Sheet {
  constructor(wb, name, rows = 1000, cols = 26, id = uid('s')) {
    this.wb = wb; this.id = id; this.name = name;
    this.rows = rows; this.cols = cols;
    this.cells = new Map();
    this.colW = new Map(); this.rowH = new Map(); this.manualH = new Set();
    this.hiddenR = new Set(); this.hiddenC = new Set(); this.filteredR = new Set();
    this.merges = []; this.mergeIx = null;
    this.freezeR = 0; this.freezeC = 0;
    this.cf = []; this.dv = []; this.filter = null; this.charts = []; this.banding = [];
    this.tab = null; this.grid = true; this.protect = false; this.hidden = false; this.pivot = null;
    // protectExcept: cells left unlocked in a protected sheet; lock: {salt, hash} password; ve: cells viewers of a share link may edit
    this.protectExcept = []; this.lock = null; this.ve = [];
    // page: print setup (paper, margins, scale, page breaks, keep-together ranges) — see pages.js
    this.page = null;
    this.spillMap = new Map(); this.spillPrev = new Map(); this.spillOrigins = new Map();
    this.maxR = -1; this.maxC = -1;
    this.layoutV = 0;
  }

  get(r, c) { return this.cells.get(key(r, c)); }
  colWidth(c) { return this.hiddenC.has(c) ? 0 : this.colW.get(c) ?? COL_W; }
  rowHeight(r) { return this.hiddenR.has(r) || this.filteredR.has(r) ? 0 : this.rowH.get(r) ?? ROW_H; }
  isRowHidden(r) { return this.hiddenR.has(r) || this.filteredR.has(r); }

  dataRows() {
    let m = this.maxR;
    for (const [k, s] of this.spillOrigins) m = Math.max(m, keyR(k) + s.rows - 1);
    return m + 1;
  }
  dataCols() {
    let m = this.maxC;
    for (const [k, s] of this.spillOrigins) m = Math.max(m, keyC(k) + s.cols - 1);
    return m + 1;
  }
  recomputeExtent() {
    this.maxR = -1; this.maxC = -1;
    for (const [k, c] of this.cells) {
      if (c.v == null && c.f == null) continue;
      const r = keyR(k), cc = keyC(k);
      if (r > this.maxR) this.maxR = r;
      if (cc > this.maxC) this.maxC = cc;
    }
  }
  // last row/col with any content or style (for Ctrl+End, export)
  usedRange() {
    let r2 = -1, c2 = -1;
    for (const [k, c] of this.cells) {
      if (c.empty) continue;
      r2 = Math.max(r2, keyR(k)); c2 = Math.max(c2, keyC(k));
    }
    for (const [k, s] of this.spillOrigins) { r2 = Math.max(r2, keyR(k) + s.rows - 1); c2 = Math.max(c2, keyC(k) + s.cols - 1); }
    return r2 < 0 ? null : { r1: 0, c1: 0, r2, c2 };
  }

  growTo(rows, cols) {
    if (rows > this.rows) { this.rows = Math.min(rows, 1048576); this.layoutV++; }
    if (cols > this.cols) { this.cols = Math.min(cols, MAXC); this.layoutV++; }
  }

  // Low-level content change, history-tracked.
  put(r, c, patch) {
    const k = key(r, c);
    this.wb.hist.touchCell(this, k);
    let cell = this.cells.get(k);
    if (!cell) { cell = new Cell(); this.cells.set(k, cell); }
    if ('f' in patch && patch.f !== cell.f) { cell.f = patch.f; cell.ast = undefined; cell.cid = 0; cell.val = null; cell.sv = null; }
    if ('v' in patch) cell.v = patch.v;
    if (cell.f != null && 'f' in patch) cell.v = null;
    if ('s' in patch) cell.s = patch.s || 0;
    if ('note' in patch) cell.note = patch.note || null;
    if ('link' in patch) cell.link = patch.link || null;
    if (cell.empty) this.cells.delete(k);
    else if (cell.v != null || cell.f != null) {
      if (r > this.maxR) this.maxR = r;
      if (c > this.maxC) this.maxC = c;
      this.growTo(r + 1, c + 1);
    }
    this.wb.dirty = true;
    return cell;
  }
  putSnap(k, snap) {
    if (!snap) { this.cells.delete(k); return; }
    const cell = new Cell();
    Object.assign(cell, { v: snap.v, f: snap.f, s: snap.s || 0, note: snap.note || null, link: snap.link || null });
    this.cells.set(k, cell);
    const r = keyR(k), c = keyC(k);
    if (r > this.maxR) this.maxR = r;
    if (c > this.maxC) this.maxC = c;
  }

  styleOf(r, c) { const cell = this.cells.get(key(r, c)); return this.wb.styles[cell ? cell.s : 0] || {}; }
  setStyle(r, c, st) { const idx = this.wb.intern(st); const cell = this.get(r, c); if ((cell ? cell.s : 0) !== idx) this.put(r, c, { s: idx }); }

  // ---- merges ----
  mergeAt(r, c) {
    if (!this.merges.length) return null;
    if (!this.mergeIx) {
      this.mergeIx = new Map();
      for (const m of this.merges) for (let i = m.r1; i <= m.r2; i++) for (let j = m.c1; j <= m.c2; j++) this.mergeIx.set(key(i, j), m);
    }
    return this.mergeIx.get(key(r, c)) || null;
  }
  setMerges(list) { this.wb.hist.touchMeta(this); this.merges = list; this.mergeIx = null; this.layoutV++; }

  metaSnapshot() {
    return clone({
      name: this.name, rows: this.rows, cols: this.cols, colW: [...this.colW], rowH: [...this.rowH], manualH: [...this.manualH],
      hiddenR: [...this.hiddenR], hiddenC: [...this.hiddenC], filteredR: [...this.filteredR], merges: this.merges,
      freezeR: this.freezeR, freezeC: this.freezeC, cf: this.cf, dv: this.dv, filter: this.filter, charts: this.charts,
      banding: this.banding, tab: this.tab, grid: this.grid, protect: this.protect, hidden: this.hidden, pivot: this.pivot,
      protectExcept: this.protectExcept, lock: this.lock, ve: this.ve, page: this.page,
    });
  }
  loadMeta(m) {
    m = clone(m);
    this.name = m.name; this.rows = m.rows; this.cols = m.cols;
    this.colW = new Map(m.colW); this.rowH = new Map(m.rowH); this.manualH = new Set(m.manualH || []);
    this.hiddenR = new Set(m.hiddenR); this.hiddenC = new Set(m.hiddenC); this.filteredR = new Set(m.filteredR || []);
    this.merges = m.merges || []; this.mergeIx = null;
    this.freezeR = m.freezeR || 0; this.freezeC = m.freezeC || 0;
    this.cf = m.cf || []; this.dv = m.dv || []; this.filter = m.filter || null; this.charts = m.charts || [];
    this.banding = m.banding || []; this.tab = m.tab || null; this.grid = m.grid !== false; this.protect = !!m.protect;
    this.hidden = !!m.hidden; this.pivot = m.pivot || null;
    this.protectExcept = m.protectExcept || []; this.lock = m.lock || null; this.ve = m.ve || [];
    this.page = m.page || null;
    this.layoutV++;
  }
  toJSON() {
    const cells = [];
    for (const [k, c] of this.cells) {
      const row = [k, c.v, c.f, c.s];
      if (c.note || c.link) row.push(c.note, c.link);
      cells.push(row);
    }
    return { id: this.id, ...this.metaSnapshot(), cells };
  }
  loadJSON(j) {
    this.loadMeta(j);
    this.cells = new Map();
    for (const [k, v, f, s, note, link] of j.cells || []) this.putSnap(k, { v, f, s, note, link });
    this.recomputeExtent();
  }
}

export class History {
  constructor(wb) { this.wb = wb; this.undoS = []; this.redoS = []; this.cur = null; this.depth = 0; this.paused = 0; }
  begin(label, sel) {
    if (this.depth++ === 0) this.cur = { label, cells: new Map(), meta: new Map(), full: new Map(), wb: null, sel };
  }
  end(selAfter) {
    if (this.depth === 0) return;
    if (--this.depth > 0) return;
    const e = this.cur;
    this.cur = null;
    if (!e.cells.size && !e.meta.size && !e.full.size && !e.wb) return;
    e.selAfter = selAfter;
    this.wb.dirty = true;   // meta-only changes (tab colour, freeze, protection…) must be saved too
    this.undoS.push(e);
    if (this.undoS.length > 300) this.undoS.shift();
    this.redoS = [];
  }
  get active() { return !!this.cur && !this.paused; }
  touchCell(sh, k) {
    if (!this.active) return;
    // a full snapshot of this sheet already holds the state before the transaction (insert/delete rows then
    // rewriting formulas would otherwise restore moved cells over it on undo)
    if (this.cur.full.has(sh.id)) return;
    let m = this.cur.cells.get(sh.id);
    if (!m) { m = new Map(); this.cur.cells.set(sh.id, m); }
    if (!m.has(k)) m.set(k, snapCell(sh.cells.get(k)));
  }
  touchMeta(sh) { if (this.active && !this.cur.meta.has(sh.id)) this.cur.meta.set(sh.id, sh.metaSnapshot()); }
  touchFull(sh) { if (this.active && !this.cur.full.has(sh.id)) this.cur.full.set(sh.id, sh.toJSON()); }
  touchWb() { if (this.active && !this.cur.wb) this.cur.wb = this.wb.metaSnapshot(); }

  capture(e) {
    const wb = this.wb;
    const inv = { label: e.label, cells: new Map(), meta: new Map(), full: new Map(), wb: null, sel: e.selAfter, selAfter: e.sel };
    if (e.wb) inv.wb = wb.metaSnapshot();
    for (const id of e.full.keys()) { const sh = wb.sheetById(id); if (sh) inv.full.set(id, sh.toJSON()); }
    for (const id of e.meta.keys()) { const sh = wb.sheetById(id); if (sh) inv.meta.set(id, sh.metaSnapshot()); }
    for (const [id, m] of e.cells) {
      const sh = wb.sheetById(id); if (!sh) continue;
      const cm = new Map();
      for (const k of m.keys()) cm.set(k, snapCell(sh.cells.get(k)));
      inv.cells.set(id, cm);
    }
    // sheets that exist now but will vanish need full snapshots for redo
    if (e.wb) for (const sh of wb.sheets) if (!e.wb.order.includes(sh.id) && !inv.full.has(sh.id)) inv.full.set(sh.id, sh.toJSON());
    return inv;
  }
  apply(e) {
    const wb = this.wb;
    if (e.wb) wb.restoreMeta(e.wb, e.full);
    for (const [id, j] of e.full) { const sh = wb.sheetById(id); if (sh) sh.loadJSON(j); }
    for (const [id, m] of e.meta) { const sh = wb.sheetById(id); if (sh) sh.loadMeta(m); }
    for (const [id, cm] of e.cells) {
      const sh = wb.sheetById(id); if (!sh) continue;
      for (const [k, snap] of cm) sh.putSnap(k, snap);
      sh.recomputeExtent();
    }
    wb.dirty = true;
  }
  undo() {
    const e = this.undoS.pop();
    if (!e) return null;
    const inv = this.capture(e);
    this.apply(e);
    this.redoS.push(inv);
    return e;
  }
  redo() {
    const e = this.redoS.pop();
    if (!e) return null;
    const inv = this.capture(e);
    this.apply(e);
    this.undoS.push(inv);
    return e;
  }
}

export class Workbook {
  constructor() {
    this.id = uid('wb');
    this.name = 'Spreadsheet tanpa judul';
    this.sheets = [];
    this.styles = [{}];
    this.styleMap = new Map([['{}', 0]]);
    this.names = new Map();
    this.settings = { locale: 'id-ID' };
    this.active = 0;
    this.eng = new Engine(this);
    this.hist = new History(this);
    this.dirty = false;
  }
  sheetByName(n) { const u = String(n).toLowerCase(); return this.sheets.find(s => s.name.toLowerCase() === u) || null; }
  sheetById(id) { return this.sheets.find(s => s.id === id) || null; }
  uniqueSheetName(base = 'Sheet') {
    let i = this.sheets.length + 1;
    while (this.sheetByName(base + i)) i++;
    return base + i;
  }
  addSheet(name, at = this.sheets.length, rows, cols) {
    this.hist.touchWb();
    const sh = new Sheet(this, name || this.uniqueSheetName(), rows, cols);
    this.sheets.splice(at, 0, sh);
    return sh;
  }
  intern(st) {
    const clean = {};
    for (const k of Object.keys(st).sort()) { const v = st[k]; if (v !== undefined && v !== null && v !== '' && v !== false) clean[k] = v; }
    const j = JSON.stringify(clean);
    let idx = this.styleMap.get(j);
    if (idx === undefined) { idx = this.styles.length; this.styles.push(Object.freeze(clean)); this.styleMap.set(j, idx); }
    return idx;
  }
  metaSnapshot() {
    return clone({ order: this.sheets.map(s => s.id), names: [...this.names.values()].map(n => ({ name: n.name, ref: n.ref, note: n.note })), active: this.active, name: this.name });
  }
  restoreMeta(m, full) {
    const byId = new Map(this.sheets.map(s => [s.id, s]));
    this.sheets = m.order.map(id => {
      let sh = byId.get(id);
      if (!sh && full && full.get(id)) { const j = full.get(id); sh = new Sheet(this, j.name, j.rows, j.cols, id); sh.loadJSON(j); }
      return sh;
    }).filter(Boolean);
    this.names = new Map(m.names.map(n => [n.name.toUpperCase(), { ...n }]));
    this.active = Math.min(m.active, this.sheets.length - 1);
    this.name = m.name;
  }
  toJSON() {
    return {
      app: 'amatu-sheets', v: 1, id: this.id, name: this.name, active: this.active, settings: this.settings,
      styles: this.styles, names: [...this.names.values()].map(n => ({ name: n.name, ref: n.ref, note: n.note })),
      sheets: this.sheets.map(s => s.toJSON()),
    };
  }
  static fromJSON(j) {
    const wb = new Workbook();
    wb.id = j.id || wb.id;
    wb.name = j.name || wb.name;
    wb.share = j.share || null;   // live share link state (store.loadDoc), never part of the document itself
    wb.lock = j.lock || null;     // file password (lock.js), stored beside the document like share
    wb.settings = { ...wb.settings, ...(j.settings || {}) };
    wb.styles = (j.styles || [{}]).map(s => Object.freeze({ ...s }));
    wb.styleMap = new Map(wb.styles.map((s, i) => [JSON.stringify(Object.fromEntries(Object.entries(s).sort())), i]));
    for (const n of j.names || []) wb.names.set(n.name.toUpperCase(), { ...n });
    for (const sj of j.sheets || []) { const sh = new Sheet(wb, sj.name, sj.rows, sj.cols, sj.id); sh.loadJSON(sj); wb.sheets.push(sh); }
    if (!wb.sheets.length) wb.addSheet('Sheet1');
    wb.active = Math.min(j.active || 0, wb.sheets.length - 1);
    return wb;
  }

  // ---- reference maintenance across the workbook ----
  eachFormula(fn) {
    for (const sh of this.sheets) for (const [k, c] of sh.cells) if (c.f != null) fn(sh, k, c);
  }
  rewriteFormulas(mapper) {
    this.eachFormula((sh, k, c) => {
      const nf = mapRefs(c.f, (a, sheetName) => mapper(a, sheetName, sh));
      if (nf !== c.f) sh.put(keyR(k), keyC(k), { f: nf });
    });
    for (const n of this.names.values()) {
      const t = String(n.ref).replace(/^=/, '');
      const nr = mapRefs(t, (a, sheetName) => mapper(a, sheetName, null));
      if (nr !== t) { this.hist.touchWb(); n.ref = nr; n.ast = undefined; }
    }
    for (const sh of this.sheets) {
      for (const ch of sh.charts) {
        if (!ch.range) continue;
        const nr = mapRefs(ch.range, (a, sheetName) => mapper(a, sheetName, sh));
        if (nr !== ch.range) { this.hist.touchMeta(sh); ch.range = nr; }
      }
    }
  }
  renameSheet(sh, name) {
    const old = sh.name;
    this.hist.touchMeta(sh);
    sh.name = name;
    this.rewriteFormulas((a, sheetName) => (sheetName != null && sheetName.toLowerCase() === old.toLowerCase() ? { ...a, sheet: name } : null));
  }
  deleteSheet(sh) {
    this.hist.touchWb();
    this.hist.touchFull(sh);
    const name = sh.name;
    this.sheets.splice(this.sheets.indexOf(sh), 1);
    this.rewriteFormulas((a, sheetName) => (sheetName != null && sheetName.toLowerCase() === name.toLowerCase() ? { ...a, err: true, sheet: null } : null));
    if (this.active >= this.sheets.length) this.active = this.sheets.length - 1;
  }

  // axis 'r' | 'c'; n > 0 inserts n before `at`, n < 0 deletes -n starting at `at`
  shiftRefs(target, axis, at, n) {
    this.rewriteFormulas((a, sheetName, fsh) => {
      const sh = sheetName != null ? this.sheetByName(sheetName) : fsh;
      if (sh !== target) return null;
      if (axis === 'r' && a.kind === 'col') return null;
      if (axis === 'c' && a.kind === 'row') return null;
      const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
      let p1 = a[k1], p2 = a[k2];
      if (n > 0) {
        if (p1 >= at) p1 += n;
        if (p2 !== Infinity && p2 >= at) p2 += n;
      } else {
        const d = -n, end = at + d - 1;
        if (p1 > end) p1 -= d; else if (p1 >= at) p1 = at;
        if (p2 !== Infinity) {
          if (p2 > end) p2 -= d; else if (p2 >= at) p2 = at - 1;
          if (p2 < p1) return { ...a, err: true };
        }
      }
      if (p1 === a[k1] && p2 === a[k2]) return null;
      return { ...a, [k1]: p1, [k2]: p2 };
    });
  }
}

// ---- range helpers for structural edits ----
function shiftRange(g, axis, at, n) {
  const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
  const o = { ...g };
  if (n > 0) {
    if (o[k1] >= at) o[k1] += n;
    if (o[k2] >= at) o[k2] += n;
  } else {
    const d = -n, end = at + d - 1;
    if (o[k1] > end) o[k1] -= d; else if (o[k1] >= at) o[k1] = at;
    if (o[k2] > end) o[k2] -= d; else if (o[k2] >= at) o[k2] = at - 1;
    if (o[k2] < o[k1]) return null;
  }
  return o;
}
const shiftMap = (m, at, n) => {
  const out = new Map();
  for (const [i, v] of m) {
    if (i < at) out.set(i, v);
    else if (n > 0) out.set(i + n, v);
    else if (i >= at - n) out.set(i + n, v);
  }
  return out;
};
const shiftSet = (s, at, n) => new Set([...shiftMap(new Map([...s].map(i => [i, 1])), at, n).keys()]);

export function insertDelete(wb, sh, axis, at, n) {
  wb.hist.touchFull(sh);
  const moved = new Map();
  for (const [k, c] of sh.cells) {
    let r = keyR(k), cc = keyC(k);
    let p = axis === 'r' ? r : cc;
    if (n < 0 && p >= at && p < at - n) continue;
    if (p >= at) p += n;
    if (axis === 'r') r = p; else cc = p;
    if (r < 0 || cc < 0) continue;
    moved.set(key(r, cc), c);
  }
  sh.cells = moved;
  if (axis === 'r') {
    sh.rowH = shiftMap(sh.rowH, at, n); sh.hiddenR = shiftSet(sh.hiddenR, at, n);
    sh.manualH = shiftSet(sh.manualH, at, n); sh.filteredR = shiftSet(sh.filteredR, at, n);
    sh.rows = Math.max(1, sh.rows + n);
    if (sh.freezeR > at) sh.freezeR = Math.max(at, sh.freezeR + n);
  } else {
    sh.colW = shiftMap(sh.colW, at, n); sh.hiddenC = shiftSet(sh.hiddenC, at, n);
    sh.cols = Math.max(1, sh.cols + n);
    if (sh.freezeC > at) sh.freezeC = Math.max(at, sh.freezeC + n);
  }
  const fix = list => list.map(g => shiftRange(g, axis, at, n)).filter(Boolean);
  sh.merges = fix(sh.merges).filter(m => m.r2 > m.r1 || m.c2 > m.c1);
  sh.mergeIx = null;
  for (const rule of sh.cf) rule.ranges = fix(rule.ranges);
  sh.cf = sh.cf.filter(r => r.ranges.length);
  for (const rule of sh.dv) rule.ranges = fix(rule.ranges);
  sh.dv = sh.dv.filter(r => r.ranges.length);
  sh.protectExcept = fix(sh.protectExcept || []);
  sh.ve = fix(sh.ve || []);
  if (sh.page) {
    const pg = sh.page, sh1 = i => (i >= at ? (n < 0 && i < at - n ? null : i + n) : i);
    if (axis === 'r') {
      pg.rb = (pg.rb || []).map(sh1).filter(i => i != null && i > 0);
      pg.keep = (pg.keep || []).map(k => shiftRange({ c1: 0, c2: 0, ...k }, axis, at, n)).filter(Boolean);
      if (pg.tr) pg.tr = shiftRange({ c1: 0, c2: 0, ...pg.tr }, axis, at, n);
    } else pg.cb = (pg.cb || []).map(sh1).filter(i => i != null && i > 0);
    if (pg.area) pg.area = shiftRange(pg.area, axis, at, n);
  }
  sh.banding = sh.banding.map(b => { const g = shiftRange(b, axis, at, n); return g ? { ...b, ...g } : null; }).filter(Boolean);
  if (sh.filter) { const g = shiftRange(sh.filter, axis, at, n); sh.filter = g ? { ...sh.filter, ...g } : null; }
  for (const ch of sh.charts) {
    const p = axis === 'r' ? 'r' : 'c';
    if (ch.anchor[p] >= at) ch.anchor[p] = Math.max(at, ch.anchor[p] + n);
  }
  sh.recomputeExtent();
  sh.layoutV++;
  wb.shiftRefs(sh, axis, at, n);
  wb.dirty = true;
}

export function rangeCells(sh, g, fn) {
  const r2 = Math.min(g.r2, sh.rows - 1), c2 = Math.min(g.c2, sh.cols - 1);
  const area = (r2 - g.r1 + 1) * (c2 - g.c1 + 1);
  if (area > sh.cells.size * 4 && area > 5000) {
    // sparse iteration
    const ks = [];
    for (const k of sh.cells.keys()) { const r = keyR(k), c = keyC(k); if (r >= g.r1 && r <= r2 && c >= g.c1 && c <= c2) ks.push(k); }
    for (const k of ks) fn(keyR(k), keyC(k), sh.cells.get(k));
    return;
  }
  for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) fn(r, c, sh.cells.get(key(r, c)));
}

export function setInput(wb, sh, r, c, text) {
  const cur = sh.get(r, c);
  const st = wb.styles[cur ? cur.s : 0] || {};
  if (st.fmt === '@' && text && text[0] !== '=') { sh.put(r, c, { v: text, f: null }); return; }
  const p = parseInput(text);
  if (p.f != null) { sh.put(r, c, { f: p.f, v: null }); return; }
  sh.put(r, c, { v: p.v, f: null });
  if (p.fmt && (!st.fmt || st.fmt === 'General' || (compile(st.fmt)?.isDate && !compile(p.fmt)?.isDate && typeof p.v === 'number' && false))) {
    sh.setStyle(r, c, { ...st, fmt: p.fmt });
  }
}

export function newWorkbook(name) {
  const wb = new Workbook();
  if (name) wb.name = name;
  wb.hist.paused++;
  wb.addSheet('Sheet1');
  wb.hist.paused--;
  return wb;
}

export { intersects };
