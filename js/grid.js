// Canvas grid: rendering, selection, pointer/keyboard interaction and the in-cell editor.
import { key, keyR, keyC, colName, clamp, inRange, el, norm, rangeAddr } from './util.js';
import { format, general } from './format.js';
import { FErr } from './formula/core.js';
import { Special } from './formula/functions.js';
import { lex, autoClose } from './formula/parse.js';
import { dvAt, dvItems, CHIP_COLORS } from './cf.js';

export const HH = 24;
export const REF_COLORS = ['#1a73e8', '#d93025', '#8430ce', '#188038', '#e37400', '#c5221f', '#129eaf', '#b06000'];
const FALLBACK = 'Arial, Helvetica, sans-serif';
const URL_RE = /^(https?:\/\/|www\.)[^\s]+$/i;

export function fontOf(st, def) {
  const px = ((st.fs || def.fs) * 4) / 3;
  return `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}${px}px "${st.ff || def.ff}", ${FALLBACK}`;
}

export class Grid {
  constructor(app, host) {
    this.app = app;
    this.host = host;
    this.canvas = el('canvas.grid-canvas');
    this.ctx = this.canvas.getContext('2d', { alpha: false });
    this.spacer = el('div.grid-spacer');
    this.scroller = el('div.grid-scroller', { tabIndex: -1 }, this.spacer);
    this.layer = el('div.grid-layer');
    this.hilite = el('div.cell-hilite');
    this.input = el('textarea.cell-input', { spellcheck: false, autocomplete: 'off', autocapitalize: 'off', 'aria-label': 'Editor sel' });
    this.addRows = el('div.add-rows',
      el('button.btn.sm', { type: 'button', onclick: () => this.app.cmd('addRows', +this.addRowsN.value || 1000) }, 'Tambah'),
      this.addRowsN = el('input.inp.sm', { type: 'number', value: 1000, min: 1, max: 100000, 'aria-label': 'Jumlah baris' }),
      el('span', 'baris lagi di bagian bawah'));
    this.spacer.append(this.addRows);
    host.append(this.canvas, this.scroller, this.layer, this.hilite, this.input);
    this.z = 1;
    this.dpr = window.devicePixelRatio || 1;
    this.sel = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], r: 0, c: 0, er: 0, ec: 0 };
    this.edit = null;
    this.drag = null;
    this.measureCache = new Map();
    this.images = new Map();
    this.copyRange = null;
    this.dash = 0;
    this.colors = {};
    this.raf = 0;
    this.bind();
    new ResizeObserver(() => this.resize()).observe(host);
    this.resize();
  }

  get sh() { return this.app.sheet; }
  get wb() { return this.app.wb; }

  readTheme() {
    const cs = getComputedStyle(document.documentElement);
    const v = n => cs.getPropertyValue(n).trim();
    this.colors = {
      bg: v('--g-bg') || '#fff', line: v('--g-line') || '#e1e3e6', hdr: v('--g-hdr') || '#f8f9fa', hdrText: v('--g-hdr-text') || '#5f6368',
      hdrLine: v('--g-hdr-line') || '#c4c7c5', hdrSel: v('--g-hdr-sel') || '#d3e3fd', hdrSelText: v('--g-hdr-sel-text') || '#0b57d0',
      hdrAll: v('--g-hdr-all') || '#0b57d0', text: v('--g-text') || '#1f1f1f', accent: v('--accent') || '#1a73e8',
      selFill: v('--g-sel') || 'rgba(26,115,232,.1)', freeze: v('--g-freeze') || '#c4c7c5', link: v('--g-link') || '#1155cc',
    };
    this.render();
  }

  resize() {
    const r = this.host.getBoundingClientRect();
    this.W = r.width; this.H = r.height;
    this.dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.max(1, Math.round(r.width * this.dpr));
    this.canvas.height = Math.max(1, Math.round(r.height * this.dpr));
    this.canvas.style.width = r.width + 'px';
    this.canvas.style.height = r.height + 'px';
    this.render();
  }

  // ---------- geometry ----------
  geo() {
    const sh = this.sh;
    const sig = sh.id + ':' + sh.layoutV + ':' + sh.rows + ':' + sh.cols;
    if (this.g && this.g.sig === sig) return this.g;
    const colX = new Float64Array(sh.cols + 1), rowY = new Float64Array(sh.rows + 1);
    for (let c = 0; c < sh.cols; c++) colX[c + 1] = colX[c] + sh.colWidth(c);
    for (let r = 0; r < sh.rows; r++) rowY[r + 1] = rowY[r] + sh.rowHeight(r);
    const hw = Math.max(46, String(sh.rows).length * 8 + 14);
    this.g = { sig, colX, rowY, hw };
    this.syncSpacer();
    return this.g;
  }
  syncSpacer() {
    const g = this.g, sh = this.sh, z = this.z;
    const w = (g.hw + g.colX[sh.cols]) * z + 40;
    const h = (HH + g.rowY[sh.rows]) * z + 64;
    this.spacer.style.width = w + 'px';
    this.spacer.style.height = h + 'px';
    this.addRows.style.top = ((HH + g.rowY[sh.rows]) * z + 12) + 'px';
    this.addRows.style.left = (g.hw * z + 8 + this.scroller.scrollLeft) + 'px';
  }
  get sx() { return this.scroller.scrollLeft / this.z; }
  get sy() { return this.scroller.scrollTop / this.z; }
  get vw() { return this.scroller.clientWidth / this.z; }
  get vh() { return this.scroller.clientHeight / this.z; }
  fw() { return this.geo().colX[this.sh.freezeC]; }
  fh() { return this.geo().rowY[this.sh.freezeR]; }
  X(c) { const g = this.geo(); return g.hw + g.colX[c] - (c >= this.sh.freezeC ? this.sx : 0); }
  Y(r) { return HH + this.geo().rowY[r] - (r >= this.sh.freezeR ? this.sy : 0); }
  cw(c) { const g = this.geo(); return g.colX[c + 1] - g.colX[c]; }
  rh(r) { const g = this.geo(); return g.rowY[r + 1] - g.rowY[r]; }

  static bs(arr, v, lo, hi) {
    // largest i in [lo, hi) with arr[i] <= v
    let ans = lo;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (arr[m] <= v) { ans = m; lo = m + 1; } else hi = m;
    }
    return ans;
  }
  colAt(x) {
    const g = this.geo(), sh = this.sh, fC = sh.freezeC;
    if (fC && x < g.hw + g.colX[fC]) return Grid.bs(g.colX, Math.max(0, x - g.hw), 0, fC);
    const lx = x - g.hw + this.sx;
    const c = Grid.bs(g.colX, lx, fC, sh.cols);
    return Math.min(c, sh.cols - 1);
  }
  rowAt(y) {
    const g = this.geo(), sh = this.sh, fR = sh.freezeR;
    if (fR && y < HH + g.rowY[fR]) return Grid.bs(g.rowY, Math.max(0, y - HH), 0, fR);
    const ly = y - HH + this.sy;
    return Math.min(Grid.bs(g.rowY, ly, fR, sh.rows), sh.rows - 1);
  }
  cellRect(r, c) {
    const m = this.sh.mergeAt(r, c);
    if (m) return this.rangeRect(m);
    return { x: this.X(c), y: this.Y(r), w: this.cw(c), h: this.rh(r) };
  }
  rangeRect(g) {
    const x = this.X(g.c1), y = this.Y(g.r1);
    return { x, y, w: this.X(g.c2) + this.cw(g.c2) - x, h: this.Y(g.r2) + this.rh(g.r2) - y };
  }

  // ---------- selection ----------
  expand(g) {
    const sh = this.sh;
    let o = norm(g);
    if (!sh.merges.length) return o;
    for (let changed = true; changed;) {
      changed = false;
      for (const m of sh.merges) {
        if (m.r1 <= o.r2 && m.r2 >= o.r1 && m.c1 <= o.c2 && m.c2 >= o.c1) {
          const n = { r1: Math.min(o.r1, m.r1), c1: Math.min(o.c1, m.c1), r2: Math.max(o.r2, m.r2), c2: Math.max(o.c2, m.c2) };
          if (n.r1 !== o.r1 || n.c1 !== o.c1 || n.r2 !== o.r2 || n.c2 !== o.c2) { o = n; changed = true; }
        }
      }
    }
    return o;
  }
  get range() { return this.sel.ranges[this.sel.ranges.length - 1]; }
  select(r, c, opts = {}) {
    const sh = this.sh;
    r = clamp(r, 0, sh.rows - 1); c = clamp(c, 0, sh.cols - 1);
    const m = sh.mergeAt(r, c);
    if (m) { r = m.r1; c = m.c1; }
    const g = this.expand({ r1: r, c1: c, r2: r, c2: c });
    if (opts.add) this.sel.ranges.push(g); else this.sel.ranges = [g];
    Object.assign(this.sel, { r, c, er: r, ec: c });
    if (opts.scroll !== false) this.scrollTo(r, c);
    this.changed();
  }
  selectRange(g, r = g.r1, c = g.c1, opts = {}) {
    const n = this.expand(g);
    if (opts.add) this.sel.ranges.push(n); else this.sel.ranges = [n];
    Object.assign(this.sel, { r, c, er: opts.er ?? n.r2, ec: opts.ec ?? n.c2 });
    if (opts.scroll) this.scrollTo(opts.er ?? r, opts.ec ?? c);
    this.changed();
  }
  extendTo(r, c) {
    const sh = this.sh;
    r = clamp(r, 0, sh.rows - 1); c = clamp(c, 0, sh.cols - 1);
    const g = this.expand({ r1: this.sel.r, c1: this.sel.c, r2: r, c2: c });
    this.sel.ranges[this.sel.ranges.length - 1] = g;
    this.sel.er = r; this.sel.ec = c;
    this.changed();
  }
  selectAll() { const sh = this.sh; this.selectRange({ r1: 0, c1: 0, r2: sh.rows - 1, c2: sh.cols - 1 }, this.sel.r, this.sel.c); }
  changed() { this.render(); this.placeInput(); this.app.onSelChange(); }

  scrollTo(r, c) {
    const sh = this.sh, g = this.geo(), z = this.z;
    const vw = this.vw - g.hw, vh = this.vh - HH;
    const fw = this.fw(), fh = this.fh();
    if (c >= sh.freezeC) {
      const x1 = g.colX[c] - fw, x2 = g.colX[c + 1] - fw;
      const view = vw - fw;
      if (x1 < this.sx) this.scroller.scrollLeft = x1 * z;
      else if (x2 > this.sx + view) this.scroller.scrollLeft = Math.min(x1, x2 - view) * z;
    }
    if (r >= sh.freezeR) {
      const y1 = g.rowY[r] - fh, y2 = g.rowY[r + 1] - fh;
      const view = vh - fh;
      if (y1 < this.sy) this.scroller.scrollTop = y1 * z;
      else if (y2 > this.sy + view) this.scroller.scrollTop = Math.min(y1, y2 - view) * z;
    }
  }

  hasValue(r, c) {
    const sh = this.sh;
    const cell = sh.cells.get(key(r, c));
    if (cell && (cell.v != null && cell.v !== '' || cell.f != null)) return true;
    return sh.spillMap.has(key(r, c));
  }
  visibleStep(pos, d, axis) {
    const sh = this.sh, max = axis === 'r' ? sh.rows - 1 : sh.cols - 1;
    let p = pos + d;
    while (p >= 0 && p <= max && (axis === 'r' ? sh.rowHeight(p) : sh.colWidth(p)) === 0) p += d;
    return p < 0 || p > max ? pos : p;
  }
  jump(r, c, dr, dc) {
    const sh = this.sh;
    const max = dr ? sh.rows - 1 : sh.cols - 1;
    const at = (p) => (dr ? this.hasValue(p, c) : this.hasValue(r, p));
    let p = dr ? r : c;
    const step = dr || dc;
    const cur = at(p), nxt = p + step >= 0 && p + step <= max && at(p + step);
    if (cur && nxt) { while (p + step >= 0 && p + step <= max && at(p + step)) p += step; }
    else { p += step; while (p >= 0 && p <= max && !at(p)) p += step; if (p < 0) p = 0; if (p > max) p = max; }
    return dr ? [p, c] : [r, p];
  }
  move(dr, dc, extend = false, jump = false) {
    const s = this.sel;
    let r = extend ? s.er : s.r, c = extend ? s.ec : s.c;
    if (!extend) {
      const m = this.sh.mergeAt(r, c);
      if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; }
    }
    if (jump) [r, c] = this.jump(r, c, dr, dc);
    else { if (dr) r = this.visibleStep(r, dr, 'r'); if (dc) c = this.visibleStep(c, dc, 'c'); }
    if (extend) { this.extendTo(r, c); this.scrollTo(r, c); }
    else this.select(r, c);
  }
  // Enter/Tab move inside a multi-cell selection
  cycle(dr, dc) {
    const g = this.range;
    if (g.r1 === g.r2 && g.c1 === g.c2 && this.sel.ranges.length === 1) return this.move(dr, dc);
    let { r, c } = this.sel;
    if (dc) { c += dc; if (c > g.c2) { c = g.c1; r = r + 1 > g.r2 ? g.r1 : r + 1; } else if (c < g.c1) { c = g.c2; r = r - 1 < g.r1 ? g.r2 : r - 1; } }
    if (dr) { r += dr; if (r > g.r2) { r = g.r1; c = c + 1 > g.c2 ? g.c1 : c + 1; } else if (r < g.r1) { r = g.r2; c = c - 1 < g.c1 ? g.c2 : c - 1; } }
    this.sel.r = r; this.sel.c = c;
    this.scrollTo(r, c);
    this.changed();
  }

  // ---------- rendering ----------
  render() {
    if (this.raf) return;
    this.raf = requestAnimationFrame(() => { this.raf = 0; this.draw(); });
  }
  drawNow() { if (this.raf) { cancelAnimationFrame(this.raf); this.raf = 0; } this.draw(); }

  measure(font, text) {
    const k = font + '\u0001' + text;
    let w = this.measureCache.get(k);
    if (w === undefined) {
      if (this.measureCache.size > 8000) this.measureCache.clear();
      this.ctx.font = font;
      w = this.ctx.measureText(text).width;
      this.measureCache.set(k, w);
    }
    return w;
  }

  display(r, c, cell) {
    const sh = this.sh, k = key(r, c);
    let v, fmt = null;
    const st = this.wb.styles[cell ? cell.s : 0] || {};
    if (cell && cell.f != null) {
      if (this.app.showFormulas) return { text: '=' + cell.f, st, v: null, formula: true };
      v = cell.sv; fmt = st.fmt || cell.af;
    } else if (cell && cell.v != null) { v = cell.v; fmt = st.fmt; }
    else {
      const sp = sh.spillMap.get(k);
      if (!sp) return null;
      v = sp.v;
      const oc = sh.cells.get(sp.o);
      fmt = st.fmt || (oc && oc.af);
    }
    if (v == null) return null;
    if (v instanceof Special) return { special: v, st, v };
    if (v instanceof FErr) return { text: v.code, st, v, err: v };
    const f = format(v, fmt);
    return { text: f.t, color: f.color, st, v, fmt };
  }

  draw() {
    const ctx = this.ctx, sh = this.sh;
    if (!sh || !this.W) return;
    const g = this.geo();
    this.syncSpacer();
    const z = this.z, S = this.dpr * z;
    this.S = S;
    const W = this.W / z, H = this.H / z;
    ctx.setTransform(S, 0, 0, S, 0, 0);
    ctx.fillStyle = this.colors.bg;
    ctx.fillRect(0, 0, W, H);
    const fC = sh.freezeC, fR = sh.freezeR, fw = this.fw(), fh = this.fh();
    const vr0 = this.rowAt(HH + fh + 0.01), vr1 = this.rowAt(H);
    const vc0 = this.colAt(g.hw + fw + 0.01), vc1 = this.colAt(W);
    this.visible = { r0: vr0, r1: vr1, c0: vc0, c1: vc1 };
    const quads = [
      { r0: vr0, r1: vr1, c0: vc0, c1: vc1, x: g.hw + fw, y: HH + fh, w: W - g.hw - fw, h: H - HH - fh, cs: fC },
    ];
    if (fR) quads.push({ r0: 0, r1: fR - 1, c0: vc0, c1: vc1, x: g.hw + fw, y: HH, w: W - g.hw - fw, h: fh, cs: fC });
    if (fC) quads.push({ r0: vr0, r1: vr1, c0: 0, c1: fC - 1, x: g.hw, y: HH + fh, w: fw, h: H - HH - fh, cs: 0 });
    if (fR && fC) quads.push({ r0: 0, r1: fR - 1, c0: 0, c1: fC - 1, x: g.hw, y: HH, w: fw, h: fh, cs: 0 });
    for (const q of quads) {
      if (q.w <= 0 || q.h <= 0) continue;
      ctx.save();
      ctx.beginPath(); ctx.rect(q.x, q.y, q.w, q.h); ctx.clip();
      this.drawCells(q);
      this.drawOverlays(q);
      ctx.restore();
    }
    this.drawHeaders(W, H);
    // freeze bars
    ctx.fillStyle = this.colors.freeze;
    if (fR || this.drag?.mode === 'freezeR') { const y = this.drag?.mode === 'freezeR' ? this.drag.y : HH + fh; ctx.fillRect(0, y - 1.5, W, 3); }
    else ctx.fillRect(0, HH - 1, g.hw, 3);
    if (fC || this.drag?.mode === 'freezeC') { const x = this.drag?.mode === 'freezeC' ? this.drag.x : g.hw + fw; ctx.fillRect(x - 1.5, 0, 3, H); }
    else ctx.fillRect(g.hw - 1, 0, 3, HH);
    if (this.drag?.mode === 'colresize') { ctx.fillStyle = this.colors.accent; ctx.fillRect(this.drag.x - 1, 0, 2, H); }
    if (this.drag?.mode === 'rowresize') { ctx.fillStyle = this.colors.accent; ctx.fillRect(0, this.drag.y - 1, W, 2); }
    this.placeEditor();
    this.app.charts?.position();
  }

  px(v) { return (Math.round(v * this.S) + 0.5) / this.S; }

  drawCells(q) {
    const ctx = this.ctx, sh = this.sh, wb = this.wb, cells = sh.cells, styles = wb.styles;
    const def = this.app.defFont;
    const cf = this.app.cf;
    const rows = [], cols = [];
    for (let r = q.r0; r <= q.r1; r++) if (this.rh(r) > 0) rows.push(r);
    for (let c = q.c0; c <= q.c1; c++) if (this.cw(c) > 0) cols.push(c);
    if (!rows.length || !cols.length) return;
    const xs = new Map(cols.map(c => [c, this.X(c)]));
    const band = this.app.bandingFor;
    // backgrounds
    const texts = [];
    const lineSkip = new Set();
    for (const r of rows) {
      const y = this.Y(r), h = this.rh(r);
      for (const c of cols) {
        if (sh.merges.length && sh.mergeAt(r, c)) continue;
        const cell = cells.get(r * 16384 + c);
        const st = cell ? styles[cell.s] || {} : {};
        const cfr = sh.cf.length ? cf.result(sh, r, c) : null;
        const bg = (cfr && cfr.bg) || st.bg || band(sh, r, c);
        if (bg) { ctx.fillStyle = bg; ctx.fillRect(xs.get(c), y, this.cw(c), h); }
        if (cfr && cfr.bar) this.drawBar(xs.get(c), y, this.cw(c), h, cfr.bar);
      }
    }
    // gridlines
    if (sh.grid) {
      ctx.strokeStyle = this.colors.line;
      ctx.lineWidth = 1 / this.S;
      ctx.beginPath();
      const yTop = this.Y(rows[0]), yBot = this.Y(rows[rows.length - 1]) + this.rh(rows[rows.length - 1]);
      const xL = xs.get(cols[0]), xR = xs.get(cols[cols.length - 1]) + this.cw(cols[cols.length - 1]);
      for (const c of cols) { const x = this.px(xs.get(c) + this.cw(c)); ctx.moveTo(x, yTop); ctx.lineTo(x, yBot); }
      for (const r of rows) { const y = this.px(this.Y(r) + this.rh(r)); ctx.moveTo(xL, y); ctx.lineTo(xR, y); }
      ctx.stroke();
    }
    // merged cells: background over gridlines
    const merges = sh.merges.filter(m => m.r2 >= q.r0 && m.r1 <= q.r1 && m.c2 >= q.c0 && m.c1 <= q.c1);
    for (const m of merges) {
      const cell = cells.get(key(m.r1, m.c1));
      const st = cell ? styles[cell.s] || {} : {};
      const cfr = sh.cf.length ? cf.result(sh, m.r1, m.c1) : null;
      const rc = this.rangeRect(m);
      ctx.fillStyle = (cfr && cfr.bg) || st.bg || band(sh, m.r1, m.c1) || this.colors.bg;
      ctx.fillRect(rc.x + 0.5 / this.S, rc.y + 0.5 / this.S, rc.w - 1 / this.S, rc.h - 1 / this.S);
      if (sh.grid) { ctx.strokeStyle = this.colors.line; ctx.lineWidth = 1 / this.S; ctx.strokeRect(this.px(rc.x), this.px(rc.y), rc.w, rc.h); }
      texts.push({ r: m.r1, c: m.c1, x: rc.x, y: rc.y, w: rc.w, h: rc.h, cell, merged: true });
    }
    // text items
    const cStart = Math.max(q.cs, q.c0 - 12);
    for (const r of rows) {
      const y = this.Y(r), h = this.rh(r);
      for (let c = cStart; c <= q.c1; c++) {
        if (this.cw(c) === 0) continue;
        const k = r * 16384 + c;
        const cell = cells.get(k);
        if (!cell && !sh.spillMap.has(k)) continue;
        if (sh.merges.length && sh.mergeAt(r, c)) continue;
        if (c < q.c0 && !(cell && typeof cell.v === 'string') && !(cell && typeof cell.sv === 'string')) continue;
        texts.push({ r, c, x: this.X(c), y, w: this.cw(c), h, cell });
      }
    }
    for (const t of texts) this.drawCellText(t, q, lineSkip);
    // borders
    for (const r of rows) {
      for (const c of cols) {
        const cell = cells.get(r * 16384 + c);
        if (!cell || !cell.s) continue;
        const st = styles[cell.s];
        if (!st || !(st.bt || st.bb || st.bl || st.br)) continue;
        const m = sh.merges.length ? sh.mergeAt(r, c) : null;
        const rc = m ? this.rangeRect(m) : { x: xs.get(c), y: this.Y(r), w: this.cw(c), h: this.rh(r) };
        this.drawBorders(rc, st);
      }
    }
    // indicators
    for (const r of rows) {
      for (const c of cols) {
        const cell = cells.get(r * 16384 + c);
        if (!cell) continue;
        const rc = { x: xs.get(c), y: this.Y(r), w: this.cw(c), h: this.rh(r) };
        if (cell.note) { ctx.fillStyle = '#f29900'; ctx.beginPath(); ctx.moveTo(rc.x + rc.w - 7, rc.y); ctx.lineTo(rc.x + rc.w, rc.y); ctx.lineTo(rc.x + rc.w, rc.y + 7); ctx.fill(); }
        if (cell.f != null && cell.sv instanceof FErr && cell.sv.msg && !this.app.showFormulas) { ctx.fillStyle = '#d93025'; ctx.beginPath(); ctx.moveTo(rc.x + rc.w - 7, rc.y); ctx.lineTo(rc.x + rc.w, rc.y); ctx.lineTo(rc.x + rc.w, rc.y + 7); ctx.fill(); }
        if (cell.bad) { ctx.fillStyle = '#d93025'; ctx.beginPath(); ctx.moveTo(rc.x + rc.w - 7, rc.y); ctx.lineTo(rc.x + rc.w, rc.y); ctx.lineTo(rc.x + rc.w, rc.y + 7); ctx.fill(); }
      }
    }
    // filter buttons
    const f = sh.filter;
    if (f && rows.includes(f.r1)) {
      for (const c of cols) {
        if (c < f.c1 || c > f.c2) continue;
        const x = xs.get(c), y = this.Y(f.r1), w = this.cw(c), h = this.rh(f.r1);
        const active = f.crit && f.crit[c];
        const bx = x + w - 18, by = y + (h - 16) / 2;
        ctx.fillStyle = active ? this.colors.accent : 'rgba(0,0,0,.06)';
        this.roundRect(bx, by, 16, 16, 3); ctx.fill();
        ctx.strokeStyle = active ? '#fff' : '#5f6368'; ctx.lineWidth = 1.3;
        ctx.beginPath();
        ctx.moveTo(bx + 4, by + 5); ctx.lineTo(bx + 12, by + 5);
        ctx.moveTo(bx + 5.5, by + 8); ctx.lineTo(bx + 10.5, by + 8);
        ctx.moveTo(bx + 7, by + 11); ctx.lineTo(bx + 9, by + 11);
        ctx.stroke();
      }
    }
  }

  roundRect(x, y, w, h, r) {
    const ctx = this.ctx;
    ctx.beginPath();
    ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath();
  }

  drawBar(x, y, w, h, bar) {
    const ctx = this.ctx;
    const pad = 2, bw = w - pad * 2;
    const x0 = x + pad + bw * bar.from, x1 = x + pad + bw * bar.to;
    const color = bar.neg ? '#e06666' : bar.color;
    if (bar.solid) ctx.fillStyle = color + 'cc';
    else { const gr = ctx.createLinearGradient(x0, 0, x1 + 1, 0); gr.addColorStop(0, color + 'e6'); gr.addColorStop(1, color + '26'); ctx.fillStyle = gr; }
    ctx.fillRect(x0, y + 3, Math.max(1, x1 - x0), h - 6);
  }

  drawBorders(rc, st) {
    const ctx = this.ctx;
    const side = (b, x1, y1, x2, y2) => {
      if (!b) return;
      const w = b.w || 1;
      ctx.strokeStyle = b.color || '#000';
      ctx.lineWidth = b.style === 'double' ? 1 : w;
      ctx.setLineDash(b.style === 'dashed' ? [4, 2] : b.style === 'dotted' ? [1, 1.5] : []);
      ctx.beginPath();
      if (b.style === 'double') {
        const off = 1.2;
        if (y1 === y2) { ctx.moveTo(x1, y1 - off); ctx.lineTo(x2, y2 - off); ctx.moveTo(x1, y1 + off); ctx.lineTo(x2, y2 + off); }
        else { ctx.moveTo(x1 - off, y1); ctx.lineTo(x2 - off, y2); ctx.moveTo(x1 + off, y1); ctx.lineTo(x2 + off, y2); }
      } else { ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); }
      ctx.stroke();
      ctx.setLineDash([]);
    };
    const x1 = this.px(rc.x), y1 = this.px(rc.y), x2 = this.px(rc.x + rc.w), y2 = this.px(rc.y + rc.h);
    side(st.bt, x1, y1, x2, y1);
    side(st.bb, x1, y2, x2, y2);
    side(st.bl, x1, y1, x1, y2);
    side(st.br, x2, y1, x2, y2);
  }

  drawCellText(t, q, lineSkip) {
    const ctx = this.ctx, sh = this.sh, def = this.app.defFont;
    const { r, c, cell } = t;
    const d = this.display(r, c, cell);
    const cfr = sh.cf.length ? this.app.cf.result(sh, r, c) : null;
    const st = d ? d.st : (cell ? this.wb.styles[cell.s] || {} : {});
    const dv = sh.dv.length ? dvAt(sh, r, c) : null;
    // checkbox
    if (dv && dv.type === 'checkbox') {
      const v = d ? d.v : null;
      const on = v === (dv.checked ?? true) || (dv.checked == null && v === true);
      this.drawCheckbox(t, on, cfr?.color || st.color);
      return;
    }
    if (!d && !(dv && (dv.type === 'list' || dv.type === 'range') && dv.chip !== 'plain')) return;
    if (d && d.special) { this.drawSpecial(t, d.special); return; }
    let text = d ? d.text : '';
    const font = fontOf({ ...st, ...(cfr && cfr.b ? { b: true } : {}), ...(cfr && cfr.i ? { i: true } : {}) }, def);
    const px = ((st.fs || def.fs) * 4) / 3;
    let color = (cfr && cfr.color) || d?.color || st.color || this.colors.text;
    const isLink = cell && (cell.link || cell.hl || (typeof cell.v === 'string' && URL_RE.test(cell.v)));
    if (isLink && !st.color) color = this.colors.link;
    const isNum = d && typeof d.v === 'number' && !d.formula;
    let ha = st.ha || (d && d.err ? 'center' : isNum ? 'right' : d && typeof d.v === 'boolean' ? 'center' : 'left');
    const va = st.va || 'bottom';
    const pad = 3, ind = (st.ind || 0) * 10;
    const wrap = st.wrap === 'wrap' || (typeof text === 'string' && text.includes('\n') && st.wrap !== 'clip' && st.wrap !== 'overflow');
    // chip-style dropdown
    if (dv && (dv.type === 'list' || dv.type === 'range') && dv.chip !== 'plain') {
      this.drawChips(t, dv, text, font, px, st);
      return;
    }
    if (dv && (dv.type === 'list' || dv.type === 'range') && dv.chip === 'plain' && this.sel.r === r && this.sel.c === c) this.drawArrow(t);
    if (!text) return;
    ctx.font = font;
    ctx.fillStyle = color;
    let x = t.x, w = t.w;
    const tw = this.measure(font, text);
    let iconW = 0;
    if (cfr && cfr.icon) { this.drawIcon(t, cfr.icon); iconW = 16; }
    const avail = w - pad * 2 - ind - iconW;
    if (isNum && tw > avail && !wrap && !st.rot) {
      // shrink General numbers, else ####
      let s = null;
      if (!d.fmt || d.fmt === 'General') {
        const v = d.v;
        for (let p = 9; p >= 1 && s == null; p--) {
          const cand = Math.abs(v) >= 1e11 || (Math.abs(v) < 1e-4 && v !== 0) ? general(Number(v.toPrecision(p))).replace(/(\.\d*?)0+E/, '$1E') : general(Number(v.toPrecision(p)));
          if (this.measure(font, cand) <= avail) s = cand;
        }
      }
      text = s ?? '#'.repeat(Math.max(1, Math.floor(avail / this.measure(font, '#'))));
    }
    const rot = st.rot;
    if (rot) {
      ctx.save();
      ctx.beginPath(); ctx.rect(t.x, t.y, t.w, t.h); ctx.clip();
      ctx.translate(t.x + t.w / 2, t.y + t.h / 2);
      if (rot === 'v') {
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        const chars = [...text]; const lh = px * 1.05;
        chars.forEach((ch, i) => ctx.fillText(ch, 0, (i - (chars.length - 1) / 2) * lh));
      } else {
        ctx.rotate((-rot * Math.PI) / 180);
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(text, 0, 0);
      }
      ctx.restore();
      return;
    }
    if (wrap) {
      const lines = this.wrapLines(font, text, Math.max(10, avail));
      const lh = px * 1.25;
      const total = lines.length * lh;
      let y0 = va === 'top' ? t.y + pad : va === 'middle' ? t.y + (t.h - total) / 2 : t.y + t.h - pad - total;
      ctx.save();
      ctx.beginPath(); ctx.rect(t.x, t.y, t.w, t.h); ctx.clip();
      ctx.textBaseline = 'top';
      for (const ln of lines) {
        const lw = this.measure(font, ln);
        const lx = ha === 'right' ? t.x + t.w - pad - lw : ha === 'center' ? t.x + (t.w - lw) / 2 : t.x + pad + ind + iconW;
        ctx.fillText(ln, lx, y0 + (lh - px) / 2);
        this.decorate(st, cfr, isLink, lx, y0 + (lh - px) / 2 + px, lw, px, color);
        y0 += lh;
      }
      ctx.restore();
      return;
    }
    // overflow computation
    let clipL = t.x, clipR = t.x + t.w;
    const overflow = !t.merged && st.wrap !== 'clip' && typeof d.v === 'string' && !d.formula || (d.formula && st.wrap !== 'clip');
    if (tw > avail && overflow) {
      const need = tw - avail;
      const tryExtend = (dir) => {
        let ext = 0, cc = c;
        const lim = dir > 0 ? Math.min(this.sh.cols - 1, q.c1 + 12) : Math.max(q.cs, 0);
        while (ext < (ha === 'center' ? need / 2 : need)) {
          const nc = cc + dir;
          if (dir > 0 ? nc > lim : nc < lim) break;
          if (this.hasValue(r, nc) || (sh.merges.length && sh.mergeAt(r, nc))) break;
          if (dir > 0 && nc >= this.sh.freezeC && c < this.sh.freezeC) break;
          ext += this.cw(nc); cc = nc;
        }
        return ext;
      };
      if (ha === 'left') clipR += tryExtend(1);
      else if (ha === 'right') clipL -= tryExtend(-1);
      else { clipR += tryExtend(1); clipL -= tryExtend(-1); }
      // hide gridlines inside the overflow area
      if ((clipR > t.x + t.w || clipL < t.x) && sh.grid) {
        ctx.save();
        const y = this.Y(r) + 1 / this.S, h = this.rh(r) - 2 / this.S;
        for (let cc = this.colAt(clipL + 1); cc < this.colAt(clipR - 1); cc++) {
          if (cc < c - 40) break;
          const bx = this.X(cc) + this.cw(cc);
          const nb = this.sh.styleOf(r, cc + 1).bg || this.sh.styleOf(r, cc).bg || this.app.bandingFor(this.sh, r, cc) || this.colors.bg;
          ctx.fillStyle = nb;
          ctx.fillRect(bx - 1 / this.S, y, 2 / this.S, h);
        }
        ctx.restore();
        ctx.font = font; ctx.fillStyle = color;
      }
    }
    let tx;
    if (ha === 'right') tx = t.x + t.w - pad - tw;
    else if (ha === 'center') tx = t.x + (t.w - tw) / 2;
    else tx = t.x + pad + ind + iconW;
    let ty;
    if (va === 'top') { ctx.textBaseline = 'top'; ty = t.y + pad; }
    else if (va === 'middle') { ctx.textBaseline = 'middle'; ty = t.y + t.h / 2; }
    else { ctx.textBaseline = 'alphabetic'; ty = t.y + t.h - pad - px * 0.2; }
    const needClip = tx < clipL || tx + tw > clipR;
    if (needClip) { ctx.save(); ctx.beginPath(); ctx.rect(clipL, t.y, clipR - clipL, t.h); ctx.clip(); }
    ctx.fillText(text, tx, ty);
    const base = va === 'top' ? ty + px : va === 'middle' ? ty + px / 2 : ty + px * 0.2 - 1;
    this.decorate(st, cfr, isLink, tx, base, tw, px, color);
    if (needClip) ctx.restore();
  }

  decorate(st, cfr, link, x, baseY, w, px, color) {
    const u = st.u || (cfr && cfr.u) || link, s = st.s || (cfr && cfr.s);
    if (!u && !s) return;
    const ctx = this.ctx;
    ctx.strokeStyle = color; ctx.lineWidth = Math.max(1 / this.S, px / 14);
    ctx.beginPath();
    if (u) { ctx.moveTo(x, baseY + 1); ctx.lineTo(x + w, baseY + 1); }
    if (s) { ctx.moveTo(x, baseY - px * 0.32); ctx.lineTo(x + w, baseY - px * 0.32); }
    ctx.stroke();
  }

  wrapLines(font, text, width) {
    const out = [];
    for (const para of String(text).split('\n')) {
      const words = para.split(/(\s+)/);
      let line = '';
      for (const w of words) {
        const test = line + w;
        if (this.measure(font, test) <= width || !line) {
          if (this.measure(font, test) > width && !line) {
            // break long word
            let part = '';
            for (const ch of w) { if (this.measure(font, part + ch) > width && part) { out.push(part); part = ''; } part += ch; }
            line = part;
          } else line = test;
        } else { out.push(line.trimEnd()); line = w.trimStart(); }
      }
      out.push(line);
    }
    return out;
  }

  drawCheckbox(t, on, color) {
    const ctx = this.ctx;
    const s = 14, x = t.x + (t.w - s) / 2, y = t.y + (t.h - s) / 2;
    this.roundRect(x, y, s, s, 2.5);
    if (on) { ctx.fillStyle = color || this.colors.accent; ctx.fill(); ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.8; ctx.beginPath(); ctx.moveTo(x + 3.2, y + 7.2); ctx.lineTo(x + 6, y + 10); ctx.lineTo(x + 11, y + 4.2); ctx.stroke(); }
    else { ctx.strokeStyle = color || '#5f6368'; ctx.lineWidth = 1.5; ctx.stroke(); }
  }

  drawArrow(t) {
    const ctx = this.ctx;
    const x = t.x + t.w - 14, y = t.y + t.h / 2;
    ctx.fillStyle = '#5f6368';
    ctx.beginPath(); ctx.moveTo(x, y - 2); ctx.lineTo(x + 8, y - 2); ctx.lineTo(x + 4, y + 3); ctx.closePath(); ctx.fill();
  }

  drawChips(t, dv, text, font, px, st) {
    const ctx = this.ctx;
    const items = dvItems(this.sh, dv);
    const vals = text ? (dv.multi ? String(text).split(',').map(s => s.trim()).filter(Boolean) : [text]) : [];
    let x = t.x + 4;
    ctx.save();
    ctx.beginPath(); ctx.rect(t.x, t.y, t.w - 18, t.h); ctx.clip();
    ctx.font = font; ctx.textBaseline = 'middle';
    for (const v of vals) {
      const idx = items.findIndex(i => i.toLowerCase() === v.toLowerCase());
      const bg = idx >= 0 ? (dv.colors?.[idx] || CHIP_COLORS[0]) : '#fce8e6';
      const w = this.measure(font, v) + 16;
      const h = Math.min(t.h - 4, px + 7);
      this.roundRect(x, t.y + (t.h - h) / 2, w, h, h / 2);
      ctx.fillStyle = bg; ctx.fill();
      ctx.fillStyle = idx >= 0 ? '#202124' : '#a50e0e';
      ctx.fillText(v, x + 8, t.y + t.h / 2 + 0.5);
      x += w + 4;
    }
    ctx.restore();
    // arrow
    const ax = t.x + t.w - 13, ay = t.y + t.h / 2;
    ctx.strokeStyle = '#5f6368'; ctx.lineWidth = 1.4;
    ctx.beginPath(); ctx.moveTo(ax, ay - 2); ctx.lineTo(ax + 3.5, ay + 1.5); ctx.lineTo(ax + 7, ay - 2); ctx.stroke();
  }

  drawIcon(t, icon) {
    const ctx = this.ctx;
    const cx = t.x + 9, cy = t.y + t.h / 2;
    const colors = ['#d93025', '#f9ab00', '#188038'];
    ctx.save();
    if (icon.set === 'traffic') { ctx.fillStyle = colors[icon.idx]; ctx.beginPath(); ctx.arc(cx, cy, 5, 0, Math.PI * 2); ctx.fill(); }
    else if (icon.set === 'flags') { ctx.fillStyle = colors[icon.idx]; ctx.fillRect(cx - 4, cy - 5, 1.5, 10); ctx.beginPath(); ctx.moveTo(cx - 2.5, cy - 5); ctx.lineTo(cx + 5, cy - 2.5); ctx.lineTo(cx - 2.5, cy); ctx.fill(); }
    else if (icon.set === 'stars') {
      ctx.fillStyle = icon.idx === 2 ? '#f9ab00' : icon.idx === 1 ? '#fbd56b' : '#dadce0';
      ctx.beginPath();
      for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, rr = i % 2 ? 2.4 : 5.5; ctx.lineTo(cx + Math.cos(a) * rr, cy + Math.sin(a) * rr); }
      ctx.fill();
    } else {
      ctx.strokeStyle = colors[icon.idx]; ctx.fillStyle = colors[icon.idx]; ctx.lineWidth = 2;
      ctx.translate(cx, cy);
      ctx.rotate(icon.idx === 2 ? -Math.PI / 2 : icon.idx === 0 ? Math.PI / 2 : 0);
      ctx.beginPath(); ctx.moveTo(-5, 0); ctx.lineTo(3, 0); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(5.5, 0); ctx.lineTo(1, -4); ctx.lineTo(1, 4); ctx.closePath(); ctx.fill();
    }
    ctx.restore();
  }

  drawSpecial(t, sp) {
    const ctx = this.ctx;
    if (sp.kind === 'spark') {
      const vals = sp.data.vals, o = sp.data.opts;
      const nums = vals.filter(v => v != null);
      if (!nums.length) return;
      const type = String(o.charttype || 'line').toLowerCase();
      const color = o.color || o.color1 || this.colors.accent;
      const pad = 3;
      const x0 = t.x + pad, y0 = t.y + pad, w = t.w - pad * 2, h = t.h - pad * 2;
      let lo = o.ymin != null ? +o.ymin : Math.min(...nums), hi = o.ymax != null ? +o.ymax : Math.max(...nums);
      ctx.save(); ctx.beginPath(); ctx.rect(t.x, t.y, t.w, t.h); ctx.clip();
      if (type === 'bar') {
        const max = o.max != null ? +o.max : nums.reduce((a, b) => a + Math.abs(b), 0) || 1;
        let x = x0;
        const pal = [color, o.color2 || '#f4b400', '#0f9d58', '#db4437'];
        nums.forEach((v, i) => { const bw = w * Math.abs(v) / max; ctx.fillStyle = pal[i % pal.length]; ctx.fillRect(x, y0, bw, h); x += bw; });
      } else if (type === 'column' || type === 'winloss') {
        if (lo > 0) lo = 0; if (hi < 0) hi = 0;
        const n = vals.length, bw = w / n;
        const yz = y0 + h * (hi / (hi - lo || 1));
        vals.forEach((v, i) => {
          if (v == null) return;
          ctx.fillStyle = v < 0 ? (o.negcolor || '#d93025') : color;
          if (type === 'winloss') { ctx.fillRect(x0 + i * bw + 1, v >= 0 ? y0 : y0 + h / 2 + 1, bw - 2, h / 2 - 1); }
          else { const yv = y0 + h * ((hi - v) / (hi - lo || 1)); ctx.fillRect(x0 + i * bw + 1, Math.min(yv, yz), bw - 2, Math.max(1, Math.abs(yz - yv))); }
        });
      } else {
        ctx.strokeStyle = color; ctx.lineWidth = o.linewidth ? +o.linewidth : 1.4; ctx.lineJoin = 'round';
        ctx.beginPath();
        let started = false;
        vals.forEach((v, i) => {
          if (v == null) return;
          const x = x0 + (vals.length === 1 ? w / 2 : (w * i) / (vals.length - 1));
          const y = y0 + h * ((hi - v) / (hi - lo || 1));
          if (!started) { ctx.moveTo(x, y); started = true; } else ctx.lineTo(x, y);
        });
        ctx.stroke();
      }
      ctx.restore();
      return;
    }
    if (sp.kind === 'image') {
      const url = sp.data.url;
      let im = this.images.get(url);
      if (!im) {
        im = { img: new Image(), ok: false };
        im.img.crossOrigin = 'anonymous';
        im.img.onload = () => { im.ok = true; this.render(); };
        im.img.src = url;
        this.images.set(url, im);
      }
      if (!im.ok) return;
      const { naturalWidth: iw, naturalHeight: ih } = im.img;
      let w = t.w - 2, h = t.h - 2;
      const mode = sp.data.mode;
      if (mode === 1) { const s = Math.min(w / iw, h / ih); w = iw * s; h = ih * s; }
      else if (mode === 3) { w = iw; h = ih; }
      else if (mode === 4) { w = sp.data.w || iw; h = sp.data.h || ih; }
      ctx.save(); ctx.beginPath(); ctx.rect(t.x, t.y, t.w, t.h); ctx.clip();
      ctx.drawImage(im.img, t.x + (t.w - w) / 2, t.y + (t.h - h) / 2, w, h);
      ctx.restore();
    }
  }

  // selection, copy marquee, formula reference highlights, fill preview
  drawOverlays(q) {
    const ctx = this.ctx, sh = this.sh, acc = this.colors.accent;
    const clampQ = g => {
      const r1 = Math.max(g.r1, q.r0), r2 = Math.min(g.r2, q.r1), c1 = Math.max(g.c1, q.c0), c2 = Math.min(g.c2, q.c1);
      if (r1 > r2 || c1 > c2) return null;
      return { r1, r2, c1, c2, sides: { t: r1 === g.r1, b: r2 === g.r2, l: c1 === g.c1, r: c2 === g.c2 } };
    };
    const rectOf = g => this.rangeRect(g);
    const strokeSides = (rc, sides, lw, color, dash) => {
      ctx.strokeStyle = color; ctx.lineWidth = lw; ctx.setLineDash(dash || []);
      ctx.beginPath();
      const x1 = this.px(rc.x), y1 = this.px(rc.y), x2 = this.px(rc.x + rc.w) - (lw > 1 ? 0.5 : 0), y2 = this.px(rc.y + rc.h) - (lw > 1 ? 0.5 : 0);
      if (sides.t) { ctx.moveTo(x1, y1); ctx.lineTo(x2, y1); }
      if (sides.b) { ctx.moveTo(x1, y2); ctx.lineTo(x2, y2); }
      if (sides.l) { ctx.moveTo(x1, y1); ctx.lineTo(x1, y2); }
      if (sides.r) { ctx.moveTo(x2, y1); ctx.lineTo(x2, y2); }
      ctx.stroke(); ctx.setLineDash([]);
    };
    // permissions: cells viewers of a share link may edit (green), unlocked cells of a locked sheet (blue)
    const app = this.app;
    const marks = (list, col) => { for (const g of list || []) { const cq = clampQ(g); if (!cq) continue; const rc = rectOf(cq); ctx.fillStyle = col + '14'; ctx.fillRect(rc.x, rc.y, rc.w, rc.h); strokeSides(rc, cq.sides, 1.5 / this.z, col, [4, 3]); } };
    if (app.readOnly || app.settings.showVE) marks(sh.ve, '#1e8e3e');
    if (sh.protect && !app.readOnly) marks(sh.protectExcept, '#1a73e8');
    const edit = this.edit;
    // formula reference highlights
    if (edit && edit.refs && edit.sh === sh) {
      edit.refs.forEach((rf, i) => {
        if (rf.sheet && rf.sheet.toLowerCase() !== sh.name.toLowerCase()) return;
        const g = { r1: rf.r1, c1: rf.c1, r2: Math.min(rf.r2, sh.rows - 1), c2: Math.min(rf.c2, sh.cols - 1) };
        const cq = clampQ(g); if (!cq) return;
        const rc = rectOf(cq);
        const col = REF_COLORS[rf.ci % REF_COLORS.length];
        ctx.fillStyle = col + '1f'; ctx.fillRect(rc.x, rc.y, rc.w, rc.h);
        strokeSides(rc, cq.sides, 2 / this.z, col, edit.point && edit.point.idx === i ? [4, 3] : null);
      });
    }
    if (edit && edit.pointSel && edit.pointSel.sh === sh) {
      const cq = clampQ(edit.pointSel.g);
      if (cq) strokeSides(rectOf(cq), cq.sides, 2 / this.z, REF_COLORS[(edit.pointSel.ci || 0) % REF_COLORS.length], [5, 3]);
    }
    // selection
    const s = this.sel;
    const showSel = !(edit && edit.pointSel);
    if (showSel) {
      for (const g of s.ranges) {
        const cq = clampQ(g); if (!cq) continue;
        const rc = rectOf(cq);
        const single = g.r1 === g.r2 && g.c1 === g.c2 || (sh.mergeAt(g.r1, g.c1) && (() => { const m = sh.mergeAt(g.r1, g.c1); return m.r1 === g.r1 && m.c1 === g.c1 && m.r2 === g.r2 && m.c2 === g.c2; })());
        if (!single || s.ranges.length > 1) {
          ctx.fillStyle = this.colors.selFill;
          ctx.beginPath();
          ctx.rect(rc.x, rc.y, rc.w, rc.h);
          if (inRange(g, s.r, s.c) && g === this.range) {
            const ac = this.cellRect(s.r, s.c);
            ctx.rect(ac.x + ac.w, ac.y, -ac.w, ac.h);
          }
          ctx.fill('evenodd');
          strokeSides(rc, cq.sides, 1 / this.z, acc);
        }
      }
      // active cell
      if (s.r >= q.r0 && s.r <= q.r1 && s.c >= q.c0 && s.c <= q.c1 || sh.mergeAt(s.r, s.c)) {
        const ac = this.cellRect(s.r, s.c);
        ctx.strokeStyle = acc; ctx.lineWidth = 2 / this.z;
        ctx.strokeRect(this.px(ac.x) + 0.5 / this.S, this.px(ac.y) + 0.5 / this.S, ac.w - 1 / this.S, ac.h - 1 / this.S);
      }
      // fill handle
      if (!edit && s.ranges.length === 1) {
        const g = this.range;
        if (g.r2 >= q.r0 && g.r2 <= q.r1 && g.c2 >= q.c0 && g.c2 <= q.c1) {
          const rc = rectOf(g);
          const hx = rc.x + rc.w, hy = rc.y + rc.h;
          ctx.fillStyle = '#fff'; ctx.fillRect(hx - 4.5, hy - 4.5, 9, 9);
          ctx.fillStyle = acc; ctx.fillRect(hx - 3.5, hy - 3.5, 7, 7);
        }
      }
      // spill range outline
      const ok = key(s.r, s.c), sp = sh.spillOrigins.get(ok);
      if (sp) {
        const g = { r1: s.r, c1: s.c, r2: s.r + sp.rows - 1, c2: s.c + sp.cols - 1 };
        const cq = clampQ(g);
        if (cq) strokeSides(rectOf(cq), cq.sides, 1 / this.z, acc, [3, 2]);
      }
    }
    // copy marquee
    if (this.copyRange && this.copyRange.sh === sh) {
      const cq = clampQ(this.copyRange.g);
      if (cq) {
        const rc = rectOf(cq);
        ctx.lineDashOffset = -this.dash;
        strokeSides(rc, cq.sides, 2 / this.z, acc, [5, 4]);
        ctx.lineDashOffset = 0;
      }
    }
    // drag previews
    const d = this.drag;
    if (d && (d.mode === 'fill' || d.mode === 'move') && d.target) {
      const cq = clampQ(d.target);
      if (cq) strokeSides(rectOf(cq), cq.sides, (d.mode === 'move' ? 2 : 1) / this.z, d.mode === 'move' ? acc : '#5f6368', [4, 3]);
    }
    // find highlights
    if (this.app.findHits && this.app.findHits.sh === sh) {
      for (const h of this.app.findHits.list) {
        if (h.r < q.r0 || h.r > q.r1 || h.c < q.c0 || h.c > q.c1) continue;
        const rc = this.cellRect(h.r, h.c);
        ctx.fillStyle = 'rgba(251,188,4,.28)'; ctx.fillRect(rc.x, rc.y, rc.w, rc.h);
      }
    }
  }

  drawHeaders(W, H) {
    const ctx = this.ctx, sh = this.sh, g = this.geo(), C = this.colors;
    const fw = this.fw(), fh = this.fh();
    const s = this.sel;
    const selCols = new Set(), selRows = new Set(), fullCols = new Set(), fullRows = new Set();
    for (const r of s.ranges) {
      const fullC = r.r1 === 0 && r.r2 >= sh.rows - 1, fullR = r.c1 === 0 && r.c2 >= sh.cols - 1;
      for (let c = Math.max(r.c1, 0); c <= Math.min(r.c2, sh.cols - 1) && c <= this.visible.c1 + 1; c++) { selCols.add(c); if (fullC) fullCols.add(c); }
      for (let rr = Math.max(r.r1, 0); rr <= Math.min(r.r2, sh.rows - 1) && rr <= this.visible.r1 + 1; rr++) { selRows.add(rr); if (fullR) fullRows.add(rr); }
    }
    ctx.font = `500 11px Inter, ${FALLBACK}`;
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    const drawCols = (c0, c1, x0, x1) => {
      ctx.save(); ctx.beginPath(); ctx.rect(x0, 0, x1 - x0, HH); ctx.clip();
      ctx.fillStyle = C.hdr; ctx.fillRect(x0, 0, x1 - x0, HH);
      ctx.strokeStyle = C.hdrLine; ctx.lineWidth = 1 / this.S;
      for (let c = c0; c <= c1; c++) {
        const w = this.cw(c); if (!w) {
          // hidden column marker
          const x = this.X(c); ctx.fillStyle = C.hdrText; ctx.fillRect(x - 1.5, 4, 3, HH - 8); continue;
        }
        const x = this.X(c);
        if (fullCols.has(c)) { ctx.fillStyle = C.hdrAll; ctx.fillRect(x, 0, w, HH); }
        else if (selCols.has(c)) { ctx.fillStyle = C.hdrSel; ctx.fillRect(x, 0, w, HH); }
        ctx.fillStyle = fullCols.has(c) ? '#fff' : selCols.has(c) ? C.hdrSelText : C.hdrText;
        ctx.font = (selCols.has(c) ? '600' : '500') + ` 11px Inter, ${FALLBACK}`;
        ctx.fillText(colName(c), x + w / 2, HH / 2 + 0.5);
        ctx.beginPath(); ctx.moveTo(this.px(x + w), 0); ctx.lineTo(this.px(x + w), HH); ctx.stroke();
      }
      ctx.beginPath(); ctx.moveTo(x0, this.px(HH)); ctx.lineTo(x1, this.px(HH)); ctx.stroke();
      ctx.restore();
    };
    const drawRows = (r0, r1, y0, y1) => {
      ctx.save(); ctx.beginPath(); ctx.rect(0, y0, g.hw, y1 - y0); ctx.clip();
      ctx.fillStyle = C.hdr; ctx.fillRect(0, y0, g.hw, y1 - y0);
      ctx.strokeStyle = C.hdrLine; ctx.lineWidth = 1 / this.S;
      for (let r = r0; r <= r1; r++) {
        const h = this.rh(r);
        if (!h) { const y = this.Y(r); ctx.fillStyle = C.hdrText; ctx.fillRect(4, y - 1.5, g.hw - 8, 3); continue; }
        const y = this.Y(r);
        if (fullRows.has(r)) { ctx.fillStyle = C.hdrAll; ctx.fillRect(0, y, g.hw, h); }
        else if (selRows.has(r)) { ctx.fillStyle = C.hdrSel; ctx.fillRect(0, y, g.hw, h); }
        ctx.fillStyle = fullRows.has(r) ? '#fff' : selRows.has(r) ? C.hdrSelText : C.hdrText;
        ctx.font = (selRows.has(r) ? '600' : '500') + ` 11px Inter, ${FALLBACK}`;
        if (h >= 9) ctx.fillText(String(r + 1), g.hw / 2, y + h / 2 + 0.5);
        ctx.beginPath(); ctx.moveTo(0, this.px(y + h)); ctx.lineTo(g.hw, this.px(y + h)); ctx.stroke();
      }
      ctx.beginPath(); ctx.moveTo(this.px(g.hw), y0); ctx.lineTo(this.px(g.hw), y1); ctx.stroke();
      ctx.restore();
    };
    drawCols(this.visible.c0, this.visible.c1, g.hw + fw, W);
    if (sh.freezeC) drawCols(0, sh.freezeC - 1, g.hw, g.hw + fw);
    drawRows(this.visible.r0, this.visible.r1, HH + fh, H);
    if (sh.freezeR) drawRows(0, sh.freezeR - 1, HH, HH + fh);
    // corner
    ctx.fillStyle = C.hdr; ctx.fillRect(0, 0, g.hw, HH);
    ctx.strokeStyle = C.hdrLine; ctx.lineWidth = 1 / this.S;
    ctx.beginPath(); ctx.moveTo(this.px(g.hw), 0); ctx.lineTo(this.px(g.hw), HH); ctx.moveTo(0, this.px(HH)); ctx.lineTo(g.hw, this.px(HH)); ctx.stroke();
    ctx.fillStyle = C.hdrLine;
    ctx.beginPath(); ctx.moveTo(g.hw - 4, HH - 13); ctx.lineTo(g.hw - 4, HH - 4); ctx.lineTo(g.hw - 13, HH - 4); ctx.closePath(); ctx.fill();
    ctx.textAlign = 'left';
  }

  // ---------- hit testing ----------
  pos(e) {
    const r = this.host.getBoundingClientRect();
    return { x: (e.clientX - r.left) / this.z, y: (e.clientY - r.top) / this.z };
  }
  hit(x, y) {
    const g = this.geo(), sh = this.sh;
    const fw = this.fw(), fh = this.fh();
    if (x < g.hw && Math.abs(y - (HH + fh)) <= 3 && (sh.freezeR || y <= HH + 2)) return { area: 'freezeR' };
    if (y < HH && Math.abs(x - (g.hw + fw)) <= 3 && (sh.freezeC || x <= g.hw + 2)) return { area: 'freezeC' };
    if (x < g.hw && y < HH) return { area: 'corner' };
    if (y < HH) {
      const c = this.colAt(x);
      const xr = this.X(c) + this.cw(c);
      if (Math.abs(x - xr) <= 4) return { area: 'colhdr', c, resize: c };
      if (Math.abs(x - this.X(c)) <= 3 && c > 0) { let p = c - 1; while (p > 0 && !this.cw(p)) p--; return { area: 'colhdr', c, resize: this.cw(p) ? p : c - 1 }; }
      return { area: 'colhdr', c };
    }
    if (x < g.hw) {
      const r = this.rowAt(y);
      const yb = this.Y(r) + this.rh(r);
      if (Math.abs(y - yb) <= 3) return { area: 'rowhdr', r, resize: r };
      if (Math.abs(y - this.Y(r)) <= 2 && r > 0) { let p = r - 1; while (p > 0 && !this.rh(p)) p--; return { area: 'rowhdr', r, resize: p }; }
      return { area: 'rowhdr', r };
    }
    const r = this.rowAt(y), c = this.colAt(x);
    const out = { area: 'cell', r, c, x, y };
    // fill handle
    if (this.sel.ranges.length === 1 && !this.edit) {
      const rc = this.rangeRect(this.range);
      if (Math.abs(x - (rc.x + rc.w)) <= 5 && Math.abs(y - (rc.y + rc.h)) <= 5) return { ...out, area: 'fill' };
      const onBorder = (Math.abs(x - rc.x) <= 2 || Math.abs(x - (rc.x + rc.w)) <= 2) && y >= rc.y && y <= rc.y + rc.h ||
        (Math.abs(y - rc.y) <= 2 || Math.abs(y - (rc.y + rc.h)) <= 2) && x >= rc.x && x <= rc.x + rc.w;
      if (onBorder) return { ...out, area: 'border' };
    }
    const m = sh.mergeAt(r, c);
    const ar = m ? m.r1 : r, ac = m ? m.c1 : c;
    const rc = this.cellRect(ar, ac);
    // filter button
    if (sh.filter && ar === sh.filter.r1 && ac >= sh.filter.c1 && ac <= sh.filter.c2 && x >= rc.x + rc.w - 19) return { ...out, area: 'filterbtn', c: ac };
    const dv = sh.dv.length ? dvAt(sh, ar, ac) : null;
    if (dv) {
      if (dv.type === 'checkbox' && Math.abs(x - (rc.x + rc.w / 2)) <= 10 && Math.abs(y - (rc.y + rc.h / 2)) <= 10) return { ...out, area: 'checkbox', r: ar, c: ac };
      if ((dv.type === 'list' || dv.type === 'range') && x >= rc.x + rc.w - 20) return { ...out, area: 'dropdown', r: ar, c: ac };
      if ((dv.type === 'list' || dv.type === 'range') && dv.chip !== 'plain' && this.hasValue(ar, ac)) return { ...out, area: 'chip', r: ar, c: ac };
    }
    return out;
  }

  // ---------- input binding ----------
  bind() {
    const sc = this.scroller;
    sc.addEventListener('scroll', () => { this.render(); this.hideTip(); this.addRows.style.left = (this.geo().hw * this.z + 8 + sc.scrollLeft) + 'px'; });
    sc.addEventListener('pointerdown', e => this.onDown(e));
    sc.addEventListener('pointermove', e => this.onMove(e));
    sc.addEventListener('pointerup', e => this.onUp(e));
    sc.addEventListener('pointercancel', e => this.onUp(e));
    sc.addEventListener('dblclick', e => this.onDbl(e));
    sc.addEventListener('contextmenu', e => this.onContext(e));
    sc.addEventListener('pointerleave', () => this.hideTip());
    sc.addEventListener('wheel', e => {
      if (e.ctrlKey || e.metaKey) { e.preventDefault(); this.app.setZoom(this.z * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); }
    }, { passive: false });
    const inp = this.input;
    inp.addEventListener('keydown', e => this.app.onKey(e));
    inp.addEventListener('input', () => this.onInput());
    inp.addEventListener('compositionstart', () => { if (!this.edit) this.startEdit('', 'enter'); });
    inp.addEventListener('click', () => this.app.fx?.refreshAssist());
    inp.addEventListener('keyup', e => { if (this.edit && ['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) this.app.fx?.refreshAssist(); });
    inp.addEventListener('blur', () => setTimeout(() => this.app.onInputBlur?.(), 0));
    this.dashTimer = null;
  }

  focus() { if (document.activeElement !== this.input) this.input.focus({ preventScroll: true }); }

  onDown(e) {
    if (e.button === 2) return;
    const { x, y } = this.pos(e);
    const h = this.hit(x, y);
    this.app.closePopups?.();
    const edit = this.edit;
    // formula point-and-click
    if (edit && h.area === 'cell' && this.isFormulaEdit() && this.canInsertRef()) {
      e.preventDefault();
      this.drag = { mode: 'point', r: h.r, c: h.c };
      this.scroller.setPointerCapture(e.pointerId);
      this.insertRef({ r1: h.r, c1: h.c, r2: h.r, c2: h.c });
      return;
    }
    if (edit && !(h.area === 'cell' && h.r === edit.r && h.c === edit.c && this.sh === edit.sh)) {
      if (!this.commitEdit(0, 0)) return;
    }
    this.focus();
    e.preventDefault();
    const sh = this.sh;
    switch (h.area) {
      case 'freezeR': this.drag = { mode: 'freezeR', y }; break;
      case 'freezeC': this.drag = { mode: 'freezeC', x }; break;
      case 'corner': this.selectAll(); break;
      case 'colhdr':
        if (h.resize !== undefined) { this.drag = { mode: 'colresize', c: h.resize, x, start: x, w0: this.cw(h.resize) }; break; }
        if (e.shiftKey) this.selectRange({ r1: 0, c1: Math.min(this.sel.c, h.c), r2: sh.rows - 1, c2: Math.max(this.sel.c, h.c) }, 0, this.sel.c, { er: sh.rows - 1, ec: h.c });
        else this.selectRange({ r1: 0, c1: h.c, r2: sh.rows - 1, c2: h.c }, Math.max(this.rowAt(HH + this.fh() + 1), 0), h.c, { add: e.ctrlKey || e.metaKey, er: sh.rows - 1, ec: h.c });
        this.drag = { mode: 'colsel', c0: e.shiftKey ? this.sel.c : h.c };
        break;
      case 'rowhdr':
        if (h.resize !== undefined) { this.drag = { mode: 'rowresize', r: h.resize, y, start: y, h0: this.rh(h.resize) }; break; }
        if (e.shiftKey) this.selectRange({ r1: Math.min(this.sel.r, h.r), c1: 0, r2: Math.max(this.sel.r, h.r), c2: sh.cols - 1 }, this.sel.r, 0, { er: h.r, ec: sh.cols - 1 });
        else this.selectRange({ r1: h.r, c1: 0, r2: h.r, c2: sh.cols - 1 }, h.r, Math.max(this.colAt(this.geo().hw + this.fw() + 1), 0), { add: e.ctrlKey || e.metaKey, er: h.r, ec: sh.cols - 1 });
        this.drag = { mode: 'rowsel', r0: e.shiftKey ? this.sel.r : h.r };
        break;
      case 'fill': this.drag = { mode: 'fill', src: { ...this.range } }; break;
      case 'border': this.drag = { mode: 'move', src: { ...this.range }, r0: h.r, c0: h.c, copy: e.ctrlKey || e.altKey || e.metaKey }; break;
      case 'filterbtn': this.app.openFilterMenu(h.c); return;
      case 'checkbox': this.select(h.r, h.c); this.app.cmd('toggleCheckbox'); return;
      case 'dropdown': this.select(h.r, h.c); this.app.openDropdown(h.r, h.c); return;
      default: {
        if (e.shiftKey) this.extendTo(h.r, h.c);
        else {
          const wasSel = this.sel.r === h.r && this.sel.c === h.c && this.sel.ranges.length === 1;
          this.select(h.r, h.c, { add: e.ctrlKey || e.metaKey, scroll: false });
          if (h.area === 'chip' && wasSel) { this.app.openDropdown(h.r, h.c); return; }
          const cell = sh.get(h.r, h.c);
          const link = cell && (cell.link || cell.hl || (typeof cell.v === 'string' && URL_RE.test(cell.v) ? cell.v : null));
          if (link && (e.ctrlKey || e.metaKey)) { this.app.openLink(link); return; }
        }
        this.drag = { mode: 'sel' };
      }
    }
    this.scroller.setPointerCapture(e.pointerId);
  }

  onMove(e) {
    const { x, y } = this.pos(e);
    const d = this.drag;
    if (!d) { this.hover(x, y, e); return; }
    this.lastPos = { x, y };
    this.autoScroll(x, y);
    const sh = this.sh;
    switch (d.mode) {
      case 'sel': { const r = this.rowAt(y), c = this.colAt(x); if (r !== this.sel.er || c !== this.sel.ec) this.extendTo(r, c); break; }
      case 'colsel': { const c = this.colAt(Math.max(x, this.geo().hw)); this.selectRange({ r1: 0, c1: Math.min(d.c0, c), r2: sh.rows - 1, c2: Math.max(d.c0, c) }, this.sel.r, this.sel.c, { er: sh.rows - 1, ec: c }); break; }
      case 'rowsel': { const r = this.rowAt(Math.max(y, HH)); this.selectRange({ r1: Math.min(d.r0, r), c1: 0, r2: Math.max(d.r0, r), c2: sh.cols - 1 }, this.sel.r, this.sel.c, { er: r, ec: sh.cols - 1 }); break; }
      case 'colresize': d.x = Math.max(this.X(d.c) + 2, x); this.render(); break;
      case 'rowresize': d.y = Math.max(this.Y(d.r) + 2, y); this.render(); break;
      case 'freezeR': d.y = clamp(y, HH, this.vh - 20); this.render(); break;
      case 'freezeC': d.x = clamp(x, this.geo().hw, this.vw - 20); this.render(); break;
      case 'fill': {
        const r = this.rowAt(y), c = this.colAt(x), g = d.src;
        let t = { ...g };
        const dDown = r - g.r2, dUp = g.r1 - r, dRight = c - g.c2, dLeft = g.c1 - c;
        const vert = Math.max(dDown, dUp), hor = Math.max(dRight, dLeft);
        if (vert <= 0 && hor <= 0) { t = null; }
        else if (vert >= hor) { if (dDown > 0) t.r2 = r; else t.r1 = r; }
        else { if (dRight > 0) t.c2 = c; else t.c1 = c; }
        d.target = t; this.render(); break;
      }
      case 'move': {
        const r = this.rowAt(y), c = this.colAt(x), g = d.src;
        const dr = clamp(r - d.r0, -g.r1, sh.rows - 1 - g.r2), dc = clamp(c - d.c0, -g.c1, sh.cols - 1 - g.c2);
        d.dr = dr; d.dc = dc;
        d.target = { r1: g.r1 + dr, c1: g.c1 + dc, r2: g.r2 + dr, c2: g.c2 + dc };
        this.render(); break;
      }
      case 'point': {
        const r = this.rowAt(y), c = this.colAt(x);
        this.insertRef(norm({ r1: d.r, c1: d.c, r2: r, c2: c }));
        break;
      }
    }
  }

  autoScroll(x, y) {
    clearInterval(this.asTimer);
    const vw = this.vw, vh = this.vh, g = this.geo();
    let dx = 0, dy = 0;
    if (x > vw - 4) dx = 1; else if (x < g.hw + this.fw() && this.drag?.mode !== 'rowsel' && this.drag?.mode !== 'colresize') dx = this.sx > 0 && x < g.hw ? -1 : 0;
    if (y > vh - 4) dy = 1; else if (y < HH && this.drag?.mode !== 'colsel' && this.drag?.mode !== 'rowresize') dy = this.sy > 0 ? -1 : 0;
    if (!dx && !dy) return;
    if (!['sel', 'colsel', 'rowsel', 'fill', 'move', 'point'].includes(this.drag?.mode)) return;
    this.asTimer = setInterval(() => {
      if (!this.drag) { clearInterval(this.asTimer); return; }
      this.scroller.scrollLeft += dx * 30; this.scroller.scrollTop += dy * 22;
      const p = this.lastPos;
      this.onMove({ clientX: p.x * this.z + this.host.getBoundingClientRect().left, clientY: p.y * this.z + this.host.getBoundingClientRect().top, fake: true });
    }, 50);
  }

  onUp(e) {
    clearInterval(this.asTimer);
    const d = this.drag;
    this.drag = null;
    try { this.scroller.releasePointerCapture(e.pointerId); } catch { /* ignore */ }
    if (!d) return;
    const sh = this.sh;
    switch (d.mode) {
      case 'colresize': {
        const w = Math.max(2, Math.round(d.x - this.X(d.c)));
        if (Math.abs(d.x - d.start) > 1) this.app.cmd('resizeCols', this.colsFor(d.c), w);
        break;
      }
      case 'rowresize': {
        const h = Math.max(2, Math.round(d.y - this.Y(d.r)));
        if (Math.abs(d.y - d.start) > 1) this.app.cmd('resizeRows', this.rowsFor(d.r), h);
        break;
      }
      case 'freezeR': { const r = d.y <= HH + 4 ? 0 : this.rowAt(d.y + 4); this.app.cmd('freeze', r, sh.freezeC); break; }
      case 'freezeC': { const c = d.x <= this.geo().hw + 4 ? 0 : this.colAt(d.x + 4); this.app.cmd('freeze', sh.freezeR, c); break; }
      case 'fill': if (d.target) this.app.cmd('autofill', d.src, d.target); break;
      case 'move': if (d.target && (d.dr || d.dc)) this.app.cmd('moveRange', d.src, d.dr, d.dc, d.copy); break;
      case 'point': if (this.edit) this.edit.pointDone = true; this.input.focus({ preventScroll: true }); return;
    }
    this.render();
  }

  colsFor(c) {
    const out = [];
    const sh = this.sh;
    for (const g of this.sel.ranges) if (g.r1 === 0 && g.r2 >= sh.rows - 1 && c >= g.c1 && c <= g.c2) { for (const g2 of this.sel.ranges) for (let i = g2.c1; i <= g2.c2; i++) out.push(i); return out; }
    return [c];
  }
  rowsFor(r) {
    const out = [];
    const sh = this.sh;
    for (const g of this.sel.ranges) if (g.c1 === 0 && g.c2 >= sh.cols - 1 && r >= g.r1 && r <= g.r2) { for (const g2 of this.sel.ranges) for (let i = g2.r1; i <= g2.r2; i++) out.push(i); return out; }
    return [r];
  }

  onDbl(e) {
    const { x, y } = this.pos(e);
    const h = this.hit(x, y);
    if (h.area === 'colhdr' && h.resize !== undefined) { this.app.cmd('autofitCols', this.colsFor(h.resize)); return; }
    if (h.area === 'rowhdr' && h.resize !== undefined) { this.app.cmd('autofitRows', this.rowsFor(h.resize)); return; }
    if (h.area === 'fill') { this.app.cmd('fillDownDouble'); return; }
    if (h.area === 'cell' || h.area === 'chip') {
      if (this.edit) return;
      this.startEdit(null, 'edit');
    }
  }

  onContext(e) {
    e.preventDefault();
    const { x, y } = this.pos(e);
    const h = this.hit(x, y);
    if (this.edit) this.commitEdit(0, 0);
    const sh = this.sh;
    const inSel = (r, c) => this.sel.ranges.some(g => inRange(g, r, c));
    if (h.area === 'cell' || h.area === 'chip' || h.area === 'border' || h.area === 'fill' || h.area === 'checkbox' || h.area === 'dropdown') {
      if (!inSel(h.r, h.c)) this.select(h.r, h.c, { scroll: false });
      this.app.contextMenu(e, 'cell');
    } else if (h.area === 'colhdr') {
      if (!this.sel.ranges.some(g => h.c >= g.c1 && h.c <= g.c2 && g.r1 === 0 && g.r2 >= sh.rows - 1)) this.selectRange({ r1: 0, c1: h.c, r2: sh.rows - 1, c2: h.c }, 0, h.c);
      this.app.contextMenu(e, 'col');
    } else if (h.area === 'rowhdr') {
      if (!this.sel.ranges.some(g => h.r >= g.r1 && h.r <= g.r2 && g.c1 === 0 && g.c2 >= sh.cols - 1)) this.selectRange({ r1: h.r, c1: 0, r2: h.r, c2: sh.cols - 1 }, h.r, 0);
      this.app.contextMenu(e, 'row');
    } else this.app.contextMenu(e, 'cell');
  }

  hover(x, y, e) {
    const h = this.hit(x, y);
    const cur = {
      colhdr: h.resize !== undefined ? 'col-resize' : 'default', rowhdr: h.resize !== undefined ? 'row-resize' : 'default',
      fill: 'crosshair', border: 'move', filterbtn: 'pointer', checkbox: 'pointer', dropdown: 'pointer', chip: 'pointer',
      freezeR: 'row-resize', freezeC: 'col-resize', corner: 'default',
    }[h.area] || 'cell';
    this.scroller.style.cursor = cur;
    clearTimeout(this.tipTimer);
    if (h.area !== 'cell' && h.area !== 'chip') { this.hideTip(); return; }
    const sh = this.sh;
    const m = sh.mergeAt(h.r, h.c);
    const r = m ? m.r1 : h.r, c = m ? m.c1 : h.c;
    if (this.tipAt && this.tipAt.r === r && this.tipAt.c === c) return;
    this.hideTip();
    const cell = sh.get(r, c);
    if (!cell) return;
    let html = null;
    if (cell.note) html = `<div class="tip-note">${escapeHtml(cell.note)}</div>`;
    else if (cell.f != null && cell.sv instanceof FErr && cell.sv.msg) html = `<div class="tip-err"><b>Error</b><span>${escapeHtml(cell.sv.msg)}</span></div>`;
    else if (cell.bad) html = `<div class="tip-err"><b>Data tidak valid</b><span>${escapeHtml(cell.bad)}</span></div>`;
    if (!html) return;
    this.tipTimer = setTimeout(() => {
      const rc = this.cellRect(r, c);
      this.tipAt = { r, c };
      this.app.showCellTip(html, (rc.x + rc.w) * this.z + 4, rc.y * this.z);
    }, 280);
  }
  hideTip() { clearTimeout(this.tipTimer); if (this.tipAt) { this.tipAt = null; this.app.hideCellTip(); } }

  // ---------- editor ----------
  isFormulaEdit() { return this.edit && this.input.value.startsWith('='); }
  canInsertRef() {
    const ed = this.edit;
    const src = this.app.fx?.activeInput() || this.input;
    const t = src.value, caret = src.selectionStart;
    if (!t.startsWith('=')) return false;
    if (ed.point && ed.point.end === caret && !ed.pointDone) return true;
    if (ed.point && ed.point.end === caret && ed.pointDone) return true;
    const before = t.slice(0, caret).replace(/\s+$/, '');
    return /[=(,;+\-*/^&<>:]$/.test(before);
  }
  insertRef(g) {
    const ed = this.edit;
    const src = this.app.fx?.activeInput() || this.input;
    const sh = this.sh;
    let text = rangeAddr(g);
    if (sh !== ed.sh) text = rangeAddr(g, sh.name);
    let t = src.value, s, e2;
    if (ed.point && ed.point.end === src.selectionStart) { s = ed.point.start; e2 = ed.point.end; }
    else { s = e2 = src.selectionStart; }
    t = t.slice(0, s) + text + t.slice(e2);
    ed.point = { start: s, end: s + text.length };
    ed.pointDone = false;
    ed.pointSel = { sh, g };
    src.value = t;
    src.setSelectionRange(s + text.length, s + text.length);
    this.onInput(src);
  }

  editTextOf(r, c) { return this.app.editTextOf(this.sh, r, c); }

  startEdit(initial, mode = 'edit', src = 'cell') {
    if (this.edit) return;
    const { r, c } = this.sel;
    if (!this.app.canEdit(r, c)) { this.input.value = ''; return; }
    const orig = this.editTextOf(r, c);
    this.edit = { r, c, sh: this.sh, orig, mode, src, point: null, refs: [] };
    this.input.value = initial == null ? orig : initial;
    this.input.classList.add('editing');
    this.placeEditor();
    if (src === 'cell') {
      this.focus();
      const L = this.input.value.length;
      this.input.setSelectionRange(L, L);
    }
    this.onInput();
  }

  onInput(srcEl) {
    if (!this.edit) {
      // text typed without an active edit (e.g. after IME) — start one
      if (this.input.value) { const v = this.input.value; this.startEdit(v, 'enter'); }
      return;
    }
    const src = srcEl || this.app.fx?.activeInput() || this.input;
    const text = src.value;
    if (src === this.input) this.app.fx?.mirror(text); else this.input.value = text;
    if (this.edit.point && src.selectionStart !== this.edit.point.end) { this.edit.point = null; this.edit.pointSel = null; }
    this.edit.refs = text.startsWith('=') ? refsOf(text) : [];
    if (text.startsWith('=')) this.edit.mode = this.edit.mode === 'edit' ? 'edit' : 'enter';
    this.placeEditor();
    this.app.fx?.refreshAssist();
    this.render();
  }

  placeEditor() {
    const inp = this.input, ed = this.edit;
    if (!ed || ed.sh !== this.sh) {
      if (ed && ed.sh !== this.sh) { inp.classList.add('away'); }
      else { inp.classList.remove('away'); this.placeInput(); }
      this.hilite.style.display = 'none';
      return;
    }
    inp.classList.remove('away');
    const z = this.z, def = this.app.defFont;
    const rc = this.cellRect(ed.r, ed.c);
    const st = this.sh.styleOf(ed.r, ed.c);
    const font = fontOf(st, def);
    const lines = inp.value.split('\n');
    let tw = 0;
    for (const ln of lines) tw = Math.max(tw, this.measure(font, ln || ' '));
    const minW = rc.w, maxW = Math.max(rc.w, this.vw - rc.x - 4);
    let w = Math.max(minW, tw + 14);
    let wrapMode = st.wrap === 'wrap';
    if (w > maxW) { w = maxW; wrapMode = true; }
    if (wrapMode) w = Math.max(minW, Math.min(maxW, w));
    const px = ((st.fs || def.fs) * 4) / 3;
    const isF = inp.value.startsWith('=');
    Object.assign(inp.style, {
      left: rc.x * z + 'px', top: rc.y * z + 'px', width: w * z + 'px',
      font: font.replace(/[\d.]+px/, (px * z).toFixed(2) + 'px'),
      lineHeight: (px * 1.25 * z).toFixed(2) + 'px',
      textAlign: isF ? 'left' : st.ha || 'left',
      whiteSpace: wrapMode ? 'pre-wrap' : 'pre',
      color: isF ? 'transparent' : st.color || '',
      background: st.bg || '',
      minHeight: rc.h * z + 'px', height: 'auto',
    });
    inp.style.height = Math.max(rc.h * z, inp.scrollHeight) + 'px';
    if (isF) {
      const hl = this.hilite;
      hl.style.display = 'block';
      Object.assign(hl.style, { left: inp.style.left, top: inp.style.top, width: inp.style.width, height: inp.style.height, font: inp.style.font, lineHeight: inp.style.lineHeight, whiteSpace: inp.style.whiteSpace });
      hl.innerHTML = highlightFormula(inp.value, ed.refs) + '​';
    } else this.hilite.style.display = 'none';
  }

  placeInput() {
    if (this.edit) return;
    const inp = this.input;
    const rc = this.cellRect(this.sel.r, this.sel.c);
    inp.style.left = clamp(rc.x * this.z, 0, this.W - 4) + 'px';
    inp.style.top = clamp(rc.y * this.z, 0, this.H - 4) + 'px';
    inp.style.width = '2px';
    inp.style.height = rc.h * this.z + 'px';
  }

  endEdit() {
    this.edit = null;
    this.input.value = '';
    this.input.classList.remove('editing', 'away');
    this.input.removeAttribute('style');
    this.hilite.style.display = 'none';
    this.app.fx?.endEdit();
    this.placeInput();
    this.render();
  }

  commitEdit(dr = 1, dc = 0, opts = {}) {
    const ed = this.edit;
    if (!ed) return true;
    let text = autoClose(this.input.value);
    if (opts.array && text.startsWith('=') && !/^=ARRAYFORMULA\(/i.test(text)) text = '=ARRAYFORMULA(' + text.slice(1) + ')';
    const back = ed.sh !== this.sh;
    if (back) this.app.switchSheet(this.wb.sheets.indexOf(ed.sh), true);
    const ok = this.app.commitCell(ed.sh, ed.r, ed.c, text, ed.orig, opts);
    if (ok === false) return false;
    this.endEdit();
    if (opts.cycle) this.cycle(dr, dc);
    else if (dr || dc) this.move(dr, dc);
    else this.changed();
    return true;
  }

  cancelEdit() {
    const ed = this.edit;
    if (!ed) return;
    if (ed.sh !== this.sh) this.app.switchSheet(this.wb.sheets.indexOf(ed.sh), true);
    this.endEdit();
    this.changed();
  }

  setZoom(z) {
    const cx = this.sx, cy = this.sy;
    this.z = clamp(z, 0.5, 2);
    this.g = null;
    this.measureCache.clear();
    this.geo();
    this.scroller.scrollLeft = cx * this.z;
    this.scroller.scrollTop = cy * this.z;
    this.render();
  }

  startMarquee(sh, g) {
    this.copyRange = { sh, g };
    clearInterval(this.dashTimer);
    this.dashTimer = setInterval(() => { this.dash = (this.dash + 1) % 18; this.render(); }, 60);
  }
  stopMarquee() { this.copyRange = null; clearInterval(this.dashTimer); this.render(); }
}

function escapeHtml(s) { return String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }

// Reference list with color indexes for a formula text
export function refsOf(text) {
  const out = [];
  const seen = new Map();
  for (const t of lex(text.slice(1))) {
    if (t.t !== 'ref') continue;
    const k = (t.sheet || '').toLowerCase() + '!' + t.s.replace(/\$/g, '').toUpperCase().replace(/^.*!/, '');
    if (!seen.has(k)) seen.set(k, seen.size);
    const a = t.a;
    out.push({ sheet: t.sheet, r1: a.r1, c1: a.c1, r2: a.r2 === Infinity ? 1048575 : a.r2, c2: a.c2 === Infinity ? 16383 : a.c2, ci: seen.get(k), p: t.p + 1, len: t.s.length });
  }
  return out;
}

export function highlightFormula(text, refs) {
  if (!text.startsWith('=')) return escapeHtml(text);
  let out = '', i = 0;
  const toks = lex(text.slice(1));
  out += '=';
  i = 1;
  let ri = 0;
  for (const t of toks) {
    const s = escapeHtml(t.s);
    if (t.t === 'ref') { const rf = refs[ri++]; out += `<span style="color:${REF_COLORS[(rf ? rf.ci : 0) % REF_COLORS.length]}">${s}</span>`; }
    else if (t.t === 'str') out += `<span class="hl-str">${s}</span>`;
    else if (t.t === 'num') out += `<span class="hl-num">${s}</span>`;
    else if (t.t === 'id') out += `<span class="hl-fn">${s}</span>`;
    else out += s;
    i += t.s.length;
  }
  return out;
}
