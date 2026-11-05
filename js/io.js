// Import/export, printing and data helpers (current region, flash fill, shuffle).
import { el, icon, key, keyR, keyC, colName, loadScript, LIBS, download, esc, rangeAddr, parseRange, uid } from './util.js';
import * as PG from './pages.js';
import { Workbook, newWorkbook, setInput } from './model.js';
import { parseInput, format, compile } from './format.js';
import { FErr } from './formula/core.js';
import { Special } from './formula/functions.js';
import { dialog, toast, select, checkbox, field } from './ui/ui.js';
import { snapRange, writeSnap } from './ops.js';

export function peekValue(text) {
  const p = parseInput(text);
  return p.f != null ? null : p.v;
}

// ---------- region & helpers ----------
export function currentRegion(app, r, c) {
  const sh = app.sheet, has = (rr, cc) => rr >= 0 && cc >= 0 && rr < sh.rows && cc < sh.cols && app.grid.hasValue(rr, cc);
  let g = { r1: r, c1: c, r2: r, c2: c };
  for (let changed = true, guard = 0; changed && guard < 2000; guard++) {
    changed = false;
    const rowHas = rr => { for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) if (has(rr, cc)) return true; return false; };
    const colHas = cc => { for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) if (has(rr, cc)) return true; return false; };
    if (g.r1 > 0 && rowHas(g.r1 - 1)) { g.r1--; changed = true; }
    if (g.r2 < sh.rows - 1 && rowHas(g.r2 + 1)) { g.r2++; changed = true; }
    if (g.c1 > 0 && colHas(g.c1 - 1)) { g.c1--; changed = true; }
    if (g.c2 < sh.cols - 1 && colHas(g.c2 + 1)) { g.c2++; changed = true; }
  }
  return g;
}

export function guessChart(app, g) {
  const sh = app.sheet;
  const c = sh.get(g.r1 + 1, g.c1);
  const fmt = app.wb.styles[c?.s || 0]?.fmt || c?.af || '';
  if (typeof app.wb.eng.value(sh, g.r1 + 1, g.c1) === 'number' && compile(fmt)?.isDate) return 'line';
  if (g.c2 - g.c1 === 1 && g.r2 - g.r1 <= 7 && typeof app.wb.eng.value(sh, g.r1 + 1, g.c1) === 'string') return 'pie';
  return 'column';
}

export function shuffleRange(app, g) {
  const sh = app.sheet;
  const rows = [];
  for (let r = g.r1; r <= g.r2; r++) rows.push(snapRange(sh, { r1: r, c1: g.c1, r2: r, c2: g.c2 })[0]);
  for (let i = rows.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [rows[i], rows[j]] = [rows[j], rows[i]]; }
  writeSnap(sh, g.r1, g.c1, rows);
}

// ---------- flash fill (Ctrl+E) ----------
const DELIMS = [' ', ',', '-', '@', '.', '/', '_', ';', '|', ':'];
const CASES = {
  as: s => s, up: s => s.toUpperCase(), lo: s => s.toLowerCase(),
  pr: s => s.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (m, a, b) => a + b.toUpperCase()),
  ini: s => s.slice(0, 1).toUpperCase(),
};
function atoms(srcs) {
  const out = [];
  srcs.forEach((s, j) => {
    if (s == null || s === '') return;
    const t = String(s);
    for (const [cn, cf] of Object.entries(CASES)) out.push({ k: `${j}|*|0|${cn}`, v: cf(t) });
    for (const d of DELIMS) {
      if (!t.includes(d)) continue;
      const parts = t.split(d).filter(x => x !== '');
      parts.forEach((p, i) => {
        for (const [cn, cf] of Object.entries(CASES)) {
          out.push({ k: `${j}|${d}|${i}|${cn}`, v: cf(p) });
          out.push({ k: `${j}|${d}|${i - parts.length}|${cn}`, v: cf(p) });
        }
      });
      const first = t.indexOf(d);
      out.push({ k: `${j}|${d}|before|as`, v: t.slice(0, first) }, { k: `${j}|${d}|after|as`, v: t.slice(first + d.length) });
    }
  });
  return out;
}
function evalAtom(k, srcs) {
  const [j, d, i, cn] = k.split('|');
  const t = srcs[+j];
  if (t == null || t === '') return null;
  const s = String(t);
  if (d === '*') return CASES[cn](s);
  if (i === 'before') { const p = s.indexOf(d); return p < 0 ? null : s.slice(0, p); }
  if (i === 'after') { const p = s.indexOf(d); return p < 0 ? null : s.slice(p + d.length); }
  const parts = s.split(d).filter(x => x !== '');
  const n = +i;
  const p = n >= 0 ? parts[n] : parts[parts.length + n];
  return p == null ? null : CASES[cn](p);
}
function programsFor(out, srcs) {
  const as = atoms(srcs);
  const progs = new Set();
  for (const a of as) if (a.v === out) progs.add(JSON.stringify([a.k]));
  if (progs.size) return progs;
  for (const a of as) {
    if (!a.v || !out.startsWith(a.v)) continue;
    const rest = out.slice(a.v.length);
    for (const b of as) {
      if (!b.v || !rest.endsWith(b.v)) continue;
      const lit = rest.slice(0, rest.length - b.v.length);
      if (lit.length <= 3 && !/[\p{L}\d]/u.test(lit)) progs.add(JSON.stringify([a.k, lit, b.k]));
      const mid = rest.slice(0, rest.length - b.v.length);
      for (const c2 of as) {
        if (!c2.v || !mid.includes(c2.v)) continue;
        const idx = mid.indexOf(c2.v);
        const l1 = mid.slice(0, idx), l2 = mid.slice(idx + c2.v.length);
        if (l1.length <= 3 && l2.length <= 3 && !/[\p{L}\d]/u.test(l1 + l2)) progs.add(JSON.stringify([a.k, l1, c2.k, l2, b.k]));
      }
    }
  }
  return progs;
}
function runProg(p, srcs) {
  let s = '';
  for (let i = 0; i < p.length; i++) {
    if (i % 2) { s += p[i]; continue; }
    const v = evalAtom(p[i], srcs);
    if (v == null) return null;
    s += v;
  }
  return s;
}
export function flashFill(app) {
  const sh = app.sheet, eng = app.wb.eng, c = app.grid.sel.c;
  const reg = currentRegion(app, app.grid.sel.r, Math.max(0, c - 1));
  const r1 = reg.r1, r2 = Math.max(reg.r2, app.grid.sel.r);
  const srcCols = [];
  for (let cc = reg.c1; cc <= Math.max(reg.c2, c); cc++) if (cc !== c) srcCols.push(cc);
  const srcOf = r => srcCols.map(cc => { const v = eng.value(sh, r, cc); return v == null || v instanceof FErr ? null : typeof v === 'number' ? format(v, sh.styleOf(r, cc).fmt).t : String(v); });
  let header = typeof eng.value(sh, r1, c) === 'string' && srcCols.every(cc => typeof eng.value(sh, r1, cc) === 'string') && r1 === 0;
  const examples = [], targets = [];
  for (let r = r1; r <= r2; r++) {
    const v = eng.value(sh, r, c);
    if (v != null && v !== '') examples.push({ r, out: String(v) });
    else if (srcOf(r).some(x => x != null)) targets.push(r);
  }
  const ex = header ? examples.filter(e => e.r !== r1) : examples;
  const usable = ex.length ? ex : examples;
  if (!usable.length || !targets.length) return 0;
  let progs = null;
  for (const e of usable) {
    const p = programsFor(e.out, srcOf(e.r));
    if (!p.size) continue;
    progs = progs ? new Set([...progs].filter(x => p.has(x))) : p;
    if (!progs.size) return 0;
  }
  if (!progs || !progs.size) return 0;
  // prefer the simplest program
  const best = JSON.parse([...progs].sort((a, b) => a.length - b.length)[0]);
  let n = 0;
  for (const r of targets) {
    if (r <= usable[0].r) continue;
    const out = runProg(best, srcOf(r));
    if (out != null) { setInput(app.wb, sh, r, c, out.startsWith('=') ? "'" + out : out); n++; }
  }
  return n;
}

// ---------- CSV ----------
export function parseCSV(text, delim) {
  const t = text.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
  if (!delim) {
    const sample = t.split('\n').slice(0, 10).join('\n');
    const cnt = d => (sample.match(new RegExp('\\' + d, 'g')) || []).length;
    delim = cnt('\t') > cnt(',') && cnt('\t') >= cnt(';') ? '\t' : cnt(';') > cnt(',') ? ';' : ',';
  }
  const rows = [];
  let row = [], cur = '', q = false;
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) { if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
    if (ch === '"' && cur === '') { q = true; continue; }
    if (ch === delim) { row.push(cur); cur = ''; continue; }
    if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; continue; }
    cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function csvWorkbook(name, rows) {
  const wb = newWorkbook(name.replace(/\.[^.]+$/, ''));
  const sh = wb.sheets[0];
  sh.name = name.replace(/\.[^.]+$/, '').slice(0, 60) || 'Sheet1';
  wb.hist.paused++;
  sh.growTo(Math.max(1000, rows.length + 100), Math.max(26, ...rows.map(r => r.length + 1)));
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') setInput(wb, sh, r, c, v); }));
  wb.hist.paused--;
  return wb;
}

// ---------- import ----------
export async function importFile(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (ext === 'json') {
    const j = JSON.parse(await file.text());
    if (!j.sheets) throw new Error('File JSON bukan spreadsheet Amatu Sheets.');
    const wb = Workbook.fromJSON(j);
    wb.id = 'wb' + Date.now().toString(36);
    wb.share = null;
    return wb;
  }
  if (ext === 'csv' || ext === 'tsv' || ext === 'txt') {
    const text = await file.text();
    return csvWorkbook(file.name, parseCSV(text, ext === 'tsv' ? '\t' : null));
  }
  await loadScript(LIBS.xlsx);
  const X = window.XLSX;
  const data = new Uint8Array(await file.arrayBuffer());
  const xb = X.read(data, { type: 'array', cellStyles: true, cellFormula: true, cellNF: true, cellDates: false });
  const wb = fromSheetJS(xb, file.name.replace(/\.[^.]+$/, ''));
  if (ext === 'xlsx' || ext === 'xlsm') {
    try { readValidations(X, data, wb); } catch (e) { console.warn('Validasi data tidak dapat dibaca', e); }   // the cells are already in
  }
  return wb;
}

// SheetJS CE drops data validation, so it is read from the file's XML: dropdowns and number/date/length/custom rules,
// including the x14 form Excel uses when a list points at another sheet.
const DV_OPS = { between: 'between', notBetween: 'notBetween', greaterThan: 'gt', greaterThanOrEqual: 'ge', lessThan: 'lt', lessThanOrEqual: 'le', equal: 'eq', notEqual: 'ne' };
const dvOperand = t => (t == null || t === '' ? undefined : !isNaN(Number(t)) ? t : '=' + t.replace(/^=/, ''));
const childText = (node, name) => { for (const k of node.children) if (k.localName === name) return k.textContent.trim(); return null; };

function dvRule(d) {
  const type = d.getAttribute('type'), f1 = childText(d, 'formula1'), f2 = childText(d, 'formula2');
  const rule = { id: uid('dv'), strict: !['warning', 'information'].includes(d.getAttribute('errorStyle')) };
  if (d.getAttribute('error')) rule.help = d.getAttribute('error');
  if (type === 'list') {
    if (!f1) return null;
    if (f1.startsWith('"')) Object.assign(rule, { type: 'list', items: f1.slice(1, -1).replace(/""/g, '"').split(',').map(x => x.trim()) });
    else Object.assign(rule, { type: 'range', src: f1.replace(/^=/, '') });
    // in the file format, showDropDown="1" oddly means the arrow is hidden
    rule.chip = d.getAttribute('showDropDown') === '1' ? 'plain' : 'arrow';
    return rule;
  }
  if (type === 'custom') return f1 ? Object.assign(rule, { type: 'custom', v1: dvOperand(f1) }) : null;
  const kind = { decimal: 'number', whole: 'number', date: 'date', textLength: 'textlen' }[type];
  if (!kind || !f1) return null;
  return Object.assign(rule, { type: kind, op: DV_OPS[d.getAttribute('operator') || 'between'] || 'between', v1: dvOperand(f1), v2: dvOperand(f2) });
}

function readValidations(X, data, wb) {
  const zip = X.CFB.read(data, { type: 'array' });
  const xml = path => { const e = X.CFB.find(zip, '/' + path); return e ? new DOMParser().parseFromString(new TextDecoder().decode(e.content), 'application/xml') : null; };
  const book = xml('xl/workbook.xml'), rels = xml('xl/_rels/workbook.xml.rels');
  if (!book || !rels) return;
  const target = new Map([...rels.getElementsByTagNameNS('*', 'Relationship')].map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
  for (const s of book.getElementsByTagNameNS('*', 'sheet')) {
    const sh = wb.sheets.find(x => x.name === s.getAttribute('name'));
    const rid = [...s.attributes].find(a => a.localName === 'id')?.value;
    const t = target.get(rid);
    if (!sh || !t) continue;
    const doc = xml(t.startsWith('/') ? t.slice(1) : 'xl/' + t);
    if (!doc) continue;
    for (const d of doc.getElementsByTagNameNS('*', 'dataValidation')) {
      const rule = dvRule(d);
      if (!rule) continue;
      const refs = (d.getAttribute('sqref') || childText(d, 'sqref') || '').split(/\s+/).filter(Boolean);
      rule.ranges = refs.map(parseRange).filter(g => g && g.r1 < sh.rows && g.c1 < sh.cols)
        .map(g => ({ r1: g.r1, c1: g.c1, r2: Math.min(g.r2, sh.rows - 1), c2: Math.min(g.c2, sh.cols - 1) }));
      if (rule.ranges.length) sh.dv.push(rule);
    }
  }
}

const XL_BORDER = { thin: [1, 'solid'], hair: [1, 'dotted'], dotted: [1, 'dotted'], dashed: [1, 'dashed'], medium: [2, 'solid'], mediumDashed: [2, 'dashed'], thick: [3, 'solid'], double: [3, 'double'], dashDot: [1, 'dashed'], mediumDashDot: [2, 'dashed'], slantDashDot: [2, 'dashed'], dashDotDot: [1, 'dotted'], mediumDashDotDot: [2, 'dotted'] };
const rgb = c => (c && c.rgb && c.rgb.length >= 6 ? '#' + c.rgb.slice(-6).toLowerCase() : null);

function fromSheetJS(xb, name) {
  const X = window.XLSX;
  const wb = new Workbook();
  wb.name = name;
  wb.hist.paused++;
  for (const sn of xb.SheetNames) {
    const ws = xb.Sheets[sn];
    const ref = ws['!ref'] ? X.utils.decode_range(ws['!ref']) : { e: { r: 0, c: 0 } };
    const sh = wb.addSheet(sn, wb.sheets.length, Math.max(1000, ref.e.r + 101), Math.max(26, ref.e.c + 2));
    for (const k of Object.keys(ws)) {
      if (k[0] === '!') continue;
      const c = ws[k];
      const p = X.utils.decode_cell(k);
      const st = {};
      if (c.z && c.z !== 'General') st.fmt = String(c.z).replace(/\\-/g, '-').replace(/\\ /g, ' ');
      const s = c.s;
      if (s) {
        const font = s.font || {};
        if (font.bold) st.b = true; if (font.italic) st.i = true; if (font.underline) st.u = true; if (font.strike) st.s = true;
        const fc = rgb(font.color); if (fc && fc !== '#000000') st.color = fc;
        if (font.sz && +font.sz !== 11 && +font.sz !== 10) st.fs = +font.sz;
        if (font.name && !/^(Calibri|Arial)$/i.test(font.name)) st.ff = font.name;
        const fill = s.fill || s;
        const bg = rgb(fill.fgColor) || (fill.patternType === 'solid' ? rgb(fill.bgColor) : null);
        if (bg && (fill.patternType === 'solid' || s.patternType === 'solid') && bg !== '#ffffff') st.bg = bg;
        const al = s.alignment || {};
        if (al.horizontal && ['left', 'center', 'right'].includes(al.horizontal)) st.ha = al.horizontal;
        if (al.vertical) st.va = al.vertical === 'center' ? 'middle' : al.vertical;
        if (al.wrapText) st.wrap = 'wrap';
        if (al.textRotation) st.rot = al.textRotation === 255 ? 'v' : al.textRotation > 90 ? 90 - al.textRotation : al.textRotation;
        const bd = s.border || {};
        for (const [side, keyN] of [['top', 'bt'], ['bottom', 'bb'], ['left', 'bl'], ['right', 'br']]) {
          const b = bd[side];
          if (b && b.style && XL_BORDER[b.style]) st[keyN] = { w: XL_BORDER[b.style][0], style: XL_BORDER[b.style][1], color: rgb(b.color) || '#000000' };
        }
      }
      const patch = { s: wb.intern(st) };
      if (c.f) { patch.f = String(c.f).replace(/^=/, ''); patch.v = null; }
      else if (c.t === 'n') patch.v = c.v;
      else if (c.t === 'b') patch.v = !!c.v;
      else if (c.t === 'e') patch.v = c.w || '#N/A';
      else if (c.t === 'd') patch.v = (new Date(c.v).getTime() - Date.UTC(1899, 11, 30)) / 86400000;
      else if (c.t === 's' || c.t === 'str') patch.v = c.v == null ? null : String(c.v);
      else patch.v = null;
      if (c.l && c.l.Target) patch.link = c.l.Target;
      if (c.c && c.c.length) patch.note = c.c.map(x => x.t).join('\n');
      sh.put(p.r, p.c, patch);
    }
    sh.merges = (ws['!merges'] || []).map(m => ({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }));
    (ws['!cols'] || []).forEach((col, i) => {
      if (!col) return;
      if (col.hidden) sh.hiddenC.add(i);
      const w = col.wpx || (col.wch ? Math.round(col.wch * 7 + 5) : col.width ? Math.round(col.width * 7 + 5) : null);
      if (w) sh.colW.set(i, w);
    });
    (ws['!rows'] || []).forEach((row, i) => {
      if (!row) return;
      if (row.hidden) sh.hiddenR.add(i);
      const h = row.hpx || (row.hpt ? Math.round(row.hpt * 4 / 3) : null);
      if (h) { sh.rowH.set(i, h); sh.manualH.add(i); }
    });
    if (ws['!autofilter']?.ref) { const r = X.utils.decode_range(ws['!autofilter'].ref); sh.filter = { r1: r.s.r, c1: r.s.c, r2: r.e.r, c2: r.e.c, crit: {} }; }
    sh.recomputeExtent();
  }
  for (const n of xb.Workbook?.Names || []) {
    if (!n.Name || n.Name.startsWith('_xlnm')) continue;
    wb.names.set(n.Name.toUpperCase(), { name: n.Name, ref: String(n.Ref).replace(/^=/, '') });
  }
  if (!wb.sheets.length) wb.addSheet('Sheet1');
  wb.hist.paused--;
  return wb;
}

// ---------- export ----------
function cellDisplay(app, sh, r, c) {
  const cell = sh.get(r, c);
  let v = cell ? (cell.f != null ? cell.sv : cell.v) : sh.spillMap.get(key(r, c))?.v ?? null;
  if (v == null || v instanceof Special) return '';
  if (v instanceof FErr) return v.code;
  const st = app.wb.styles[cell ? cell.s : 0] || {};
  return format(v, st.fmt || cell?.af).t;
}

export function exportCSV(app, delim) {
  const sh = app.sheet;
  const u = sh.usedRange();
  const lines = [];
  if (u) for (let r = 0; r <= u.r2; r++) {
    const row = [];
    for (let c = 0; c <= u.c2; c++) {
      const t = cellDisplay(app, sh, r, c);
      row.push(new RegExp(`["\\n${delim === '\t' ? '\\t' : delim}]`).test(t) ? '"' + t.replace(/"/g, '""') + '"' : t);
    }
    lines.push(row.join(delim));
  }
  download(`${app.wb.name} - ${sh.name}.${delim === '\t' ? 'tsv' : 'csv'}`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
}

export function exportJSON(app) {
  download(app.wb.name + '.json', JSON.stringify(app.wb.toJSON()), 'application/json');
}

const hex6 = c => (c ? String(c).replace('#', '').toUpperCase().padEnd(6, '0').slice(0, 6) : undefined);
const XL_STYLE_OUT = b => (b.style === 'double' ? 'double' : b.style === 'dashed' ? (b.w >= 2 ? 'mediumDashed' : 'dashed') : b.style === 'dotted' ? 'dotted' : b.w >= 3 ? 'thick' : b.w === 2 ? 'medium' : 'thin');

export async function exportXLSX(app, bookType = 'xlsx') {
  try { await loadScript(LIBS.xlsx); } catch { toast('Pustaka ekspor gagal dimuat. Periksa koneksi internet.'); return; }
  const X = window.XLSX, wb = app.wb;
  const out = X.utils.book_new();
  for (const sh of wb.sheets) {
    const ws = {};
    let R = 0, C = 0;
    for (const [k, cell] of sh.cells) {
      const r = keyR(k), c = keyC(k);
      const st = wb.styles[cell.s] || {};
      const v = cell.f != null ? cell.sv : cell.v;
      const o = {};
      if (typeof v === 'number') { o.t = 'n'; o.v = v; }
      else if (typeof v === 'boolean') { o.t = 'b'; o.v = v; }
      else if (v instanceof FErr) { o.t = 'e'; o.v = { '#NULL!': 0, '#DIV/0!': 7, '#VALUE!': 15, '#REF!': 23, '#NAME?': 29, '#NUM!': 36, '#N/A': 42 }[v.code] ?? 42; o.w = v.code; }
      else if (v == null || v instanceof Special) { o.t = 's'; o.v = ''; }
      else { o.t = 's'; o.v = String(v); }
      if (cell.f != null) o.f = cell.f;
      const fmt = st.fmt || cell.af;
      if (fmt) o.z = fmt;
      const s = {};
      const font = {};
      if (st.b) font.bold = true; if (st.i) font.italic = true; if (st.u) font.underline = true; if (st.s) font.strike = true;
      if (st.color) font.color = { rgb: hex6(st.color) };
      if (st.fs) font.sz = st.fs; if (st.ff) font.name = st.ff;
      if (cell.link) { font.color = font.color || { rgb: '1155CC' }; font.underline = true; o.l = { Target: cell.link }; }
      if (Object.keys(font).length) s.font = font;
      if (st.bg) s.fill = { patternType: 'solid', fgColor: { rgb: hex6(st.bg) } };
      const al = {};
      if (st.ha) al.horizontal = st.ha; if (st.va) al.vertical = st.va === 'middle' ? 'center' : st.va;
      if (st.wrap === 'wrap') al.wrapText = true;
      if (st.rot) al.textRotation = st.rot === 'v' ? 255 : st.rot < 0 ? 90 - st.rot : st.rot;
      if (st.ind) al.indent = st.ind;
      if (Object.keys(al).length) s.alignment = al;
      const bd = {};
      for (const [side, kk] of [['top', 'bt'], ['bottom', 'bb'], ['left', 'bl'], ['right', 'br']]) if (st[kk]) bd[side] = { style: XL_STYLE_OUT(st[kk]), color: { rgb: hex6(st[kk].color || '#000000') } };
      if (Object.keys(bd).length) s.border = bd;
      if (fmt) s.numFmt = fmt;
      if (Object.keys(s).length) o.s = s;
      if (cell.note) o.c = [{ a: 'Amatu Sheets', t: cell.note }];
      ws[X.utils.encode_cell({ r, c })] = o;
      R = Math.max(R, r); C = Math.max(C, c);
    }
    // spilled values are written as plain values so other apps see them
    for (const [k, sp] of sh.spillMap) {
      const r = keyR(k), c = keyC(k);
      const a = X.utils.encode_cell({ r, c });
      if (ws[a] && ws[a].f) continue;
      const v = sp.v;
      ws[a] = typeof v === 'number' ? { t: 'n', v } : typeof v === 'boolean' ? { t: 'b', v } : { t: 's', v: v == null ? '' : v instanceof FErr ? v.code : String(v) };
      R = Math.max(R, r); C = Math.max(C, c);
    }
    ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: R, c: C } });
    if (sh.merges.length) ws['!merges'] = sh.merges.map(m => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
    const cols = [];
    for (let c = 0; c <= Math.max(C, ...sh.colW.keys(), 0); c++) cols.push({ wpx: sh.colW.get(c) ?? 100, hidden: sh.hiddenC.has(c) || undefined });
    ws['!cols'] = cols;
    const rows = [];
    for (let r = 0; r <= Math.max(R, ...sh.rowH.keys(), 0); r++) rows.push(sh.rowH.has(r) || sh.hiddenR.has(r) ? { hpx: sh.rowH.get(r) ?? 21, hidden: sh.hiddenR.has(r) || undefined } : {});
    ws['!rows'] = rows;
    if (sh.page?.m) { const m = sh.page.m, i = v => +(v / 2.54).toFixed(3); ws['!margins'] = { left: i(m.l), right: i(m.r), top: i(m.t), bottom: i(m.b), header: 0.3, footer: 0.3 }; }
    if (sh.filter) ws['!autofilter'] = { ref: X.utils.encode_range({ s: { r: sh.filter.r1, c: sh.filter.c1 }, e: { r: sh.filter.r2, c: sh.filter.c2 } }) };
    let name = sh.name.replace(/[[\]*?/\\:]/g, ' ').slice(0, 31) || 'Sheet';
    let n = 2; while (out.SheetNames.includes(name)) name = name.slice(0, 28) + ' ' + n++;
    X.utils.book_append_sheet(out, ws, name);
  }
  if (wb.names.size) out.Workbook = { Names: [...wb.names.values()].map(n => ({ Name: n.name, Ref: n.ref })) };
  X.writeFile(out, `${wb.name}.${bookType}`, { bookType, cellStyles: true, compression: true });
}

function cssBorder(b) { return b ? `${b.style === 'double' ? 3 : b.w || 1}px ${b.style || 'solid'} ${b.color || '#000'}` : ''; }

export function sheetHTML(app, sh, g, opts = {}) {
  const wb = app.wb;
  // opts.clip: merges are cut to the range g (a printed page), the cut part shows the merged cell's content
  const merged = new Map();
  const skip = new Set();
  for (const m0 of sh.merges) {
    const m = opts.clip ? { r1: Math.max(m0.r1, g.r1), c1: Math.max(m0.c1, g.c1), r2: Math.min(m0.r2, g.r2), c2: Math.min(m0.c2, g.c2), src: m0 } : m0;
    if (m.r1 > m.r2 || m.c1 > m.c2) continue;
    merged.set(key(m.r1, m.c1), m);
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) skip.add(key(r, c));
  }
  let tw = 0;
  for (let c = g.c1; c <= g.c2; c++) tw += sh.colWidth(c);
  let html = `<table class="sheet${opts.grid ? ' grid' : ''}"${opts.clip ? ` style="width:${tw}px"` : ''}><colgroup>`;
  for (let c = g.c1; c <= g.c2; c++) if (!sh.hiddenC.has(c)) html += `<col style="width:${sh.colWidth(c)}px">`;
  html += '</colgroup>';
  for (let r = g.r1; r <= g.r2; r++) {
    if (sh.isRowHidden(r)) continue;
    const h = opts.clip ? sh.rowHeight(r) : sh.rowH.get(r);
    html += `<tr${h ? ` style="height:${h}px"` : ''}>`;
    for (let c = g.c1; c <= g.c2; c++) {
      if (sh.hiddenC.has(c) || skip.has(key(r, c))) continue;
      const mm = merged.get(key(r, c));
      const [sr, sc] = mm?.src ? [mm.src.r1, mm.src.c1] : [r, c];
      const cell = sh.get(sr, sc);
      const st = wb.styles[cell ? cell.s : 0] || {};
      const cfr = app.cf.result(sh, sr, sc) || {};
      const m = mm;
      const v = cell ? (cell.f != null ? cell.sv : cell.v) : sh.spillMap.get(key(r, c))?.v;
      const css = [];
      const bg = cfr.bg || st.bg || app.bandingFor(sh, r, c);
      if (bg) css.push('background:' + bg);
      const color = cfr.color || st.color || (v != null && typeof v === 'number' && format(v, st.fmt || cell?.af).color);
      if (color) css.push('color:' + color);
      if (st.b || cfr.b) css.push('font-weight:700'); if (st.i || cfr.i) css.push('font-style:italic');
      if (st.u || st.s) css.push('text-decoration:' + [st.u ? 'underline' : '', st.s ? 'line-through' : ''].join(' '));
      if (st.ff) css.push(`font-family:'${st.ff}'`); if (st.fs) css.push(`font-size:${st.fs}pt`);
      css.push('text-align:' + (st.ha || (typeof v === 'number' ? 'right' : typeof v === 'boolean' ? 'center' : 'left')));
      css.push('vertical-align:' + (st.va === 'middle' ? 'middle' : st.va || 'bottom'));
      if (st.wrap === 'wrap') css.push('white-space:pre-wrap');
      // text spills into empty cells on its right, as on the grid
      else if (!st.wrap && !m && typeof v === 'string' && (st.ha || 'left') === 'left' && c < g.c2 && !sh.get(r, c + 1)?.v && sh.get(r, c + 1)?.f == null && !merged.has(key(r, c + 1)) && !skip.has(key(r, c + 1))) css.push('overflow:visible;position:relative;z-index:1');
      if (st.bt) css.push('border-top:' + cssBorder(st.bt)); if (st.bb) css.push('border-bottom:' + cssBorder(st.bb));
      if (st.bl) css.push('border-left:' + cssBorder(st.bl)); if (st.br) css.push('border-right:' + cssBorder(st.br));
      const span = m ? ` rowspan="${m.r2 - m.r1 + 1}" colspan="${m.c2 - m.c1 + 1}"` : '';
      let text = esc(cellDisplay(app, sh, sr, sc)).replace(/\n/g, '<br>');
      if (cell?.link) text = `<a href="${esc(cell.link)}">${text}</a>`;
      if (v instanceof Special && v.kind === 'image') text = `<img src="${esc(v.data.url)}" style="max-width:100%;max-height:100%">`;
      html += `<td${span} style="${css.join(';')}">${text}</td>`;
    }
    html += '</tr>';
  }
  return html + '</table>';
}

const PRINT_CSS = `
*{box-sizing:border-box}body{margin:0;font-family:Arial,Helvetica,sans-serif;font-size:10pt;color:#000}
h1{font:600 13pt Inter,Arial,sans-serif;margin:0 0 8px}h2{font:600 11pt Inter,Arial,sans-serif;margin:18px 0 6px;page-break-before:always}
h2:first-of-type{page-break-before:auto}
table.sheet{border-collapse:collapse;table-layout:fixed}table.sheet td{padding:2px 4px;overflow:hidden;white-space:nowrap;line-height:1.25;font-size:10pt}
table.sheet.grid td{outline:1px solid #d9d9d9;outline-offset:-0.5px}
.meta{color:#666;font-size:8pt;margin-bottom:10px}
@media print{.noprint{display:none}}
.noprint{position:fixed;top:12px;right:12px;font:500 13px Inter,Arial;background:#1a73e8;color:#fff;border:0;border-radius:18px;padding:9px 18px;cursor:pointer}
`;

export function printView(app) {
  const sh = app.sheet;
  const selG = app.grid.range;
  const hasSel = selG.r1 !== selG.r2 || selG.c1 !== selG.c2;
  let scope = hasSel ? 'sel' : 'sheet';
  const L = PG.layout(app, sh), P = L.p;
  const info = el('div.muted.small', `${PG.PAPER[P.size]?.n || 'A4'} ${P.orient === 'landscape' ? 'lanskap' : 'potret'} · margin ${[P.m.t, P.m.b, P.m.l, P.m.r].map(x => String(+x.toFixed(2)).replace('.', ',')).join(' / ')} cm · ${P.fit === 'width' ? 'sesuai lebar' : P.fit === 'page' ? '1 halaman' : P.scale + '%'}${P.auto ? ' · tabel tidak dipotong' : ''}`);
  dialog({
    title: 'Cetak', width: 440,
    body: el('div.stack',
      field('Cetak', select([{ value: 'sheet', label: 'Sheet saat ini' }, { value: 'all', label: 'Seluruh workbook' }, ...(hasSel ? [{ value: 'sel', label: `Sel yang dipilih (${rangeAddr(selG)})` }] : [])], scope, v => (scope = v))),
      info),
    buttons: [{ label: 'Kertas & margin…', run: () => { PG.setupDialog(app); } }, { spacer: true }, { label: 'Batal' }, { label: 'Lanjut', primary: true, run: () => {
      if (scope === 'all') PG.printPages(app, { sheets: app.wb.sheets.filter(s => !s.hidden), auto: true });
      else PG.printPages(app, { sheets: [sh], area: scope === 'sel' ? { r1: selG.r1, c1: selG.c1, r2: selG.r2, c2: selG.c2 } : null, auto: true });
    } }],
  });
}

export function exportHTML(app) {
  const parts = app.wb.sheets.filter(s => !s.hidden).map(s => `<h2>${esc(s.name)}</h2>` + app.withSheet(s, () => sheetHTML(app, s, s.usedRange() || { r1: 0, c1: 0, r2: 0, c2: 0 }, { grid: true })));
  const html = `<!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(app.wb.name)}</title><style>${PRINT_CSS}body{padding:24px}h2{page-break-before:auto}</style></head><body><h1>${esc(app.wb.name)}</h1>${parts.join('')}</body></html>`;
  download(app.wb.name + '.html', html, 'text/html;charset=utf-8');
}
