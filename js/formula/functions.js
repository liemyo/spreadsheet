// Function library part 1: math, logic, lookup & reference, text, date, info.
import {
  FErr, ERR, err, Ref, Lambda, unref, scalar, toArr, num, int, str, bool, opt, check, nums, flat,
  compare, criteria, wildcardRe, dims, mapArr, transpose, vec, parseNumStr,
} from './core.js';
import { serial, parts, colName, quoteSheet, parseRange, jsDateToSerial, colIndex } from '../util.js';
import { format, L, parseDate, general } from '../format.js';

export const F = Object.create(null);
export function def(names, f, o = {}) { for (const n of names.split(' ')) F[n] = { f, ...o }; }
export const S = (names, f, o = {}) => def(names, f, { scalar: true, ...o });
export const E = (n, ctx) => (!n || n.type === 'empty' ? undefined : ctx.eng.ev(n, ctx));

export class Special { constructor(kind, data) { this.kind = kind; this.data = data; this.code = ''; } }

// grid view over a Ref or array: {R, C, get(i,j)}
export function grid(v) {
  if (v instanceof Ref) {
    const eng = v.sheet.wb.eng, sh = v.sheet;
    return { R: v.lastRow - v.r1 + 1, C: v.lastCol - v.c1 + 1, get: (i, j) => eng.value(sh, v.r1 + i, v.c1 + j), ref: v };
  }
  const a = toArr(v);
  return { R: a.length, C: a[0].length, get: (i, j) => (a[i] ? (a[i][j] === undefined ? null : a[i][j]) : null) };
}

const sizedLike = (v, R, C) => (v instanceof Ref ? new Ref(v.sheet, v.r1, v.c1, v.r1 + R - 1, v.c1 + C - 1) : v);

export function rnd(x, d = 0) {
  d = Math.trunc(d);
  const s = x < 0 ? -1 : 1, a = Math.abs(x);
  let r;
  if (d >= 0 && d <= 15) {
    const p = a.toPrecision(15);
    r = p.includes('e') ? Math.round(a * 10 ** d) / 10 ** d : +(Math.round(+(p + 'e' + d)) + 'e-' + d);
  } else if (d > 15) r = a;
  else { const p = 10 ** -d; r = Math.round(a / p) * p; }
  return s * r;
}
function rndDir(x, d, up) {
  d = Math.trunc(d);
  const s = x < 0 ? -1 : 1, a = Math.abs(x);
  const p = 10 ** d;
  const y = +(a * p).toPrecision(15);
  return s * (up ? Math.ceil(y) : Math.floor(y)) / p;
}
const fixf = x => (Number.isInteger(x) ? x : +x.toPrecision(15));

// ---------- MATH ----------
S('ABS', x => Math.abs(num(x)), { fmtArg: 0 });
S('ACOS', x => check(Math.acos(num(x))));
S('ACOSH', x => check(Math.acosh(num(x))));
S('ACOT', x => Math.PI / 2 - Math.atan(num(x)));
S('ACOTH', x => { const v = num(x); if (Math.abs(v) <= 1) throw ERR.NUM; return 0.5 * Math.log((v + 1) / (v - 1)); });
S('ASIN', x => check(Math.asin(num(x))));
S('ASINH', x => Math.asinh(num(x)));
S('ATAN', x => Math.atan(num(x)));
S('ATAN2', (x, y) => { const a = num(x), b = num(y); if (!a && !b) throw ERR.DIV0; return Math.atan2(b, a); });
S('ATANH', x => check(Math.atanh(num(x))));
S('COS', x => Math.cos(num(x)));
S('COSH', x => Math.cosh(num(x)));
S('COT', x => { const t = Math.tan(num(x)); if (!t) throw ERR.DIV0; return 1 / t; });
S('COTH', x => { const t = Math.tanh(num(x)); if (!t) throw ERR.DIV0; return 1 / t; });
S('CSC', x => { const t = Math.sin(num(x)); if (!t) throw ERR.DIV0; return 1 / t; });
S('CSCH', x => { const t = Math.sinh(num(x)); if (!t) throw ERR.DIV0; return 1 / t; });
S('SEC', x => 1 / Math.cos(num(x)));
S('SECH', x => 1 / Math.cosh(num(x)));
S('SIN', x => Math.sin(num(x)));
S('SINH', x => Math.sinh(num(x)));
S('TAN', x => Math.tan(num(x)));
S('TANH', x => Math.tanh(num(x)));
S('DEGREES', x => num(x) * 180 / Math.PI);
S('RADIANS', x => num(x) * Math.PI / 180);
S('EXP', x => check(Math.exp(num(x))));
S('LN', x => { const v = num(x); if (v <= 0) throw ERR.NUM; return Math.log(v); });
S('LOG', (x, b) => { const v = num(x), base = opt(b, 10) === null ? 10 : num(opt(b, 10)); if (v <= 0 || base <= 0 || base === 1) throw ERR.NUM; return fixf(Math.log(v) / Math.log(base)); });
S('LOG10', x => { const v = num(x); if (v <= 0) throw ERR.NUM; return Math.log10(v); });
S('SQRT', x => { const v = num(x); if (v < 0) throw ERR.NUM; return Math.sqrt(v); });
S('SQRTPI', x => { const v = num(x); if (v < 0) throw ERR.NUM; return Math.sqrt(v * Math.PI); });
def('PI', () => Math.PI, { max: 0 });
S('POWER POW', (x, y) => { const a = num(x), b = num(y); if (!a && !b) throw ERR.NUM; return check(Math.pow(a, b)); });
S('SIGN', x => Math.sign(num(x)));
S('INT', x => Math.floor(num(x)), { fmtArg: 0 });
S('TRUNC', (x, d) => rndDir(num(x), num(opt(d, 0)), false), { fmtArg: 0 });
S('MOD', (n, d) => { const a = num(n), b = num(d); if (b === 0) throw ERR.DIV0; return fixf(a - b * Math.floor(a / b)); });
S('ROUND', (x, d) => rnd(num(x), num(opt(d, 0))), { fmtArg: 0 });
S('ROUNDUP', (x, d) => rndDir(num(x), num(opt(d, 0)), true), { fmtArg: 0 });
S('ROUNDDOWN', (x, d) => rndDir(num(x), num(opt(d, 0)), false), { fmtArg: 0 });
S('MROUND', (x, m) => { const a = num(x), b = num(m); if (!b) return 0; if (a * b < 0) throw ERR.NUM; return fixf(rnd(a / b) * b); }, { fmtArg: 0 });
S('CEILING', (x, s) => { const a = num(x), b = num(opt(s, 1)); if (!b) return 0; if (a > 0 && b < 0) throw ERR.NUM; return fixf(Math.ceil(fixf(a / b)) * b); }, { fmtArg: 0 });
S('CEILING.MATH', (x, s, m) => { const a = num(x), b = Math.abs(num(opt(s, 1))) || 1; if (a < 0 && num(opt(m, 0))) return -fixf(Math.ceil(fixf(-a / b)) * b); return fixf(Math.ceil(fixf(a / b)) * b); }, { fmtArg: 0 });
S('CEILING.PRECISE ISO.CEILING', (x, s) => { const a = num(x), b = Math.abs(num(opt(s, 1))); if (!b) return 0; return fixf(Math.ceil(fixf(a / b)) * b); }, { fmtArg: 0 });
S('FLOOR', (x, s) => { const a = num(x), b = num(opt(s, 1)); if (!b) throw ERR.DIV0; if (a > 0 && b < 0) throw ERR.NUM; return fixf(Math.floor(fixf(a / b)) * b); }, { fmtArg: 0 });
S('FLOOR.MATH', (x, s, m) => { const a = num(x), b = Math.abs(num(opt(s, 1))) || 1; if (a < 0 && num(opt(m, 0))) return -fixf(Math.floor(fixf(-a / b)) * b); return fixf(Math.floor(fixf(a / b)) * b); }, { fmtArg: 0 });
S('FLOOR.PRECISE', (x, s) => { const a = num(x), b = Math.abs(num(opt(s, 1))); if (!b) return 0; return fixf(Math.floor(fixf(a / b)) * b); }, { fmtArg: 0 });
S('EVEN', x => { const a = num(x), s = a < 0 ? -1 : 1; let v = Math.ceil(Math.abs(a)); if (v % 2) v++; return s * v; });
S('ODD', x => { const a = num(x), s = a < 0 ? -1 : 1; let v = Math.ceil(Math.abs(a)); if (v % 2 === 0) v++; return s * v; });
export function fact(n) { n = Math.floor(n); if (n < 0) throw ERR.NUM; if (n > 170) throw ERR.NUM; let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; }
S('FACT', x => fact(num(x)));
S('FACTDOUBLE', x => { let n = Math.floor(num(x)); if (n < -1) throw ERR.NUM; let r = 1; for (; n > 1; n -= 2) r *= n; return r; });
export function combin(n, k) { n = Math.floor(n); k = Math.floor(k); if (n < 0 || k < 0 || k > n) throw ERR.NUM; let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); }
S('COMBIN', (n, k) => combin(num(n), num(k)));
S('COMBINA', (n, k) => { const a = Math.floor(num(n)), b = Math.floor(num(k)); return combin(a + b - 1, b); });
S('PERMUT', (n, k) => { const a = Math.floor(num(n)), b = Math.floor(num(k)); if (a < 0 || b < 0 || b > a) throw ERR.NUM; let r = 1; for (let i = 0; i < b; i++) r *= a - i; return r; });
S('PERMUTATIONA', (n, k) => Math.pow(Math.floor(num(n)), Math.floor(num(k))));
const gcd2 = (a, b) => { while (b) [a, b] = [b, a % b]; return a; };
def('GCD', args => { const v = nums(args).map(Math.floor); if (v.some(x => x < 0)) throw ERR.NUM; return v.reduce(gcd2, 0); }, { min: 1 });
def('LCM', args => { const v = nums(args).map(Math.floor); if (v.some(x => x < 0)) throw ERR.NUM; if (v.some(x => x === 0)) return 0; return v.reduce((a, b) => a * b / gcd2(a, b), 1); }, { min: 1 });
def('MULTINOMIAL', args => { const v = nums(args).map(Math.floor); if (v.some(x => x < 0)) throw ERR.NUM; let s = 0, d = 1; for (const x of v) { s += x; d *= fact(x); } return Math.round(fact(s) / d); }, { min: 1 });
S('QUOTIENT', (n, d) => { const b = num(d); if (!b) throw ERR.DIV0; return Math.trunc(num(n) / b); });
def('RAND', () => Math.random(), { max: 0, volatile: true });
S('RANDBETWEEN', (a, b) => { const lo = Math.ceil(num(a)), hi = Math.floor(num(b)); if (lo > hi) throw ERR.NUM; return lo + Math.floor(Math.random() * (hi - lo + 1)); }, { volatile: true });
def('RANDARRAY', ([r, c, mn, mx, w]) => {
  const R = int(opt(r, 1)), C = int(opt(c, 1)), lo = num(opt(mn, 0)), hi = num(opt(mx, 1)), whole = bool(opt(w, false));
  if (R < 1 || C < 1 || lo > hi) throw ERR.VALUE;
  const out = [];
  for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) row.push(whole ? lo + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo)); out.push(row); }
  return out;
}, { max: 5, volatile: true });
S('ROMAN', (n, f) => {
  let v = Math.floor(num(n)); if (v < 0 || v > 3999) throw ERR.VALUE; if (!v) return '';
  const T = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = ''; for (const [k, r] of T) while (v >= k) { s += r; v -= k; } return s;
});
S('ARABIC', s => {
  const t = str(s).trim().toUpperCase(); const neg = t.startsWith('-'); const u = neg ? t.slice(1) : t;
  const V = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 }; let r = 0;
  for (let i = 0; i < u.length; i++) { const a = V[u[i]], b = V[u[i + 1]]; if (!a) throw ERR.VALUE; r += b > a ? -a : a; }
  return neg ? -r : r;
});
S('BASE', (n, r, m) => { const v = Math.floor(num(n)), rad = Math.floor(num(r)); if (v < 0 || rad < 2 || rad > 36) throw ERR.NUM; return v.toString(rad).toUpperCase().padStart(int(opt(m, 0)), '0'); });
S('DECIMAL', (s, r) => { const rad = Math.floor(num(r)); const v = parseInt(str(s), rad); if (isNaN(v) || rad < 2 || rad > 36) throw ERR.NUM; return v; });
def('SUM', args => { let s = 0; for (const x of nums(args)) s += x; return fixf(s); }, { fmtArg: 0 });
def('SUMSQ', args => nums(args).reduce((s, x) => s + x * x, 0));
def('PRODUCT', args => { const v = nums(args); return v.length ? v.reduce((p, x) => p * x, 1) : 0; });
def('SUMPRODUCT', args => {
  const arrs = args.filter(a => a !== undefined).map(toArr);
  if (!arrs.length) throw ERR.VALUE;
  const R = arrs[0].length, C = arrs[0][0].length;
  if (arrs.some(a => a.length !== R || a[0].length !== C)) throw ERR.VALUE;
  let s = 0;
  for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
    let p = 1;
    for (const a of arrs) { const v = a[i][j]; if (v instanceof FErr) throw v; p *= typeof v === 'number' ? v : typeof v === 'boolean' && arrs.length > 1 ? +v : 0; }
    s += p;
  }
  return fixf(s);
}, { min: 1 });
function pairs2(a, b) {
  const x = flat([a], true), y = flat([b], true);
  if (x.length !== y.length) throw ERR.NA;
  const out = [];
  for (let i = 0; i < x.length; i++) { if (x[i] instanceof FErr) throw x[i]; if (y[i] instanceof FErr) throw y[i]; if (typeof x[i] === 'number' && typeof y[i] === 'number') out.push([x[i], y[i]]); }
  return out;
}
export { pairs2 };
def('SUMX2MY2', ([a, b]) => pairs2(a, b).reduce((s, [x, y]) => s + x * x - y * y, 0), { min: 2, max: 2 });
def('SUMX2PY2', ([a, b]) => pairs2(a, b).reduce((s, [x, y]) => s + x * x + y * y, 0), { min: 2, max: 2 });
def('SUMXMY2', ([a, b]) => pairs2(a, b).reduce((s, [x, y]) => s + (x - y) ** 2, 0), { min: 2, max: 2 });
def('SERIESSUM', ([x, n, m, c]) => { const X = num(x), N = num(n), M = num(m); return nums([c]).reduce((s, a, i) => s + a * Math.pow(X, N + i * M), 0); }, { min: 4, max: 4 });

function ifAgg(range, crit, target, agg) {
  const g = grid(range);
  const p = criteria(crit);
  const t = target === undefined ? g : grid(sizedLike(target, g.R, g.C));
  const vals = [];
  for (let i = 0; i < g.R; i++) for (let j = 0; j < g.C; j++) {
    if (!p(g.get(i, j))) continue;
    const v = t.get(i, j);
    if (v instanceof FErr) throw v;
    if (typeof v === 'number') vals.push(v);
  }
  return agg(vals);
}
function ifsAgg(target, list, agg) {
  if (list.length % 2) throw ERR.NA;
  const t = target ? grid(target) : null;
  const conds = [];
  for (let k = 0; k < list.length; k += 2) conds.push([grid(list[k]), criteria(list[k + 1])]);
  const R = t ? t.R : conds[0][0].R, C = t ? t.C : conds[0][0].C;
  for (const [g] of conds) if (g.R !== R || g.C !== C) throw err('#VALUE!', 'Ukuran rentang kriteria tidak sama');
  const vals = [];
  for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
    let ok = true;
    for (const [g, p] of conds) if (!p(g.get(i, j))) { ok = false; break; }
    if (!ok) continue;
    if (!t) { vals.push(1); continue; }
    const v = t.get(i, j);
    if (v instanceof FErr) throw v;
    if (typeof v === 'number') vals.push(v);
  }
  return agg(vals);
}
const sumA = v => fixf(v.reduce((s, x) => s + x, 0));
const avgA = v => { if (!v.length) throw ERR.DIV0; return sumA(v) / v.length; };
def('SUMIF', (r, c, s) => ifAgg(r, c, s, sumA), { scalar: [1], min: 2, max: 3, fmtArg: 2 });
def('SUMIFS', ([t, ...l]) => ifsAgg(t, l, sumA), { min: 3, fmtArg: 0 });
def('AVERAGEIF', (r, c, s) => ifAgg(r, c, s, avgA), { scalar: [1], min: 2, max: 3, fmtArg: 2 });
def('AVERAGEIFS', ([t, ...l]) => ifsAgg(t, l, avgA), { min: 3, fmtArg: 0 });
def('COUNTIF', (r, c) => { const g = grid(r), p = criteria(c); let n = 0; for (let i = 0; i < g.R; i++) for (let j = 0; j < g.C; j++) if (p(g.get(i, j))) n++; return n; }, { scalar: [1], min: 2, max: 2 });
def('COUNTIFS', args => ifsAgg(null, args, v => v.length), { min: 2 });
def('MAXIFS', ([t, ...l]) => ifsAgg(t, l, v => (v.length ? Math.max(...v) : 0)), { min: 3, fmtArg: 0 });
def('MINIFS', ([t, ...l]) => ifsAgg(t, l, v => (v.length ? Math.min(...v) : 0)), { min: 3, fmtArg: 0 });

// matrices
function numMatrix(v) { return toArr(v).map(r => r.map(x => { if (x instanceof FErr) throw x; if (typeof x !== 'number') throw ERR.VALUE; return x; })); }
def('MMULT', ([a, b]) => {
  const A = numMatrix(a), B = numMatrix(b);
  if (A[0].length !== B.length) throw ERR.VALUE;
  return A.map(row => B[0].map((_, j) => row.reduce((s, x, k) => s + x * B[k][j], 0)));
}, { min: 2, max: 2 });
function det(M) {
  const n = M.length, A = M.map(r => r.slice());
  let d = 1;
  for (let i = 0; i < n; i++) {
    let p = i; for (let k = i + 1; k < n; k++) if (Math.abs(A[k][i]) > Math.abs(A[p][i])) p = k;
    if (!A[p][i]) return 0;
    if (p !== i) { [A[p], A[i]] = [A[i], A[p]]; d = -d; }
    d *= A[i][i];
    for (let k = i + 1; k < n; k++) { const f = A[k][i] / A[i][i]; for (let j = i; j < n; j++) A[k][j] -= f * A[i][j]; }
  }
  return fixf(d);
}
def('MDETERM', ([a]) => { const M = numMatrix(a); if (M.length !== M[0].length) throw ERR.VALUE; return det(M); }, { min: 1, max: 1 });
def('MINVERSE', ([a]) => {
  const M = numMatrix(a), n = M.length; if (n !== M[0].length) throw ERR.VALUE;
  const A = M.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let i = 0; i < n; i++) {
    let p = i; for (let k = i + 1; k < n; k++) if (Math.abs(A[k][i]) > Math.abs(A[p][i])) p = k;
    if (Math.abs(A[p][i]) < 1e-15) throw ERR.NUM;
    [A[p], A[i]] = [A[i], A[p]];
    const pv = A[i][i]; for (let j = 0; j < 2 * n; j++) A[i][j] /= pv;
    for (let k = 0; k < n; k++) if (k !== i) { const f = A[k][i]; for (let j = 0; j < 2 * n; j++) A[k][j] -= f * A[i][j]; }
  }
  return A.map(r => r.slice(n).map(fixf));
}, { min: 1, max: 1 });
def('MUNIT', ([n]) => { const k = int(n); if (k < 1) throw ERR.VALUE; return Array.from({ length: k }, (_, i) => Array.from({ length: k }, (_, j) => (i === j ? 1 : 0))); }, { min: 1, max: 1 });
def('TRANSPOSE', ([a]) => transpose(toArr(a)), { min: 1, max: 1 });

// SUBTOTAL / AGGREGATE
function visibleNums(args, skipHidden) {
  const out = [];
  for (const a of args) {
    if (a instanceof Ref) {
      const sh = a.sheet;
      a.forEach((v, r, c) => {
        if ((skipHidden && sh.hiddenR.has(r)) || sh.filteredR.has(r)) return;
        const cell = sh.cells.get(r * 16384 + c);
        if (cell && cell.ast && cell.ast.type === 'call' && (cell.ast.name === 'SUBTOTAL' || cell.ast.name === 'AGGREGATE')) return;
        if (v instanceof FErr) throw v;
        out.push(v);
      });
    } else out.push(...flat([a]));
  }
  return out;
}
export const AGG = {
  1: v => { const n = v.filter(x => typeof x === 'number'); if (!n.length) throw ERR.DIV0; return sumA(n) / n.length; },
  2: v => v.filter(x => typeof x === 'number').length,
  3: v => v.filter(x => x != null && x !== '').length,
  4: v => { const n = v.filter(x => typeof x === 'number'); return n.length ? Math.max(...n) : 0; },
  5: v => { const n = v.filter(x => typeof x === 'number'); return n.length ? Math.min(...n) : 0; },
  6: v => v.filter(x => typeof x === 'number').reduce((p, x) => p * x, 1),
  7: v => F['STDEV.S'].f([v.filter(x => typeof x === 'number').map(x => [x])]),
  8: v => F['STDEV.P'].f([v.filter(x => typeof x === 'number').map(x => [x])]),
  9: v => sumA(v.filter(x => typeof x === 'number')),
  10: v => F['VAR.S'].f([v.filter(x => typeof x === 'number').map(x => [x])]),
  11: v => F['VAR.P'].f([v.filter(x => typeof x === 'number').map(x => [x])]),
};
def('SUBTOTAL', ([fn, ...refs]) => {
  const k = int(fn), base = k > 100 ? k - 100 : k;
  if (!AGG[base]) throw ERR.VALUE;
  return AGG[base](visibleNums(refs, k > 100));
}, { min: 2, fmtArg: 1 });
def('AGGREGATE', ([fn, op, ...refs]) => {
  const k = int(fn), o = int(opt(op, 0));
  const skipHidden = [1, 3, 5, 7].includes(o), skipErr = [2, 3, 6, 7].includes(o);
  let vals;
  try { vals = visibleNums(refs.slice(0, k >= 14 ? 1 : refs.length), skipHidden); }
  catch (e) { if (!(e instanceof FErr) || !skipErr) throw e; vals = flat(refs.slice(0, 1)).filter(x => !(x instanceof FErr)); }
  if (skipErr) vals = vals.filter(x => !(x instanceof FErr));
  if (k <= 11) return AGG[k](vals);
  const n = vals.filter(x => typeof x === 'number').map(x => [x]);
  const K = refs[1];
  switch (k) {
    case 12: return F.MEDIAN.f([n]); case 13: return F['MODE.SNGL'].f([n]);
    case 14: return F.LARGE.f(n, K); case 15: return F.SMALL.f(n, K);
    case 16: return F['PERCENTILE.INC'].f([n, K]); case 17: return F['QUARTILE.INC'].f([n, K]);
    case 18: return F['PERCENTILE.EXC'].f([n, K]); case 19: return F['QUARTILE.EXC'].f([n, K]);
  }
  throw ERR.VALUE;
}, { min: 3 });

// ---------- LOGICAL ----------
function logicals(args) {
  const out = [];
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof Ref || Array.isArray(a)) {
      for (const v of flat([a])) { if (v instanceof FErr) throw v; if (typeof v === 'boolean') out.push(v); else if (typeof v === 'number') out.push(v !== 0); }
    } else out.push(bool(a));
  }
  if (!out.length) throw ERR.VALUE;
  return out;
}
def('AND', args => logicals(args).every(Boolean), { min: 1 });
def('OR', args => logicals(args).some(Boolean), { min: 1 });
def('XOR', args => logicals(args).filter(Boolean).length % 2 === 1, { min: 1 });
S('NOT', x => !bool(x), { min: 1, max: 1 });
def('TRUE', () => true, { max: 0 });
def('FALSE', () => false, { max: 0 });
def('IF', (a, ctx) => {
  const c = unref(E(a[0], ctx));
  const yes = () => (a.length > 1 ? (a[1].type === 'empty' ? 0 : E(a[1], ctx)) : true);
  const no = () => (a.length > 2 ? (a[2].type === 'empty' ? 0 : E(a[2], ctx)) : false);
  if (Array.isArray(c)) {
    const t = unref(yes()), f = unref(no());
    const R = c.length, C = c[0].length;
    return c.map((row, i) => row.map((v, j) => {
      if (v instanceof FErr) return v;
      let b; try { b = bool(v); } catch (e) { return e; }
      const src = b ? t : f;
      return Array.isArray(src) ? (src[src.length === 1 ? 0 : i]?.[src[0].length === 1 ? 0 : j] ?? ERR.NA) : src;
    })).slice(0, R).map(r => r.slice(0, C));
  }
  if (c instanceof FErr) return c;
  return bool(c) ? yes() : no();
}, { lazy: true, min: 1, max: 3, fmtArg: 1 });
def('IFS', (a, ctx) => {
  if (a.length % 2) throw ERR.NA;
  for (let i = 0; i < a.length; i += 2) {
    const c = scalar(E(a[i], ctx));
    if (c instanceof FErr) return c;
    if (bool(c)) return E(a[i + 1], ctx);
  }
  return err('#N/A', 'Tidak ada kondisi IFS yang terpenuhi');
}, { lazy: true, min: 2, fmtArg: 1 });
def('IFERROR', (a, ctx) => {
  const v = E(a[0], ctx);
  const alt = () => (a[1] && a[1].type !== 'empty' ? E(a[1], ctx) : '');
  const u = unref(v);
  if (Array.isArray(u)) { let fb; return u.map(row => row.map(x => (x instanceof FErr ? scalar(fb === undefined ? (fb = alt()) : fb) : x))); }
  return u instanceof FErr ? alt() : v;
}, { lazy: true, min: 1, max: 2, fmtArg: 0 });
def('IFNA', (a, ctx) => {
  const v = E(a[0], ctx);
  const u = unref(v);
  const isNA = x => x instanceof FErr && x.code === '#N/A';
  if (Array.isArray(u)) { const fb = scalar(E(a[1], ctx)); return u.map(row => row.map(x => (isNA(x) ? fb : x))); }
  return isNA(u) ? E(a[1], ctx) : v;
}, { lazy: true, min: 2, max: 2, fmtArg: 0 });
def('SWITCH', (a, ctx) => {
  const v = scalar(E(a[0], ctx));
  if (v instanceof FErr) return v;
  let i = 1;
  for (; i + 1 < a.length; i += 2) {
    const k = scalar(E(a[i], ctx));
    if (k instanceof FErr) return k;
    if (compare(v, k) === 0 && typeof v === typeof k) return E(a[i + 1], ctx);
  }
  return i < a.length ? E(a[i], ctx) : err('#N/A', 'Tidak ada nilai SWITCH yang cocok');
}, { lazy: true, min: 3 });
def('LET', (a, ctx) => {
  if (a.length < 3 || a.length % 2 === 0) throw ERR.VALUE;
  const vars = new Map();
  const c2 = { ...ctx, env: { vars, parent: ctx.env } };
  for (let i = 0; i + 1 < a.length; i += 2) {
    if (a[i].type !== 'name') throw err('#VALUE!', 'Nama LET tidak valid');
    vars.set(a[i].name.toUpperCase(), E(a[i + 1], c2));
  }
  return E(a[a.length - 1], c2);
}, { lazy: true, min: 3 });
def('LAMBDA', (a, ctx) => {
  const params = a.slice(0, -1).map(n => { if (n.type !== 'name') throw err('#VALUE!', 'Parameter LAMBDA tidak valid'); return n.name.toUpperCase(); });
  return new Lambda(params, a[a.length - 1], ctx.env, ctx);
}, { lazy: true, min: 1 });
const asLam = v => { if (!(v instanceof Lambda)) throw err('#VALUE!', 'Argumen terakhir harus LAMBDA'); return v; };
def('MAP', (args, ctx) => {
  const lam = asLam(args[args.length - 1]);
  const arrs = args.slice(0, -1).map(x => toArr(unref(x)));
  const R = arrs[0].length, C = arrs[0][0].length;
  return Array.from({ length: R }, (_, i) => Array.from({ length: C }, (_, j) => scalar(ctx.eng.apply(lam, arrs.map(a => a[i]?.[j] ?? ERR.NA)))));
}, { min: 2 });
def('REDUCE', ([init, arr, l], ctx) => {
  const lam = asLam(l);
  let acc = init === undefined ? 0 : unref(init);
  for (const v of flat([arr], true)) acc = unref(ctx.eng.apply(lam, [acc, v]));
  return acc;
}, { min: 3, max: 3 });
def('SCAN', ([init, arr, l], ctx) => {
  const lam = asLam(l);
  let acc = init === undefined ? 0 : unref(init);
  return toArr(arr).map(row => row.map(v => (acc = scalar(ctx.eng.apply(lam, [acc, v])))));
}, { min: 3, max: 3 });
def('BYROW', ([arr, l], ctx) => { const lam = asLam(l); return toArr(arr).map(row => [scalar(ctx.eng.apply(lam, [[row]]))]); }, { min: 2, max: 2 });
def('BYCOL', ([arr, l], ctx) => { const lam = asLam(l); return [transpose(toArr(arr)).map(col => scalar(ctx.eng.apply(lam, [col.map(x => [x])])))]; }, { min: 2, max: 2 });
def('MAKEARRAY', ([r, c, l], ctx) => {
  const lam = asLam(l), R = int(r), C = int(c);
  if (R < 1 || C < 1) throw ERR.VALUE;
  return Array.from({ length: R }, (_, i) => Array.from({ length: C }, (_, j) => scalar(ctx.eng.apply(lam, [i + 1, j + 1]))));
}, { min: 3, max: 3 });
def('ISOMITTED', (a, ctx) => E(a[0], ctx) === undefined, { lazy: true, min: 1, max: 1 });
// Lazy like Sheets and Excel: an error in a branch that was not picked must not affect the result.
def('CHOOSE', (a, ctx) => {
  const ix = unref(E(a[0], ctx));
  if (ix instanceof FErr) return ix;
  const pick = n => { const v = E(a[n], ctx); return v === undefined ? 0 : v; };
  if (Array.isArray(ix)) {
    const seen = new Map();
    return ix.map(row => row.map(k => {
      if (k instanceof FErr) return k;
      let n; try { n = Math.floor(num(k)); } catch (e) { return e; }
      if (n < 1 || n >= a.length) return ERR.VALUE;
      if (!seen.has(n)) seen.set(n, scalar(pick(n)));
      return seen.get(n);
    }));
  }
  const n = Math.floor(num(ix));
  if (n < 1 || n >= a.length) throw ERR.VALUE;
  return pick(n);
}, { lazy: true, min: 2 });

// ---------- LOOKUP ----------
function eqLookup(key) {
  if (typeof key === 'string' && /[*?~]/.test(key)) { const re = wildcardRe(key); return v => typeof v === 'string' && re.test(v); }
  return v => v != null && typeof v === typeof key && compare(v, key) === 0;
}
// binary search: index of last element <= key (ascending), same type only
function bsearch(list, key) {
  let lo = 0, hi = list.length - 1, ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let v = list[mid];
    let m = mid;
    // skip blanks / other types like Excel
    while (m >= lo && (v == null || typeof v !== typeof key)) { m--; v = list[m]; }
    if (m < lo) { lo = mid + 1; continue; }
    if (compare(v, key) <= 0) { ans = m; lo = mid + 1; } else hi = m - 1;
  }
  return ans;
}
def('VLOOKUP', (key, range, col, approx) => {
  if (key instanceof FErr) return key;
  const g = grid(range), c = int(col);
  if (c < 1) throw ERR.VALUE;
  if (c > g.C) throw err('#REF!', 'Indeks kolom VLOOKUP di luar rentang');
  const keyCol = Array.from({ length: g.R }, (_, i) => g.get(i, 0));
  let i;
  if (approx === undefined || bool(approx)) i = bsearch(keyCol, key);
  else { const p = eqLookup(key); i = keyCol.findIndex(p); }
  if (i < 0) throw err('#N/A', `Nilai '${str(key)}' tidak ditemukan dalam evaluasi VLOOKUP`);
  return g.get(i, c - 1);
}, { scalar: [0], min: 3, max: 4, errOk: true, fmtArg: 1 });
def('HLOOKUP', (key, range, row, approx) => {
  if (key instanceof FErr) return key;
  const g = grid(range), r = int(row);
  if (r < 1) throw ERR.VALUE;
  if (r > g.R) throw ERR.REF;
  const keyRow = Array.from({ length: g.C }, (_, j) => g.get(0, j));
  let j;
  if (approx === undefined || bool(approx)) j = bsearch(keyRow, key);
  else j = keyRow.findIndex(eqLookup(key));
  if (j < 0) throw err('#N/A', `Nilai '${str(key)}' tidak ditemukan dalam evaluasi HLOOKUP`);
  return g.get(r - 1, j);
}, { scalar: [0], min: 3, max: 4, errOk: true, fmtArg: 1 });
def('LOOKUP', (key, look, res) => {
  if (key instanceof FErr) return key;
  const g = grid(look);
  let keys, vals;
  if (res === undefined) {
    if (g.C > g.R) { keys = Array.from({ length: g.C }, (_, j) => g.get(0, j)); vals = Array.from({ length: g.C }, (_, j) => g.get(g.R - 1, j)); }
    else { keys = Array.from({ length: g.R }, (_, i) => g.get(i, 0)); vals = Array.from({ length: g.R }, (_, i) => g.get(i, g.C - 1)); }
  } else { keys = vec(look); vals = vec(res); }
  const i = bsearch(keys, key);
  if (i < 0) throw ERR.NA;
  return vals[i] ?? ERR.NA;
}, { scalar: [0], min: 2, max: 3, errOk: true });
function matchIndex(key, list, type) {
  if (type === 0) return list.findIndex(eqLookup(key));
  if (type === 1) return bsearch(list, key);
  let ans = -1;
  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (v == null || typeof v !== typeof key) continue;
    if (compare(v, key) >= 0) ans = i; else break;
  }
  return ans;
}
def('MATCH', (key, arr, t) => {
  if (key instanceof FErr) return key;
  const type = t === undefined ? 1 : Math.sign(num(t));
  const i = matchIndex(key, vec(arr), type);
  if (i < 0) throw err('#N/A', `Nilai '${str(key)}' tidak ditemukan dalam evaluasi MATCH`);
  return i + 1;
}, { scalar: [0], min: 2, max: 3, errOk: true });
function xfind(key, list, mode, search) {
  const order = [];
  if (search === -1 || search === -2) for (let i = list.length - 1; i >= 0; i--) order.push(i);
  else for (let i = 0; i < list.length; i++) order.push(i);
  if (mode === 2) { const re = typeof key === 'string' ? wildcardRe(key) : null; return order.find(i => (re ? typeof list[i] === 'string' && re.test(list[i]) : compare(list[i], key) === 0 && list[i] != null)) ?? -1; }
  let best = -1;
  for (const i of order) {
    const v = list[i];
    if (v == null) continue;
    if (typeof v === typeof key && compare(v, key) === 0) return i;
    if (mode === -1 && typeof v === typeof key && compare(v, key) < 0 && (best < 0 || compare(v, list[best]) > 0)) best = i;
    if (mode === 1 && typeof v === typeof key && compare(v, key) > 0 && (best < 0 || compare(v, list[best]) < 0)) best = i;
  }
  return mode === 0 ? -1 : best;
}
def('XMATCH', (key, arr, m, s) => {
  if (key instanceof FErr) return key;
  const i = xfind(key, vec(arr), int(opt(m, 0)), int(opt(s, 1)));
  if (i < 0) throw ERR.NA;
  return i + 1;
}, { scalar: [0], min: 2, max: 4, errOk: true });
def('XLOOKUP', (key, look, ret, nf, m, s) => {
  if (key instanceof FErr) return key;
  const [lr, lc] = dims(look);
  const byCol = lr >= lc || lc === 1;
  const list = vec(look);
  const i = xfind(key, list, int(opt(m, 0)), int(opt(s, 1)));
  if (i < 0) { if (nf !== undefined) return nf; throw err('#N/A', `Nilai '${str(key)}' tidak ditemukan dalam evaluasi XLOOKUP`); }
  const [rr, rc] = dims(ret);
  if (ret instanceof Ref) return byCol ? ret.sub(i, 0, i, rc - 1) : ret.sub(0, i, rr - 1, i);
  const a = toArr(ret);
  return byCol ? [a[i] || [ERR.NA]] : a.map(row => [row[i]]);
}, { scalar: [0], min: 3, max: 6, errOk: true, fmtArg: 2 });
def('INDEX', ([ref, row, col]) => {
  const [R, C] = dims(ref);
  let r = row === undefined ? 0 : int(row), c = col === undefined ? 0 : int(col);
  if (col === undefined && R === 1 && C > 1) { c = r; r = 1; }
  else if (col === undefined && C === 1) c = 1;
  if (r < 0 || c < 0 || r > R || c > C) throw err('#REF!', 'Indeks INDEX di luar rentang');
  if (ref instanceof Ref) {
    if (r && c) return ref.sub(r - 1, c - 1, r - 1, c - 1);
    if (r) return ref.sub(r - 1, 0, r - 1, C - 1);
    if (c) return ref.sub(0, c - 1, R - 1, c - 1);
    return ref;
  }
  const a = toArr(ref);
  if (r && c) return a[r - 1][c - 1];
  if (r) return [a[r - 1]];
  if (c) return a.map(x => [x[c - 1]]);
  return a;
}, { min: 1, max: 4, fmtArg: 0 });
def('OFFSET', ([ref, rows, cols, h, w]) => {
  if (!(ref instanceof Ref)) throw ERR.VALUE;
  const r1 = ref.r1 + int(rows), c1 = ref.c1 + int(cols);
  const H = h === undefined ? ref.rows : int(h), W = w === undefined ? ref.cols : int(w);
  if (r1 < 0 || c1 < 0 || H < 1 || W < 1) throw ERR.REF;
  return new Ref(ref.sheet, r1, c1, r1 + H - 1, c1 + W - 1);
}, { min: 3, max: 5, volatile: true });
def('INDIRECT', ([t, a1], ctx) => {
  let text = str(t).trim();
  if (a1 !== undefined && !bool(a1)) {
    const m = /^(?:(.+)!)?R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(text);
    if (!m) throw ERR.REF;
    const pre = m[1] ? m[1] + '!' : '';
    text = pre + colName(+m[3] - 1) + m[2] + (m[4] ? ':' + colName(+m[5] - 1) + m[4] : '');
  }
  const nm = ctx.eng.wb.names.get(text.toUpperCase());
  if (nm) return ctx.eng.evalName({ name: text }, ctx);
  const g = parseRange(text);
  if (!g) throw err('#REF!', `Referensi "${text}" tidak valid`);
  const sh = g.sheet ? ctx.eng.wb.sheetByName(g.sheet) : ctx.sh;
  if (!sh) throw ERR.REF;
  const open = g.r2 === Infinity || g.c2 === Infinity;
  return new Ref(sh, g.r1, g.c1, g.r2 === Infinity ? sh.rows - 1 : g.r2, g.c2 === Infinity ? sh.cols - 1 : g.c2, open);
}, { min: 1, max: 2, volatile: true });
def('ROW', ([ref], ctx) => {
  if (ref === undefined) return ctx.r + 1 + (ctx.dr || 0);
  if (!(ref instanceof Ref)) throw ERR.VALUE;
  return ref.rows === 1 ? ref.r1 + 1 : Array.from({ length: ref.lastRow - ref.r1 + 1 }, (_, i) => [ref.r1 + i + 1]);
}, { max: 1 });
def('COLUMN', ([ref], ctx) => {
  if (ref === undefined) return ctx.c + 1 + (ctx.dc || 0);
  if (!(ref instanceof Ref)) throw ERR.VALUE;
  return ref.cols === 1 ? ref.c1 + 1 : [Array.from({ length: ref.cols }, (_, i) => ref.c1 + i + 1)];
}, { max: 1 });
def('ROWS', ([a]) => dims(a)[0], { min: 1, max: 1 });
def('COLUMNS', ([a]) => dims(a)[1], { min: 1, max: 1 });
def('AREAS', ([a]) => (a instanceof Ref ? 1 : (() => { throw ERR.VALUE; })()), { min: 1, max: 1 });
S('ADDRESS', (r, c, a, s, sh) => {
  const R = int(r), C = int(c), abs = int(opt(a, 1)), style = opt(s, true);
  if (R < 1 || C < 1) throw ERR.VALUE;
  const pre = sh !== undefined && sh !== null ? quoteSheet(str(sh)) + '!' : '';
  if (style === false || style === 0) {
    const rr = abs === 1 || abs === 2 ? 'R' + R : 'R[' + R + ']', cc = abs === 1 || abs === 3 ? 'C' + C : 'C[' + C + ']';
    return pre + rr + cc;
  }
  const ar = abs === 1 || abs === 2, ac = abs === 1 || abs === 3;
  return pre + (ac ? '$' : '') + colName(C - 1) + (ar ? '$' : '') + R;
}, { min: 2, max: 5 });
def('FORMULATEXT', ([ref]) => {
  if (!(ref instanceof Ref)) throw ERR.NA;
  const cell = ref.sheet.cells.get(ref.r1 * 16384 + ref.c1);
  if (!cell || cell.f == null) throw ERR.NA;
  return '=' + cell.f;
}, { min: 1, max: 1 });
def('ISFORMULA', ([ref]) => { if (!(ref instanceof Ref)) throw ERR.VALUE; const cell = ref.sheet.cells.get(ref.r1 * 16384 + ref.c1); return !!(cell && cell.f != null); }, { min: 1, max: 1 });
def('ISREF', ([v]) => v instanceof Ref, { min: 1, max: 1, errOk: true });
def('SHEET', ([v], ctx) => {
  const wb = ctx.eng.wb;
  if (v === undefined) return wb.sheets.indexOf(ctx.sh) + 1;
  if (v instanceof Ref) return wb.sheets.indexOf(v.sheet) + 1;
  const sh = wb.sheetByName(str(v)); if (!sh) throw ERR.NA; return wb.sheets.indexOf(sh) + 1;
}, { max: 1 });
def('SHEETS', (_, ctx) => ctx.eng.wb.sheets.length, { max: 1 });
def('CELL', ([info, ref], ctx) => {
  const t = str(info).toLowerCase();
  const r = ref instanceof Ref ? ref : null;
  const sh = r ? r.sheet : ctx.sh, row = r ? r.r1 : ctx.r, col = r ? r.c1 : ctx.c;
  const v = ctx.eng.value(sh, row, col);
  switch (t) {
    case 'address': return '$' + colName(col) + '$' + (row + 1);
    case 'col': return col + 1;
    case 'row': return row + 1;
    case 'contents': return v;
    case 'type': return v == null ? 'b' : typeof v === 'string' ? 'l' : 'v';
    case 'width': return Math.round(sh.colWidth(col) / 7);
    case 'filename': return ctx.eng.wb.name;
    case 'format': { const s = ctx.eng.wb.styles[sh.cells.get(row * 16384 + col)?.s]; return s?.fmt || 'G'; }
    default: throw ERR.VALUE;
  }
}, { min: 1, max: 2 });
S('HYPERLINK', (u, l) => (l === undefined || l === null ? str(u) : l), { min: 1, max: 2 });

// dynamic arrays
function boolArr(v) { return toArr(v).map(r => r.map(x => { if (x instanceof FErr) throw x; if (x == null || x === '') return false; if (typeof x === 'string') { const n = parseNumStr(x); if (n === null) throw ERR.VALUE; return n !== 0; } return bool(x); })); }
def('FILTER', ([arr, inc, ...more]) => {
  const a = toArr(arr);
  const R = a.length, C = a[0].length;
  const conds = [inc];
  let ifEmpty;
  for (const m of more) {
    const [mr, mc] = dims(m);
    if ((m instanceof Ref || Array.isArray(m)) && ((mr === R && mc === 1) || (mc === C && mr === 1)) && !(mr === 1 && mc === 1)) conds.push(m);
    else ifEmpty = m;
  }
  let keepR = null, keepC = null;
  for (const cnd of conds) {
    const b = boolArr(cnd);
    if (b.length === R && b[0].length === 1) keepR = (keepR || Array(R).fill(true)).map((k, i) => k && b[i][0]);
    else if (b.length === 1 && b[0].length === C) keepC = (keepC || Array(C).fill(true)).map((k, j) => k && b[0][j]);
    else if (R === 1 && C === 1 && b.length === 1) keepR = [b[0][0]];
    else throw err('#VALUE!', 'Ukuran kondisi FILTER tidak cocok dengan rentang');
  }
  let out = a;
  if (keepR) out = out.filter((_, i) => keepR[i]);
  if (keepC) out = out.map(row => row.filter((_, j) => keepC[j]));
  if (!out.length || !out[0].length) {
    if (ifEmpty !== undefined) return ifEmpty;
    throw err('#CALC!', 'Tidak ada kecocokan dalam evaluasi FILTER');
  }
  return out;
}, { min: 2 });
function sortRows(rows, keys) {
  return rows.map((r, i) => [r, i]).sort((x, y) => {
    for (const [k, asc] of keys) {
      const a = x[0][k], b = y[0][k];
      if (a == null && b == null) continue;
      if (a == null) return 1;
      if (b == null) return -1;
      const c = compare(a, b);
      if (c) return asc ? c : -c;
    }
    return x[1] - y[1];
  }).map(p => p[0]);
}
def('SORT', ([arr, ...rest]) => {
  let a = toArr(arr);
  const ascOf = v => (v === undefined ? true : typeof v === 'boolean' ? v : num(v) >= 0);
  let byCol = false;
  const keys = [];
  if (rest.length <= 3 && (rest.length < 3 || typeof scalar(rest[2]) === 'boolean' || rest[2] === undefined)) {
    byCol = rest[2] !== undefined && bool(rest[2]);
    const idx = rest[0] === undefined ? [[1]] : toArr(rest[0]);
    const ord = rest[1] === undefined ? [[1]] : toArr(rest[1]);
    idx.flat().forEach((k, i) => keys.push([int(k) - 1, ascOf(ord.flat()[i] ?? ord.flat()[0])]));
  } else {
    for (let i = 0; i < rest.length; i += 2) keys.push([int(rest[i]) - 1, ascOf(rest[i + 1])]);
  }
  if (byCol) a = transpose(a);
  const W = a[0].length;
  if (keys.some(([k]) => k < 0 || k >= W)) throw err('#VALUE!', 'Indeks sortir di luar rentang');
  let out = sortRows(a, keys);
  return byCol ? transpose(out) : out;
}, { min: 1 });
def('SORTBY', ([arr, ...rest]) => {
  const a = toArr(arr);
  const keys = [];
  for (let i = 0; i < rest.length; i += 2) keys.push([vec(rest[i]), rest[i + 1] === undefined ? 1 : num(rest[i + 1])]);
  if (keys.some(([k]) => k.length !== a.length)) throw ERR.VALUE;
  return a.map((r, i) => [r, i]).sort((x, y) => {
    for (const [k, o] of keys) { const c = compare(k[x[1]], k[y[1]]); if (c) return o < 0 ? -c : c; }
    return x[1] - y[1];
  }).map(p => p[0]);
}, { min: 2 });
def('SORTN', ([arr, n, mode, ...rest]) => {
  const a = toArr(arr);
  const keys = [];
  for (let i = 0; i < rest.length; i += 2) keys.push([int(rest[i]) - 1, rest[i + 1] === undefined ? true : bool(rest[i + 1])]);
  if (!keys.length) keys.push([0, true]);
  const sorted = sortRows(a, keys);
  return sorted.slice(0, n === undefined ? 1 : int(n));
}, { min: 1 });
const rowKey = r => r.map(v => (v == null ? '' : typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + String(v))).join('\u0001');
def('UNIQUE', ([arr, bc, once]) => {
  let a = toArr(arr);
  const byCol = bc !== undefined && bool(bc), exactly = once !== undefined && bool(once);
  if (byCol) a = transpose(a);
  const count = new Map(), first = [];
  for (const r of a) { const k = rowKey(r); if (!count.has(k)) { count.set(k, 0); first.push([k, r]); } count.set(k, count.get(k) + 1); }
  let out = first.filter(([k]) => !exactly || count.get(k) === 1).map(p => p[1]);
  if (!out.length) throw err('#CALC!', 'Tidak ada nilai unik');
  return byCol ? transpose(out) : out;
}, { min: 1, max: 3 });
def('SEQUENCE', ([r, c, s, st]) => {
  const R = int(r), C = int(opt(c, 1)), start = num(opt(s, 1)), step = num(opt(st, 1));
  if (R < 1 || C < 1) throw err('#CALC!', 'SEQUENCE memerlukan ukuran positif');
  if (R * C > 2e6) throw ERR.NUM;
  return Array.from({ length: R }, (_, i) => Array.from({ length: C }, (_, j) => fixf(start + (i * C + j) * step)));
}, { min: 1, max: 4 });
def('ARRAYFORMULA', ([v]) => v, { min: 1, max: 1, fmtArg: 0 });
def('ARRAY_CONSTRAIN', ([a, r, c]) => toArr(a).slice(0, int(r)).map(row => row.slice(0, int(c))), { min: 3, max: 3 });
const pickIdx = (n, list) => list.map(k => { const i = int(k); if (!i || Math.abs(i) > n) throw ERR.VALUE; return i > 0 ? i - 1 : n + i; });
def('CHOOSECOLS', ([a, ...cols]) => { const A = toArr(a); const idx = pickIdx(A[0].length, flat(cols)); return A.map(r => idx.map(i => r[i])); }, { min: 2 });
def('CHOOSEROWS', ([a, ...rows]) => { const A = toArr(a); const idx = pickIdx(A.length, flat(rows)); return idx.map(i => A[i]); }, { min: 2 });
def('TAKE', ([a, r, c]) => {
  let A = toArr(a);
  if (r !== undefined) { const n = int(r); A = n >= 0 ? A.slice(0, n) : A.slice(n); }
  if (c !== undefined) { const n = int(c); A = A.map(row => (n >= 0 ? row.slice(0, n) : row.slice(n))); }
  if (!A.length || !A[0].length) throw ERR.CALC;
  return A;
}, { min: 2, max: 3 });
def('DROP', ([a, r, c]) => {
  let A = toArr(a);
  if (r !== undefined) { const n = int(r); A = n >= 0 ? A.slice(n) : A.slice(0, n); }
  if (c !== undefined) { const n = int(c); A = A.map(row => (n >= 0 ? row.slice(n) : row.slice(0, n))); }
  if (!A.length || !A[0].length) throw ERR.CALC;
  return A;
}, { min: 2, max: 3 });
def('EXPAND', ([a, r, c, p]) => {
  const A = toArr(a), R = r === undefined ? A.length : int(r), C = c === undefined ? A[0].length : int(c);
  if (R < A.length || C < A[0].length) throw ERR.VALUE;
  const pad = p === undefined ? ERR.NA : scalar(p);
  return Array.from({ length: R }, (_, i) => Array.from({ length: C }, (_, j) => (A[i] && j < A[i].length ? A[i][j] : pad)));
}, { min: 2, max: 4 });
def('VSTACK', args => {
  const arrs = args.filter(x => x !== undefined).map(toArr);
  const C = Math.max(...arrs.map(a => a[0].length));
  return arrs.flatMap(a => a.map(r => (r.length < C ? [...r, ...Array(C - r.length).fill(ERR.NA)] : r)));
}, { min: 1 });
def('HSTACK', args => {
  const arrs = args.filter(x => x !== undefined).map(toArr);
  const R = Math.max(...arrs.map(a => a.length));
  return Array.from({ length: R }, (_, i) => arrs.flatMap(a => a[i] || Array(a[0].length).fill(ERR.NA)));
}, { min: 1 });
function toList(a, ignore, byCol) {
  let A = toArr(a);
  if (byCol) A = transpose(A);
  const ig = int(opt(ignore, 0));
  return A.flat().filter(v => !((ig === 1 || ig === 3) && v == null) && !((ig === 2 || ig === 3) && v instanceof FErr));
}
def('TOCOL', ([a, ig, bc]) => toList(a, ig, bc !== undefined && bool(bc)).map(v => [v]), { min: 1, max: 3 });
def('TOROW', ([a, ig, bc]) => [toList(a, ig, bc !== undefined && bool(bc))], { min: 1, max: 3 });
def('FLATTEN', args => flat(args).map(v => [v]), { min: 1 });
function wrap(v, n, pad, rowsFirst) {
  const list = vec(v), k = int(n);
  if (k < 1) throw ERR.NUM;
  const p = pad === undefined ? ERR.NA : scalar(pad);
  const out = [];
  for (let i = 0; i < list.length; i += k) { const chunk = list.slice(i, i + k); while (chunk.length < k) chunk.push(p); out.push(chunk); }
  return rowsFirst ? out : transpose(out);
}
def('WRAPROWS', ([v, n, p]) => wrap(v, n, p, true), { min: 2, max: 3 });
def('WRAPCOLS', ([v, n, p]) => wrap(v, n, p, false), { min: 2, max: 3 });

// ---------- TEXT ----------
S('CHAR', n => { const v = int(n); if (v < 1 || v > 255) throw ERR.VALUE; return String.fromCharCode(v); });
S('UNICHAR', n => { const v = int(n); if (v < 1) throw ERR.VALUE; return String.fromCodePoint(v); });
S('CODE', s => { const t = str(s); if (!t) throw ERR.VALUE; return t.charCodeAt(0); });
S('UNICODE', s => { const t = str(s); if (!t) throw ERR.VALUE; return t.codePointAt(0); });
S('CLEAN', s => str(s).replace(/[\x00-\x1f]/g, ''));
def('CONCAT CONCATENATE', args => flat(args).map(v => { if (v instanceof FErr) throw v; return str(v); }).join(''));
S('DOLLAR', (n, d) => { const k = int(opt(d, 2)); const v = rnd(num(n), k); return format(v, '"' + L.cur + '"#,##0' + (k > 0 ? '.' + '0'.repeat(k) : '') + ';-"' + L.cur + '"#,##0' + (k > 0 ? '.' + '0'.repeat(k) : '')).t; });
S('FIXED', (n, d, nc) => { const k = int(opt(d, 2)); const v = rnd(num(n), k); return format(v, (nc !== undefined && bool(nc) ? '0' : '#,##0') + (k > 0 ? '.' + '0'.repeat(k) : '')).t; });
S('EXACT', (a, b) => str(a) === str(b));
S('FIND FINDB', (f, w, s) => { const t = str(w), q = str(f), st = int(opt(s, 1)); if (st < 1 || st > t.length + 1) throw ERR.VALUE; const i = t.indexOf(q, st - 1); if (i < 0) throw err('#VALUE!', `"${q}" tidak ditemukan`); return i + 1; });
S('SEARCH SEARCHB', (f, w, s) => {
  const t = str(w), q = str(f), st = int(opt(s, 1)); if (st < 1 || st > t.length + 1) throw ERR.VALUE;
  const re = wildcardRe(q, false); const m = re.exec(t.slice(st - 1));
  if (!m) throw err('#VALUE!', `"${q}" tidak ditemukan`); return m.index + st;
});
S('LEFT LEFTB', (s, n) => { const k = int(opt(n, 1)); if (k < 0) throw ERR.VALUE; return str(s).slice(0, k); });
S('RIGHT RIGHTB', (s, n) => { const k = int(opt(n, 1)); if (k < 0) throw ERR.VALUE; const t = str(s); return k ? t.slice(-k) : ''; });
S('MID MIDB', (s, st, n) => { const a = int(st), k = int(n); if (a < 1 || k < 0) throw ERR.VALUE; return str(s).substr(a - 1, k); });
S('LEN LENB', s => str(s).length);
S('LOWER', s => str(s).toLowerCase());
S('UPPER', s => str(s).toUpperCase());
S('PROPER', s => str(s).toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (m, a, b) => a + b.toUpperCase()));
S('REPLACE REPLACEB', (s, st, n, r) => { const t = str(s), a = int(st), k = int(n); if (a < 1 || k < 0) throw ERR.VALUE; return t.slice(0, a - 1) + str(r) + t.slice(a - 1 + k); });
S('REPT', (s, n) => { const k = int(n); if (k < 0 || k * str(s).length > 32767) throw ERR.VALUE; return str(s).repeat(k); });
S('SUBSTITUTE', (s, o, n, i) => {
  const t = str(s), old = str(o), nw = str(n);
  if (!old) return t;
  if (i === undefined) return t.split(old).join(nw);
  const k = int(i); if (k < 1) throw ERR.VALUE;
  let pos = -1; for (let c = 0; c < k; c++) { pos = t.indexOf(old, pos + 1); if (pos < 0) return t; }
  return t.slice(0, pos) + nw + t.slice(pos + old.length);
});
S('T', v => (typeof v === 'string' ? v : ''));
S('TRIM', s => str(s).replace(/ +/g, ' ').trim());
S('TEXT', (v, f) => { const x = typeof v === 'string' && parseNumStr(v) !== null ? parseNumStr(v) : v; return format(x, str(f)).t; });
def('TEXTJOIN', ([d, ig, ...vals]) => {
  const delims = flat([d]).map(str);
  const ignore = bool(ig);
  const items = flat(vals, true).filter(v => !(ignore && (v == null || v === ''))).map(v => { if (v instanceof FErr) throw v; return str(v); });
  if (!delims.length) return items.join('');
  return items.reduce((s, x, i) => (i ? s + delims[(i - 1) % delims.length] + x : x), '');
}, { min: 3 });
S('VALUE', v => { if (typeof v === 'number') return v; if (v == null || v === '') return 0; if (typeof v === 'boolean') throw ERR.VALUE; const n = parseNumStr(str(v)); if (n === null) throw err('#VALUE!', `"${str(v)}" bukan angka`); return n; });
S('NUMBERVALUE', (v, d, g) => {
  let t = str(v).replace(/\s/g, '');
  const dec = d === undefined ? L.dec : str(d)[0], grp = g === undefined ? L.grp : str(g)[0];
  t = t.split(grp).join('');
  if (dec !== '.') t = t.replace(dec, '.');
  let pct = 0; while (t.endsWith('%')) { pct++; t = t.slice(0, -1); }
  const n = Number(t); if (t === '' ) return 0; if (isNaN(n)) throw ERR.VALUE; return n / Math.pow(100, pct);
});
function findNth(t, delims, n, caseIns, fromEnd) {
  const T = caseIns ? t.toLowerCase() : t;
  const ds = delims.map(x => (caseIns ? x.toLowerCase() : x)).filter(Boolean);
  const hits = [];
  for (let i = 0; i < T.length; i++) for (const d of ds) if (T.startsWith(d, i)) { hits.push([i, d.length]); break; }
  if (!hits.length) return null;
  const k = n > 0 ? n - 1 : hits.length + n;
  return hits[k] || null;
}
S('TEXTBEFORE', (t, d, n, m, e, nf) => {
  const s = str(t), inst = int(opt(n, 1)); if (!inst) throw ERR.VALUE;
  const h = findNth(s, flat([d]).map(str), inst, int(opt(m, 0)) === 1);
  if (!h) { if (opt(e, 0) && int(e)) return inst > 0 ? s : ''; if (nf !== undefined) return nf; throw ERR.NA; }
  return s.slice(0, h[0]);
}, { scalar: [0, 2, 3, 4] });
S('TEXTAFTER', (t, d, n, m, e, nf) => {
  const s = str(t), inst = int(opt(n, 1)); if (!inst) throw ERR.VALUE;
  const h = findNth(s, flat([d]).map(str), inst, int(opt(m, 0)) === 1);
  if (!h) { if (opt(e, 0) && int(e)) return inst > 0 ? '' : s; if (nf !== undefined) return nf; throw ERR.NA; }
  return s.slice(h[0] + h[1]);
}, { scalar: [0, 2, 3, 4] });
function splitBy(s, delims, caseIns) {
  const ds = delims.filter(Boolean);
  if (!ds.length) return [s];
  const re = new RegExp(ds.map(d => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), caseIns ? 'gi' : 'g');
  return s.split(re);
}
def('TEXTSPLIT', ([t, cd, rd, ie, mm, pad]) => {
  const s = str(t), ci = int(opt(mm, 0)) === 1, ignore = ie !== undefined && bool(ie);
  const rows = rd === undefined ? [s] : splitBy(s, flat([rd]).map(str), ci);
  let out = rows.map(r => (cd === undefined ? [r] : splitBy(r, flat([cd]).map(str), ci)));
  if (ignore) out = out.map(r => r.filter(x => x !== '')).filter(r => r.length);
  const C = Math.max(1, ...out.map(r => r.length));
  const p = pad === undefined ? ERR.NA : scalar(pad);
  return out.length ? out.map(r => [...r, ...Array(C - r.length).fill(p)]) : [['']];
}, { min: 2, max: 6 });
def('SPLIT', ([t, d, each, rm]) => {
  const s = str(t), del = str(d);
  const byEach = each === undefined || bool(each), removeEmpty = rm === undefined || bool(rm);
  let parts = byEach ? s.split(new RegExp('[' + del.replace(/[\]\\^-]/g, '\\$&') + ']')) : s.split(del);
  if (removeEmpty) parts = parts.filter(x => x !== '');
  return [parts.map(p => { const n = parseNumStr(p); return n !== null && /^\s*[-+]?[\d.,]+\s*$/.test(p) ? n : p; })];
}, { min: 2, max: 4 });
def('JOIN', ([d, ...vals]) => flat(vals).map(str).join(str(d)), { min: 2 });
function mkRe(p, flags = '') { try { return new RegExp(str(p), flags); } catch { throw err('#REF!', 'Ekspresi reguler tidak valid'); } }
S('REGEXMATCH', (t, p) => mkRe(p).test(str(t)));
def('REGEXEXTRACT', (t, p) => {
  const m = mkRe(p).exec(str(t));
  if (!m) throw err('#N/A', 'Ekspresi reguler tidak cocok');
  if (m.length > 2) return [m.slice(1)];
  return m.length === 2 ? m[1] : m[0];
}, { scalar: [0], min: 2, max: 2 });
S('REGEXREPLACE', (t, p, r) => str(t).replace(mkRe(p, 'g'), str(r)));
S('VALUETOTEXT', (v, f) => (int(opt(f, 0)) === 1 && typeof v === 'string' ? '"' + v + '"' : v instanceof FErr ? v.code : str(v)), { errOk: true });
def('ARRAYTOTEXT', ([a, f]) => {
  const strict = int(opt(f, 0)) === 1;
  const A = toArr(a);
  const s = v => (v instanceof FErr ? v.code : strict && typeof v === 'string' ? '"' + v + '"' : str(v));
  return strict ? '{' + A.map(r => r.map(s).join(',')).join(';') + '}' : A.flat().map(s).join(', ');
}, { min: 1, max: 2, errOk: true });
S('TO_TEXT', v => (typeof v === 'number' ? general(v) : str(v)));
S('TO_PERCENT', v => num(v), { fmt: '0.00%' });
S('TO_DOLLARS', v => num(v), { fmt: 'cur' });
S('TO_DATE', v => num(v), { fmt: 'date' });
S('TO_PURE_NUMBER', v => num(v), { fmt: 'General' });
S('ISEMAIL', v => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(str(v)));
S('ISURL', v => /^(https?:\/\/)?([\w-]+\.)+[a-z]{2,}(\/\S*)?$/i.test(str(v)));
def('SPARKLINE', ([data, opts]) => {
  const vals = flat([data], true).map(v => (typeof v === 'number' ? v : null));
  const o = {};
  if (opts !== undefined) for (const row of toArr(opts)) if (row.length >= 2 && row[0] != null) o[str(row[0]).toLowerCase()] = row[1];
  return new Special('spark', { vals, opts: o });
}, { min: 1, max: 2 });
def('IMAGE', ([url, mode, h, w]) => new Special('image', { url: str(url), mode: int(opt(mode, 1)), h: h === undefined ? null : num(h), w: w === undefined ? null : num(w) }), { min: 1, max: 4 });

// ---------- DATE & TIME ----------
export function serialNum(v) {
  if (typeof v === 'string') {
    const d = parseDate(v);
    if (d) return d.v;
  }
  const n = num(v);
  if (n < 0) throw ERR.NUM;
  return n;
}
const P = v => parts(serialNum(v));
S('DATE', (y, m, d) => {
  let Y = int(y); if (Y < 1900) Y += 1900; if (Y < 0 || Y > 9999) throw ERR.NUM;
  const s = serial(Y, int(m), int(d)); if (s < 0) throw ERR.NUM; return s;
}, { fmt: 'date' });
S('DATEVALUE', s => { const d = parseDate(str(s)); if (!d) throw err('#VALUE!', `DATEVALUE tidak dapat mengurai "${str(s)}"`); return Math.floor(d.v); }, { fmt: 'date' });
S('TIMEVALUE', s => { const d = parseDate(str(s)); if (!d) throw ERR.VALUE; return d.v - Math.floor(d.v); }, { fmt: 'time' });
S('DAY', v => P(v).d);
S('MONTH', v => P(v).m);
S('YEAR', v => P(v).y);
S('HOUR', v => P(v).H);
S('MINUTE', v => P(v).M);
S('SECOND', v => P(v).S);
S('WEEKDAY', (v, t) => {
  const dow = P(v).dow, k = int(opt(t, 1));
  switch (k) {
    case 1: case 17: return dow + 1;
    case 2: case 11: return ((dow + 6) % 7) + 1;
    case 3: return (dow + 6) % 7;
    case 12: return ((dow + 5) % 7) + 1; case 13: return ((dow + 4) % 7) + 1; case 14: return ((dow + 3) % 7) + 1;
    case 15: return ((dow + 2) % 7) + 1; case 16: return ((dow + 1) % 7) + 1;
    default: throw ERR.NUM;
  }
});
function isoWeek(sv) {
  const p = parts(sv);
  const d = new Date(Date.UTC(p.y, p.m - 1, p.d));
  const day = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - day + 3);
  const fy = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((d - fy) / 86400000 - 3 + ((fy.getUTCDay() + 6) % 7)) / 7);
}
S('ISOWEEKNUM', v => isoWeek(serialNum(v)));
S('WEEKNUM', (v, t) => {
  const sv = serialNum(v), k = int(opt(t, 1));
  if (k === 21) return isoWeek(sv);
  const start = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 }[k];
  if (start === undefined) throw ERR.NUM;
  const p = parts(sv);
  const jan1 = serial(p.y, 1, 1), jdow = parts(jan1).dow;
  const off = (jdow - start + 7) % 7;
  return Math.floor((Math.floor(sv) - jan1 + off) / 7) + 1;
});
S('DAYS', (e, s) => Math.floor(serialNum(e)) - Math.floor(serialNum(s)));
S('DAYS360', (s, e, m) => {
  const a = P(s), b = P(e), eu = m !== undefined && bool(m);
  let d1 = a.d, d2 = b.d;
  if (eu) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
  else {
    const lastFeb = p => p.m === 2 && parts(serial(p.y, 3, 1) - 1).d === p.d;
    if (lastFeb(a)) d1 = 30;
    if (d1 === 31) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
  }
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1);
});
S('DATEDIF', (s, e, u) => {
  const sv = Math.floor(serialNum(s)), ev = Math.floor(serialNum(e));
  if (sv > ev) throw ERR.NUM;
  const a = parts(sv), b = parts(ev), unit = str(u).toUpperCase();
  let months = (b.y - a.y) * 12 + (b.m - a.m); if (b.d < a.d) months--;
  switch (unit) {
    case 'D': return ev - sv;
    case 'M': return months;
    case 'Y': return Math.floor(months / 12);
    case 'YM': return months % 12;
    case 'MD': { let d = b.d - a.d; if (d < 0) { const pm = parts(serial(b.y, b.m, 1) - 1); d += pm.d; } return d; }
    case 'YD': { let ys = serial(b.y, a.m, a.d); if (ys > ev) ys = serial(b.y - 1, a.m, a.d); return ev - ys; }
    default: throw ERR.NUM;
  }
});
function addMonths(sv, n, eom) {
  const p = parts(sv);
  const y = p.y, m = p.m - 1 + n;
  const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return serial(y, m + 1, eom ? last : Math.min(p.d, last));
}
S('EDATE', (s, m) => addMonths(Math.floor(serialNum(s)), int(m), false), { fmt: 'date' });
S('EOMONTH', (s, m) => addMonths(Math.floor(serialNum(s)), int(m), true), { fmt: 'date' });
def('NOW', () => jsDateToSerial(new Date()), { max: 0, fmt: 'datetime', volatile: true });
def('TODAY', () => Math.floor(jsDateToSerial(new Date())), { max: 0, fmt: 'date', volatile: true });
S('TIME', (h, m, s) => { const t = (int(h) * 3600 + int(m) * 60 + int(s)) / 86400; if (t < 0) throw ERR.NUM; return t % 1; }, { fmt: 'time' });
function weekendMask(w) {
  if (w === undefined || w === null) return [true, false, false, false, false, false, true];
  if (typeof w === 'string') { if (!/^[01]{7}$/.test(w)) throw ERR.VALUE; const m = w.split('').map(x => x === '1'); return [m[6], ...m.slice(0, 6)]; }
  const k = int(w);
  const pairs = { 1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6] };
  const single = { 11: 0, 12: 1, 13: 2, 14: 3, 15: 4, 16: 5, 17: 6 };
  const m = Array(7).fill(false);
  if (pairs[k]) pairs[k].forEach(d => (m[d] = true));
  else if (single[k] !== undefined) m[single[k]] = true;
  else throw ERR.NUM;
  return m;
}
const holidaySet = h => new Set(h === undefined ? [] : flat([h]).filter(x => x != null).map(x => Math.floor(serialNum(x))));
function networkdays(s, e, mask, hol) {
  let a = Math.floor(serialNum(s)), b = Math.floor(serialNum(e)), sign = 1;
  if (a > b) { [a, b] = [b, a]; sign = -1; }
  let n = 0;
  for (let d = a; d <= b; d++) if (!mask[parts(d).dow] && !hol.has(d)) n++;
  return n * sign;
}
function workday(s, days, mask, hol) {
  let d = Math.floor(serialNum(s)), k = int(days);
  const step = k < 0 ? -1 : 1;
  if (mask.every(Boolean)) throw ERR.VALUE;
  while (k) { d += step; if (!mask[parts(d).dow] && !hol.has(d)) k -= step; }
  return d;
}
def('NETWORKDAYS', ([s, e, h]) => networkdays(s, e, weekendMask(), holidaySet(h)), { min: 2, max: 3 });
def('NETWORKDAYS.INTL', ([s, e, w, h]) => networkdays(s, e, weekendMask(w === undefined ? undefined : scalar(w)), holidaySet(h)), { min: 2, max: 4 });
def('WORKDAY', ([s, d, h]) => workday(s, d, weekendMask(), holidaySet(h)), { min: 2, max: 3, fmt: 'date' });
def('WORKDAY.INTL', ([s, d, w, h]) => workday(s, d, weekendMask(w === undefined ? undefined : scalar(w)), holidaySet(h)), { min: 2, max: 4, fmt: 'date' });
S('YEARFRAC', (s, e, b) => {
  let a = Math.floor(serialNum(s)), c = Math.floor(serialNum(e));
  if (a > c) [a, c] = [c, a];
  const basis = int(opt(b, 0));
  const pa = parts(a), pc = parts(c);
  switch (basis) {
    case 0: {
      let d1 = pa.d, d2 = pc.d;
      if (d1 === 31) d1 = 30;
      if (d2 === 31 && d1 >= 30) d2 = 30;
      return ((pc.y - pa.y) * 360 + (pc.m - pa.m) * 30 + (d2 - d1)) / 360;
    }
    case 1: {
      const yl = y => (new Date(Date.UTC(y, 1, 29)).getUTCMonth() === 1 ? 366 : 365);
      if (pa.y === pc.y) return (c - a) / yl(pa.y);
      let tot = 0; for (let y = pa.y; y <= pc.y; y++) tot += yl(y);
      return (c - a) / (tot / (pc.y - pa.y + 1));
    }
    case 2: return (c - a) / 360;
    case 3: return (c - a) / 365;
    case 4: return ((pc.y - pa.y) * 360 + (pc.m - pa.m) * 30 + (Math.min(pc.d, 30) - Math.min(pa.d, 30))) / 360;
    default: throw ERR.NUM;
  }
});
S('EPOCHTODATE', (t, u) => { const k = int(opt(u, 1)); const ms = num(t) * (k === 1 ? 1000 : k === 2 ? 1 : 0.001); return (ms + 2209161600000) / 86400000; }, { fmt: 'datetime' });

// ---------- INFO ----------
S('ISBLANK', v => v == null, { errOk: true });
S('ISERR', v => v instanceof FErr && v.code !== '#N/A', { errOk: true });
S('ISERROR', v => v instanceof FErr, { errOk: true });
S('ISNA', v => v instanceof FErr && v.code === '#N/A', { errOk: true });
S('ISEVEN', v => Math.floor(Math.abs(num(v))) % 2 === 0);
S('ISODD', v => Math.floor(Math.abs(num(v))) % 2 === 1);
S('ISLOGICAL', v => typeof v === 'boolean', { errOk: true });
S('ISNONTEXT', v => typeof v !== 'string', { errOk: true });
S('ISNUMBER', v => typeof v === 'number', { errOk: true });
S('ISTEXT', v => typeof v === 'string', { errOk: true });
S('ISDATE', v => typeof v === 'number' && v >= 0, { errOk: true });
S('N', v => (typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0));
def('NA', () => ERR.NA, { max: 0 });
def('TYPE', ([v]) => {
  if (v instanceof FErr) return 16;
  if (Array.isArray(v) || (v instanceof Ref && !v.single)) return 64;
  const x = scalar(v);
  if (x instanceof FErr) return 16;
  return typeof x === 'number' || x == null ? 1 : typeof x === 'string' ? 2 : typeof x === 'boolean' ? 4 : 128;
}, { min: 1, max: 1, errOk: true });
S('ERROR.TYPE', v => {
  if (!(v instanceof FErr)) throw ERR.NA;
  return { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#SPILL!': 9, '#CALC!': 14 }[v.code] || 8;
}, { errOk: true });

// Google-style operator functions
S('ADD', (a, b) => num(a) + num(b));
S('MINUS', (a, b) => num(a) - num(b));
S('MULTIPLY', (a, b) => num(a) * num(b));
S('DIVIDE', (a, b) => { const d = num(b); if (!d) throw ERR.DIV0; return num(a) / d; });
S('EQ', (a, b) => compare(a, b) === 0);
S('NE', (a, b) => compare(a, b) !== 0);
S('GT', (a, b) => compare(a, b) > 0);
S('GTE', (a, b) => compare(a, b) >= 0);
S('LT', (a, b) => compare(a, b) < 0);
S('LTE', (a, b) => compare(a, b) <= 0);
S('UMINUS', a => -num(a));
S('UPLUS', a => num(a));
S('UNARY_PERCENT', a => num(a) / 100, { fmt: '0%' });

export { colIndex };
