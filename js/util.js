// Shared helpers: addressing, DOM, dates.

export const MAXC = 16384;
export const MAXR = 1048576;
export const key = (r, c) => r * MAXC + c;
export const keyR = k => Math.floor(k / MAXC);
export const keyC = k => k % MAXC;

export function colName(c) {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colIndex(s) {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const addr = (r, c) => colName(c) + (r + 1);

export function rangeAddr(g, sheetName) {
  const pre = sheetName ? quoteSheet(sheetName) + '!' : '';
  if (g.r1 === g.r2 && g.c1 === g.c2) return pre + addr(g.r1, g.c1);
  return pre + addr(g.r1, g.c1) + ':' + addr(g.r2, g.c2);
}

export function quoteSheet(name) {
  return /^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name : "'" + name.replace(/'/g, "''") + "'";
}

// Parses "A1", "A1:B5", "Sheet1!A1:B2", "A:C", "3:5". Returns {sheet, r1,c1,r2,c2} or null.
export function parseRange(text) {
  if (!text) return null;
  let s = text.trim(), sheet = null;
  const bang = s.lastIndexOf('!');
  if (bang > 0) {
    sheet = s.slice(0, bang);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    s = s.slice(bang + 1);
  }
  s = s.replace(/\$/g, '').toUpperCase();
  let m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(s);
  if (m) {
    const r1 = +m[2] - 1, c1 = colIndex(m[1]);
    const r2 = m[3] ? +m[4] - 1 : r1, c2 = m[3] ? colIndex(m[3]) : c1;
    if (r1 < 0 || r2 < 0) return null;
    return norm({ sheet, r1, c1, r2, c2 });
  }
  m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(s);
  if (m) return norm({ sheet, r1: 0, c1: colIndex(m[1]), r2: Infinity, c2: colIndex(m[2]), wholeCol: true });
  m = /^(\d+):(\d+)$/.exec(s);
  if (m) return norm({ sheet, r1: +m[1] - 1, c1: 0, r2: +m[2] - 1, c2: Infinity, wholeRow: true });
  return null;
}

export function norm(g) {
  const o = { ...g };
  if (o.r1 > o.r2) [o.r1, o.r2] = [o.r2, o.r1];
  if (o.c1 > o.c2) [o.c1, o.c2] = [o.c2, o.c1];
  return o;
}

export const inRange = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
export const intersects = (a, b) => a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;
export const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

let uidN = 0;
export const uid = (p = 'id') => p + Date.now().toString(36) + (uidN++).toString(36) + Math.random().toString(36).slice(2, 6);

export function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

// Tiny hyperscript helper. el('div.cls#id', {attrs}, children...)
export function el(sel, attrs, ...kids) {
  const m = /^([a-z0-9-]+)?((?:[.#][\w-]+)*)$/i.exec(sel) || [];
  const node = document.createElement(m[1] || 'div');
  (m[2] || '').replace(/([.#])([\w-]+)/g, (_, t, n) => { if (t === '.') node.classList.add(n); else node.id = n; });
  if (attrs && (typeof attrs !== 'object' || attrs instanceof Node || Array.isArray(attrs))) { kids.unshift(attrs); attrs = null; }
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v == null || v === false) continue;
      if (k.startsWith('on')) node.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'style' && typeof v === 'object') { for (const [p, x] of Object.entries(v)) { if (p.startsWith('--')) node.style.setProperty(p, x); else node.style[p] = x; } }
      else if (k === 'html') node.innerHTML = v;
      else if (k === 'dataset') Object.assign(node.dataset, v);
      else if (k in node && typeof v !== 'string') node[k] = v;
      else node.setAttribute(k, v === true ? '' : v);
    }
  }
  for (const k of kids.flat(Infinity)) {
    if (k == null || k === false) continue;
    node.append(k instanceof Node ? k : document.createTextNode(String(k)));
  }
  return node;
}

export const icon = (name, cls = '') => el('span.ms' + (cls ? '.' + cls : ''), { 'aria-hidden': 'true' }, name);

// ---- Dates (serial numbers, 1899-12-30 epoch like Google Sheets) ----
export const EPOCH = Date.UTC(1899, 11, 30);
export const DAYMS = 86400000;

export function serial(y, m, d, H = 0, M = 0, S = 0) {
  return (Date.UTC(y, m - 1, d) - EPOCH) / DAYMS + (H * 3600 + M * 60 + S) / 86400;
}

export function parts(sv) {
  let days = Math.floor(sv);
  let secs = Math.round((sv - days) * 86400 * 1000) / 1000;
  if (secs >= 86400) { days += 1; secs -= 86400; }
  const dt = new Date(EPOCH + days * DAYMS);
  const whole = Math.floor(secs);
  return {
    y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: dt.getUTCDay(),
    H: Math.floor(whole / 3600), M: Math.floor(whole / 60) % 60, S: whole % 60, ms: secs - whole,
  };
}

export function jsDateToSerial(d) {
  return serial(d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds() + d.getMilliseconds() / 1000);
}

export const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

export function download(name, data, type = 'application/octet-stream') {
  const blob = data instanceof Blob ? data : new Blob([data], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}

const loaded = new Map();
export function loadScript(src) {
  if (!loaded.has(src)) {
    loaded.set(src, new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = src; s.async = true;
      s.onload = res; s.onerror = () => { loaded.delete(src); rej(new Error('Gagal memuat ' + src)); };
      document.head.append(s);
    }));
  }
  return loaded.get(src);
}

export const LIBS = {
  xlsx: 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js',
  chart: 'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js',
};
