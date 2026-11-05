// Generic UI primitives: menus, popups, dialogs, toasts, color picker, tooltips.
import { el, icon } from '../util.js';

let openMenus = [];
const layer = () => document.getElementById('popups');

export function closeMenus(except) {
  for (const m of openMenus.slice()) if (m !== except && !m.contains(except)) { m.remove(); openMenus.splice(openMenus.indexOf(m), 1); }
  document.querySelectorAll('.menubar .mb-item.open, .tb-btn.open').forEach(b => { if (!openMenus.length) b.classList.remove('open'); });
}

function place(node, at, opts = {}) {
  node.style.visibility = 'hidden';
  layer().append(node);
  const r = node.getBoundingClientRect();
  const vw = window.innerWidth, vh = window.innerHeight;
  let x, y;
  if (at instanceof DOMRect || (at && at.width !== undefined && at.left !== undefined)) {
    if (opts.side === 'right') { x = at.right - 2; y = at.top - 6; if (x + r.width > vw - 6) x = at.left - r.width + 2; }
    else { x = opts.alignRight ? at.right - r.width : at.left; y = at.bottom + 4; if (y + r.height > vh - 6) y = Math.max(6, at.top - r.height - 4); }
  } else { x = at.x; y = at.y; }
  if (x + r.width > vw - 6) x = vw - r.width - 6;
  if (y + r.height > vh - 6) y = Math.max(6, vh - r.height - 6);
  node.style.left = Math.max(6, x) + 'px';
  node.style.top = Math.max(6, y) + 'px';
  node.style.visibility = '';
}

const isOn = f => (typeof f === 'function' ? f() : f);

export function openMenu(items, at, opts = {}) {
  if (!opts.sub) closeMenus();
  const list = typeof items === 'function' ? items() : items;
  const menu = el('div.menu.glass', { role: 'menu' });
  let subTimer = null, subOpen = null;
  const rows = [];
  for (const it of list) {
    if (!it || isOn(it.hidden)) continue;
    if (it.sep) { menu.append(el('div.menu-sep')); continue; }
    if (it.head) { menu.append(el('div.menu-head', it.head)); continue; }
    if (it.node) { menu.append(typeof it.node === 'function' ? it.node(() => closeMenus()) : it.node); continue; }
    const dis = isOn(it.disabled), chk = isOn(it.checked);
    const row = el('div.menu-item' + (dis ? '.disabled' : '') + (it.sub ? '.has-sub' : ''), { role: 'menuitem', tabIndex: -1 },
      el('span.mi-icon', chk ? icon('check') : it.icon ? icon(it.icon) : it.swatch ? el('span.swatch', { style: { background: it.swatch } }) : ''),
      el('span.mi-label', it.label, it.hint ? el('span.mi-hint', it.hint) : null),
      it.kbd ? el('span.mi-kbd', it.kbd) : null,
      it.sub ? icon('chevron_right', 'mi-arrow') : null);
    if (it.preview) row.append(el('span.mi-prev', it.preview));
    const openSub = () => {
      if (subOpen) { subOpen.remove(); openMenus.splice(openMenus.indexOf(subOpen), 1); subOpen = null; }
      menu.querySelectorAll('.menu-item.sub-open').forEach(x => x.classList.remove('sub-open'));
      if (!it.sub || dis) return;
      row.classList.add('sub-open');
      subOpen = openMenu(it.sub, row.getBoundingClientRect(), { sub: true, side: 'right', parent: menu });
    };
    row.addEventListener('mouseenter', () => { clearTimeout(subTimer); subTimer = setTimeout(openSub, it.sub ? 90 : 160); rows.forEach(x => x.classList.remove('kb')); row.classList.add('kb'); });
    row.addEventListener('click', e => {
      e.stopPropagation();
      if (dis) return;
      if (it.sub) { openSub(); return; }
      closeMenus();
      opts.onPick?.();
      it.run?.();
    });
    rows.push(row);
    row._it = it;
    menu.append(row);
  }
  openMenus.push(menu);
  menu._rows = rows;
  menu._openSub = () => { const k = rows.find(r => r.classList.contains('kb')); if (k && k._it.sub) k.dispatchEvent(new Event('click')); };
  place(menu, at, opts);
  return menu;
}

// keyboard navigation inside the top-most menu
document.addEventListener('keydown', e => {
  if (!openMenus.length) return;
  const m = openMenus[openMenus.length - 1];
  const rows = (m._rows || []).filter(r => !r.classList.contains('disabled'));
  if (!rows.length && e.key !== 'Escape') return;
  const i = rows.findIndex(r => r.classList.contains('kb'));
  const setKb = j => { rows.forEach(r => r.classList.remove('kb')); rows[j]?.classList.add('kb'); rows[j]?.scrollIntoView({ block: 'nearest' }); };
  if (e.key === 'ArrowDown') { e.preventDefault(); e.stopPropagation(); setKb((i + 1) % rows.length); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); setKb((i - 1 + rows.length) % rows.length); }
  else if (e.key === 'ArrowRight' && i >= 0 && rows[i]._it.sub) { e.preventDefault(); e.stopPropagation(); rows[i].click(); setTimeout(() => { const sm = openMenus[openMenus.length - 1]; sm?._rows?.[0]?.classList.add('kb'); }, 120); }
  else if (e.key === 'ArrowLeft' && openMenus.length > 1) { e.preventDefault(); e.stopPropagation(); m.remove(); openMenus.pop(); }
  else if (e.key === 'Enter' && i >= 0) { e.preventDefault(); e.stopPropagation(); rows[i].click(); }
  else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeMenus(); }
}, true);
document.addEventListener('pointerdown', e => {
  if (!openMenus.length) return;
  if (openMenus.some(m => m.contains(e.target))) return;
  if (e.target.closest('.mb-item, [data-menu-anchor]')) return;
  closeMenus();
}, true);

// Free-form popup anchored to a rect (dropdowns, pickers). Closes on outside click.
export function popup(content, at, opts = {}) {
  closeMenus();
  const p = el('div.menu.popup.glass' + (opts.cls ? '.' + opts.cls : ''), content);
  openMenus.push(p);
  p._rows = [];
  place(p, at, opts);
  return { el: p, close: () => closeMenus() };
}

// ---------- dialogs ----------
export function dialog({ title, body, buttons = [], width = 460, onClose, cls = '', noBackdropClose = false }) {
  const back = el('div.dlg-back');
  const close = () => {
    back.remove();
    document.removeEventListener('keydown', key, true);
    onClose?.();
    if (!document.querySelector('.dlg-back') && !document.activeElement?.closest?.('input,textarea,select')) document.querySelector('.cell-input')?.focus({ preventScroll: true });
  };
  const foot = el('div.dlg-foot');
  for (const b of buttons) {
    if (b.spacer) { foot.append(el('span.flex1')); continue; }
    foot.append(el('button.btn' + (b.primary ? '.primary' : '') + (b.danger ? '.danger' : ''), {
      type: 'button',
      onclick: async () => { const r = await b.run?.(); if (r !== false) close(); },
    }, b.label));
  }
  const box = el('div.dlg.glass' + (cls ? '.' + cls : ''), { role: 'dialog', 'aria-modal': 'true', style: { width: typeof width === 'number' ? width + 'px' : width } },
    el('div.dlg-head', el('h2', title), el('button.icon-btn', { type: 'button', 'aria-label': 'Tutup', onclick: close }, icon('close'))),
    el('div.dlg-body', body),
    buttons.length ? foot : null);
  back.append(box);
  back.addEventListener('pointerdown', e => { if (e.target === back && !noBackdropClose) close(); });
  const key = e => {
    if (e.key === 'Escape') { e.stopPropagation(); close(); }
    if (e.key === 'Enter' && !e.shiftKey && !e.target.dataset?.enter && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') {
      const p = buttons.find(b => b.primary);
      if (p) { e.preventDefault(); e.stopPropagation(); foot.querySelector('.primary')?.click(); }
    }
  };
  document.addEventListener('keydown', key, true);
  document.getElementById('dialogs').append(back);
  setTimeout(() => box.querySelector('input:not([type=checkbox]):not([type=radio]),select,textarea')?.focus(), 30);
  return { close, el: box };
}

// Password box for sheet/file passwords. Chrome's password manager must not touch these: it
// drops the account password into the first box (even with autocomplete=new-password), which then
// never matches the repeat box, and the user cannot see why behind the dots. So it is a text input
// masked by CSS, which password managers ignore; browsers without the CSS fall back to type=password.
export function secretInput(placeholder = '') {
  const masked = !!window.CSS?.supports?.('-webkit-text-security', 'disc');
  const inp = el('input.inp' + (masked ? '.secret' : ''), {
    type: masked ? 'text' : 'password', autocomplete: masked ? 'off' : 'new-password', placeholder,
    name: 'as-' + Math.random().toString(36).slice(2), spellcheck: 'false', autocapitalize: 'off', autocorrect: 'off',
    'data-lpignore': 'true', 'data-1p-ignore': 'true', 'data-bwignore': 'true', 'data-form-type': 'other',
  });
  inp._masked = masked;
  return inp;
}

// The box plus an eye button, so what was actually typed can be checked.
export function secretBox(inp) {
  const eye = icon('visibility');
  const btn = el('button.icon-btn.sm', { type: 'button', tabIndex: -1, 'data-tip': 'Tampilkan kata sandi', 'aria-label': 'Tampilkan kata sandi', onclick: () => {
    const show = eye.textContent === 'visibility';
    if (inp._masked) inp.classList.toggle('show', show); else inp.type = show ? 'text' : 'password';
    eye.textContent = show ? 'visibility_off' : 'visibility';
    btn.setAttribute('aria-label', show ? 'Sembunyikan kata sandi' : 'Tampilkan kata sandi');
    btn.dataset.tip = show ? 'Sembunyikan kata sandi' : 'Tampilkan kata sandi';
    inp.focus();
  } }, eye);
  return el('div.pw-box', inp, btn);
}

// Live "do they match" line under a new-password + repeat pair; returns the text for a mismatch.
export function secretMatch(a, b) {
  const note = el('div.pw-note', { role: 'status' });
  const diff = () => `Kata sandi tidak sama: kolom pertama ${a.value.length} karakter, kolom ulangi ${b.value.length} karakter.`;
  const same = () => a.value.trim() === b.value.trim();
  const update = () => {
    const on = a.value && b.value;
    note.className = 'pw-note' + (on ? (same() ? ' ok' : ' bad') : '');
    note.textContent = !on ? '' : same() ? 'Kata sandi sudah sama.' : diff();
  };
  a.addEventListener('input', update); b.addEventListener('input', update);
  return { note, same, diff, update };
}

export function promptDlg(title, label, value = '', opts = {}) {
  return new Promise(res => {
    const secret = opts.type === 'password';
    const inp = secret ? secretInput(opts.placeholder || '') : el('input.inp', { value, type: opts.type || 'text', placeholder: opts.placeholder || '' });
    let done = false;
    dialog({
      title, width: opts.width || 400,
      body: el('label.field', el('span', label), secret ? secretBox(inp) : inp, opts.hint ? el('small.muted', opts.hint) : null),
      buttons: [{ label: 'Batal' }, { label: opts.ok || 'OK', primary: true, run: () => { done = true; res(inp.value); } }],
      onClose: () => { if (!done) res(null); },
    });
    setTimeout(() => inp.select(), 40);
  });
}

export function confirmDlg(title, text, ok = 'OK', danger = false) {
  return new Promise(res => {
    let done = false;
    dialog({
      title, width: 420, body: el('p.dlg-text', text),
      buttons: [{ label: 'Batal' }, { label: ok, primary: !danger, danger, run: () => { done = true; res(true); } }],
      onClose: () => { if (!done) res(false); },
    });
  });
}

export function alertDlg(title, text) {
  return new Promise(res => dialog({ title, width: 420, body: el('p.dlg-text', text), buttons: [{ label: 'OK', primary: true }], onClose: res }));
}

// ---------- toast ----------
export function toast(msg, opts = {}) {
  const host = document.getElementById('toasts');
  const t = el('div.toast.glass', el('span', msg), opts.action ? el('button.btn.link', { type: 'button', onclick: () => { opts.action.run(); t.remove(); } }, opts.action.label) : null);
  host.append(t);
  requestAnimationFrame(() => t.classList.add('in'));
  setTimeout(() => { t.classList.remove('in'); setTimeout(() => t.remove(), 300); }, opts.ms || 3200);
}

// ---------- color picker ----------
export const PALETTE = [
  ['#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff'],
  ['#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff'],
  ['#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc'],
  ['#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd'],
  ['#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0'],
  ['#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79'],
  ['#85200c', '#990000', '#b45f06', '#bf9000', '#38761d', '#134f5c', '#1155cc', '#0b5394', '#351c75', '#741b47'],
  ['#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130'],
];

function customColors() { try { return JSON.parse(localStorage.getItem('as.customColors') || '[]'); } catch { return []; } }
function saveCustom(c) { try { const l = [c, ...customColors().filter(x => x !== c)].slice(0, 10); localStorage.setItem('as.customColors', JSON.stringify(l)); } catch { /* storage unavailable */ } }

export function colorGrid(onPick, opts = {}) {
  const wrap = el('div.cpick');
  if (opts.reset !== false) wrap.append(el('button.cp-reset', { type: 'button', onclick: () => onPick(null) }, icon('format_color_reset'), opts.resetLabel || 'Reset'));
  const grid = el('div.cp-grid');
  PALETTE.forEach((row, i) => {
    for (const c of row) grid.append(el('button.cp-sw' + (i === 1 ? '.gap' : ''), { type: 'button', title: c, style: { background: c }, onclick: () => onPick(c) }));
  });
  wrap.append(grid);
  const cust = el('div.cp-custom', el('span.cp-label', 'KUSTOM'));
  const row = el('div.cp-row');
  for (const c of customColors()) row.append(el('button.cp-sw', { type: 'button', title: c, style: { background: c }, onclick: () => onPick(c) }));
  const inp = el('input', { type: 'color', value: opts.value || '#1a73e8', 'aria-label': 'Warna kustom' });
  inp.addEventListener('change', () => { saveCustom(inp.value); onPick(inp.value); });
  row.append(el('label.cp-add', { title: 'Warna kustom' }, icon('add'), inp));
  cust.append(row);
  wrap.append(cust);
  return wrap;
}

export function pickColor(at, onPick, opts = {}) {
  let p;
  p = popup(colorGrid(c => { p.close(); onPick(c); }, opts), at);
  return p;
}

// ---------- tooltips ----------
let tipEl = null, tipT = null;
export function initTooltips() {
  tipEl = el('div.tooltip');
  document.body.append(tipEl);
  document.addEventListener('pointerover', e => {
    const t = e.target.closest('[data-tip]');
    clearTimeout(tipT);
    if (!t) { tipEl.classList.remove('show'); return; }
    tipT = setTimeout(() => {
      const r = t.getBoundingClientRect();
      tipEl.innerHTML = '';
      const [a, b] = t.dataset.tip.split('|');
      tipEl.append(el('span', a), ...(b ? [el('span.kbd', b)] : []));
      tipEl.classList.add('show');
      const tr = tipEl.getBoundingClientRect();
      let x = r.left + r.width / 2 - tr.width / 2;
      x = Math.max(6, Math.min(window.innerWidth - tr.width - 6, x));
      let y = r.bottom + 6;
      if (y + tr.height > window.innerHeight - 4) y = r.top - tr.height - 6;
      tipEl.style.left = x + 'px'; tipEl.style.top = y + 'px';
    }, 450);
  });
  document.addEventListener('pointerdown', () => { clearTimeout(tipT); tipEl.classList.remove('show'); });
}

export function field(label, input, hint) {
  return el('label.field', el('span', label), input, hint ? el('small.muted', hint) : null);
}

export function select(options, value, onchange, attrs = {}) {
  const s = el('select.inp', attrs);
  for (const o of options) {
    if (o.group) { const g = el('optgroup', { label: o.group }); for (const x of o.items) g.append(el('option', { value: x.value ?? x.id ?? x }, x.label ?? x)); s.append(g); continue; }
    const v = o.value ?? o.id ?? o;
    s.append(el('option', { value: v }, o.label ?? o));
  }
  if (value !== undefined) s.value = value;
  if (onchange) s.addEventListener('change', () => onchange(s.value));
  return s;
}

export function checkbox(label, checked, onchange) {
  const c = el('input', { type: 'checkbox', checked: !!checked });
  if (onchange) c.addEventListener('change', () => onchange(c.checked));
  return el('label.check', c, el('span', label));
}

export function segmented(options, value, onchange) {
  const w = el('div.seg');
  for (const o of options) {
    const b = el('button' + (o.value === value ? '.on' : ''), { type: 'button', onclick: () => { w.querySelectorAll('button').forEach(x => x.classList.remove('on')); b.classList.add('on'); onchange(o.value); }, 'data-tip': o.tip || '' }, o.icon ? icon(o.icon) : null, o.label || null);
    w.append(b);
  }
  return w;
}
