// Toolbar with state sync and responsive overflow.
import { el, icon } from '../util.js';
import { openMenu, popup, pickColor, colorGrid, closeMenus } from './ui.js';
import { numberFormatItems, functionItems } from './menus.js';
import { L } from '../format.js';
import * as D from './dialogs.js';
import * as io from '../io.js';

export const THEME_FONTS = ['Arial', 'Calibri', 'Inter', 'Roboto', 'Lato', 'Montserrat', 'Open Sans', 'Poppins', 'Nunito', 'Source Sans 3', 'Merriweather', 'Playfair Display', 'Georgia', 'Times New Roman', 'Roboto Mono', 'Courier New', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Comic Sans MS'];
const GOOGLE = new Set(['Inter', 'Roboto', 'Lato', 'Montserrat', 'Open Sans', 'Poppins', 'Nunito', 'Source Sans 3', 'Merriweather', 'Playfair Display', 'Roboto Mono']);
const loadedFonts = new Set(['Inter']);
export function ensureFont(name, app) {
  if (!GOOGLE.has(name) || loadedFonts.has(name)) return;
  loadedFonts.add(name);
  const link = el('link', { rel: 'stylesheet', href: `https://fonts.googleapis.com/css2?family=${encodeURIComponent(name).replace(/%20/g, '+')}:ital,wght@0,400;0,700;1,400;1,700&display=swap` });
  document.head.append(link);
  link.onload = () => document.fonts.load(`12px "${name}"`).then(() => { app.grid.measureCache.clear(); app.grid.render(); });
}

const SIZES = [6, 7, 8, 9, 10, 11, 12, 14, 18, 24, 36, 48, 72];
const BORDER_TYPES = [
  ['all', 'border_all', 'Semua batas'], ['inner', 'border_inner', 'Batas dalam'], ['horizontal', 'border_horizontal', 'Batas horizontal'], ['vertical', 'border_vertical', 'Batas vertikal'], ['outer', 'border_outer', 'Batas luar'],
  ['left', 'border_left', 'Batas kiri'], ['top', 'border_top', 'Batas atas'], ['right', 'border_right', 'Batas kanan'], ['bottom', 'border_bottom', 'Batas bawah'], ['clear', 'border_clear', 'Hapus batas'],
];

export function buildToolbar(app, host) {
  const M = app.MOD;
  const refs = {};
  const btn = (ic, tip, run, opts = {}) => {
    const b = el('button.tb-btn' + (opts.cls ? '.' + opts.cls : ''), { type: 'button', 'data-tip': tip, 'aria-label': tip.split('|')[0], onmousedown: e => e.preventDefault(), onclick: e => run(e.currentTarget, e) },
      typeof ic === 'string' ? icon(ic) : ic, opts.caret ? icon('arrow_drop_down', 'caret') : null);
    if (opts.ref) refs[opts.ref] = b;
    return b;
  };
  const groups = () => {
    const sel = () => app.sheet.styleOf(app.grid.sel.r, app.grid.sel.c);
    const colorBtn = (ic, tip, prop) => {
      const bar = el('span.color-bar');
      const b = btn(el('span.ic-stack', icon(ic), bar), tip, a => pickColor(a.getBoundingClientRect(), c => app.cmd(prop === 'color' ? 'color' : 'fill', c), { value: sel()[prop] || undefined }), { ref: prop + 'Btn', caret: true });
      b._bar = bar;
      return b;
    };
    const fontSel = el('button.tb-btn.tb-font', { type: 'button', 'data-tip': 'Font', onmousedown: e => e.preventDefault(), onclick: e => openMenu([
      { head: 'Font' },
      ...THEME_FONTS.map(f => ({ label: f, node: null, checked: () => (sel().ff || app.defFont.ff) === f, run: () => { ensureFont(f, app); app.cmd('font', f); } })),
    ].map(it => (it.label && it.head === undefined ? { ...it, label: it.label } : it)), e.currentTarget.getBoundingClientRect()) }, el('span.tb-font-name', 'Arial'), icon('arrow_drop_down', 'caret'));
    refs.font = fontSel;
    const sizeInp = el('input.tb-size', { type: 'text', inputMode: 'numeric', value: '10', 'aria-label': 'Ukuran font' });
    sizeInp.addEventListener('keydown', e => { if (e.key === 'Enter') { const v = parseFloat(sizeInp.value); if (v > 0 && v < 400) app.cmd('fontSize', v); app.grid.focus(); } if (e.key === 'Escape') app.grid.focus(); });
    sizeInp.addEventListener('focus', () => { sizeInp.select(); openMenu(SIZES.map(s => ({ label: String(s), run: () => app.cmd('fontSize', s) })), sizeInp.getBoundingClientRect()); });
    refs.size = sizeInp;
    const zoomBtn = el('button.tb-btn.tb-zoom', { type: 'button', 'data-tip': 'Zoom', onmousedown: e => e.preventDefault(), onclick: e => openMenu(['Sesuaikan', 50, 75, 90, 100, 125, 150, 200].map(z => (z === 'Sesuaikan' ? { label: 'Sesuaikan lebar', run: () => fitZoom(app) } : { label: z + '%', checked: () => Math.round(app.grid.z * 100) === z, run: () => app.setZoom(z / 100) })), e.currentTarget.getBoundingClientRect()) }, el('span.tb-zoom-val', '100%'), icon('arrow_drop_down', 'caret'));
    refs.zoom = zoomBtn;
    return [
      { id: 'hist', pri: 0, items: [
        btn('undo', `Urungkan|${M}+Z`, () => app.undo(), { ref: 'undo' }),
        btn('redo', `Ulangi|${M}+Y`, () => app.redo(), { ref: 'redo' }),
        btn('print', `Cetak|${M}+P`, () => io.printView(app)),
        btn('format_paint', 'Format cat', () => app.cmd('paintFormat'), { ref: 'paint' }),
      ] },
      { id: 'zoom', pri: 5, items: [zoomBtn] },
      { id: 'num', pri: 1, items: [
        btn(el('span.tb-text', L.cur), 'Format sebagai mata uang', () => app.cmd('numberFormat', '"' + L.cur + '"#,##0.00')),
        btn('percent', 'Format sebagai persen', () => app.cmd('numberFormat', '0.00%')),
        btn(el('span.tb-text.dec', '.0', icon('arrow_back', 'tiny')), 'Kurangi tempat desimal', () => app.cmd('decimals', -1)),
        btn(el('span.tb-text.dec', '.00', icon('arrow_forward', 'tiny')), 'Tambah tempat desimal', () => app.cmd('decimals', 1)),
        btn(el('span.tb-text', '123'), 'Format lainnya', a => openMenu(numberFormatItems(app), a.getBoundingClientRect()), { caret: true }),
      ] },
      { id: 'font', pri: 2, items: [fontSel] },
      { id: 'size', pri: 2, items: [
        btn('remove', 'Kurangi ukuran font', () => { const s = sel().fs || app.defFont.fs; app.cmd('fontSize', Math.max(1, s - 1)); }, { cls: 'sm' }),
        sizeInp,
        btn('add', 'Tambah ukuran font', () => { const s = sel().fs || app.defFont.fs; app.cmd('fontSize', s + 1); }, { cls: 'sm' }),
      ] },
      { id: 'text', pri: 0, items: [
        btn('format_bold', `Tebal|${M}+B`, () => app.cmd('toggleStyle', 'b'), { ref: 'b' }),
        btn('format_italic', `Miring|${M}+I`, () => app.cmd('toggleStyle', 'i'), { ref: 'i' }),
        btn('strikethrough_s', 'Coret|Alt+Shift+5', () => app.cmd('toggleStyle', 's'), { ref: 's' }),
        btn('format_underlined', `Garis bawah|${M}+U`, () => app.cmd('toggleStyle', 'u'), { ref: 'u' }),
        colorBtn('format_color_text', 'Warna teks', 'color'),
      ] },
      { id: 'cell', pri: 1, items: [
        colorBtn('format_color_fill', 'Warna isian', 'bg'),
        btn('border_all', 'Batas', a => bordersPopup(app, a), { caret: true }),
        btn('cell_merge', 'Gabungkan sel', () => app.cmd('merge', 'all'), { ref: 'merge' }),
        btn('arrow_drop_down', 'Jenis penggabungan', a => openMenu([
          { label: 'Gabungkan semua', run: () => app.cmd('merge', 'all') },
          { label: 'Gabungkan secara vertikal', run: () => app.cmd('merge', 'v') },
          { label: 'Gabungkan secara horizontal', run: () => app.cmd('merge', 'h') },
          { label: 'Pisahkan', run: () => app.cmd('merge', 'none') },
        ], a.getBoundingClientRect()), { cls: 'narrow' }),
      ] },
      { id: 'align', pri: 1, items: [
        btn('format_align_left', 'Perataan horizontal', a => iconMenu(a, [['left', 'format_align_left', 'Kiri'], ['center', 'format_align_center', 'Tengah'], ['right', 'format_align_right', 'Kanan']], v => app.cmd('align', v), sel().ha || 'left'), { caret: true, ref: 'ha' }),
        btn('vertical_align_bottom', 'Perataan vertikal', a => iconMenu(a, [['top', 'vertical_align_top', 'Atas'], ['middle', 'vertical_align_center', 'Tengah'], ['bottom', 'vertical_align_bottom', 'Bawah']], v => app.cmd('valign', v), sel().va || 'bottom'), { caret: true, ref: 'va' }),
        btn('wrap_text', 'Pengemasan teks', a => iconMenu(a, [['overflow', 'format_overline', 'Luapan'], ['wrap', 'wrap_text', 'Kemas'], ['clip', 'content_cut', 'Potong']], v => app.cmd('wrap', v), sel().wrap || 'overflow'), { caret: true, ref: 'wrap' }),
        btn('text_rotate_up', 'Rotasi teks', a => iconMenu(a, [[0, 'text_rotation_none', 'Tidak ada'], [45, 'text_rotate_up', 'Miring ke atas'], [-45, 'text_rotation_down', 'Miring ke bawah'], ['v', 'text_rotate_vertical', 'Vertikal'], [90, 'text_rotation_angleup', 'Putar ke atas'], [-90, 'text_rotation_angledown', 'Putar ke bawah']], v => app.cmd('rotate', v), sel().rot || 0), { caret: true }),
      ] },
      { id: 'insert', pri: 3, items: [
        btn('link', `Sisipkan link|${M}+K`, () => D.linkDialog(app)),
        btn('add_comment', 'Sisipkan catatan|Shift+F2', () => D.noteDialog(app)),
        btn('insert_chart', 'Sisipkan diagram', () => app.sidebar.open('chart', { create: true })),
        btn('signature', 'Sisipkan tanda tangan', () => app.images.insertDialog()),
        btn('filter_alt', `Buat filter|${M}+Shift+F`, () => app.cmd('toggleFilter'), { ref: 'filter' }),
        btn('table_chart', 'Warna bergantian / Format sebagai tabel', () => app.sidebar.open('banding')),
        btn('functions', 'Fungsi', a => openMenu(functionItems(app), a.getBoundingClientRect()), { caret: true }),
      ] },
      { id: 'excel', pri: 4, items: [
        btn('format_color_fill', 'Format bersyarat', () => app.sidebar.open('cf'), { cls: 'cf' }),
        btn('rule', 'Validasi data', () => app.sidebar.open('dv')),
        btn('pivot_table_chart', 'Tabel pivot', () => D.pivotDialog(app)),
        btn('search', `Cari dan ganti|${M}+H`, () => D.replaceDialog(app, true)),
      ] },
    ];
  };

  const render = () => {
    host.innerHTML = '';
    Object.keys(refs).forEach(k => delete refs[k]);
    const gs = groups();
    const wrap = el('div.tb-groups');
    for (const g of gs) {
      const ge = el('div.tb-group', { dataset: { id: g.id, pri: g.pri } }, ...g.items);
      wrap.append(ge, el('div.tb-sep', { dataset: { for: g.id } }));
    }
    const more = el('button.tb-btn.tb-more', { type: 'button', 'data-tip': 'Lainnya', onmousedown: e => e.preventDefault(), onclick: e => {
      const hidden = [...wrap.querySelectorAll('.tb-group.ovf')].map(x => x.dataset.id);
      const saved = { ...refs };
      const extra = groups().filter(g => hidden.includes(g.id));
      Object.assign(refs, saved);
      const box = el('div.tb-overflow', ...extra.map(g => el('div.tb-group', ...g.items)));
      popup(box, e.currentTarget.getBoundingClientRect(), { alignRight: true });
    } }, icon('more_vert'));
    const right = el('div.tb-right', btn('keyboard_arrow_up', 'Sembunyikan menu', () => { document.body.classList.toggle('compact'); app.grid.resize(); }, { cls: 'collapse' }));
    host.append(wrap, more, right);
    fit();
  };
  const fit = () => {
    const wrap = host.querySelector('.tb-groups'), more = host.querySelector('.tb-more');
    if (!wrap) return;
    const gs = [...wrap.querySelectorAll('.tb-group')];
    gs.forEach(g => { g.classList.remove('ovf'); wrap.querySelector(`.tb-sep[data-for="${g.dataset.id}"]`).classList.remove('ovf'); });
    more.style.display = 'none';
    const avail = host.clientWidth - 84;
    const used = () => { let w = 0; for (const ch of wrap.children) if (ch.offsetWidth) w += ch.offsetWidth + 12 * (ch.classList.contains('tb-sep') ? 1 : 0); return w; };
    const order = [...gs].sort((a, b) => b.dataset.pri - a.dataset.pri || gs.indexOf(b) - gs.indexOf(a));
    let i = 0;
    while (used() > avail && i < order.length - 1) {
      order[i].classList.add('ovf');
      wrap.querySelector(`.tb-sep[data-for="${order[i].dataset.id}"]`).classList.add('ovf');
      more.style.display = '';
      i++;
    }
  };
  new ResizeObserver(() => fit()).observe(host);
  render();

  const api = {
    sync() {
      if (!app.wb) return;
      const sh = app.sheet, g = app.grid;
      const st = sh.styleOf(g.sel.r, g.sel.c);
      for (const k of ['b', 'i', 's', 'u']) refs[k]?.classList.toggle('on', !!st[k]);
      refs.undo?.toggleAttribute('disabled', !app.wb.hist.undoS.length);
      refs.redo?.toggleAttribute('disabled', !app.wb.hist.redoS.length);
      refs.filter?.classList.toggle('on', !!sh.filter);
      refs.paint?.classList.toggle('on', !!app.painter);
      refs.merge?.classList.toggle('on', !!sh.mergeAt(g.sel.r, g.sel.c));
      if (refs.font) refs.font.querySelector('.tb-font-name').textContent = st.ff || app.defFont.ff;
      if (refs.size && document.activeElement !== refs.size) refs.size.value = String(st.fs || app.defFont.fs);
      if (refs.zoom) refs.zoom.querySelector('.tb-zoom-val').textContent = Math.round(g.z * 100) + '%';
      if (refs.colorBtn) refs.colorBtn._bar.style.background = st.color || 'var(--text)';
      if (refs.bgBtn) refs.bgBtn._bar.style.background = st.bg || 'transparent';
      const ha = { left: 'format_align_left', center: 'format_align_center', right: 'format_align_right' }[st.ha || 'left'];
      if (refs.ha) refs.ha.querySelector('.ms').textContent = ha;
      const va = { top: 'vertical_align_top', middle: 'vertical_align_center', bottom: 'vertical_align_bottom' }[st.va || 'bottom'];
      if (refs.va) refs.va.querySelector('.ms').textContent = va;
      const wr = { wrap: 'wrap_text', clip: 'content_cut', overflow: 'format_overline' }[st.wrap || 'overflow'];
      if (refs.wrap) refs.wrap.querySelector('.ms').textContent = wr;
      if (st.ff) ensureFont(st.ff, app);
    },
    render,
  };
  return api;
}

function iconMenu(anchor, opts, onPick, current) {
  const box = el('div.icon-menu', ...opts.map(([v, ic, label]) => el('button.tb-btn' + (v === current ? '.on' : ''), { type: 'button', 'data-tip': label, onclick: () => { closeMenus(); onPick(v); } }, icon(ic))));
  popup(box, anchor.getBoundingClientRect());
}

function bordersPopup(app, anchor) {
  let color = app.lastBorder?.color || '#000000', style = app.lastBorder?.style || 'solid', w = app.lastBorder?.w || 1;
  const grid = el('div.border-grid', ...BORDER_TYPES.map(([t, ic, label]) => el('button.tb-btn', { type: 'button', 'data-tip': label, onclick: () => { closeMenus(); app.cmd('borders', t, { color, style, w }); } }, icon(ic))));
  const colorBtn = el('button.tb-btn', { type: 'button', 'data-tip': 'Warna batas', onclick: () => {
    const inner = colorGrid(c => { color = c || '#000000'; bar.style.background = color; sub.remove(); }, { reset: false });
    const sub = el('div.menu.glass.sub-pop', inner);
    document.getElementById('popups').append(sub);
    const r = colorBtn.getBoundingClientRect();
    sub.style.left = r.right + 4 + 'px'; sub.style.top = r.top + 'px';
    const off = e => { if (!sub.contains(e.target)) { sub.remove(); document.removeEventListener('pointerdown', off, true); } };
    setTimeout(() => document.addEventListener('pointerdown', off, true));
  } }, el('span.ic-stack', icon('border_color'), el('span.color-bar')));
  const bar = colorBtn.querySelector('.color-bar');
  bar.style.background = color;
  const styles = [['solid', 1, 'Garis tipis'], ['solid', 2, 'Garis sedang'], ['solid', 3, 'Garis tebal'], ['dashed', 1, 'Putus-putus'], ['dotted', 1, 'Titik-titik'], ['double', 3, 'Garis ganda']];
  const styleBtn = el('button.tb-btn', { type: 'button', 'data-tip': 'Gaya batas', onclick: () => {
    const list = el('div.border-styles', ...styles.map(([s, ww, label]) => el('button.bs-item' + (s === style && ww === w ? '.on' : ''), { type: 'button', title: label, onclick: () => { style = s; w = ww; list.remove(); } }, el('span.bs-line', { style: { borderTop: `${s === 'double' ? 3 : ww}px ${s} currentColor` } }))));
    const sub = el('div.menu.glass.sub-pop', list);
    document.getElementById('popups').append(sub);
    const r = styleBtn.getBoundingClientRect();
    sub.style.left = r.right + 4 + 'px'; sub.style.top = r.top + 'px';
    const off = e => { if (!sub.contains(e.target)) { sub.remove(); document.removeEventListener('pointerdown', off, true); } };
    setTimeout(() => document.addEventListener('pointerdown', off, true));
  } }, icon('line_style'));
  popup(el('div.borders-pop', grid, el('div.border-side', colorBtn, styleBtn)), anchor.getBoundingClientRect());
}

function fitZoom(app) {
  const sh = app.sheet, g = app.grid;
  const u = sh.usedRange();
  if (!u) return app.setZoom(1);
  const geo = g.geo();
  const w = geo.hw + geo.colX[u.c2 + 1];
  app.setZoom(Math.max(0.5, Math.min(2, (g.scroller.clientWidth - 20) / w)));
}
