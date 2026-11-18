// Buka spreadsheet: the file list (favourites always on top, colour categories, sort, search), the template
// gallery and Sampah — files moved to the trash stay 30 days (restore, or delete for good / empty the trash).
// Favourite, colour and trash state live in store.js docMeta (localStorage), not in the workbook records.
import { el, icon, uid } from '../util.js';
import { dialog, openMenu, toast, confirmDlg, promptDlg } from './ui.js';
import { TEMPLATES } from '../templates.js';
import * as store from '../store.js';

const DAY = 86400000;
const STAR = '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M12 3.6l2.47 5.2 5.7.68-4.2 3.9 1.1 5.64L12 16.25 6.93 19.02l1.1-5.64-4.2-3.9 5.7-.68z"/></svg>';
export const starEl = (on = false) => el('span.star-ic' + (on ? '.on' : ''), { html: STAR });
const dot = (c, cls = '') => el('span.fb-dot' + (cls ? '.' + cls : ''), { style: { '--fc': c } });

// "Hari ini, 12.36" / "Kemarin, 09.10" / "3 Nov, 08.00" / "28 Des 2025, 17.45"
function when(t) {
  const d = new Date(t), now = new Date();
  const time = d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
  const day0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  if (t >= day0 && t < day0 + DAY) return 'Hari ini, ' + time;
  if (t >= day0 - DAY && t < day0) return 'Kemarin, ' + time;
  const opt = { day: 'numeric', month: 'short' };
  if (d.getFullYear() !== now.getFullYear()) opt.year = 'numeric';
  return d.toLocaleDateString('id-ID', opt) + ', ' + time;
}
const full = t => new Date(t).toLocaleString('id-ID', { dateStyle: 'full', timeStyle: 'short' });
const leftText = n => (n <= 1 ? 'Kurang dari 1 hari' : `${n} hari lagi`);

// ---------- sort (saved) + filter (while the page is open) ----------
const SORTS = ['updated', 'name', 'sheets', 'color'];
function getSort() {
  try { const s = JSON.parse(localStorage.getItem('as.docsort') || 'null'); if (s && SORTS.includes(s.by) && (s.dir === 1 || s.dir === -1)) return s; } catch { /* default */ }
  return { by: 'updated', dir: -1 };
}
function saveSort(s) { try { localStorage.setItem('as.docsort', JSON.stringify(s)); } catch { /* ignore */ } }
let filter = 'all';   // all | fav | <colour id>
const coll = new Intl.Collator('id', { numeric: true, sensitivity: 'base' });
const rank = x => { const i = store.LABELS.findIndex(L => L.id === x.color); return i < 0 ? 99 : i; };
function compare(sort) {
  return (a, b) => {
    let r;
    if (sort.by === 'name') r = coll.compare(a.name, b.name) * sort.dir;
    else if (sort.by === 'sheets') r = (a.sheets - b.sheets) * sort.dir;
    else if (sort.by === 'color') { const ra = rank(a), rb = rank(b); r = ra === 99 || rb === 99 ? ra - rb : (ra - rb) * sort.dir; }   // files without a colour stay last
    else r = (a.updated - b.updated) * sort.dir;
    return r || b.updated - a.updated;
  };
}

// ---------- colour category menu (file list rows and File › Kategori warna) ----------
export function colorItems(cur, pick, after) {
  const m = store.readMeta();
  return [
    { head: 'Kategori warna' },
    ...store.LABELS.map(L => ({ label: store.labelName(L.id, m), swatch: L.c, round: true, preview: cur === L.id ? icon('check') : null, run: () => pick(L.id) })),
    { sep: true },
    { label: 'Tanpa kategori', icon: 'format_color_reset', disabled: !cur, run: () => pick(null) },
    { label: 'Ubah nama kategori…', icon: 'edit', run: () => labelsDialog(after) },
  ];
}

export function labelsDialog(onDone) {
  const m = store.readMeta();
  const inputs = {};
  const rows = store.LABELS.map(L => {
    const inp = el('input.inp', { value: store.labelName(L.id, m), placeholder: L.name, maxLength: 40, 'aria-label': 'Nama kategori ' + L.name });
    inputs[L.id] = inp;
    return el('label.fb-lab', dot(L.c, 'lg'), inp);
  });
  dialog({
    title: 'Nama kategori warna', width: 400,
    body: el('div.stack.sm', el('p.dlg-text.small', 'Beri nama setiap warna, misalnya Kantor, Pribadi, Proyek, atau Penting.'), ...rows),
    buttons: [
      { label: 'Nama bawaan', run: () => { for (const L of store.LABELS) inputs[L.id].value = L.name; return false; } },
      { spacer: true },
      { label: 'Batal' },
      { label: 'Simpan', primary: true, run: () => { store.setLabelNames(Object.fromEntries(Object.entries(inputs).map(([k, i]) => [k, i.value]))); onDone?.(); } },
    ],
  });
}

// ---------- the dialog ----------
export async function openDialog(app, tab = 'recent') {
  await store.purgeTrash().catch(() => 0);
  let docs = await store.listDocs();
  let sort = getSort();
  let d = null;
  const curId = () => app.wb?.id;
  const { Workbook } = await import('../model.js');

  const body = el('div.home.fb');
  const tabs = el('div.tabs-h', { role: 'tablist' });
  const pages = {};
  const tBadge = el('span.fb-badge');
  const tabBtn = (id, label, extra) => el('button' + (id === 'trash' ? '.fb-trash-tab' : ''), { type: 'button', role: 'tab', dataset: { id }, onclick: () => show(id) }, label, extra || null);
  tabs.append(tabBtn('recent', 'Terbaru'), tabBtn('templates', 'Template'), tabBtn('trash', [icon('delete'), 'Sampah'], tBadge));
  const show = id => {
    tabs.querySelectorAll('button').forEach(b => { const on = b.dataset.id === id; b.classList.toggle('on', on); b.setAttribute('aria-selected', String(on)); });
    for (const [k, p] of Object.entries(pages)) p.style.display = k === id ? '' : 'none';
  };

  // ----- actions -----
  const open = async x => {
    d.close();
    if (x.id === curId()) return;
    await app.saveBeforeSwitch();
    const j = await store.loadDoc(x.id);
    if (j) app.openWorkbook(Workbook.fromJSON(j));
    else toast('Spreadsheet tidak ditemukan di perangkat ini');
  };
  const toggleFav = x => {
    x.fav = !x.fav;
    store.setDocMeta(x.id, { f: x.fav ? 1 : null });
    drawAll(x.id, '.fb-star');
  };
  const setColor = (x, c) => { x.color = c; store.setDocMeta(x.id, { c }); drawAll(x.id); };
  const rename = async x => {
    const v = (await promptDlg('Ganti nama', 'Nama spreadsheet', x.name, { ok: 'Simpan' }))?.trim();
    if (!v || v === x.name) return;
    if (x.id === curId()) app.renameDoc(v); else await store.renameDoc(x.id, v);
    x.name = v;
    drawAll(x.id);
  };
  const copy = async x => {
    const j = x.id === curId() ? { ...app.wb.toJSON(), lock: app.wb.lock || null } : await store.loadDoc(x.id);
    if (!j) return;
    const wb = Workbook.fromJSON({ ...j, share: null });
    wb.id = uid('wb'); wb.name = 'Salinan dari ' + x.name; wb.share = null; wb.lock = j.lock || null;
    await store.saveDoc(wb);
    if (x.color) store.setDocMeta(wb.id, { c: x.color });
    docs = await store.listDocs();
    drawAll(wb.id);
    toast(`Salinan "${x.name}" dibuat`);
  };
  const trash = async x => {
    const wasOpen = x.id === curId();
    if (wasOpen) await app.save();
    store.trashDoc(x.id);
    x.trashed = Date.now();
    if (wasOpen) await app.openNewest(x.id);
    drawAll();
    toast(`"${x.name}" dipindahkan ke sampah`, { ms: 6000, action: { label: 'Urungkan', run: async () => {
      store.restoreDoc(x.id); x.trashed = null;
      if (wasOpen && curId() !== x.id) await app.openDocById(x.id);
      if (d.el.isConnected) drawAll(x.id);
    } } });
  };
  const restore = async (x, thenOpen = false) => {
    store.restoreDoc(x.id); x.trashed = null;
    if (thenOpen) return open(x);
    drawAll();
    toast(`"${x.name}" dipulihkan`);
  };
  const forget = async list => {   // delete for good; the open workbook moves to another file first
    if (list.some(x => x.id === curId())) await app.openNewest(list.map(x => x.id));
    for (const x of list) await store.deleteForever(x.id);
    const gone = new Set(list.map(x => x.id));
    docs = docs.filter(x => !gone.has(x.id));
    drawAll();
  };
  const deleteForever = async x => {
    if (!(await confirmDlg('Hapus permanen', `Hapus "${x.name}" untuk selamanya? Spreadsheet ini tidak dapat dipulihkan lagi.`, 'Hapus permanen', true))) return;
    await forget([x]);
    toast('Spreadsheet dihapus permanen');
  };
  const emptyTrash = async () => {
    const list = docs.filter(x => x.trashed);
    if (!list.length) return;
    if (!(await confirmDlg('Kosongkan sampah', `Hapus permanen ${list.length} spreadsheet di sampah? Tindakan ini tidak dapat diurungkan.`, 'Kosongkan sampah', true))) return;
    await forget(list);
    toast('Sampah dikosongkan');
  };

  const rowMenu = x => [
    { label: 'Buka', icon: 'open_in_browser', run: () => open(x) },
    { sep: true },
    { label: x.fav ? 'Hapus dari favorit' : 'Tambahkan ke favorit', icon: 'star', run: () => toggleFav(x) },
    { label: 'Kategori warna', icon: 'palette', sub: () => colorItems(x.color, c => setColor(x, c), () => drawAll(x.id)) },
    { label: 'Ganti nama…', icon: 'drive_file_rename_outline', run: () => rename(x) },
    { label: 'Buat salinan', icon: 'file_copy', run: () => copy(x) },
    { sep: true },
    { label: 'Pindahkan ke sampah', icon: 'delete', run: () => trash(x) },
  ];
  const trashMenu = x => [
    { label: 'Pulihkan', icon: 'restore_from_trash', run: () => restore(x) },
    { label: 'Pulihkan dan buka', icon: 'open_in_browser', run: () => restore(x, true) },
    { sep: true },
    { label: 'Hapus permanen', icon: 'delete_forever', run: () => deleteForever(x) },
  ];

  // rows are focusable: ↑/↓ move, Enter opens, Delete moves to the trash (or deletes for good there)
  const keyNav = (r, onEnter, onDelete) => r.addEventListener('keydown', e => {
    if (e.target !== r) return;
    if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); onEnter(); }
    else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); e.stopPropagation(); onDelete(); }
    else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const rows = [...r.parentNode.querySelectorAll('.fb-row')], i = rows.indexOf(r);
      rows[i + (e.key === 'ArrowDown' ? 1 : -1)]?.focus();
    }
  });
  const refocus = (host, id, sel, had) => {
    if (!id) return;
    const r = host.querySelector(`.fb-row[data-id="${CSS.escape(id)}"]`);
    if (!r) return;
    r.scrollIntoView({ block: 'nearest' });
    if (had) (sel && r.querySelector(sel) || r).focus({ preventScroll: true });
  };
  const focusIn = host => host.contains(document.activeElement) || document.activeElement === document.body;

  // ----- Terbaru: list view -----
  const search = el('input.inp', { type: 'search', placeholder: 'Telusuri spreadsheet', 'aria-label': 'Telusuri spreadsheet', autocomplete: 'off' });
  const chips = el('div.fb-chips');
  const head = el('div.fb-head', { role: 'row' });
  const list = el('div.fb-list', { role: 'list' });
  const sortBtn = (by, label, dir0, short) => {
    const on = sort.by === by;
    return el('button.fb-sort' + (on ? '.on' : '') + (short ? '.has-s' : ''), { type: 'button', 'data-tip': 'Urutkan menurut ' + label.toLowerCase(),
      onclick: () => { sort = on ? { by, dir: -sort.dir } : { by, dir: dir0 }; saveSort(sort); drawRecent(); } },
    el('span.fb-l', label), short ? el('span.fb-s', short) : null, on ? icon(sort.dir > 0 ? 'arrow_upward' : 'arrow_downward') : null);
  };
  const catBtn = (x, meta) => {
    const col = store.labelColor(x.color);
    return el('button.fb-catbtn' + (col ? '.set' : ''), { type: 'button', style: col ? { '--fc': col } : null, 'data-tip': col ? 'Ganti kategori warna' : 'Beri kategori warna',
      onclick: e => { e.stopPropagation(); openMenu(colorItems(x.color, c => setColor(x, c), () => drawAll(x.id)), e.currentTarget.getBoundingClientRect()); } },
    col ? dot(col) : icon('add'), el('span.lbl', col ? store.labelName(x.color, meta) : 'Kategori'));
  };
  const row = (x, meta) => {
    const isCur = x.id === curId(), col = store.labelColor(x.color);
    const favTip = x.fav ? 'Hapus dari favorit' : 'Tambahkan ke favorit';
    const r = el('div.fb-row' + (isCur ? '.current' : ''), { role: 'listitem', tabIndex: 0, dataset: { id: x.id }, style: col ? { '--fc': col } : null, 'aria-label': x.name },
      el('button.fb-star' + (x.fav ? '.on' : ''), { type: 'button', 'aria-pressed': String(x.fav), 'aria-label': favTip, 'data-tip': favTip, onclick: e => { e.stopPropagation(); toggleFav(x); } }, starEl(x.fav)),
      el('span.fb-ic' + (col ? '.c' : ''), icon(x.locked ? 'lock' : 'grid_on')),
      el('div.fb-name',
        el('div.fb-title', el('span.fb-tt', { title: x.name }, x.name), isCur ? el('span.fb-tag', 'Terbuka') : null, x.share ? el('span.fb-tag.sh', { 'data-tip': 'Dibagikan dengan tautan' }, icon('group')) : null),
        el('div.fb-sub', col ? el('span.fb-subcat', dot(col), store.labelName(x.color, meta)) : null, el('span', when(x.updated)), el('span', x.sheets + ' sheet'))),
      el('div.fb-cat', catBtn(x, meta)),
      el('div.fb-date', { title: full(x.updated) }, when(x.updated)),
      el('div.fb-n', String(x.sheets)),
      el('div.fb-act',
        el('button.icon-btn.sm.fb-del', { type: 'button', 'data-tip': 'Pindahkan ke sampah', 'aria-label': 'Pindahkan ke sampah', onclick: e => { e.stopPropagation(); trash(x); } }, icon('delete')),
        el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Lainnya', 'aria-label': 'Tindakan lainnya', onclick: e => { e.stopPropagation(); openMenu(rowMenu(x), e.currentTarget.getBoundingClientRect(), { alignRight: true }); } }, icon('more_vert'))));
    r.addEventListener('click', () => open(x));
    r.addEventListener('contextmenu', e => { e.preventDefault(); openMenu(rowMenu(x), { x: e.clientX, y: e.clientY }); });
    keyNav(r, () => open(x), () => trash(x));
    return r;
  };
  const drawRecent = (focusId, focusSel) => {
    const meta = store.readMeta();
    const live = docs.filter(x => !x.trashed);
    const nFav = live.filter(x => x.fav).length;
    const used = store.LABELS.filter(L => live.some(x => x.color === L.id));
    if ((filter === 'fav' && !nFav) || (filter !== 'all' && filter !== 'fav' && !used.some(L => L.id === filter))) filter = 'all';
    // filter chips: Semua · Favorit · each colour in use
    const chip = (id, kids, n) => el('button.fb-chip' + (filter === id ? '.on' : ''), { type: 'button', 'aria-pressed': String(filter === id), onclick: () => { filter = filter === id ? 'all' : id; drawRecent(); } }, ...kids, el('span.n', String(n)));
    chips.replaceChildren(chip('all', [el('span', 'Semua')], live.length));
    if (nFav) chips.append(chip('fav', [starEl(true), el('span', 'Favorit')], nFav));
    for (const L of used) chips.append(chip(L.id, [dot(L.c), el('span', store.labelName(L.id, meta))], live.filter(x => x.color === L.id).length));
    chips.append(el('button.icon-btn.sm.fb-chip-edit', { type: 'button', 'data-tip': 'Ubah nama kategori warna', 'aria-label': 'Ubah nama kategori warna', onclick: () => labelsDialog(() => drawAll()) }, icon('edit')));
    chips.hidden = !nFav && !used.length;
    head.replaceChildren(el('span'), el('span'), el('div.h-name', sortBtn('name', 'Nama', 1)), el('div.h-cat', sortBtn('color', 'Kategori', 1)), el('div.h-date', sortBtn('updated', 'Terakhir diubah', -1, 'Diubah')), el('div.h-n', sortBtn('sheets', 'Sheet', -1)), el('span.h-act'));
    const q = search.value.trim().toLowerCase();
    const items = live.filter(x => (!q || x.name.toLowerCase().includes(q)) && (filter === 'all' || (filter === 'fav' ? x.fav : x.color === filter))).sort(compare(sort));
    const favs = items.filter(x => x.fav), rest = items.filter(x => !x.fav);
    const had = focusIn(list);
    list.replaceChildren();
    if (!items.length) {
      list.append(live.length
        ? el('div.empty.sm', icon('search_off'), el('span', 'Tidak ada spreadsheet yang cocok'))
        : el('div.empty', icon('folder_off'), el('span', 'Belum ada spreadsheet tersimpan')));
    }
    if (favs.length && filter !== 'fav') list.append(el('div.fb-sec', starEl(true), 'Favorit'));
    for (const x of favs) list.append(row(x, meta));
    if (favs.length && rest.length) list.append(el('div.fb-sec', 'Spreadsheet lainnya'));
    for (const x of rest) list.append(row(x, meta));
    refocus(list, focusId, focusSel, had);
  };
  search.addEventListener('input', () => drawRecent());
  search.addEventListener('keydown', e => {
    if (e.key === 'ArrowDown') { e.preventDefault(); list.querySelector('.fb-row')?.focus(); }
    else if (e.key === 'Enter') { const r = list.querySelector('.fb-row'); if (r) { e.preventDefault(); r.click(); } }
  });
  pages.recent = el('div.stack.fb-page',
    el('div.fb-bar', el('div.fb-search', icon('search'), search),
      el('button.btn', { type: 'button', onclick: async () => { d.close(); (await import('./dialogs.js')).importDialog(app); } }, icon('upload_file'), el('span.fb-btn-t', 'Impor file'))),
    chips,
    el('div.fb-table', head, list));

  // ----- Template -----
  pages.templates = el('div.tpl-grid', ...TEMPLATES.filter(t => t.id !== 'welcome').map(t => el('button.tpl', { type: 'button', onclick: () => { d.close(); app.cmd('newDoc', t.id); } },
    el('div.tpl-thumb', { dataset: { kind: t.id } }, icon(t.icon)), el('div.tpl-name', t.name), el('div.muted.small', t.desc))));

  // ----- Sampah -----
  const emptyBtn = el('button.btn.fb-empty', { type: 'button', onclick: emptyTrash }, icon('delete_forever'), el('span', 'Kosongkan sampah'));
  const tList = el('div.fb-list', { role: 'list' });
  const tHead = el('div.fb-head', { role: 'row' }, el('span'), el('div.h-name', 'Nama'), el('div.h-date', 'Dihapus'), el('div.h-left', 'Hapus permanen'), el('span.h-act'));
  const trow = x => {
    const left = store.trashDaysLeft(x.trashed), soon = left <= 3, col = store.labelColor(x.color);
    const r = el('div.fb-row' + (x.id === curId() ? '.current' : ''), { role: 'listitem', tabIndex: 0, dataset: { id: x.id }, style: col ? { '--fc': col } : null, 'aria-label': x.name },
      el('span.fb-ic.gone', icon(x.locked ? 'lock' : 'grid_on')),
      el('div.fb-name',
        el('div.fb-title', el('span.fb-tt', { title: x.name }, x.name), x.id === curId() ? el('span.fb-tag', 'Terbuka') : null),
        el('div.fb-sub', el('span', 'Dihapus ' + when(x.trashed)), el('span' + (soon ? '.soon' : ''), 'hapus permanen ' + leftText(left).toLowerCase()))),
      el('div.fb-date', { title: full(x.trashed) }, when(x.trashed)),
      el('div.fb-left' + (soon ? '.soon' : ''), { title: 'Dihapus permanen ' + full(x.trashed + store.TRASH_DAYS * DAY) }, leftText(left)),
      el('div.fb-act',
        el('button.fb-restore', { type: 'button', 'data-tip': 'Kembalikan ke daftar spreadsheet', 'aria-label': 'Pulihkan', onclick: e => { e.stopPropagation(); restore(x); } }, icon('restore_from_trash'), el('span', 'Pulihkan')),
        el('button.icon-btn.sm.fb-del', { type: 'button', 'data-tip': 'Hapus permanen', 'aria-label': 'Hapus permanen', onclick: e => { e.stopPropagation(); deleteForever(x); } }, icon('delete_forever'))));
    const menu = at => openMenu(trashMenu(x), at);
    r.addEventListener('click', e => menu({ x: e.clientX, y: e.clientY }));
    r.addEventListener('contextmenu', e => { e.preventDefault(); menu({ x: e.clientX, y: e.clientY }); });
    keyNav(r, () => menu(r.getBoundingClientRect()), () => deleteForever(x));
    return r;
  };
  const drawTrash = focusId => {
    const items = docs.filter(x => x.trashed).sort((a, b) => b.trashed - a.trashed);
    tBadge.textContent = items.length ? String(items.length) : '';
    tBadge.hidden = !items.length;
    emptyBtn.disabled = !items.length;
    const had = focusIn(tList);
    tList.replaceChildren(...(items.length ? items.map(trow)
      : [el('div.empty', icon('delete'), el('span', 'Sampah kosong'), el('span.small', `Spreadsheet yang dihapus disimpan di sini selama ${store.TRASH_DAYS} hari.`))]));
    tHead.hidden = !items.length;
    refocus(tList, focusId, null, had);
  };
  pages.trash = el('div.stack.fb-page',
    el('div.fb-note', icon('auto_delete'), el('span', `Spreadsheet di sampah dihapus permanen otomatis setelah ${store.TRASH_DAYS} hari. Pulihkan untuk memakainya lagi.`), emptyBtn),
    el('div.fb-table.trash', tHead, tList));

  const drawAll = (focusId, focusSel) => { drawRecent(focusId, focusSel); drawTrash(focusId); };
  drawAll();
  body.append(tabs, pages.recent, pages.templates, pages.trash);
  show(tab);
  d = dialog({ title: 'Buka spreadsheet', width: 860, cls: 'dlg-home', body });
}
