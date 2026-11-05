// Sharing and permissions of Amatu Sheets: live share links (see live.js), viewer-editable cells,
// password-locked sheets with unlocked cells, and the transaction guard that enforces them.
import { el, icon, inRange, keyR, keyC, rangeAddr, parseRange } from './util.js';
import { dialog, toast, field, promptDlg, secretInput, secretBox, secretMatch } from './ui/ui.js';
import * as io from './io.js';
import * as store from './store.js';
import { Live, rpc, account, linkFor, errText, merge3, eq } from './live.js';

// ---------- documents on the server ----------
export function payload(wb) { const j = wb.toJSON(); delete j.id; delete j.active; return j; }

// Cells keyed by address with their style written out, so two copies that interned styles differently still compare/merge.
function norm(j) {
  if (!j || !j.sheets) return { order: [], sheets: {} };
  const styles = j.styles || [{}], sheets = {};
  for (const s of j.sheets) {
    const { cells, id, ...meta } = s;
    const cm = {};
    for (const [k, v, f, si, note, link] of cells || []) cm[k] = { v: v ?? null, f: f ?? null, st: JSON.stringify(styles[si] || {}), note: note ?? null, link: link ?? null };
    sheets[id] = { ...meta, cells: cm };
  }
  return { name: j.name, settings: j.settings, names: j.names || [], order: j.sheets.map(s => s.id), sheets };
}
function denorm(o) {
  const styles = [{}], ix = new Map([['{}', 0]]);
  const ids = Object.keys(o.sheets || {});
  const order = (o.order || []).filter(id => o.sheets[id]);
  for (const id of ids) if (!order.includes(id)) order.push(id);
  const sheets = order.map(id => {
    const { cells, ...meta } = o.sheets[id];
    const arr = [];
    for (const [k, c] of Object.entries(cells || {})) {
      let si = ix.get(c.st);
      if (si === undefined) { si = styles.length; styles.push(JSON.parse(c.st)); ix.set(c.st, si); }
      const row = [+k, c.v, c.f, si];
      if (c.note || c.link) row.push(c.note, c.link);
      arr.push(row);
    }
    return { id, ...meta, cells: arr };
  });
  return { app: 'amatu-sheets', v: 1, name: o.name, settings: o.settings, names: o.names || [], styles, sheets };
}
export const sameBooks = (a, b) => eq(norm(a), norm(b));
export const mergeBooks = (b, m, t) => denorm(merge3(norm(b), norm(m), norm(t)));

// ---------- permissions ----------
const inAny = (list, r, c) => (list || []).some(g => inRange(g, r, c));
// Viewer of a "Lihat saja" link may edit only the sheet's viewer cells (the owner's explicit grant, also in a locked sheet).
// Everyone else: a locked sheet allows only its unlocked cells.
export function cellBlock(app, sh, r, c) {
  if (app.readOnly) return app.share?.revoked ? 'Tautan berbagi tidak berlaku lagi.' : inAny(sh.ve, r, c) ? null : 'Anda hanya dapat melihat. Sel ini tidak diizinkan untuk diedit.';
  if (sh.protect && !inAny(sh.protectExcept, r, c)) return sh.lock ? 'Sheet ini dikunci dengan kata sandi. Hanya sel yang dibuka yang dapat diedit.' : 'Sheet ini dilindungi. Buka di Data › Proteksi & izin.';
  return null;
}
const LAYOUT_ONLY = ['rowH', 'manualH', 'filteredR'];
const metaChanged = (a, b) => { const x = { ...a }, y = { ...b }; for (const k of LAYOUT_ONLY) { delete x[k]; delete y[k]; } return !eq(x, y); };
// Checks the history entry of the transaction that just ran. Returns null when allowed, else the reason.
// app._permit = true lets the protection commands themselves through.
export function checkEntry(app, e) {
  if (app._permit) return null;
  const wb = app.wb;
  if (app.readOnly) {
    if (e.wb || e.full.size) return 'Anda hanya dapat melihat spreadsheet ini.';
    for (const [id, m] of e.meta) { const sh = wb.sheetById(id); if (sh && metaChanged(m, sh.metaSnapshot())) return 'Anda hanya dapat melihat spreadsheet ini.'; }
  }
  if (e.wb) for (const id of e.wb.order) { const sh = wb.sheetById(id); const was = e.full.get(id); if (!sh && was?.protect) return 'Sheet yang dikunci tidak dapat dihapus.'; }
  for (const id of e.full.keys()) { const sh = wb.sheetById(id); if (sh?.protect) return 'Sheet ini dikunci. Struktur baris/kolom tidak dapat diubah.'; }
  for (const [id, m] of e.meta) { const sh = wb.sheetById(id); if (sh?.protect && metaChanged(m, sh.metaSnapshot())) return 'Sheet ini dikunci. Pengaturan sheet tidak dapat diubah.'; }
  for (const [id, cm] of e.cells) {
    const sh = wb.sheetById(id); if (!sh) continue;
    for (const [k, snap] of cm) {
      const r = keyR(k), c = keyC(k);
      const why = cellBlock(app, sh, r, c);
      if (why) { const cur = sh.cells.get(k); if (!sameCell(snap, cur)) return why; continue; }
      if (app.readOnly) { const cur = sh.cells.get(k); if ((snap?.s || 0) !== (cur?.s || 0) || (snap?.note || null) !== (cur?.note || null) || (snap?.link || null) !== (cur?.link || null)) return 'Pelihat hanya dapat mengubah isi sel, bukan format.'; }
    }
  }
  return null;
}
const sameCell = (a, b) => (a?.v ?? null) === (b?.v ?? null) && (a?.f ?? null) === (b?.f ?? null) && (a?.s || 0) === (b?.s || 0) && (a?.note || null) === (b?.note || null) && (a?.link || null) === (b?.link || null);
// Cells a viewer's transaction changed, for amt_share_cells.
// all = every touched cell (undo/redo re-sends them), else only the ones that differ from the entry's snapshot.
export function viewerCells(app, e, all = false) {
  const out = [];
  for (const [id, cm] of e.cells) {
    const sh = app.wb.sheetById(id); if (!sh) continue;
    for (const [k, snap] of cm) {
      const cur = sh.cells.get(k);
      if (!all && sameCell(snap, cur)) continue;
      if (all && cellBlock(app, sh, keyR(k), keyC(k))) continue;
      let v = cur?.v ?? null;
      if (v !== null && typeof v === 'object') v = String(v);
      out.push({ s: id, k, v, f: cur?.f ?? null });
    }
  }
  return out;
}

// ---------- password lock ----------
async function sha(text) { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join(''); }
export async function makeLock(pw) { const salt = crypto.getRandomValues(new Uint32Array(2)).join('-'); return { salt, hash: await sha(salt + '|' + pw) }; }
export async function checkLock(lock, pw) {
  if (!lock) return false;
  if ((await sha(lock.salt + '|' + pw)) === lock.hash) return true;
  return pw !== pw.trim() && (await sha(lock.salt + '|' + pw.trim())) === lock.hash;   // stray space from a phone keyboard
}

// Editable list of ranges ("A1:B5"), starting from the current selection.
function rangeList(app, list) {
  const rows = el('div.rl-rows');
  const draw = () => {
    rows.replaceChildren(...list.map((g, i) => el('div.rl-row', icon('crop_free', 'sm'), el('span.mono', rangeAddr(g)),
      el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Hapus', onclick: () => { list.splice(i, 1); draw(); } }, icon('close')))));
    if (!list.length) rows.append(el('div.muted.small', 'Belum ada sel. Pilih sel di sheet lalu klik "Tambah pilihan", atau ketik rentang.'));
  };
  const inp = el('input.inp.mono', { placeholder: 'mis. B2:D20', dataset: { enter: 'local' } });   // Enter adds, not submits
  const add = text => {
    let gs = [];
    if (text == null) gs = app.grid.sel.ranges.map(g => ({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }));
    else for (const part of text.split(/[,;\s]+/).filter(Boolean)) { const g = parseRange(part); if (!g) { toast(`Rentang "${part}" tidak valid`); return; } gs.push({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }); }
    for (const g of gs) if (!list.some(x => x.r1 === g.r1 && x.c1 === g.c1 && x.r2 === g.r2 && x.c2 === g.c2)) list.push(g);
    inp.value = ''; draw();
  };
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); if (inp.value.trim()) add(inp.value); } });
  draw();
  return el('div.stack.rl', rows, el('div.row',
    el('button.btn', { type: 'button', onclick: () => add(null) }, icon('add'), 'Tambah pilihan (' + app.grid.sel.ranges.map(g => rangeAddr(g)).join(', ') + ')'),
    inp, el('button.btn', { type: 'button', onclick: () => inp.value.trim() && add(inp.value) }, 'Tambah')));
}

export async function lockDialog(app, sh = app.sheet) {
  if (sh.protect) return unlockDialog(app, sh);
  const open = (sh.protectExcept || []).map(g => ({ ...g }));
  const pw = secretInput('Kata sandi'), pw2 = secretInput('Ulangi kata sandi');
  const match = secretMatch(pw, pw2);
  dialog({
    title: `Kunci sheet "${sh.name}"`, width: 520,
    body: el('div.stack',
      el('p.dlg-text', 'Semua sel di sheet ini dikunci: isi, format, baris/kolom dan pengaturan sheet tidak dapat diubah oleh siapa pun (termasuk tautan edit) sampai kunci dibuka dengan kata sandi. Sel izin pelihat (tautan Lihat saja) tetap bisa diisi pelihat.'),
      el('div.grid2', field('Kata sandi', secretBox(pw)), field('Ulangi', secretBox(pw2))),
      match.note,
      el('h4', 'Sel yang tetap terbuka untuk diedit'),
      rangeList(app, open)),
    buttons: [{ label: 'Batal' }, { label: 'Kunci sheet', primary: true, run: async () => {
      if (!pw.value.trim()) { toast('Isi kata sandi'); pw.focus(); return false; }
      if (!match.same()) { match.update(); toast(match.diff()); return false; }
      const lock = await makeLock(pw.value.trim());
      app.permit('Kunci sheet', () => { app.wb.hist.touchMeta(sh); sh.protect = true; sh.lock = lock; sh.protectExcept = open; });
      toast(`Sheet "${sh.name}" dikunci` + (open.length ? ` · ${open.length} rentang tetap terbuka` : ''));
    } }],
  });
}

export async function unlockDialog(app, sh = app.sheet) {
  if (!sh.protect) return lockDialog(app, sh);
  if (app.readOnly) { toast('Hanya pemilik atau editor yang dapat membuka kunci'); return; }
  if (sh.lock) {
    const pw = await promptDlg(`Buka kunci "${sh.name}"`, 'Kata sandi', '', { type: 'password', ok: 'Buka kunci' });
    if (pw == null) return;
    if (!(await checkLock(sh.lock, pw))) { toast('Kata sandi salah'); return; }
  }
  app.permit('Buka kunci sheet', () => { app.wb.hist.touchMeta(sh); sh.protect = false; sh.lock = null; });
  toast(`Kunci sheet "${sh.name}" dibuka`);
}

// Unlocked cells of a locked sheet can be changed after entering the password.
export async function lockedCellsDialog(app, sh = app.sheet) {
  if (!sh.protect) return lockDialog(app, sh);
  if (sh.lock) {
    const pw = await promptDlg('Ubah sel yang terbuka', 'Kata sandi sheet', '', { type: 'password' });
    if (pw == null) return;
    if (!(await checkLock(sh.lock, pw))) { toast('Kata sandi salah'); return; }
  }
  const open = (sh.protectExcept || []).map(g => ({ ...g }));
  dialog({
    title: `Sel terbuka di "${sh.name}"`, width: 500,
    body: el('div.stack', el('p.dlg-text', 'Sel di daftar ini tetap dapat diedit walaupun sheet dikunci.'), rangeList(app, open)),
    buttons: [{ label: 'Batal' }, { label: 'Simpan', primary: true, run: () => { app.permit('Sel terbuka', () => { app.wb.hist.touchMeta(sh); sh.protectExcept = open; }); } }],
  });
}

export function viewerCellsDialog(app, sh = app.sheet) {
  if (app.readOnly) { toast('Hanya pemilik atau editor yang dapat mengatur izin ini'); return; }
  const list = (sh.ve || []).map(g => ({ ...g }));
  dialog({
    title: `Sel yang bisa diedit pelihat — "${sh.name}"`, width: 520,
    body: el('div.stack',
      el('p.dlg-text', 'Orang yang membuka tautan "Lihat saja" hanya dapat mengisi sel di daftar ini. Sel lainnya tetap hanya bisa dilihat. Server memeriksa setiap perubahan pelihat.'),
      rangeList(app, list),
      el('p.muted.small', 'Sel ini ditandai garis hijau putus-putus (Lihat › Tandai sel yang bisa diedit pelihat).')),
    buttons: [{ label: 'Batal' }, { label: 'Simpan', primary: true, run: () => {
      app.tx('Izin pelihat', () => { app.wb.hist.touchMeta(sh); sh.ve = list; });
      if (list.length && !app.settings.showVE) { app.settings.showVE = true; app.saveSettings(); app.grid.render(); }
      toast(list.length ? `${list.length} rentang dapat diedit pelihat` : 'Pelihat tidak dapat mengedit sel apa pun');
    } }],
  });
}

// ---------- live session ----------
export function startLive(app) {
  app.live?.stop(); app.live = null;
  const s = app.wb.share;
  app.setRole(s && s.role ? s : null);
  if (!s?.key) return;
  const wbId = app.wb.id;
  const mine = () => app.wb && app.wb.id === wbId;
  const live = new Live(s.key, {
    payload: () => (mine() ? payload(app.wb) : null),
    merge: mergeBooks,
    same: sameBooks,
    viewerEdits: true,
    apply: (doc, info) => { if (mine()) app.applyRemote(doc, info); },
    busy: () => !!app.grid.edit,
    status: (st, info, e) => { if (mine()) app.shareStatus(st, info, e); },
    role: info => { if (!mine()) return; Object.assign(app.wb.share, { role: info.role, name: info.name, by: info.by, ownerName: info.owner, owner: s.owner || false }); app.setRole(app.wb.share); },
    name: () => app.wb.name,
    by: () => account()?.u || 'Tamu',
    persist: st => { if (mine()) { Object.assign(app.wb.share, st); app.saveLater(); } },
  }, s);
  app.live = live;
  live.start().catch(e => app.shareStatus(e.code === 'notfound' ? 'revoked' : 'offline', null, e));
}

// ---------- share dialog ----------
export async function shareDialog(app) {
  const wb = app.wb, s = wb.share;
  const acc = account();
  const body = el('div.stack.share-dlg');
  const close = () => box.close();
  const copyRow = (label, desc, url, on, onToggle, onReset) => {
    const inp = el('input.inp.mono', { readOnly: true, value: on ? url : 'Tautan dimatikan', disabled: !on });
    return el('div.share-row' + (on ? '' : '.off'),
      el('div.share-head', el('div', el('b', label), el('div.muted.small', desc)),
        onToggle ? el('label.switch', el('input', { type: 'checkbox', checked: on, onchange: e => onToggle(e.target.checked) }), el('span')) : null),
      el('div.row', inp,
        el('button.btn', { type: 'button', disabled: !on, onclick: () => { inp.select(); navigator.clipboard?.writeText(url); toast('Tautan disalin'); } }, icon('content_copy'), 'Salin'),
        onReset && on ? el('button.icon-btn', { type: 'button', 'data-tip': 'Buat tautan baru (tautan lama berhenti berfungsi)', onclick: onReset }, icon('autorenew')) : null));
  };
  const files = el('div.row', el('button.btn', { type: 'button', onclick: () => io.exportXLSX(app) }, icon('download'), 'Download .xlsx'), el('button.btn', { type: 'button', onclick: () => io.exportCSV(app, ',') }, icon('download'), 'Download .csv'));
  const veCount = wb.sheets.reduce((n, sh) => n + (sh.ve || []).length, 0);
  const veNote = () => el('div.share-note', icon('edit_note'), el('div',
    el('b', veCount ? `Pelihat dapat mengisi ${veCount} rentang sel` : 'Pelihat tidak dapat mengedit sel apa pun'),
    el('div.muted.small', 'Atur di Data › Proteksi › Sel yang bisa diedit pelihat.'),
    el('button.btn.sm', { type: 'button', onclick: () => { close(); viewerCellsDialog(app); } }, icon('tune'), 'Atur sel')));

  const render = () => {
    const s = wb.share;
    body.replaceChildren();
    if (s && !s.owner) {   // opened through someone's link
      body.append(
        el('div.share-note', icon(s.role === 'edit' ? 'edit' : 'visibility'), el('div', el('b', s.role === 'edit' ? 'Anda dapat mengedit spreadsheet ini' : 'Anda hanya dapat melihat spreadsheet ini'),
          el('div.muted.small', `Dibagikan oleh ${s.ownerName || 'pemilik'}. Perubahan tersimpan langsung untuk semua yang membuka tautan.`))),
        copyRow('Tautan ini', 'Teruskan hanya ke orang yang Anda percaya.', linkFor('spreadsheet', s.key), true),
        files);
      return;
    }
    if (!acc) {
      body.append(el('div.share-note', icon('login'), el('div', el('b', 'Masuk ke Portal AMATU untuk membuat tautan berbagi'), el('div.muted.small', 'Tautan "Lihat saja" dan "Bisa mengedit" disimpan di akun AMATU Anda.'))), files);
      return;
    }
    if (!s) {
      body.append(
        el('div.share-note', icon('group_add'), el('div', el('b', 'Bagikan dengan tautan'), el('div.muted.small', 'Spreadsheet diunggah ke server AMATU. Anda mendapat dua tautan: "Lihat saja" dan "Bisa mengedit". Perubahan tersinkron otomatis untuk semua yang membukanya.'))),
        el('div.row', el('button.btn.primary', { type: 'button', onclick: create }, icon('link'), 'Buat tautan berbagi')),
        files);
      return;
    }
    body.append(
      copyRow('Lihat saja', 'Siapa pun yang memiliki tautan dapat melihat. Hanya sel yang diizinkan yang bisa diisi.', linkFor('spreadsheet', s.vkey), s.view_on !== false,
        on => set({ p_view: on }), () => set({ p_reset: 'view' })),
      veNote(),
      copyRow('Bisa mengedit', 'Siapa pun yang memiliki tautan dapat mengedit seluruh spreadsheet.', linkFor('spreadsheet', s.ekey), s.edit_on !== false,
        on => set({ p_edit: on }), () => set({ p_reset: 'edit' })),
      el('div.row', el('span.muted.small.flex1', s.by ? `Terakhir diubah oleh ${s.by}` : ''),
        el('button.btn.danger', { type: 'button', onclick: stop }, icon('link_off'), 'Berhenti berbagi')),
      files);
  };
  async function create() {
    try {
      const r = await rpc('amt_share_create', { p_token: acc.t, p_app: 'spreadsheet', p_doc: wb.id, p_name: wb.name, p_data: payload(wb) });
      wb.share = { owner: true, id: r.id, key: r.okey, vkey: r.vkey, ekey: r.ekey, view_on: r.view_on, edit_on: r.edit_on, role: 'edit', rev: r.existing ? 0 : r.rev, clean: !r.existing };
      await store.saveDoc(wb);
      startLive(app);
      render();
      toast('Tautan berbagi dibuat');
    } catch (e) { toast(errText(e)); }
  }
  async function set(args) {
    try {
      const r = await rpc('amt_share_set', { p_token: acc.t, p_id: wb.share.id, p_view: null, p_edit: null, p_reset: null, p_delete: false, ...args });
      Object.assign(wb.share, { vkey: r.vkey, ekey: r.ekey, view_on: r.view_on, edit_on: r.edit_on, key: r.okey });
      await store.saveDoc(wb);
      render();
      if (args.p_reset) toast('Tautan baru dibuat. Tautan lama tidak berfungsi lagi.');
    } catch (e) { toast(errText(e)); }
  }
  async function stop() {
    try {
      await rpc('amt_share_set', { p_token: acc.t, p_id: wb.share.id, p_view: null, p_edit: null, p_reset: null, p_delete: true });
      app.live?.stop(); app.live = null;
      wb.share = null;
      await store.saveDoc(wb);
      app.setRole(null);
      render();
      toast('Berhenti berbagi. Kedua tautan tidak berfungsi lagi.');
    } catch (e) { toast(errText(e)); }
  }
  // Owner on another device: the share may exist already (amt_share_mine).
  if (acc && !s) rpc('amt_share_mine', { p_token: acc.t, p_app: 'spreadsheet', p_doc: wb.id }).then(r => {
    if (r && !wb.share) { wb.share = { owner: true, id: r.id, key: r.okey, vkey: r.vkey, ekey: r.ekey, view_on: r.view_on, edit_on: r.edit_on, role: 'edit', rev: 0, clean: false }; store.saveDoc(wb); startLive(app); render(); }
  }).catch(() => { });
  render();
  const box = dialog({ title: `Bagikan "${wb.name}"`, width: 560, body, buttons: [{ label: 'Selesai', primary: true }] });
}
