// Local persistence: IndexedDB documents + version history, with an in-memory fallback.
let db = null;
const mem = { docs: new Map(), versions: [] };
let vid = 1;

export function ready() {
  return new Promise(res => {
    try {
      const req = indexedDB.open('amatu-sheets', 1);
      req.onupgradeneeded = () => {
        const d = req.result;
        if (!d.objectStoreNames.contains('docs')) d.createObjectStore('docs', { keyPath: 'id' });
        if (!d.objectStoreNames.contains('versions')) { const s = d.createObjectStore('versions', { keyPath: 'vid', autoIncrement: true }); s.createIndex('doc', 'doc'); }
      };
      req.onsuccess = () => { db = req.result; res(); };
      req.onerror = () => res();
      req.onblocked = () => res();
    } catch { res(); }
  });
}

function tx(store, mode, fn) {
  return new Promise((res, rej) => {
    const t = db.transaction(store, mode);
    const s = t.objectStore(store);
    let out;
    Promise.resolve(fn(s)).then(v => (out = v));
    t.oncomplete = () => res(out);
    t.onerror = () => rej(t.error);
  });
}
const reqP = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });

export async function saveDoc(wb) {
  const j = wb.toJSON();
  // share: live share link state of this copy ({key, role, owner, id, vkey, ekey, rev, clean}), kept beside the workbook
  // lock: file password ({salt, hash}) — kept beside the workbook so share links never carry it
  const rec = { id: wb.id, name: wb.name, updated: Date.now(), sheets: wb.sheets.length, json: j, share: wb.share || null, lock: wb.lock || null };
  if (!db) { mem.docs.set(wb.id, rec); return; }
  await tx('docs', 'readwrite', s => s.put(rec));
}
export async function loadDoc(id) {
  if (!db) { const rec = mem.docs.get(id); return rec ? { ...rec.json, share: rec.share || null, lock: rec.lock || null } : null; }
  const rec = await tx('docs', 'readonly', s => reqP(s.get(id)));
  return rec ? { ...rec.json, share: rec.share || null, lock: rec.lock || null } : null;
}
export async function listDocs() {
  let all;
  if (!db) all = [...mem.docs.values()];
  else all = await tx('docs', 'readonly', s => reqP(s.getAll()));
  const meta = readMeta().d;
  return all.map(d => {
    const x = meta[d.id] || {};
    return { id: d.id, name: d.name, updated: d.updated, sheets: d.sheets, share: d.share ? d.share.role : null, locked: !!d.lock, fav: !!x.f, color: LABELS.some(L => L.id === x.c) ? x.c : null, trashed: x.t || null };
  }).sort((a, b) => b.updated - a.updated);
}
export async function deleteDoc(id) {
  if (!db) { mem.docs.delete(id); mem.versions = mem.versions.filter(v => v.doc !== id); return; }
  await tx('docs', 'readwrite', s => s.delete(id));
  const vs = await listVersions(id);
  if (vs.length) await tx('versions', 'readwrite', s => { for (const v of vs) s.delete(v.vid); });
}

export function lastDoc() { try { return localStorage.getItem('as.lastDoc'); } catch { return null; } }
export function setLastDoc(id) { try { localStorage.setItem('as.lastDoc', id); } catch { /* ignore */ } }

// ---------- file list: favourites, colour categories, trash ----------
// Kept beside the documents in one small localStorage value instead of inside the IndexedDB records:
// starring a file or moving it to the trash must not rewrite a whole workbook (the portal syncs every
// record as one big file), and the autosave of an open workbook rewrites its record without these fields.
// as.docmeta = { d: { <doc id>: { f: 1 favourite, c: colour id, t: moved to the trash at (ms) } }, l: { <colour id>: label } }
const META = 'as.docmeta';
export const TRASH_DAYS = 30;
const DAY = 86400000;
export const LABELS = [
  { id: 'red', name: 'Merah', c: '#e5484d' },
  { id: 'orange', name: 'Oranye', c: '#f2711c' },
  { id: 'yellow', name: 'Kuning', c: '#e0a800' },
  { id: 'green', name: 'Hijau', c: '#2e9d5b' },
  { id: 'teal', name: 'Toska', c: '#0f9ba0' },
  { id: 'blue', name: 'Biru', c: '#2f6fe4' },
  { id: 'purple', name: 'Ungu', c: '#8a4fd8' },
  { id: 'gray', name: 'Abu-abu', c: '#7c838c' },
];
let memMeta = null;   // localStorage unavailable (private mode, blocked storage)
export function readMeta() {
  let m = memMeta;
  try { m = JSON.parse(localStorage.getItem(META) || 'null'); } catch { /* keep memMeta */ }
  if (!m || typeof m !== 'object') m = {};
  if (!m.d || typeof m.d !== 'object') m.d = {};
  if (!m.l || typeof m.l !== 'object') m.l = {};
  return m;
}
function writeMeta(m) {
  for (const [id, x] of Object.entries(m.d)) if (!x || (!x.f && !x.c && !x.t)) delete m.d[id];
  for (const [k, v] of Object.entries(m.l)) if (!v) delete m.l[k];
  memMeta = m;
  try { localStorage.setItem(META, JSON.stringify(m)); } catch { /* memory only */ }
  try { dispatchEvent(new CustomEvent('as:docmeta')); } catch { /* ignore */ }
}
export const docMeta = id => readMeta().d[id] || {};
export function setDocMeta(id, patch) {
  const m = readMeta();
  const x = { ...(m.d[id] || {}) };
  for (const [k, v] of Object.entries(patch)) { if (v == null || v === false) delete x[k]; else x[k] = v; }
  m.d[id] = x;
  writeMeta(m);
}
export const isTrashed = id => !!docMeta(id).t;
export const labelColor = cid => LABELS.find(x => x.id === cid)?.c || null;
export const labelName = (cid, m = readMeta()) => m.l[cid] || LABELS.find(x => x.id === cid)?.name || '';
export function setLabelNames(map) {
  const m = readMeta();
  for (const L of LABELS) { const v = String(map[L.id] ?? '').trim().slice(0, 40); if (v && v !== L.name) m.l[L.id] = v; else delete m.l[L.id]; }
  writeMeta(m);
}
export const trashDoc = id => setDocMeta(id, { t: Date.now() });
export const restoreDoc = id => setDocMeta(id, { t: null });
// days left before a trashed file goes for good (0 = today)
export const trashDaysLeft = (t, now = Date.now()) => Math.max(0, Math.ceil((t + TRASH_DAYS * DAY - now) / DAY));
export async function deleteForever(id) {
  await deleteDoc(id);
  const m = readMeta(); delete m.d[id]; writeMeta(m);
}
// Files that sat in the trash for 30 days are deleted for good (startup and whenever the file list opens).
export async function purgeTrash(now = Date.now()) {
  const old = Object.entries(readMeta().d).filter(([, x]) => x.t && now - x.t >= TRASH_DAYS * DAY).map(([id]) => id);
  for (const id of old) { try { await deleteDoc(id); } catch (e) { console.warn('purge', id, e); } }
  if (old.length) { const m = readMeta(); for (const id of old) delete m.d[id]; writeMeta(m); }
  return old.length;
}
// Rename a stored workbook that is not open here (the open one is renamed through the app).
export async function renameDoc(id, name) {
  if (!db) { const rec = mem.docs.get(id); if (rec) { rec.name = name; rec.json = { ...rec.json, name }; } return; }
  await tx('docs', 'readwrite', async s => {
    const rec = await reqP(s.get(id));
    if (rec) { rec.name = name; rec.json = { ...rec.json, name }; s.put(rec); }
  });
}

// ---------- versions ----------
const lastVersionAt = new Map();
export async function saveVersion(wb, label = null) {
  const rec = { doc: wb.id, at: Date.now(), label, json: wb.toJSON() };
  lastVersionAt.set(wb.id, rec.at);
  if (!db) { mem.versions.push({ ...rec, vid: vid++ }); return; }
  await tx('versions', 'readwrite', s => s.add(rec));
  // keep the most recent 60 automatic versions
  const vs = await listVersions(wb.id);
  const auto = vs.filter(v => !v.label);
  if (auto.length > 60) await tx('versions', 'readwrite', s => { for (const v of auto.slice(60)) s.delete(v.vid); });
}
export async function maybeVersion(wb) {
  const last = lastVersionAt.get(wb.id) || 0;
  if (Date.now() - last < 10 * 60 * 1000) return;
  if (!last) { lastVersionAt.set(wb.id, Date.now()); const vs = await listVersions(wb.id); if (vs.length && Date.now() - vs[0].at < 10 * 60 * 1000) return; }
  await saveVersion(wb, null);
}
export async function listVersions(docId) {
  let all;
  if (!db) all = mem.versions.filter(v => v.doc === docId);
  else all = await tx('versions', 'readonly', s => reqP(s.index('doc').getAll(docId)));
  return all.map(v => ({ vid: v.vid, at: v.at, label: v.label })).sort((a, b) => b.at - a.at);
}
export async function getVersion(id) {
  if (!db) return mem.versions.find(v => v.vid === id)?.json || null;
  const rec = await tx('versions', 'readonly', s => reqP(s.get(id)));
  return rec ? rec.json : null;
}

// ---------- share links (deflate + base64url) ----------
const b64 = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); };
const unb64 = str => { const s = atob(str.replace(/-/g, '+').replace(/_/g, '/')); const out = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i); return out; };
export async function encodeShare(json) {
  const data = new TextEncoder().encode(JSON.stringify(json));
  if (typeof CompressionStream === 'undefined') return 'r' + b64(data);
  const cs = new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return 'z' + b64(new Uint8Array(await new Response(cs).arrayBuffer()));
}
export async function decodeShare(code) {
  const kind = code[0], bytes = unb64(code.slice(1));
  let text;
  if (kind === 'z') { const ds = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw')); text = await new Response(ds).text(); }
  else text = new TextDecoder().decode(bytes);
  return JSON.parse(text);
}
