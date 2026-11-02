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
  const rec = { id: wb.id, name: wb.name, updated: Date.now(), sheets: wb.sheets.length, json: j, share: wb.share || null };
  if (!db) { mem.docs.set(wb.id, rec); return; }
  await tx('docs', 'readwrite', s => s.put(rec));
}
export async function loadDoc(id) {
  if (!db) { const rec = mem.docs.get(id); return rec ? (rec.share ? { ...rec.json, share: rec.share } : rec.json) : null; }
  const rec = await tx('docs', 'readonly', s => reqP(s.get(id)));
  return rec ? (rec.share ? { ...rec.json, share: rec.share } : rec.json) : null;
}
export async function listDocs() {
  let all;
  if (!db) all = [...mem.docs.values()];
  else all = await tx('docs', 'readonly', s => reqP(s.getAll()));
  return all.map(d => ({ id: d.id, name: d.name, updated: d.updated, sheets: d.sheets, share: d.share ? d.share.role : null })).sort((a, b) => b.updated - a.updated);
}
export async function deleteDoc(id) {
  if (!db) { mem.docs.delete(id); mem.versions = mem.versions.filter(v => v.doc !== id); return; }
  await tx('docs', 'readwrite', s => s.delete(id));
  const vs = await listVersions(id);
  if (vs.length) await tx('versions', 'readwrite', s => { for (const v of vs) s.delete(v.vid); });
}

export function lastDoc() { try { return localStorage.getItem('as.lastDoc'); } catch { return null; } }
export function setLastDoc(id) { try { localStorage.setItem('as.lastDoc', id); } catch { /* ignore */ } }

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
