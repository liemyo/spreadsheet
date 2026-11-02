// Live share links (amatu-id.github.io): one server copy of a document, opened through a secret link.
//   vkey "Lihat saja"   – read only (Sheets: plus the cells the owner allowed, checked on the server)
//   ekey "Bisa mengedit" – full edit
// Clients poll amt_share_get(key, rev); editors save whole documents against a base revision and, when someone else
// saved in between, merge three ways (base, mine, theirs) and save again. Same file in Amatu Docs and Amatu Sheets.

const SB_URL = 'https://ieyqxmzztoxgpvwexbkj.supabase.co';
const SB_KEY = 'sb_publishable_o-PQxhkJJ8FG4mHYTUqEBw_uXv9n-eG';

export async function rpc(fn, args) {
  if (window.AMATU?.rpc) return window.AMATU.rpc(fn, args);
  const r = await fetch(SB_URL + '/rest/v1/rpc/' + fn, {
    method: 'POST',
    headers: { apikey: SB_KEY, Authorization: 'Bearer ' + SB_KEY, 'Content-Type': 'application/json' },
    body: JSON.stringify(args || {}),
  });
  const txt = await r.text();
  let j = null; try { j = txt ? JSON.parse(txt) : null; } catch { /* ignore */ }
  if (!r.ok) { const e = new Error((j && j.message) || ('HTTP ' + r.status)); e.code = (j && j.message) || 'net'; throw e; }
  return j;
}

// Signed-in AMATU account on this device (owner side). Share-link guests have none.
export const account = () => { const A = window.AMATU; if (!A || (A.guest && A.guest())) return null; const s = A.sess?.(); return s && s.t ? s : null; };
export const shareKey = () => new URLSearchParams(location.search).get('share');
export const linkFor = (app, key) => `${location.origin}/${app}/?share=${encodeURIComponent(key)}`;
export function errText(e) {
  const c = e?.code || e?.message || '';
  return {
    notfound: 'Tautan tidak berlaku lagi (dimatikan atau diganti pemilik).',
    denied: 'Anda tidak punya izin untuk perubahan ini.',
    big: 'Dokumen terlalu besar untuk dibagikan (maks. 8 MB).',
    quota: 'Kuota berbagi akun penuh (maks. 60 MB / 200 dokumen).',
    auth: 'Sesi AMATU berakhir. Masuk lagi lewat Portal AMATU.',
    net: 'Tidak ada koneksi internet.',
  }[c] || ('Gagal: ' + c);
}

// ---------- three-way merge ----------
const isObj = v => v !== null && typeof v === 'object' && !Array.isArray(v);
export function eq(a, b) {
  if (a === b) return true;
  if (typeof a !== typeof b || a === null || b === null || typeof a !== 'object') return false;
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  if (Array.isArray(a)) { if (a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (!eq(a[i], b[i])) return false; return true; }
  const ka = Object.keys(a).filter(k => a[k] !== undefined), kb = Object.keys(b).filter(k => b[k] !== undefined);
  if (ka.length !== kb.length) return false;
  for (const k of ka) if (!eq(a[k], b[k])) return false;
  return true;
}
// Objects merge per key; anything else changed on both sides keeps mine. custom[key](b, m, t) overrides a top-level key.
export function merge3(b, m, t, custom = null) {
  if (eq(m, t)) return m;
  if (eq(b, m)) return t;
  if (eq(b, t)) return m;
  if (isObj(m) && isObj(t)) {
    const o = {}, bb = isObj(b) ? b : {};
    for (const k of new Set([...Object.keys(m), ...Object.keys(t)])) {
      const v = custom && custom[k] ? custom[k](bb[k], m[k], t[k]) : merge3(bb[k], m[k], t[k]);
      if (v !== undefined) o[k] = v;
    }
    return o;
  }
  return m;
}

// Hunks turning base into side: {s, e, lines} replaces base[s, e).
function hunks(a, b) {
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let ea = a.length, eb = b.length;
  while (ea > p && eb > p && a[ea - 1] === b[eb - 1]) { ea--; eb--; }
  const n = ea - p, m = eb - p;
  if (!n && !m) return [];
  if (!n || !m || n * m > 4e6) return [{ s: p, e: ea, lines: b.slice(p, eb) }];
  const W = m + 1, L = new Uint16Array((n + 1) * W);
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--)
    L[i * W + j] = a[p + i] === b[p + j] ? L[(i + 1) * W + j + 1] + 1 : Math.max(L[(i + 1) * W + j], L[i * W + j + 1]);
  const out = [];
  let i = 0, j = 0, hs = -1, hb = -1;
  const close = () => { if (hs >= 0) { out.push({ s: p + hs, e: p + i, lines: b.slice(p + hb, p + j) }); hs = -1; } };
  while (i < n || j < m) {
    if (i < n && j < m && a[p + i] === b[p + j]) { close(); i++; j++; continue; }
    if (hs < 0) { hs = i; hb = j; }
    if (j < m && (i >= n || L[i * W + j + 1] >= L[(i + 1) * W + j])) j++; else i++;
  }
  close();
  return out;
}
// Line/block-level diff3. Both sides changed the same region differently → mine, then theirs' new blocks.
export function mergeLines(base, mine, theirs) {
  const all = [...hunks(base, mine).map(h => ({ ...h, side: 0 })), ...hunks(base, theirs).map(h => ({ ...h, side: 1 }))].sort((x, y) => x.s - y.s || x.e - y.e);
  const out = [];
  let pos = 0, i = 0;
  const render = (hs, s, e) => { const r = []; let p = s; for (const h of hs) { r.push(...base.slice(p, h.s)); r.push(...h.lines); p = h.e; } r.push(...base.slice(p, e)); return r; };
  while (i < all.length) {
    const g = [all[i++]];
    let s = g[0].s, e = g[0].e;
    while (i < all.length && all[i].s <= e) { e = Math.max(e, all[i].e); g.push(all[i++]); }
    out.push(...base.slice(pos, s));
    const gm = g.filter(h => h.side === 0), gt = g.filter(h => h.side === 1);
    if (!gt.length) out.push(...render(gm, s, e));
    else if (!gm.length) out.push(...render(gt, s, e));
    else {
      const rm = render(gm, s, e), rt = render(gt, s, e);
      out.push(...rm);
      if (!eq(rm, rt)) { const have = new Set(rm), old = new Set(base.slice(s, e)); out.push(...rt.filter(x => !have.has(x) && !old.has(x))); }
    }
    pos = e;
  }
  out.push(...base.slice(pos));
  return out;
}

// ---------- sync session ----------
// opts: payload() → current document (plain JSON), apply(doc, remote) → show doc, merge(base, mine, theirs),
//       busy() → true while the user is in the middle of an input (apply later), status(s, info), role(info),
//       name() → document title, by() → display name, persist(state) → remember {rev, clean} with the local copy,
//       same(a, b) → documents equal (default deep equal; jsonb comes back with its keys reordered).
//       viewerEdits → a view link may still change parts of the document (Sheets: allowed cells), keep them when merging.
export class Live {
  constructor(key, opts, state = {}) {
    this.key = key; this.o = opts;
    this.rev = state.rev || 0; this.clean = state.clean !== false;
    this.base = null; this.same = opts.same || eq;
    this.info = null; this.role = null;
    this.timer = 0; this.saving = null; this.dead = false; this.pending = null; this.fails = 0;
    this.saveSoon = debounce(() => this.save(), 900);
    this.onVis = () => { if (document.visibilityState === 'visible') this.poll(); else if (!this.clean) this.save(); };
    document.addEventListener('visibilitychange', this.onVis);
  }
  async start() {
    const r = await rpc('amt_share_get', { p_key: this.key, p_rev: 0 });
    this.info = r; this.role = r.role; this.o.role?.(r);
    const mine = this.o.payload();
    const remote = r.data;
    if (mine && !this.clean && this.rev === r.rev) { this.setBase(remote, r.rev); this.markDirty(); }
    else if (mine && !this.clean && this.role === 'edit') { const m = this.o.merge({}, mine, remote); this.setBase(remote, r.rev); this.o.apply(m, r); this.markDirty(); }
    else { this.setBase(remote, r.rev); this.o.apply(remote, r); this.markClean(); }
    this.loop();
    return r;
  }
  setBase(doc, rev) { this.base = doc; this.rev = rev; }
  markClean() { if (!this.clean) { this.clean = true; } this.o.persist?.({ rev: this.rev, clean: true }); this.o.status?.('synced', this.info); }
  markDirty() { this.clean = false; this.o.persist?.({ rev: this.rev, clean: false }); this.o.status?.('pending', this.info); }
  loop() {
    clearTimeout(this.timer);
    if (this.dead) return;
    const ms = document.visibilityState === 'visible' ? (this.fails ? Math.min(30000, 3000 * 2 ** this.fails) : 3000) : 20000;
    this.timer = setTimeout(() => this.poll().finally(() => this.loop()), ms);
  }
  // Local edit happened (editor side).
  changed() { if (this.role !== 'edit' || this.dead) return; if (this.clean) this.markDirty(); this.saveSoon(); }
  async poll() {
    if (this.dead || this.saving || this.o.busy?.()) return;
    try {
      const r = await rpc('amt_share_get', { p_key: this.key, p_rev: this.rev });
      this.fails = 0;
      const roleChanged = r.role !== this.role;
      this.info = r; this.role = r.role;
      if (roleChanged) this.o.role?.(r);
      if (!('data' in r) || this.dead) { if (this.clean) this.o.status?.('synced', r); return; }
      const mine = this.o.payload();
      // nothing local, or a viewer whose local copy is not meant to survive (Docs): take the server copy as is
      if (this.same(mine, this.base) || (this.role !== 'edit' && !this.o.viewerEdits)) { this.setBase(r.data, r.rev); this.o.apply(r.data, r); this.markClean(); return; }
      const merged = this.o.merge(this.base, mine, r.data);
      this.setBase(r.data, r.rev);
      this.o.apply(merged, r);
      if (this.role === 'edit' && !this.same(merged, r.data)) { this.markDirty(); this.saveSoon(); } else this.markClean();
    } catch (e) {
      this.fails = Math.min(4, this.fails + 1);
      if (e.code === 'notfound') { this.o.status?.('revoked', this.info); this.stop(); }
      else this.o.status?.('offline', this.info);
    }
  }
  save() { return this.saving || (this.saving = this.doSave().finally(() => { this.saving = null; })); }
  async doSave() {
    if (this.dead || this.role !== 'edit') return;
    for (let attempt = 0; attempt < 4; attempt++) {
      const mine = this.o.payload();
      if (!mine || this.same(mine, this.base)) { this.markClean(); return; }
      this.o.status?.('saving', this.info);
      try {
        const r = await rpc('amt_share_save', { p_key: this.key, p_base: this.rev, p_data: mine, p_name: this.o.name?.() || '', p_by: this.o.by?.() || '' });
        this.fails = 0;
        if (r.conflict) {
          const merged = this.o.merge(this.base, mine, r.data);
          this.setBase(r.data, r.rev);
          this.o.apply(merged, { ...this.info, by: r.by });
          continue;
        }
        this.setBase(mine, r.rev);
        if (this.same(this.o.payload(), mine)) this.markClean(); else this.saveSoon();
        return;
      } catch (e) {
        this.fails = Math.min(4, this.fails + 1);
        if (e.code === 'notfound') { this.o.status?.('revoked', this.info); this.stop(); }
        else if (e.code === 'denied') { this.o.status?.('denied', this.info); }
        else { this.o.status?.(e.code === 'big' || e.code === 'quota' ? 'error' : 'offline', this.info, e); setTimeout(() => !this.dead && this.saveSoon(), 5000); }
        return;
      }
    }
  }
  // Viewer: write single cells (Sheets). cells: [{s, k, v, f}]
  async cells(list) {
    if (!list.length || this.dead) return;
    this.o.status?.('saving', this.info);
    try { const r = await rpc('amt_share_cells', { p_key: this.key, p_cells: list, p_by: this.o.by?.() || '' }); this.fails = 0; this.o.status?.('synced', this.info); return r; }
    catch (e) { this.o.status?.(e.code === 'denied' ? 'denied' : 'offline', this.info, e); throw e; }
    finally { this.poll(); }
  }
  stop() { this.dead = true; clearTimeout(this.timer); this.saveSoon.cancel?.(); document.removeEventListener('visibilitychange', this.onVis); }
}

function debounce(fn, ms) {
  let t = 0;
  const d = (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
  d.cancel = () => clearTimeout(t);
  return d;
}
