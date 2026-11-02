// Application shell: state, commands, keyboard, clipboard, sheet tabs, status bar, persistence.
import { el, icon, key, keyR, keyC, addr, rangeAddr, colName, debounce, inRange, jsDateToSerial, uid, parseRange, esc } from './util.js';
import { Workbook, newWorkbook, setInput, rangeCells, COL_W, ROW_H } from './model.js';
import { setLocale, editText, L, changeDecimals, format, compile, fmtSmart } from './format.js';
import { Grid, HH } from './grid.js';
import { CFEngine, dvAt, dvItems, dvValid, CHIP_COLORS } from './cf.js';
import * as ops from './ops.js';
import { FErr } from './formula/core.js';
import { FormulaBar } from './ui/fx.js';
import { openMenu, popup, closeMenus, toast, dialog, promptDlg, confirmDlg, initTooltips, alertDlg } from './ui/ui.js';
import { buildMenubar, cellContextMenu, colContextMenu, rowContextMenu, tabContextMenu } from './ui/menus.js';
import { buildToolbar } from './ui/toolbar.js';
import * as D from './ui/dialogs.js';
import { Sidebar } from './ui/sidebar.js';
import { Charts } from './charts.js';
import * as store from './store.js';
import * as io from './io.js';
import { TEMPLATES } from './templates.js';
import * as SH from './sharing.js';
import * as PG from './pages.js';
import * as LK from './lock.js';
import { shareKey, errText } from './live.js';

const isMac = /Mac|iPhone|iPad/.test(navigator.platform);
export const MOD = isMac ? '⌘' : 'Ctrl';

const THEMES = [
  { id: 'glass-light', label: 'Kaca Terang', dark: false },
  { id: 'glass-dark', label: 'Kaca Gelap', dark: true },
  { id: 'aurora', label: 'Aurora', dark: false },
  { id: 'ocean', label: 'Samudra', dark: true },
  { id: 'classic', label: 'Klasik', dark: false },
  { id: 'excel', label: 'Hijau Excel', dark: false },
  { id: 'graphite', label: 'Grafit', dark: true },
];

export class App {
  constructor() {
    this.settings = this.loadSettings();
    this.showFormulas = false;
    this.clip = null;
    this.findHits = null;
    this.docId = null;
    this.THEMES = THEMES;
    this.MOD = MOD;
  }

  loadSettings() {
    let s = {};
    try { s = JSON.parse(localStorage.getItem('as.settings') || '{}'); } catch { /* ignore */ }
    return { theme: 'glass-light', glass: true, autocomplete: true, stats: ['sum', 'avg', 'count'], ...s };
  }
  saveSettings() { try { localStorage.setItem('as.settings', JSON.stringify(this.settings)); } catch { /* ignore */ } }

  // wb is set only after init() has loaded it; the grid may draw (ResizeObserver) before that.
  get sheet() { return this.wb ? this.wb.sheets[this.wb.active] : undefined; }
  get defFont() { const s = this.wb ? this.wb.settings : {}; return { ff: s.font || 'Arial', fs: s.fontSize || 10 }; }

  async init() {
    initTooltips();
    this.applyTheme(this.settings.theme, false);
    this.titleInput = document.getElementById('doc-title');
    this.saveState = document.getElementById('save-state');
    this.grid = new Grid(this, document.getElementById('grid'));
    this.fx = new FormulaBar(this, document.getElementById('fxbar'));
    this.cf = new CFEngine(this);
    this.sidebar = new Sidebar(this, document.getElementById('sidebar'));
    this.charts = new Charts(this, this.grid.layer);
    buildMenubar(this, document.getElementById('menubar'));
    this.toolbar = buildToolbar(this, document.getElementById('toolbar'));
    this.bindChrome();
    this.bindKeys();
    this.bindClipboard();
    this.grid.readTheme();
    this.saveLater = debounce(() => this.save(), 700);
    await store.ready();
    const params = new URLSearchParams(location.search);
    let wb = null;
    const sk = shareKey();
    if (sk) {   // opened through a share link: the server copy, cached on this device under its own id
      const cid = 'sh_' + sk.slice(0, 16).replace(/[^\w-]/g, '_');
      const j = await store.loadDoc(cid);
      wb = j ? Workbook.fromJSON(j) : newWorkbook();
      if (!j) wb.name = 'Memuat…';
      wb.id = cid;
      wb.share = j?.share?.key === sk ? j.share : { key: sk, role: 'view', owner: false, rev: 0, clean: true };
    }
    if (!wb && location.hash.startsWith('#s=')) {
      try { wb = Workbook.fromJSON(await store.decodeShare(location.hash.slice(3))); wb.id = uid('wb'); wb.name += ' (salinan)'; history.replaceState(null, '', location.pathname); toast('Spreadsheet dibuka dari tautan bagikan'); } catch (e) { console.error(e); toast('Tautan bagikan tidak valid'); }
    }
    if (!wb) {
      const id = params.get('id') || store.lastDoc();
      if (id) { const j = await store.loadDoc(id); if (j) wb = Workbook.fromJSON(j); }
    }
    if (!wb) { wb = TEMPLATES.find(t => t.id === 'welcome').build(); }
    this.openWorkbook(wb, true);
    document.body.classList.remove('loading');
    setInterval(() => { if (this.hasVolatile()) this.after({ save: false }); }, 60000);
  }

  // ---------- workbook lifecycle ----------
  openWorkbook(wb, initial = false) {
    this.live?.stop(); this.live = null;
    this.wb = wb;
    this.docId = wb.id;
    setLocale(wb.settings.locale || 'id-ID');
    this.grid.g = null;
    this.titleInput.value = wb.name;
    document.title = wb.name + ' – Amatu Sheets';
    this.grid.sel = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], r: 0, c: 0, er: 0, ec: 0 };
    this.grid.scroller.scrollLeft = 0; this.grid.scroller.scrollTop = 0;
    this.sidebar.close();
    this.findHits = null;
    this.wb.eng.recalc();
    for (const sh of wb.sheets) { this.cf.reset(); this.withSheet(sh, () => ops.filterRows(this)); }
    this.after({ save: !initial });
    this.renderTabs();
    this.charts.reset();
    const guest = wb.share && !wb.share.owner;
    if (!guest) store.setLastDoc(wb.id);
    history.replaceState(null, '', location.pathname + (guest ? '?share=' + encodeURIComponent(wb.share.key) : '?id=' + encodeURIComponent(wb.id)));
    if (!initial) this.save();
    SH.startLive(this);
    this.updateLockBadge();
    LK.gate(this, wb);
    if (!document.body.classList.contains('file-locked')) this.grid.focus();
  }
  openOther() { D.openDialog(this); }
  updateLockBadge() {
    let b = document.getElementById('lock-badge');
    if (!this.wb?.lock) { b?.remove(); return; }
    if (!b) { b = el('button.icon-btn.sm#lock-badge', { type: 'button', 'data-tip': 'Dikunci dengan kata sandi', 'aria-label': 'Keamanan file', onclick: e => openMenu(LK.menuItems(this), e.currentTarget.getBoundingClientRect()) }, icon('lock')); this.titleInput.after(b); }
  }

  // ---------- share links ----------
  // A newer server copy (merged with local edits): replace the workbook, keep sheet, selection, scroll and undo.
  applyRemote(doc) {
    const cur = this.wb;
    if (!cur || !doc || !doc.sheets) return;
    const nw = Workbook.fromJSON({ ...doc, id: cur.id });
    nw.share = cur.share; nw.lock = cur.lock; nw.hist = cur.hist; nw.hist.wb = nw;
    const aid = cur.sheets[cur.active]?.id, i = nw.sheets.findIndex(x => x.id === aid);
    nw.active = i >= 0 ? i : Math.min(cur.active, nw.sheets.length - 1);
    const sel = this.grid.sel, sx = this.grid.scroller.scrollLeft, sy = this.grid.scroller.scrollTop;
    this.wb = nw;
    setLocale(nw.settings.locale || 'id-ID');
    if (document.activeElement !== this.titleInput) this.titleInput.value = nw.name;
    document.title = nw.name + ' – Amatu Sheets';
    this.grid.g = null;
    nw.eng.recalc();
    for (const sh of nw.sheets) { this.cf.reset(); this.withSheet(sh, () => { nw.hist.paused++; ops.filterRows(this); nw.hist.paused--; }); }
    nw.dirty = true;
    const sh = this.sheet;
    if (sel.r < sh.rows && sel.c < sh.cols) this.grid.sel = sel;
    this.after({ remote: true });
    this.charts.reset();
    this.grid.scroller.scrollLeft = sx; this.grid.scroller.scrollTop = sy;
  }
  setRole(s) {
    this.share = s;
    this.readOnly = !!(s && !s.owner && (s.role === 'view' || s.revoked));
    document.body.classList.toggle('ro', this.readOnly);
    document.body.classList.toggle('shared', !!s);
    this.titleInput.readOnly = this.readOnly;
    let b = document.getElementById('share-badge');
    if (!s) { b?.remove(); }
    else {
      if (!b) { b = el('button.share-badge#share-badge', { type: 'button', onclick: () => SH.shareDialog(this) }); this.saveState.after(b); }
      const ve = this.wb.sheets.some(x => (x.ve || []).length);
      const [ic, tx, tip] = s.revoked ? ['link_off', 'Tautan tidak berlaku', 'Pemilik mematikan atau mengganti tautan ini'] : s.owner ? ['group', 'Dibagikan', 'Spreadsheet ini dibagikan dengan tautan'] : s.role === 'edit' ? ['edit', 'Bisa mengedit', 'Anda membuka tautan edit'] : ['visibility', ve ? 'Hanya melihat · isi sel hijau' : 'Hanya melihat', ve ? 'Anda hanya dapat mengisi sel yang ditandai hijau' : 'Anda membuka tautan lihat saja'];
      b.className = 'share-badge ' + (s.revoked ? 'bad' : s.owner ? 'own' : s.role);
      b.replaceChildren(icon(ic), el('span', tx));
      b.dataset.tip = tip;
    }
    if (this.wb) { this.grid.render(); this.fx?.sync(); this.toolbar?.sync(); }
  }
  shareStatus(st, info, e) {
    if (st === 'revoked') {
      if (this.wb.share && !this.wb.share.owner) {   // the owner withdrew the link: drop the copy cached on this device
        this.wb.share.revoked = true;
        store.deleteDoc(this.wb.id).catch(() => { });
        const blank = newWorkbook(); blank.id = this.wb.id; blank.share = this.wb.share; blank.name = 'Tautan tidak berlaku';
        this.live = null; this.wb = blank; this.wb.dirty = false; this.titleInput.value = blank.name;
        this.grid.g = null; this.renderTabs(); this.grid.render();
        this.setRole(this.wb.share);
      }
      toast(errText({ code: 'notfound' }));
      this.setSaveState('error', ['link_off', 'Tautan tidak berlaku']);
      return;
    }
    if (st === 'denied' || st === 'error') toast(errText(e || { code: 'denied' }));
    const t = {
      synced: ['cloud_done', 'Tersinkron', 'Tersimpan di server berbagi' + (info?.by ? ' · terakhir diubah oleh ' + info.by : '')],
      saving: ['sync', 'Menyinkronkan…'], pending: ['sync', 'Menyinkronkan…'],
      offline: ['cloud_off', 'Offline', 'Tidak terhubung. Perubahan disimpan di perangkat dan dikirim saat online.'],
      denied: ['block', 'Ditolak'], error: ['error', 'Gagal sinkron'],
    }[st];
    if (t) this.setSaveState(st === 'synced' ? 'saved' : st === 'offline' || st === 'error' || st === 'denied' ? 'error' : 'saving', t);
  }
  permit(label, fn) { this._permit = true; try { return this.tx(label, fn); } finally { this._permit = false; } }
  withSheet(sh, fn) { const prev = this.wb.active; this.wb.active = this.wb.sheets.indexOf(sh); try { return fn(); } finally { this.wb.active = prev; } }

  async save() {
    if (!this.wb || this.wb.share?.revoked) return;
    if (!this.live) this.setSaveState('saving');
    try {
      await store.saveDoc(this.wb);
      this.wb.dirty = false;
      if (!this.live) this.setSaveState('saved');
      store.maybeVersion(this.wb);
    } catch (e) { console.error(e); this.setSaveState('error'); }
  }
  setSaveState(s, custom) {
    const t = custom || { saving: ['sync', 'Menyimpan…'], saved: ['cloud_done', 'Semua perubahan disimpan di perangkat'], error: ['error', 'Gagal menyimpan'] }[s];
    this.saveState.innerHTML = '';
    this.saveState.append(icon(t[0]), el('span', custom ? t[1] : s === 'saved' ? 'Tersimpan' : t[1]));
    this.saveState.dataset.tip = t[2] || t[1];
    this.saveState.className = 'save-state ' + s;
  }

  hasVolatile() {
    for (const sh of this.wb.sheets) for (const c of sh.cells.values()) if (c.f != null && /\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(/i.test(c.f)) return true;
    return false;
  }

  // ---------- transactions ----------
  tx(label, fn, opts = {}) {
    const hist = this.wb.hist, last = hist.undoS[hist.undoS.length - 1];
    hist.begin(label, this.selSnapshot());
    let res;
    try { res = fn(); }
    finally { hist.end(this.selSnapshot()); }
    // locked sheets / view-only share links: undo whatever the transaction changed outside what is allowed
    const e = hist.depth === 0 ? hist.undoS[hist.undoS.length - 1] : null;
    if (e && e !== last) {
      const why = SH.checkEntry(this, e);
      if (why) { hist.undo(); hist.redoS.pop(); this.after({ save: false }); toast(why); return undefined; }
      if (this.readOnly) this.sendCells(SH.viewerCells(this, e));
    }
    this.after(opts);
    return res;
  }
  // Viewer edits of allowed cells go to the server one cell at a time; a refusal restores the server copy.
  sendCells(cells) {
    if (!cells.length || !this.live) return;
    this.live.cells(cells).catch(() => { if (this.live?.base) this.applyRemote(this.live.base); });
  }
  selSnapshot() { return { sheet: this.sheet.id, sel: JSON.parse(JSON.stringify(this.grid.sel)) }; }
  restoreSel(s) {
    if (!s) return;
    const i = this.wb.sheets.findIndex(x => x.id === s.sheet);
    if (i >= 0 && i !== this.wb.active) this.switchSheet(i);
    this.grid.sel = s.sel;
    this.grid.scrollTo(s.sel.r, s.sel.c);
  }
  after(opts = {}) {
    const wb = this.wb;
    wb.eng.recalc();
    if (opts.refilter !== false) for (const sh of wb.sheets) if (sh.filter) this.withSheet(sh, () => { wb.hist.paused++; ops.filterRows(this); wb.hist.paused--; });
    for (const sh of wb.sheets) if (sh.dv.length) this.withSheet(sh, () => ops.markInvalid(this, sh));
    this.cf.reset();
    this.cf.ver = -1;
    this.grid.g = null;
    this.pageCache = null;
    this.grid.render();
    this.fx.sync();
    this.toolbar?.sync();
    this.updateStatus();
    this.renderTabs();
    this.charts.update();
    this.sidebar.refresh?.();
    if (opts.save !== false && wb.dirty) { this.saveLater(); if (!opts.remote) this.live?.changed(); }
  }

  undo() { if (this.grid.edit) this.grid.cancelEdit(); const e = this.wb.hist.undo(); if (!e) return; if (this.readOnly) this.sendCells(SH.viewerCells(this, e, true)); this.after(); this.restoreSel(e.sel); this.grid.changed(); this.charts.reset(); }
  redo() { if (this.grid.edit) this.grid.cancelEdit(); const e = this.wb.hist.redo(); if (!e) return; if (this.readOnly) this.sendCells(SH.viewerCells(this, e, true)); this.after(); this.restoreSel(e.selAfter); this.grid.changed(); this.charts.reset(); }

  // ---------- editing ----------
  editTextOf(sh, r, c) {
    const cell = sh.get(r, c);
    if (!cell) return '';
    if (cell.f != null) return '=' + cell.f;
    const st = this.wb.styles[cell.s] || {};
    return editText(cell.v, st.fmt);
  }
  canEdit(r, c, quiet = false) {
    const why = SH.cellBlock(this, this.sheet, r, c);
    if (why) { if (!quiet) toast(why); return false; }
    return true;
  }
  commitCell(sh, r, c, text, orig, opts = {}) {
    if (text === orig && !opts.force && !opts.fillSel) return true;
    // validation
    const rule = dvAt(sh, r, c);
    if (rule && rule.strict && text !== '' && !text.startsWith('=')) {
      const p = io.peekValue(text);
      if (!dvValid(sh, rule, r, c, p)) {
        alertDlg('Data tidak valid', rule.help || 'Data yang Anda masukkan tidak sesuai dengan aturan validasi data yang ditetapkan untuk sel ini.');
        return false;
      }
    }
    const targets = opts.fillSel ? this.grid.sel.ranges : [{ r1: r, c1: c, r2: r, c2: c }];
    this.tx('Edit', () => {
      for (const g of targets) {
        for (let rr = g.r1; rr <= Math.min(g.r2, sh.rows - 1); rr++) for (let cc = g.c1; cc <= Math.min(g.c2, sh.cols - 1); cc++) {
          let t = text;
          if (opts.fillSel && t.startsWith('=')) t = '=' + shiftF(t.slice(1), rr - r, cc - c);
          setInput(this.wb, sh, rr, cc, t);
        }
      }
      if (text.includes('\n') || sh.styleOf(r, c).wrap === 'wrap') { this.wb.eng.recalc(); ops.fitRow(this, sh, r); }
    });
    return true;
  }

  // ---------- selection feedback ----------
  onSelChange() {
    this.fx?.sync();
    this.toolbar?.sync();
    this.updateStatus();
    this.showLinkChip();
    PG.breakChip(this);
    this.sidebar?.onSel?.();
  }

  showLinkChip() {
    const host = document.getElementById('linkchip');
    host.innerHTML = '';
    host.classList.remove('show');
    const g = this.grid;
    if (g.edit) return;
    const cell = this.sheet.get(g.sel.r, g.sel.c);
    const link = cell && (cell.link || cell.hl || (typeof cell.v === 'string' && /^(https?:\/\/|www\.)\S+$/i.test(cell.v) ? cell.v : null));
    if (!link) return;
    const rc = g.cellRect(g.sel.r, g.sel.c);
    const hr = g.host.getBoundingClientRect();
    host.append(...[
      icon('link'),
      el('a', { href: normUrl(link), target: '_blank', rel: 'noopener noreferrer' }, link.length > 46 ? link.slice(0, 46) + '…' : link),
      el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Salin link', onclick: () => { navigator.clipboard?.writeText(link); toast('Link disalin'); } }, icon('content_copy')),
      cell.link ? el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Edit link', onclick: () => D.linkDialog(this) }, icon('edit')) : null,
      cell.link ? el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Hapus link', onclick: () => this.tx('Hapus link', () => this.sheet.put(g.sel.r, g.sel.c, { link: null })) }, icon('link_off')) : null,
    ].filter(Boolean));
    host.style.left = Math.max(4, hr.left + rc.x * g.z) + 'px';
    host.style.top = (hr.top + (rc.y + rc.h) * g.z + 6) + 'px';
    host.classList.add('show');
  }
  openLink(url) { window.open(normUrl(url), '_blank', 'noopener'); }

  showCellTip(html, x, y) {
    const t = document.getElementById('celltip');
    const hr = this.grid.host.getBoundingClientRect();
    t.innerHTML = html;
    t.style.left = Math.min(window.innerWidth - 280, hr.left + x) + 'px';
    t.style.top = (hr.top + y) + 'px';
    t.classList.add('show');
  }
  hideCellTip() { document.getElementById('celltip').classList.remove('show'); }

  closePopups() { closeMenus(); document.getElementById('linkchip').classList.remove('show'); document.getElementById('pagechip')?.classList.remove('show'); }

  updateStatus() {
    const host = document.getElementById('stats');
    const sh = this.sheet, eng = this.wb.eng;
    const g = this.grid.sel;
    let n = 0, cnt = 0, sum = 0, min = Infinity, max = -Infinity, counta = 0, cells = 0;
    const seen = new Set();
    for (const rg of g.ranges) {
      const r2 = Math.min(rg.r2, sh.rows - 1, Math.max(rg.r1, sh.dataRows() - 1)), c2 = Math.min(rg.c2, sh.cols - 1, Math.max(rg.c1, sh.dataCols() - 1));
      for (let r = rg.r1; r <= r2; r++) {
        if (sh.isRowHidden(r)) continue;
        for (let c = rg.c1; c <= c2; c++) {
          const k = key(r, c);
          if (seen.has(k)) continue;
          seen.add(k);
          if (++cells > 300000) break;
          const v = eng.value(sh, r, c);
          if (v == null || v === '') continue;
          counta++;
          if (typeof v === 'number') { cnt++; sum += v; if (v < min) min = v; if (v > max) max = v; }
        }
      }
    }
    host.innerHTML = '';
    const multi = g.ranges.length > 1 || g.ranges[0].r1 !== g.ranges[0].r2 || g.ranges[0].c1 !== g.ranges[0].c2;
    if (!multi || !counta) return;
    const cell = sh.get(g.r, g.c);
    const fmt = (this.wb.styles[cell?.s || 0] || {}).fmt || cell?.af;
    const f = v => (fmt && !compile(fmt)?.isDate ? format(+v.toPrecision(12), fmt).t : fmtSmart(v));
    const all = {
      sum: ['Jumlah', cnt ? f(sum) : null], avg: ['Rata-rata', cnt ? f(sum / cnt) : null], min: ['Min', cnt ? f(min) : null], max: ['Maks', cnt ? f(max) : null],
      count: ['Hitung', String(counta)], countn: ['Hitung angka', String(cnt)],
    };
    const chosen = this.settings.stats || ['sum', 'avg', 'count'];
    const btn = el('button.stat-btn', { type: 'button', 'data-tip': 'Pilih statistik' });
    for (const k of chosen) { const [lb, v] = all[k]; if (v == null) continue; btn.append(el('span.stat', el('span.muted', lb + ': '), el('b', v))); }
    btn.append(icon('arrow_drop_down'));
    btn.addEventListener('click', () => openMenu(Object.entries(all).map(([k, [lb, v]]) => ({
      label: lb, hint: v ?? '–', checked: () => (this.settings.stats || []).includes(k),
      run: () => { const s = new Set(this.settings.stats || []); s.has(k) ? s.delete(k) : s.add(k); this.settings.stats = Object.keys(all).filter(x => s.has(x)); this.saveSettings(); this.updateStatus(); },
    })), btn.getBoundingClientRect()));
    host.append(btn);
  }

  // ---------- sheets ----------
  switchSheet(i, keepEdit = false) {
    if (i < 0 || i >= this.wb.sheets.length) return;
    const g = this.grid;
    if (g.edit && !keepEdit) {
      if (g.isFormulaEdit()) { /* keep editing across sheets for cross-sheet refs */ }
      else if (!g.commitEdit(0, 0)) return;
    }
    if (i === this.wb.active) return;
    this.saveView();
    this.wb.active = i;
    g.g = null;
    const v = this.sheet.view || { sel: { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], r: 0, c: 0, er: 0, ec: 0 }, sx: 0, sy: 0 };
    g.sel = JSON.parse(JSON.stringify(v.sel));
    g.geo();
    g.scroller.scrollLeft = v.sx; g.scroller.scrollTop = v.sy;
    this.cf.reset();
    g.render();
    g.placeEditor();
    this.renderTabs();
    this.charts.reset();
    this.onSelChange();
    this.sidebar.onSheet?.();
  }
  saveView() {
    const g = this.grid;
    this.sheet.view = { sel: JSON.parse(JSON.stringify(g.sel)), sx: g.scroller.scrollLeft, sy: g.scroller.scrollTop };
  }

  renderTabs() {
    const host = document.getElementById('tabs');
    if (!host || !this.wb) return;
    host.innerHTML = '';
    this.wb.sheets.forEach((sh, i) => {
      if (sh.hidden) return;
      const tab = el('div.tab' + (i === this.wb.active ? '.active' : ''), { draggable: true, role: 'tab', 'aria-selected': i === this.wb.active ? 'true' : 'false' },
        sh.protect ? icon('lock', 'sm') : null,
        el('span.tab-name', sh.name),
        el('button.tab-caret', { type: 'button', 'aria-label': 'Opsi sheet', onclick: e => { e.stopPropagation(); this.switchSheet(i); openMenu(tabContextMenu(this, sh), e.currentTarget.getBoundingClientRect()); } }, icon('arrow_drop_down')));
      if (sh.tab) tab.style.setProperty('--tab-color', sh.tab);
      tab.addEventListener('click', () => this.switchSheet(i));
      tab.addEventListener('dblclick', () => this.renameSheetInline(tab, sh));
      tab.addEventListener('contextmenu', e => { e.preventDefault(); this.switchSheet(i); openMenu(tabContextMenu(this, sh), { x: e.clientX, y: e.clientY }); });
      tab.addEventListener('dragstart', e => { e.dataTransfer.setData('text/x-sheet', String(i)); tab.classList.add('dragging'); });
      tab.addEventListener('dragend', () => tab.classList.remove('dragging'));
      tab.addEventListener('dragover', e => { e.preventDefault(); tab.classList.add('dropping'); });
      tab.addEventListener('dragleave', () => tab.classList.remove('dropping'));
      tab.addEventListener('drop', e => {
        e.preventDefault(); tab.classList.remove('dropping');
        const from = +e.dataTransfer.getData('text/x-sheet');
        if (Number.isFinite(from) && from !== i) this.cmd('moveSheet', from, i);
      });
      host.append(tab);
    });
    host.querySelector('.tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  renameSheetInline(tab, sh) {
    const name = tab.querySelector('.tab-name');
    const inp = el('input.tab-input', { value: sh.name, 'aria-label': 'Nama sheet' });
    name.replaceWith(inp);
    inp.focus(); inp.select();
    let done = false;
    const finish = ok => {
      if (done) return; done = true;
      const v = inp.value.trim();
      if (ok && v && v !== sh.name) this.cmd('renameSheet', sh, v);
      else this.renderTabs();
    };
    inp.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') finish(true); if (e.key === 'Escape') finish(false); });
    inp.addEventListener('blur', () => finish(true));
  }

  // ---------- themes & view ----------
  applyTheme(id, save = true) {
    const t = THEMES.find(x => x.id === id) || THEMES[0];
    document.documentElement.dataset.theme = t.id;
    document.documentElement.classList.toggle('dark', t.dark);
    document.documentElement.classList.toggle('no-glass', !this.settings.glass);
    document.documentElement.classList.toggle('liquid-ok', !!this.settings.glass && !!navigator.userAgentData?.brands?.some(b => /Chromium/.test(b.brand)));
    this.settings.theme = t.id;
    if (save) { this.saveSettings(); this.grid?.readTheme(); this.charts?.update(true); }
  }
  setGlass(on) { this.settings.glass = on; this.saveSettings(); this.applyTheme(this.settings.theme); }
  setZoom(z) {
    this.grid.setZoom(Math.round(z * 100) / 100);
    const zoom = Math.round(this.grid.z * 100);
    document.getElementById('zoom-val').textContent = zoom + '%';
    document.getElementById('zoom-range').value = zoom;
    this.toolbar?.sync();
    this.charts.position();
  }

  bandingFor = (sh, r, c) => {
    if (!sh.banding.length) return null;
    for (const b of sh.banding) {
      if (r < b.r1 || r > b.r2 || c < b.c1 || c > b.c2) continue;
      if (b.header && r === b.r1) return b.h;
      if (b.footer && r === b.r2) return b.f || b.h;
      const idx = r - b.r1 - (b.header ? 1 : 0);
      return idx % 2 ? b.b : b.a;
    }
    return null;
  };

  // ---------- commands ----------
  cmd(name, ...args) {
    const fn = COMMANDS[name];
    if (!fn) { console.warn('Unknown command', name); return; }
    if (this.grid.edit && !NO_COMMIT.has(name)) { if (!this.grid.commitEdit(0, 0)) return; }
    return fn(this, ...args);
  }
  ranges() { return this.grid.sel.ranges.map(g => ({ ...g })); }

  // ---------- keyboard ----------
  bindKeys() {
    document.addEventListener('keydown', e => {
      if (e.target.closest && e.target.closest('input,textarea,select,[contenteditable]') ) return;
      if (document.body.classList.contains('file-locked')) return;
      if (document.querySelector('.dlg-back')) return;
      this.onKey(e);
    });
  }

  onKey(e, src = 'cell') {
    const g = this.grid, mod = e.ctrlKey || e.metaKey, k = e.key;
    if (document.querySelector('.dlg-back') && src !== 'bar') return;
    if (g.edit) return this.onEditKey(e, src);
    if (src === 'bar') return;
    const stop = () => { e.preventDefault(); e.stopPropagation(); };
    const sh = this.sheet;
    // navigation
    const nav = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
    if (nav && !e.altKey) { stop(); g.move(nav[0], nav[1], e.shiftKey, mod); return; }
    if (k === 'ArrowDown' && e.altKey) { stop(); this.openDropdown(g.sel.r, g.sel.c); return; }
    if ((k === 'ArrowDown' || k === 'ArrowUp') && e.altKey && e.shiftKey) return;
    if (k === 'Tab') { stop(); g.cycle(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'Enter') {
      stop();
      const rg = g.range, single = g.sel.ranges.length === 1 && (rg.r1 === rg.r2 && rg.c1 === rg.c2 || (() => { const m = sh.mergeAt(g.sel.r, g.sel.c); return m && m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2; })());
      if (e.shiftKey) g.cycle(-1, 0);
      else if (single && !mod) {
        const dv = dvAt(sh, g.sel.r, g.sel.c);
        if (dv && dv.type === 'checkbox') { this.cmd('toggleCheckbox'); return; }
        if (dv && (dv.type === 'list' || dv.type === 'range')) { this.openDropdown(g.sel.r, g.sel.c); return; }
        g.startEdit(null, 'edit');
      } else g.cycle(1, 0);
      return;
    }
    if (k === 'F2') { stop(); g.startEdit(null, 'edit'); return; }
    if (k === 'Escape') { if (g.copyRange) { g.stopMarquee(); } this.findHits = null; this.sidebar.escape?.(); g.render(); return; }
    if ((k === 'Delete' || k === 'Backspace') && !mod) { stop(); this.cmd('clear', 'contents'); return; }
    if (k === 'PageDown' || k === 'PageUp') {
      stop();
      const dir = k === 'PageDown' ? 1 : -1;
      if (e.altKey) { const cols = Math.max(1, g.visible.c1 - g.visible.c0); g.scroller.scrollLeft += dir * cols * COL_W * g.z; g.move(0, dir * cols, e.shiftKey); return; }
      if (mod) { this.switchSheet(this.wb.active + dir); return; }
      const rows = Math.max(1, g.visible.r1 - g.visible.r0 - 1);
      g.scroller.scrollTop += dir * rows * ROW_H * g.z;
      const r = Math.max(0, Math.min(sh.rows - 1, (e.shiftKey ? g.sel.er : g.sel.r) + dir * rows));
      e.shiftKey ? g.extendTo(r, g.sel.ec) : g.select(r, g.sel.c);
      return;
    }
    if (k === 'Home') { stop(); if (mod) g.select(sh.freezeR ? sh.freezeR : 0, 0); else e.shiftKey ? g.extendTo(g.sel.er, 0) : g.select(g.sel.r, 0); return; }
    if (k === 'End') { stop(); const u = sh.usedRange(); if (mod && u) g.select(u.r2, u.c2); else if (!mod) { const [r, c] = g.jump(g.sel.r, g.sel.c, 0, 1); g.select(r, c); } return; }
    if (k === ' ' && !mod && !e.shiftKey) {
      const dv = dvAt(sh, g.sel.r, g.sel.c);
      if (dv && dv.type === 'checkbox') { stop(); this.cmd('toggleCheckbox'); return; }
    }
    if (k === ' ' && mod && !e.shiftKey) { stop(); const rg = g.range; g.selectRange({ r1: 0, c1: rg.c1, r2: sh.rows - 1, c2: rg.c2 }, g.sel.r, g.sel.c); return; }
    if (k === ' ' && e.shiftKey && !mod) { stop(); const rg = g.range; g.selectRange({ r1: rg.r1, c1: 0, r2: rg.r2, c2: sh.cols - 1 }, g.sel.r, g.sel.c); return; }
    if (k === ' ' && e.shiftKey && mod) { stop(); g.selectAll(); return; }
    if (k === 'F4' && !mod) { stop(); this.cmd('repeat'); return; }
    if (k === 'F11' && e.shiftKey) { stop(); this.cmd('addSheet'); return; }
    if (k === 'F2' && e.shiftKey) { stop(); D.noteDialog(this); return; }
    if (k === 'F9') { stop(); this.after({ save: false }); toast('Dihitung ulang'); return; }
    if (e.altKey && !mod && (k === '=' || e.code === 'Equal')) { stop(); this.cmd('autosum', 'SUM'); return; }
    if (e.altKey && e.shiftKey && (e.code === 'Digit5')) { stop(); this.cmd('toggleStyle', 's'); return; }
    if (e.altKey && !mod && (k === 'ArrowUp' || k === 'ArrowDown')) return;
    if (mod) {
      const code = e.code;
      const lower = k.length === 1 ? k.toLowerCase() : k;
      const handled = this.modKey(lower, code, e);
      if (handled) { stop(); return; }
      return;
    }
    // start typing
    if (k.length === 1 && !e.altKey) {
      if (!this.canEdit(g.sel.r, g.sel.c)) { stop(); return; }
      if (e.target !== g.input) { e.preventDefault(); g.startEdit(k, 'enter'); return; }
      g.startEdit('', 'enter');
      // character goes into the textarea by default action
    }
  }

  modKey(k, code, e) {
    const s = e.shiftKey, a = e.altKey;
    if (s && !a) {
      switch (code) {
        case 'Digit1': this.cmd('numberFormat', '#,##0.00'); return true;
        case 'Digit2': this.cmd('numberFormat', 'hh:mm:ss'); return true;
        case 'Digit3': this.cmd('numberFormat', L.date); return true;
        case 'Digit4': this.cmd('numberFormat', '"' + L.cur + '"#,##0.00'); return true;
        case 'Digit5': this.cmd('numberFormat', '0.00%'); return true;
        case 'Digit6': this.cmd('numberFormat', '0.00E+00'); return true;
        case 'Digit7': this.cmd('borders', 'outer'); return true;
        case 'Minus': this.cmd('borders', 'clear'); return true;
        case 'Digit9': this.cmd('unhideRows'); return true;
        case 'Digit0': this.cmd('unhideCols'); return true;
        case 'Semicolon': this.cmd('insertNow', 'time'); return true;
      }
      switch (k) {
        case 'z': this.redo(); return true;
        case 'e': this.cmd('align', 'center'); return true;
        case 'l': this.cmd('align', 'left'); return true;
        case 'r': this.cmd('align', 'right'); return true;
        case 'f': this.cmd('toggleFilter'); return true;
        case 'v': this.pasteMode = 'values'; this.pasteSeen = false; setTimeout(() => { if (!this.pasteSeen) this.pasteFromMenu('values'); this.pasteMode = null; }, 250); return false;
        case 'h': D.replaceDialog(this, true); return true;
        case 'x': this.cmd('toggleStyle', 's'); return true;
      }
      return false;
    }
    if (a && s) {
      if (code === 'Semicolon') { this.cmd('insertNow', 'datetime'); return true; }
      if (code === 'Digit5') { this.cmd('toggleStyle', 's'); return true; }
      return false;
    }
    if (a && !s) {
      if (k === 'm' || code === 'KeyM') { D.noteDialog(this); return true; }
      if (code === 'Equal') { D.insertCellsDialog(this); return true; }
      if (code === 'Minus') { D.deleteCellsDialog(this); return true; }
      return false;
    }
    switch (k) {
      case 'c': case 'x': {
        const ok = (() => { try { return document.execCommand(k === 'x' ? 'cut' : 'copy'); } catch { return false; } })();
        if (!ok || !this.copyFired) this.copyTo(null, k === 'x');
        this.copyFired = false;
        return true;
      }
      case 'v': {
        this.pasteSeen = false;
        setTimeout(() => { if (!this.pasteSeen) this.pasteFromMenu(this.pasteMode || 'all'); this.pasteMode = null; }, 250);
        return false;
      }
      case 'z': this.undo(); return true;
      case 'y': this.redo(); return true;
      case 'b': this.cmd('toggleStyle', 'b'); return true;
      case 'i': this.cmd('toggleStyle', 'i'); return true;
      case 'u': this.cmd('toggleStyle', 'u'); return true;
      case '5': this.cmd('toggleStyle', 's'); return true;
      case 'a': this.cmd('selectAll'); return true;
      case 'f': D.findDialog(this); return true;
      case 'h': D.replaceDialog(this, true); return true;
      case 'k': D.linkDialog(this); return true;
      case 'd': this.cmd('fillDown'); return true;
      case 'r': this.cmd('fillRight'); return true;
      case ';': this.cmd('insertNow', 'date'); return true;
      case '`': this.cmd('toggleFormulas'); return true;
      case '\\': this.cmd('clear', 'formats'); return true;
      case '1': D.formatCellsDialog(this); return true;
      case '9': this.cmd('hideRows'); return true;
      case '0': this.cmd('hideCols'); return true;
      case 'p': this.cmd('print'); return true;
      case 's': this.save(); toast('Semua perubahan disimpan di perangkat ini'); return true;
      case 'o': D.openDialog(this); return true;
      case '/': D.shortcutsDialog(this); return true;
      case 'e': this.cmd('flashFill'); return true;
      case 'enter': return false;
      case '=': case '+': if (!e.shiftKey) { this.setZoom(this.grid.z + 0.1); return true; } return false;
      case '-': this.setZoom(this.grid.z - 0.1); return true;
    }
    return false;
  }

  onEditKey(e, src) {
    const g = this.grid, k = e.key, mod = e.ctrlKey || e.metaKey;
    const stop = () => { e.preventDefault(); e.stopPropagation(); };
    if (this.fx.handleKey(e)) { stop(); return; }
    if (k === 'Enter') {
      if (e.altKey || (mod && !e.shiftKey && src === 'bar')) {
        stop();
        const inp = src === 'bar' ? this.fx.input : g.input;
        const p = inp.selectionStart;
        inp.setRangeText('\n', p, inp.selectionEnd, 'end');
        g.onInput(inp);
        return;
      }
      stop();
      if (mod && e.shiftKey) { g.commitEdit(1, 0, { array: true }); return; }
      if (mod) { g.commitEdit(0, 0, { fillSel: true }); return; }
      const multi = g.sel.ranges.length > 1 || g.range.r1 !== g.range.r2 || g.range.c1 !== g.range.c2;
      g.commitEdit(e.shiftKey ? -1 : 1, 0, { cycle: multi });
      if (src === 'bar') g.focus();
      return;
    }
    if (k === 'Tab') { stop(); const multi = g.sel.ranges.length > 1 || g.range.r1 !== g.range.r2 || g.range.c1 !== g.range.c2; g.commitEdit(0, e.shiftKey ? -1 : 1, { cycle: multi }); g.focus(); return; }
    if (k === 'Escape') { stop(); g.cancelEdit(); g.focus(); return; }
    if (k === 'F2') { stop(); g.edit.mode = g.edit.mode === 'edit' ? 'enter' : 'edit'; toast(g.edit.mode === 'edit' ? 'Mode edit' : 'Mode masukkan', { ms: 1200 }); return; }
    const nav = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
    if (nav && src === 'cell' && g.edit.mode === 'enter') {
      if (g.isFormulaEdit() && g.canInsertRef()) {
        stop();
        const ed = g.edit;
        const p = ed.pointSel && ed.pointSel.sh === this.sheet ? ed.pointSel : null;
        let r = p ? (e.shiftKey ? p.er ?? p.g.r2 : p.g.r1) : ed.r, c = p ? (e.shiftKey ? p.ec ?? p.g.c2 : p.g.c1) : ed.c;
        r = Math.max(0, Math.min(this.sheet.rows - 1, r + nav[0]));
        c = Math.max(0, Math.min(this.sheet.cols - 1, c + nav[1]));
        let gg = { r1: r, c1: c, r2: r, c2: c };
        if (e.shiftKey && p) gg = { r1: Math.min(p.g.r1, r), c1: Math.min(p.g.c1, c), r2: Math.max(p.g.r1, r), c2: Math.max(p.g.c1, c) };
        g.insertRef(gg);
        if (e.shiftKey) { ed.pointSel.er = r; ed.pointSel.ec = c; }
        g.scrollTo(r, c);
        return;
      }
      if (!g.isFormulaEdit()) { stop(); g.commitEdit(nav[0], nav[1]); return; }
    }
    if (mod && !e.shiftKey && (k === 'z' || k === 'y')) return; // native textarea undo
    if (mod && k.toLowerCase() === 'b' && !g.isFormulaEdit()) { stop(); return; }
  }

  // ---------- clipboard ----------
  bindClipboard() {
    const inGrid = () => document.activeElement === this.grid.input && !this.grid.edit;
    document.addEventListener('copy', e => { if (!inGrid()) return; e.preventDefault(); this.copyFired = true; this.copyTo(e.clipboardData, false); });
    document.addEventListener('cut', e => { if (!inGrid()) return; e.preventDefault(); this.copyFired = true; this.copyTo(e.clipboardData, true); });
    document.addEventListener('paste', e => {
      if (!inGrid()) return;
      e.preventDefault();
      this.pasteSeen = true;
      const html = e.clipboardData.getData('text/html'), text = e.clipboardData.getData('text/plain');
      const mode = this.pasteMode || 'all';
      this.pasteMode = null;
      this.pasteData(html, text, mode);
    });
  }
  copyTo(dt, cut) {
    const g = this.grid;
    if (g.sel.ranges.length > 1) { toast('Tindakan ini tidak dapat dilakukan pada beberapa pilihan'); return; }
    const rg = g.range;
    const clip = ops.cellsForClipboard(this, rg);
    clip.cut = cut;
    clip.stamp = uid('clip');
    clip.sheetRef = this.sheet;
    clip.wholeCols = rg.r1 === 0 && rg.r2 >= this.sheet.rows - 1;
    this.clip = clip;
    const tsv = ops.toTSV(clip);
    clip.tsv = tsv;
    if (dt) { dt.setData('text/plain', tsv); dt.setData('text/html', ops.toHTML(clip, clip.stamp)); }
    else navigator.clipboard?.writeText(tsv).catch(() => {});
    g.startMarquee(this.sheet, rg);
  }
  pasteData(html, text, mode = 'all', opts = {}) {
    const g = this.grid;
    const dest = { ...g.range };
    const stamp = /data-amatu="([^"]+)"/.exec(html || '')?.[1];
    const internal = this.clip && ((stamp && stamp === this.clip.stamp) || (!html && text === this.clip.tsv) || (!html && !text));
    let clip;
    if (internal) clip = this.clip;
    else if (html && /<table/i.test(html)) {
      const t = ops.parseHTMLTable(html);
      if (t) clip = { rows: t.map((row, i) => row.map((x, j) => ({ raw: x.text, st: mode === 'values' ? {} : x.st, r: i, c: j }))) };
    }
    if (!clip && text != null && text !== '') {
      const rows = ops.parseTSV(text);
      clip = { rows: rows.map((row, i) => row.map((x, j) => ({ raw: x, st: {}, r: i, c: j }))) };
      if (mode === 'all') mode = 'nofmt';
    }
    if (!clip) return;
    if (!this.canEdit(dest.r1, dest.c1)) return;
    const W = Math.max(...clip.rows.map(r => r.length));
    clip.rows = clip.rows.map(r => r.length < W ? [...r, ...Array(W - r.length).fill(null)] : r);
    const res = this.tx('Tempel', () => {
      const sh = this.sheet;
      const R = mode === 'transpose' ? W : clip.rows.length, C = mode === 'transpose' ? clip.rows.length : W;
      sh.growTo(dest.r1 + R, dest.c1 + C);
      if (clip.cut && internal && clip.sheetRef) {
        const src = clip.g;
        const srcSheet = clip.sheetRef;
        if (srcSheet === sh) { ops.moveRange(this, src, dest.r1 - src.r1, dest.c1 - src.c1, false); return { r1: dest.r1, c1: dest.c1, r2: dest.r1 + src.r2 - src.r1, c2: dest.c1 + src.c2 - src.c1 }; }
        const out = ops.pasteClip(this, { ...clip, cut: true }, dest, 'all', { noTile: true });
        this.withSheet(srcSheet, () => ops.clearRanges(this, [src], 'all'));
        return out;
      }
      return ops.pasteClip(this, clip, dest, mode, opts);
    });
    if (clip.cut && internal) { this.clip = null; g.stopMarquee(); }
    if (res) g.selectRange(res, res.r1, res.c1);
  }
  async pasteFromMenu(mode = 'all', opts = {}) {
    let text = null, html = null;
    try {
      if (navigator.clipboard?.read) {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
          if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
        }
      } else if (navigator.clipboard?.readText) text = await navigator.clipboard.readText();
    } catch { /* permission denied → internal clipboard */ }
    if (html == null && text == null && !this.clip) { toast(`Gunakan ${MOD}+V untuk menempel dari papan klip sistem`); return; }
    this.pasteData(html, text, mode, opts);
  }

  // ---------- popups for cells ----------
  openDropdown(r, c) {
    const sh = this.sheet, dv = dvAt(sh, r, c);
    if (!dv || !(dv.type === 'list' || dv.type === 'range')) return;
    if (!this.canEdit(r, c)) return;
    const items = dvItems(sh, dv);
    const g = this.grid;
    const rc = g.cellRect(r, c), hr = g.host.getBoundingClientRect();
    const at = new DOMRect(hr.left + rc.x * g.z, hr.top + rc.y * g.z, rc.w * g.z, rc.h * g.z);
    const cur = this.wb.eng.value(sh, r, c);
    const curSet = new Set(cur == null ? [] : dv.multi ? String(cur).split(',').map(s => s.trim()) : [String(cur)]);
    const search = el('input.inp.sm', { placeholder: 'Cari', 'aria-label': 'Cari opsi' });
    const list = el('div.dd-list');
    let p;
    const choose = v => {
      let val = v;
      if (dv.multi) { const s = new Set(curSet); s.has(v) ? s.delete(v) : s.add(v); val = items.filter(x => s.has(x)).join(', '); }
      this.tx('Pilih', () => setInput(this.wb, sh, r, c, val));
      if (!dv.multi) { p.close(); g.focus(); } else { curSet.clear(); String(val).split(',').map(s => s.trim()).filter(Boolean).forEach(x => curSet.add(x)); draw(); }
    };
    const draw = () => {
      list.innerHTML = '';
      const q = search.value.toLowerCase();
      items.forEach((it, i) => {
        if (q && !it.toLowerCase().includes(q)) return;
        list.append(el('button.dd-item' + (curSet.has(it) ? '.on' : ''), { type: 'button', onclick: () => choose(it) },
          dv.multi ? icon(curSet.has(it) ? 'check_box' : 'check_box_outline_blank', 'sm') : null,
          el('span.chip', { style: { background: dv.chip === 'plain' ? 'transparent' : dv.colors?.[i] || CHIP_COLORS[0] } }, it)));
      });
      if (!list.children.length) list.append(el('div.muted.pad', 'Tidak ada opsi'));
    };
    search.addEventListener('input', draw);
    draw();
    p = popup(el('div.dd', items.length > 7 ? search : null, list, el('div.menu-sep'), el('button.dd-edit', { type: 'button', onclick: () => { p.close(); this.sidebar.open('dv'); } }, icon('edit', 'sm'), 'Edit dropdown')), at);
    p.el.style.minWidth = Math.max(180, rc.w * g.z) + 'px';
    setTimeout(() => (items.length > 7 ? search : list.querySelector('button'))?.focus(), 20);
  }

  openFilterMenu(c) {
    D.filterPopup(this, c);
  }

  contextMenu(e, kind) {
    const at = { x: e.clientX, y: e.clientY };
    const items = kind === 'col' ? colContextMenu(this) : kind === 'row' ? rowContextMenu(this) : cellContextMenu(this);
    openMenu(items, at);
  }

  namedRangesMenu(at) {
    const names = [...this.wb.names.values()];
    openMenu([
      ...names.map(n => ({ label: n.name, hint: String(n.ref), icon: 'label', run: () => this.fx.gotoName(n.name) })),
      names.length ? { sep: true } : null,
      { label: 'Kelola rentang bernama', icon: 'tune', run: () => this.sidebar.open('names') },
    ], at);
  }

  bindChrome() {
    const t = document.getElementById('doc-title');
    t.addEventListener('change', () => { const v = t.value.trim() || 'Spreadsheet tanpa judul'; this.tx('Ganti nama', () => { this.wb.hist.touchWb(); this.wb.name = v; }); document.title = v + ' – Amatu Sheets'; t.blur(); });
    t.addEventListener('keydown', e => { if (e.key === 'Enter') t.blur(); if (e.key === 'Escape') { t.value = this.wb.name; t.blur(); } });
    t.addEventListener('blur', () => this.grid.focus());
    document.getElementById('add-sheet').addEventListener('click', () => this.cmd('addSheet'));
    document.getElementById('all-sheets').addEventListener('click', e => openMenu(this.wb.sheets.map((sh, i) => ({
      label: sh.name, icon: sh.hidden ? 'visibility_off' : sh.protect ? 'lock' : null, checked: () => i === this.wb.active,
      run: () => { if (sh.hidden) this.cmd('unhideSheet', sh); this.switchSheet(i); },
    })), e.currentTarget.getBoundingClientRect()));
    const zr = document.getElementById('zoom-range');
    zr.addEventListener('input', () => this.setZoom(+zr.value / 100));
    document.getElementById('zoom-out').addEventListener('click', () => this.setZoom(this.grid.z - 0.1));
    document.getElementById('zoom-in').addEventListener('click', () => this.setZoom(this.grid.z + 0.1));
    document.getElementById('zoom-val').addEventListener('click', e => openMenu([50, 75, 90, 100, 125, 150, 200].map(z => ({ label: z + '%', checked: () => Math.round(this.grid.z * 100) === z, run: () => this.setZoom(z / 100) })), e.currentTarget.getBoundingClientRect()));
    document.getElementById('btn-share').addEventListener('click', () => SH.shareDialog(this));
    document.getElementById('btn-theme').addEventListener('click', e => openMenu(this.themeMenu(), e.currentTarget.getBoundingClientRect(), { alignRight: true }));
    document.getElementById('app-logo').addEventListener('click', () => D.openDialog(this));
    document.getElementById('btn-history').addEventListener('click', () => this.sidebar.open('versions'));
    window.addEventListener('beforeunload', () => { if (this.wb?.dirty) store.saveDoc(this.wb); });
    window.addEventListener('resize', () => { this.showLinkChip(); PG.breakChip(this); });
    const chipLater = debounce(() => PG.breakChip(this), 120);
    this.grid.scroller.addEventListener('scroll', () => { document.getElementById('pagechip')?.classList.remove('show'); chipLater(); }, { passive: true });
    // drag & drop files
    document.addEventListener('dragover', e => { if (e.dataTransfer?.types?.includes('Files')) { e.preventDefault(); document.body.classList.add('dropping'); } });
    document.addEventListener('dragleave', e => { if (!e.relatedTarget) document.body.classList.remove('dropping'); });
    document.addEventListener('drop', e => {
      document.body.classList.remove('dropping');
      const f = e.dataTransfer?.files?.[0];
      if (!f) return;
      e.preventDefault();
      D.importFile(this, f);
    });
  }

  themeMenu() {
    return [
      { head: 'Tema' },
      ...THEMES.map(t => ({ label: t.label, swatch: `var(--sw-${t.id})`, checked: () => this.settings.theme === t.id, run: () => this.applyTheme(t.id) })),
      { sep: true },
      { label: 'Efek kaca cair', icon: 'blur_on', checked: () => this.settings.glass, run: () => this.setGlass(!this.settings.glass) },
    ];
  }
}

function normUrl(u) { return /^https?:\/\//i.test(u) ? u : /^mailto:/i.test(u) ? u : 'https://' + u; }
import { shiftFormula } from './formula/parse.js';
const shiftF = (f, dr, dc) => shiftFormula(f, dr, dc);

// Commands that must not force-commit an in-progress edit
const NO_COMMIT = new Set(['insertFunction']);

const sel1 = app => app.grid.range;
const fullRowsSel = app => app.grid.sel.ranges.every(g => g.c1 === 0 && g.c2 >= app.sheet.cols - 1);
const fullColsSel = app => app.grid.sel.ranges.every(g => g.r1 === 0 && g.r2 >= app.sheet.rows - 1);

export const COMMANDS = {
  selectAll: app => app.grid.selectAll(),
  clear: (app, what) => { if (!app.canEdit(app.grid.sel.r, app.grid.sel.c)) return; app.tx('Hapus', () => ops.clearRanges(app, app.ranges(), what)); },
  toggleStyle: (app, prop) => {
    const st = app.sheet.styleOf(app.grid.sel.r, app.grid.sel.c);
    const on = !st[prop];
    app.lastAction = ['toggleStyle', prop];
    app.tx('Format', () => ops.styleRanges(app, app.ranges(), s => ({ ...s, [prop]: on || undefined })));
  },
  style: (app, patch, label = 'Format') => { app.lastAction = ['style', patch, label]; app.tx(label, () => ops.styleRanges(app, app.ranges(), patch)); },
  align: (app, ha) => COMMANDS.style(app, { ha }),
  valign: (app, va) => COMMANDS.style(app, { va }),
  wrap: (app, mode) => {
    app.tx('Pengemasan teks', () => {
      ops.styleRanges(app, app.ranges(), { wrap: mode === 'overflow' ? undefined : mode });
      app.wb.eng.recalc();
      for (const g of app.ranges()) for (let r = g.r1; r <= Math.min(g.r2, app.sheet.dataRows()); r++) ops.fitRow(app, app.sheet, r);
    });
  },
  rotate: (app, rot) => app.tx('Rotasi', () => { ops.styleRanges(app, app.ranges(), { rot: rot || undefined }); app.wb.eng.recalc(); for (const g of app.ranges()) for (let r = g.r1; r <= Math.min(g.r2, app.sheet.dataRows()); r++) ops.fitRow(app, app.sheet, r); }),
  font: (app, ff) => COMMANDS.style(app, { ff }),
  fontSize: (app, fs) => app.tx('Ukuran font', () => { ops.styleRanges(app, app.ranges(), { fs: fs === 10 ? undefined : fs }); app.wb.eng.recalc(); for (const g of app.ranges()) for (let r = g.r1; r <= Math.min(g.r2, app.sheet.dataRows()); r++) ops.fitRow(app, app.sheet, r); }),
  color: (app, color) => COMMANDS.style(app, { color: color || undefined }),
  fill: (app, bg) => COMMANDS.style(app, { bg: bg || undefined }),
  numberFormat: (app, fmt) => COMMANDS.style(app, { fmt: fmt === 'General' ? undefined : fmt }, 'Format angka'),
  decimals: (app, d) => {
    const sh = app.sheet, g = app.grid;
    const st = sh.styleOf(g.sel.r, g.sel.c), cell = sh.get(g.sel.r, g.sel.c);
    const sample = app.wb.eng.value(sh, g.sel.r, g.sel.c);
    const base = st.fmt || cell?.af || 'General';
    const nf = changeDecimals(base, d, sample);
    app.tx('Desimal', () => ops.styleRanges(app, app.ranges(), s => ({ ...s, fmt: changeDecimals(s.fmt || base, d, sample) || nf })));
  },
  borders: (app, type, b) => { app.lastBorder = b || app.lastBorder; app.tx('Batas', () => ops.bordersOn(app, app.ranges(), type, b || app.lastBorder || { w: 1, style: 'solid', color: '#000000' })); },
  merge: async (app, mode) => {
    if (mode !== 'none' && !ops.canMerge(app, app.ranges())) {
      if (!(await confirmDlg('Gabungkan sel', 'Menggabungkan sel hanya akan mempertahankan nilai paling kiri atas dan menghapus nilai lainnya.', 'OK'))) return;
    }
    app.tx('Gabung sel', () => ops.mergeRanges(app, app.ranges(), mode));
    app.grid.selectRange(app.grid.range, app.grid.range.r1, app.grid.range.c1);
  },
  paintFormat: app => {
    const g = app.grid, sh = app.sheet;
    const src = { ...g.range };
    const styles = [];
    for (let r = src.r1; r <= src.r2; r++) { const row = []; for (let c = src.c1; c <= src.c2; c++) row.push(sh.styleOf(r, c)); styles.push(row); }
    app.painter = { styles, src };
    document.body.classList.add('painting');
    toast('Pilih sel untuk menerapkan format', { ms: 1800 });
    const done = () => {
      g.scroller.removeEventListener('pointerup', done);
      setTimeout(() => {
        const p = app.painter; app.painter = null; document.body.classList.remove('painting');
        if (!p) return;
        const dest = g.range;
        const R = Math.max(dest.r2 - dest.r1 + 1, p.styles.length), C = Math.max(dest.c2 - dest.c1 + 1, p.styles[0].length);
        app.tx('Format cat', () => { for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) app.sheet.setStyle(dest.r1 + i, dest.c1 + j, p.styles[i % p.styles.length][j % p.styles[0].length]); });
      }, 0);
    };
    g.scroller.addEventListener('pointerup', done);
  },
  resizeCols: (app, cols, w) => app.tx('Ubah ukuran kolom', () => ops.resizeCols(app, cols, w)),
  resizeRows: (app, rows, h) => app.tx('Ubah ukuran baris', () => ops.resizeRows(app, rows, h)),
  autofitCols: (app, cols) => app.tx('Sesuaikan kolom', () => ops.autofitCols(app, cols)),
  autofitRows: (app, rows) => app.tx('Sesuaikan baris', () => ops.autofitRows(app, rows)),
  freeze: (app, r, c) => app.tx('Bekukan', () => ops.freeze(app, r, c)),
  autofill: (app, src, target) => { if (!app.canEdit(src.r1, src.c1)) return; app.tx('Isi otomatis', () => ops.autofill(app, src, target)); app.grid.selectRange(target, app.grid.sel.r, app.grid.sel.c); },
  fillDown: app => app.tx('Isi ke bawah', () => ops.fillFrom(app, app.ranges(), 'down')),
  fillRight: app => app.tx('Isi ke kanan', () => ops.fillFrom(app, app.ranges(), 'right')),
  fillDownDouble: app => { const t = app.tx('Isi otomatis', () => ops.fillDownDouble(app)); if (t) app.grid.selectRange(t, app.grid.sel.r, app.grid.sel.c); },
  moveRange: (app, src, dr, dc, copy) => {
    app.tx(copy ? 'Salin' : 'Pindah', () => ops.moveRange(app, src, dr, dc, copy));
    app.grid.selectRange({ r1: src.r1 + dr, c1: src.c1 + dc, r2: src.r2 + dr, c2: src.c2 + dc }, app.grid.sel.r + dr, app.grid.sel.c + dc);
  },
  insertRows: (app, where) => {
    const g = app.grid.range, n = g.r2 - g.r1 + 1;
    const at = where === 'below' ? g.r2 + 1 : g.r1;
    app.tx('Sisipkan baris', () => ops.insertRows(app, at, n));
    app.grid.selectRange({ r1: at, c1: g.c1, r2: at + n - 1, c2: g.c2 }, at, app.grid.sel.c);
  },
  insertCols: (app, where) => {
    const g = app.grid.range, n = g.c2 - g.c1 + 1;
    const at = where === 'right' ? g.c2 + 1 : g.c1;
    app.tx('Sisipkan kolom', () => ops.insertCols(app, at, n));
    app.grid.selectRange({ r1: g.r1, c1: at, r2: g.r2, c2: at + n - 1 }, app.grid.sel.r, at);
  },
  deleteRows: app => {
    const rs = app.ranges().sort((a, b) => b.r1 - a.r1);
    app.tx('Hapus baris', () => { for (const g of rs) ops.deleteRows(app, g.r1, g.r2 - g.r1 + 1); });
    app.grid.select(Math.min(rs[rs.length - 1].r1, app.sheet.rows - 1), app.grid.sel.c);
  },
  deleteCols: app => {
    const rs = app.ranges().sort((a, b) => b.c1 - a.c1);
    app.tx('Hapus kolom', () => { for (const g of rs) ops.deleteCols(app, g.c1, g.c2 - g.c1 + 1); });
    app.grid.select(app.grid.sel.r, Math.min(rs[rs.length - 1].c1, app.sheet.cols - 1));
  },
  shiftCells: (app, dir) => app.tx('Geser sel', () => ops.shiftCells(app, sel1(app), dir)),
  hideRows: app => { const rows = []; for (const g of app.ranges()) for (let r = g.r1; r <= g.r2; r++) rows.push(r); app.tx('Sembunyikan baris', () => ops.hideRows(app, rows, true)); },
  hideCols: app => { const cols = []; for (const g of app.ranges()) for (let c = g.c1; c <= g.c2; c++) cols.push(c); app.tx('Sembunyikan kolom', () => ops.hideCols(app, cols, true)); },
  unhideRows: app => { const rows = []; for (const g of app.ranges()) for (let r = Math.max(0, g.r1 - 1); r <= Math.min(app.sheet.rows - 1, g.r2 + 1); r++) if (app.sheet.hiddenR.has(r)) rows.push(r); if (fullColsSel(app) || app.grid.sel.ranges[0].r1 === 0 && app.grid.sel.ranges[0].r2 >= app.sheet.rows - 1) rows.push(...app.sheet.hiddenR); app.tx('Tampilkan baris', () => ops.hideRows(app, rows, false)); },
  unhideCols: app => { const cols = []; for (const g of app.ranges()) for (let c = Math.max(0, g.c1 - 1); c <= Math.min(app.sheet.cols - 1, g.c2 + 1); c++) if (app.sheet.hiddenC.has(c)) cols.push(c); if (fullRowsSel(app)) cols.push(...app.sheet.hiddenC); app.tx('Tampilkan kolom', () => ops.hideCols(app, cols, false)); },
  addRows: (app, n) => app.tx('Tambah baris', () => { app.wb.hist.touchMeta(app.sheet); app.sheet.growTo(app.sheet.rows + Math.min(100000, Math.max(1, n | 0)), app.sheet.cols); }),
  addCols: (app, n) => app.tx('Tambah kolom', () => { app.wb.hist.touchMeta(app.sheet); app.sheet.growTo(app.sheet.rows, app.sheet.cols + n); }),
  addSheet: (app, name) => {
    const i = app.wb.active + 1;
    app.tx('Sheet baru', () => app.wb.addSheet(name, i));
    app.switchSheet(i);
  },
  duplicateSheet: app => {
    const src = app.sheet, wb = app.wb;
    const i = wb.active + 1;
    app.tx('Duplikat sheet', () => {
      let base = 'Salinan dari ' + src.name, name = base, n = 2;
      while (wb.sheetByName(name)) name = base + ' ' + n++;
      const sh = wb.addSheet(name, i, src.rows, src.cols);
      const j = src.toJSON(); j.name = name; j.charts = (j.charts || []).map(c => ({ ...c, id: uid('ch') }));
      sh.loadJSON(j);
    });
    app.switchSheet(i);
  },
  deleteSheet: async (app, sh = app.sheet) => {
    if (app.wb.sheets.filter(s => !s.hidden).length <= 1) { toast('Spreadsheet harus memiliki minimal satu sheet yang terlihat'); return; }
    if (!(await confirmDlg('Hapus sheet', `Hapus sheet "${sh.name}"? Anda dapat mengurungkan tindakan ini.`, 'Hapus', true))) return;
    app.tx('Hapus sheet', () => app.wb.deleteSheet(sh));
    let i = Math.min(app.wb.active, app.wb.sheets.length - 1);
    while (app.wb.sheets[i]?.hidden && i > 0) i--;
    app.wb.active = -1;
    app.switchSheet(i);
  },
  renameSheet: (app, sh, name) => {
    if (app.wb.sheetByName(name) && app.wb.sheetByName(name) !== sh) { toast(`Sheet bernama "${name}" sudah ada`); app.renderTabs(); return; }
    if (/[[\]*?/\\:]/.test(name)) { toast('Nama sheet tidak boleh berisi [ ] * ? / \\ :'); app.renderTabs(); return; }
    app.tx('Ganti nama sheet', () => app.wb.renameSheet(sh, name));
  },
  moveSheet: (app, from, to) => {
    const wb = app.wb, cur = app.sheet;
    app.tx('Pindah sheet', () => { wb.hist.touchWb(); const [s] = wb.sheets.splice(from, 1); wb.sheets.splice(to, 0, s); wb.active = wb.sheets.indexOf(cur); });
    app.renderTabs();
  },
  tabColor: (app, sh, color) => app.tx('Warna tab', () => { app.wb.hist.touchMeta(sh); sh.tab = color || null; }),
  hideSheet: (app, sh = app.sheet) => {
    if (app.wb.sheets.filter(s => !s.hidden).length <= 1) { toast('Tidak dapat menyembunyikan satu-satunya sheet'); return; }
    app.tx('Sembunyikan sheet', () => { app.wb.hist.touchMeta(sh); sh.hidden = true; });
    const i = app.wb.sheets.findIndex(s => !s.hidden);
    app.wb.active = -1; app.switchSheet(i);
  },
  unhideSheet: (app, sh) => app.tx('Tampilkan sheet', () => { app.wb.hist.touchMeta(sh); sh.hidden = false; }),
  protectSheet: (app, sh = app.sheet) => (sh.protect ? SH.unlockDialog(app, sh) : SH.lockDialog(app, sh)),
  toggleShowVE: app => { app.settings.showVE = !app.settings.showVE; app.saveSettings(); app.grid.render(); },
  toggleGridlines: app => app.tx('Garis kisi', () => { app.wb.hist.touchMeta(app.sheet); app.sheet.grid = !app.sheet.grid; }),
  toggleFormulas: app => { app.showFormulas = !app.showFormulas; app.grid.render(); app.toolbar.sync(); },
  toggleCheckbox: app => {
    const sh = app.sheet, g = app.grid;
    if (!app.canEdit(g.sel.r, g.sel.c)) return;
    app.tx('Kotak centang', () => {
      for (const rg of app.ranges()) for (let r = rg.r1; r <= Math.min(rg.r2, sh.rows - 1); r++) for (let c = rg.c1; c <= Math.min(rg.c2, sh.cols - 1); c++) {
        const dv = dvAt(sh, r, c);
        if (!dv || dv.type !== 'checkbox') continue;
        const cell = sh.get(r, c);
        if (cell && cell.f != null) continue;
        const on = dv.checked ?? true, off = dv.unchecked ?? false;
        const v = app.wb.eng.value(sh, r, c);
        sh.put(r, c, { v: v === on || String(v) === String(on) ? off : on, f: null });
      }
    });
  },
  insertCheckbox: app => app.tx('Sisipkan kotak centang', () => {
    const sh = app.sheet;
    app.wb.hist.touchMeta(sh);
    const ranges = app.ranges().map(g => ops.clampG(sh, g));
    sh.dv = sh.dv.map(d => ({ ...d, ranges: d.ranges.filter(x => !ranges.some(g => x.r1 >= g.r1 && x.r2 <= g.r2 && x.c1 >= g.c1 && x.c2 <= g.c2)) })).filter(d => d.ranges.length);
    sh.dv.push({ id: uid('dv'), type: 'checkbox', ranges });
    for (const g of ranges) for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const v = sh.get(r, c); if (!v || v.v == null) sh.put(r, c, { v: false }); }
  }),
  insertDropdown: app => app.sidebar.open('dv', { newRule: true }),
  insertNow: (app, kind) => {
    const now = new Date();
    const v = kind === 'time' ? jsDateToSerial(now) % 1 : kind === 'date' ? Math.floor(jsDateToSerial(now)) : jsDateToSerial(now);
    const fmt = kind === 'time' ? 'hh:mm:ss' : kind === 'date' ? L.date : L.date + ' hh:mm:ss';
    const g = app.grid;
    if (!app.canEdit(g.sel.r, g.sel.c)) return;
    app.tx('Sisipkan tanggal', () => { app.sheet.put(g.sel.r, g.sel.c, { v, f: null }); const st = app.sheet.styleOf(g.sel.r, g.sel.c); app.sheet.setStyle(g.sel.r, g.sel.c, { ...st, fmt }); });
  },
  autosum: (app, fn = 'SUM') => {
    const g = app.grid, sh = app.sheet, rg = g.range;
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    if (single) {
      // look up then left for a numeric block
      let r = g.sel.r - 1, c = g.sel.c;
      const isNum = (rr, cc) => typeof app.wb.eng.value(sh, rr, cc) === 'number';
      let text;
      if (r >= 0 && isNum(r, c)) { let r1 = r; while (r1 > 0 && isNum(r1 - 1, c)) r1--; text = `=${fn}(${rangeAddr({ r1, c1: c, r2: r, c2: c })})`; }
      else if (c > 0 && isNum(g.sel.r, c - 1)) { let c1 = c - 1; while (c1 > 0 && isNum(g.sel.r, c1 - 1)) c1--; text = `=${fn}(${rangeAddr({ r1: g.sel.r, c1, r2: g.sel.r, c2: c - 1 })})`; }
      else text = `=${fn}()`;
      g.startEdit(text, 'enter');
      if (text.endsWith('()')) { g.input.setSelectionRange(text.length - 1, text.length - 1); g.onInput(); }
      return;
    }
    // range selected: put totals below each column
    app.tx('Jumlah otomatis', () => {
      const r = rg.r2 + 1;
      sh.growTo(r + 1, sh.cols);
      for (let c = rg.c1; c <= rg.c2; c++) setInput(app.wb, sh, r, c, `=${fn}(${rangeAddr({ r1: rg.r1, c1: c, r2: rg.r2, c2: c })})`);
    });
  },
  insertFunction: (app, name) => {
    const g = app.grid;
    if (!g.edit) {
      const rg = g.range;
      const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
      const text = '=' + name + '(' + (multi ? rangeAddr(rg) : '');
      if (multi) { g.select(rg.r2 + 1 < app.sheet.rows ? rg.r2 + 1 : rg.r2, rg.c1); }
      g.startEdit(text, 'enter');
      return;
    }
    const inp = app.fx.activeInput();
    const p = inp.selectionStart;
    const before = inp.value.slice(0, p);
    const ins = (before.startsWith('=') ? '' : '=') + name + '(';
    inp.setRangeText(ins, p, inp.selectionEnd, 'end');
    g.onInput(inp);
    inp.focus();
  },
  toggleFilter: app => {
    const sh = app.sheet;
    if (sh.filter) { app.tx('Hapus filter', () => { app.wb.hist.touchMeta(sh); sh.filter = null; ops.filterRows(app); }); return; }
    let g = { ...app.grid.range };
    if (g.r1 === g.r2 && g.c1 === g.c2) g = io.currentRegion(app, g.r1, g.c1);
    if (g.r2 >= sh.rows - 1 && g.r1 === 0) g.r2 = Math.max(1, sh.dataRows() - 1);
    app.tx('Buat filter', () => { app.wb.hist.touchMeta(sh); sh.filter = { r1: g.r1, c1: g.c1, r2: g.r2, c2: Math.min(g.c2, Math.max(g.c1, sh.dataCols() - 1, g.c2)), crit: {} }; });
  },
  setFilterCrit: (app, c, crit) => {
    const sh = app.sheet;
    app.tx('Filter', () => { app.wb.hist.touchMeta(sh); sh.filter.crit = { ...sh.filter.crit }; if (crit) sh.filter.crit[c] = crit; else delete sh.filter.crit[c]; ops.filterRows(app); });
  },
  sortRange: (app, g, keys, header) => app.tx('Urutkan', () => ops.sortRange(app, g, keys, header)),
  sortSheet: (app, asc) => {
    const sh = app.sheet, c = app.grid.sel.c;
    const start = sh.freezeR || 0;
    const g = { r1: start, c1: 0, r2: Math.max(start, sh.dataRows() - 1), c2: Math.max(0, sh.dataCols() - 1) };
    app.tx('Urutkan sheet', () => ops.sortRange(app, g, [{ c, asc }], false));
  },
  sortFilterCol: (app, c, asc) => {
    const sh = app.sheet, f = sh.filter;
    const g = { r1: f.r1, c1: f.c1, r2: Math.max(f.r2, sh.dataRows() - 1), c2: f.c2 };
    app.tx('Urutkan', () => ops.sortRange(app, g, [{ c, asc }], true));
  },
  toValues: app => app.tx('Tempel nilai', () => ops.toValues(app, app.ranges())),
  trim: app => { const n = app.tx('Pangkas spasi', () => ops.trimRange(app, app.ranges())); toast(n ? `${n} sel dipangkas` : 'Tidak ada spasi berlebih'); },
  changeCase: (app, mode) => app.tx('Ubah huruf', () => ops.changeCase(app, app.ranges(), mode)),
  flashFill: app => { const n = app.tx('Isi kilat', () => io.flashFill(app)); toast(n ? `Isi kilat mengisi ${n} sel` : 'Isi kilat tidak menemukan pola. Ketik contoh pada sel pertama.'); },
  randomize: app => app.tx('Acak rentang', () => io.shuffleRange(app, sel1(app))),
  print: app => io.printView(app),
  repeat: app => { const a = app.lastAction; if (a) COMMANDS[a[0]](app, ...a.slice(1)); },
  newDoc: (app, tpl = 'blank') => { app.save(); const t = TEMPLATES.find(x => x.id === tpl) || TEMPLATES[0]; app.openWorkbook(t.build()); },
  makeCopy: app => { app.save(); const wb = Workbook.fromJSON(app.wb.toJSON()); wb.id = uid('wb'); wb.lock = app.wb.lock; wb.name = 'Salinan dari ' + app.wb.name; app.openWorkbook(wb); toast('Salinan dibuat'); },
  setLocale: (app, id) => app.tx('Lokal', () => { app.wb.hist.touchWb(); app.wb.settings.locale = id; setLocale(id); app.grid.measureCache.clear(); }),
  defaultFont: (app, ff, fs) => app.tx('Font default', () => { app.wb.settings.font = ff; app.wb.settings.fontSize = fs; app.wb.dirty = true; app.grid.measureCache.clear(); }),
  banding: (app, preset, opts = {}) => {
    const sh = app.sheet;
    let g = { ...app.grid.range };
    if (g.r1 === g.r2 && g.c1 === g.c2) g = io.currentRegion(app, g.r1, g.c1);
    app.tx('Warna bergantian', () => {
      app.wb.hist.touchMeta(sh);
      sh.banding = sh.banding.filter(b => !(b.r1 <= g.r2 && b.r2 >= g.r1 && b.c1 <= g.c2 && b.c2 >= g.c1));
      if (preset) sh.banding.push({ ...g, ...preset, header: opts.header !== false, footer: !!opts.footer });
      if (preset && opts.headerStyle !== false) ops.styleRanges(app, [{ ...g, r2: g.r1 }], s => ({ ...s, b: true, color: preset.ht || s.color }));
    });
  },
  addNamedRange: (app, name, ref) => app.tx('Rentang bernama', () => { app.wb.hist.touchWb(); app.wb.names.set(name.toUpperCase(), { name, ref }); }),
  deleteNamedRange: (app, name) => app.tx('Hapus rentang bernama', () => { app.wb.hist.touchWb(); app.wb.names.delete(name.toUpperCase()); }),
  setNote: (app, note) => { const g = app.grid; app.tx('Catatan', () => app.sheet.put(g.sel.r, g.sel.c, { note: note || null })); },
  setLink: (app, url, text) => {
    const g = app.grid, sh = app.sheet;
    app.tx('Link', () => { sh.put(g.sel.r, g.sel.c, { link: url || null }); if (text != null && text !== '' && text !== app.editTextOf(sh, g.sel.r, g.sel.c)) setInput(app.wb, sh, g.sel.r, g.sel.c, text); else if (url && !app.editTextOf(sh, g.sel.r, g.sel.c)) setInput(app.wb, sh, g.sel.r, g.sel.c, url); });
  },
  cfSave: (app, rule) => app.tx('Format bersyarat', () => {
    const sh = app.sheet; app.wb.hist.touchMeta(sh);
    const i = sh.cf.findIndex(r => r.id === rule.id);
    if (i >= 0) sh.cf[i] = rule; else sh.cf.push(rule);
    sh.cfVer = (sh.cfVer || 0) + 1;
  }),
  cfDelete: (app, id) => app.tx('Hapus format bersyarat', () => { const sh = app.sheet; app.wb.hist.touchMeta(sh); sh.cf = sh.cf.filter(r => r.id !== id); sh.cfVer = (sh.cfVer || 0) + 1; }),
  cfMove: (app, id, dir) => app.tx('Urutan aturan', () => { const sh = app.sheet; app.wb.hist.touchMeta(sh); const i = sh.cf.findIndex(r => r.id === id), j = i + dir; if (i < 0 || j < 0 || j >= sh.cf.length) return; [sh.cf[i], sh.cf[j]] = [sh.cf[j], sh.cf[i]]; sh.cfVer = (sh.cfVer || 0) + 1; }),
  dvSave: (app, rule) => app.tx('Validasi data', () => {
    const sh = app.sheet; app.wb.hist.touchMeta(sh);
    const i = sh.dv.findIndex(r => r.id === rule.id);
    if (i >= 0) sh.dv[i] = rule; else sh.dv.push(rule);
    if (rule.type === 'checkbox') for (const g of rule.ranges) for (let r = g.r1; r <= Math.min(g.r2, sh.rows - 1); r++) for (let c = g.c1; c <= Math.min(g.c2, sh.cols - 1); c++) { const v = sh.get(r, c); if (!v || v.v == null && v.f == null) sh.put(r, c, { v: rule.unchecked ?? false }); }
  }),
  dvDelete: (app, id) => app.tx('Hapus validasi', () => { const sh = app.sheet; app.wb.hist.touchMeta(sh); sh.dv = sh.dv.filter(r => r.id !== id); }),
  chartAdd: (app, chart) => app.tx('Sisipkan diagram', () => { app.wb.hist.touchMeta(app.sheet); app.sheet.charts.push(chart); }),
  chartUpdate: (app, chart) => app.tx('Edit diagram', () => { app.wb.hist.touchMeta(app.sheet); const i = app.sheet.charts.findIndex(c => c.id === chart.id); if (i >= 0) app.sheet.charts[i] = chart; }),
  chartDelete: (app, id) => app.tx('Hapus diagram', () => { app.wb.hist.touchMeta(app.sheet); app.sheet.charts = app.sheet.charts.filter(c => c.id !== id); }),
  pivotCreate: (app, def, newSheet = true) => {
    let target;
    const i = app.wb.active + 1;
    app.tx('Tabel pivot', () => {
      if (newSheet) { let n = 1; while (app.wb.sheetByName('Tabel Pivot ' + n)) n++; target = app.wb.addSheet('Tabel Pivot ' + n, i); }
      else target = app.sheet;
      app.wb.eng.recalc();
      ops.writePivot(app, target, def);
    });
    if (newSheet) app.switchSheet(i);
    return target;
  },
  pivotRefresh: app => { const sh = app.sheet; if (!sh.pivot) return; app.tx('Perbarui pivot', () => ops.writePivot(app, sh, sh.pivot)); },
  removeDuplicates: (app, g, cols, header) => { const res = app.tx('Hapus duplikat', () => ops.removeDuplicates(app, g, cols, header)); toast(`${res.removed} baris duplikat dihapus; ${res.left} baris unik tersisa`); },
  splitText: (app, delim) => app.tx('Pisahkan teks', () => ops.splitToColumns(app, sel1(app), delim)),
  transposePaste: app => app.pasteFromMenu('transpose'),
  goalSeek: (app, target, goal, changing) => app.tx('Pencarian tujuan', () => ops.goalSeek(app, app.sheet, target, goal, changing)),
  replaceAll: (app, hits, q, rep, opts) => app.tx('Ganti', () => ops.replaceIn(app, hits, q, rep, opts)),
  importSheets: (app, wbNew, mode) => {
    if (mode === 'new') { app.openWorkbook(wbNew); return; }
    app.tx('Impor', () => {
      if (mode === 'replace') { const sh = app.sheet; const src = wbNew.sheets[0]; const j = src.toJSON(); j.name = sh.name; const styleMap = remapStyles(app.wb, wbNew); j.cells = j.cells.map(([k, v, f, s, n, l]) => [k, v, f, styleMap[s] || 0, n, l]); app.wb.hist.touchFull(sh); sh.loadJSON(j); return; }
      const styleMap = remapStyles(app.wb, wbNew);
      for (const src of wbNew.sheets) {
        let name = src.name, n = 2;
        while (app.wb.sheetByName(name)) name = src.name + ' (' + n++ + ')';
        const sh = app.wb.addSheet(name, app.wb.sheets.length, src.rows, src.cols);
        const j = src.toJSON(); j.name = name;
        j.cells = j.cells.map(([k, v, f, s, nn, l]) => [k, v, f, styleMap[s] || 0, nn, l]);
        sh.loadJSON(j);
      }
    });
    if (mode === 'insert') app.switchSheet(app.wb.sheets.length - 1);
  },
};

function remapStyles(dst, src) { return src.styles.map(s => dst.intern({ ...s })); }

// ---------- boot ----------
const app = new App();
window.amatu = app;
app.init().catch(e => { console.error(e); document.body.classList.remove('loading'); alertDlg('Kesalahan', 'Aplikasi gagal dimuat: ' + e.message); });
