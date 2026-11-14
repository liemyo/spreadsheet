// Gambar mengambang di atas sel: tanda tangan / stempel (sheet.images). Dipilih, digeser dan diubah ukurannya seperti
// diagram; ikut disimpan, dibagikan, diurungkan dan dicetak. Posisi = sel jangkar + offset px, jadi tetap menempel saat
// baris/kolom disisipkan. Sheet yang dikunci tetap bisa ditandatangani (kunci menjaga isi sel); tautan Lihat saja tidak.
import { el, icon, uid, download } from './util.js';
import { HH } from './grid.js';
import { openMenu, dialog, toast } from './ui/ui.js';
import { signaturePanel, defaultSize } from './signature.js';

const MIN = 12;
const what = d => (d && d.sig === false ? 'gambar' : 'tanda tangan');   // pictures imported from Excel are plain images

// Sel yang memuat titik (x, y) dalam koordinat lembar (tanpa gulir): pencarian biner di posisi kumulatif kolom/baris.
function cellAt(pos, v, n) {
  let lo = 0, hi = n - 1;
  while (lo < hi) { const m = (lo + hi + 1) >> 1; if (pos[m] <= v) lo = m; else hi = m - 1; }
  return lo;
}

export class FloatImages {
  constructor(app, layer) {
    this.app = app;
    this.layer = layer;
    this.boxes = new Map();
    this.sel = null;
  }
  get sh() { return this.app.sheet; }

  // ---------- tampilan ----------
  reset() {
    for (const b of this.boxes.values()) b.el.remove();
    this.boxes.clear();
    this.update();
  }

  update() {
    const sh = this.sh;
    if (!sh) return;
    const list = sh.images || [];
    const ids = new Set(list.map(i => i.id));
    for (const [id, b] of this.boxes) if (!ids.has(id) || b.sheet !== sh) { b.el.remove(); this.boxes.delete(id); }
    list.forEach((im, i) => {
      let b = this.boxes.get(im.id);
      if (!b) { b = this.make(im); this.boxes.set(im.id, b); }
      b.def = im;
      if (b.img.getAttribute('src') !== im.src) b.img.setAttribute('src', im.src);
      b.el.style.zIndex = String(i + 1);
      b.el.classList.toggle('fimg-ink', im.sig !== false);   // not 'sig': that is the formula hint popup
    });
    if (this.sel && !ids.has(this.sel)) this.sel = null;
    for (const [id, b] of this.boxes) b.el.classList.toggle('selected', id === this.sel);
    this.position();
  }

  make(im) {
    const img = el('img', { alt: im.name || 'Tanda tangan', draggable: false });
    const menuBtn = el('button.icon-btn.sm.fimg-menu', { type: 'button', 'aria-label': 'Opsi tanda tangan', 'data-tip': 'Opsi', onclick: e => { e.stopPropagation(); this.menu(b.def.id, e.currentTarget.getBoundingClientRect()); } }, icon('more_vert'));
    const box = el('div.fimg-box', { tabIndex: -1, role: 'img', 'aria-label': im.name || 'Tanda tangan', title: im.name || '' }, img, menuBtn,
      ...['nw', 'ne', 'se', 'sw'].map(h => el('span.ch-h.h-' + h, { dataset: { h } })));
    const b = { el: box, img, def: im, sheet: this.sh };
    box.addEventListener('pointerdown', e => {
      if (e.target.closest('.fimg-menu') || e.button === 2) return;
      this.drag(e, b, e.target.dataset?.h || null);
    });
    box.addEventListener('dblclick', e => { e.preventDefault(); this.replace(b.def.id); });
    box.addEventListener('contextmenu', e => { e.preventDefault(); e.stopPropagation(); this.select(b.def.id); this.menu(b.def.id, { x: e.clientX, y: e.clientY }); });
    box.addEventListener('keydown', e => this.onKey(e, b));
    // roda mouse di atas gambar yang dipilih tetap menggulir lembar
    box.addEventListener('wheel', e => {
      if (e.ctrlKey || e.metaKey) return;
      e.preventDefault();
      const k = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? 400 : 1, sc = this.app.grid.scroller;
      sc.scrollLeft += (e.shiftKey && !e.deltaX ? e.deltaY : e.deltaX) * k;
      if (!e.shiftKey || e.deltaX) sc.scrollTop += e.deltaY * k;
    }, { passive: false });
    this.layer.append(box);
    return b;
  }

  // Koordinat lembar (tanpa gulir) dari gambar: posisi kumulatif sel jangkar + offset.
  abs(d) {
    const g = this.app.grid.geo(), sh = this.sh;
    const r = Math.min(d.anchor.r, sh.rows - 1), c = Math.min(d.anchor.c, sh.cols - 1);
    return { x: g.colX[c] + d.dx, y: g.rowY[r] + d.dy };
  }
  // Jangkar baru dari koordinat lembar (dipotong ke >= 0 agar tidak keluar dari lembar).
  anchorFor(x, y, d) {
    const g = this.app.grid.geo(), sh = this.sh;
    x = Math.max(0, x); y = Math.max(0, y);
    const c = cellAt(g.colX, x, sh.cols), r = cellAt(g.rowY, y, sh.rows);
    return { ...d, anchor: { r, c }, dx: x - g.colX[c], dy: y - g.rowY[r] };
  }

  place(b, d) {
    const g = this.app.grid, z = g.z, sh = this.sh, geo = g.geo();
    const r = Math.min(d.anchor.r, sh.rows - 1), c = Math.min(d.anchor.c, sh.cols - 1);
    const x = g.X(c) + d.dx, y = g.Y(r) + d.dy;
    Object.assign(b.el.style, { left: x * z + 'px', top: y * z + 'px', width: d.w * z + 'px', height: d.h * z + 'px' });
    // gambar di area yang bergulir tidak boleh menutupi baris/kolom yang dibekukan
    const top = r >= sh.freezeR ? HH + g.fh() : HH, left = c >= sh.freezeC ? geo.hw + g.fw() : geo.hw;
    const ct = Math.max(0, top - y), cl = Math.max(0, left - x);
    b.el.style.clipPath = ct || cl ? `inset(${ct * z}px 0 0 ${cl * z}px)` : '';
  }

  position() {
    if (!this.boxes.size) return;
    const g = this.app.grid, geo = g.geo();
    this.layer.style.clipPath = `inset(${HH * g.z}px 0 0 ${geo.hw * g.z}px)`;
    for (const b of this.boxes.values()) this.place(b, b.def);
  }

  // Gambar teratas di titik grid (x, y) — koordinat yang sama dengan Grid.hit().
  at(x, y) {
    const sh = this.sh, list = sh?.images;
    if (!list?.length) return null;
    const g = this.app.grid, inFrozenR = y < HH + g.fh(), inFrozenC = x < g.geo().hw + g.fw();
    for (let i = list.length - 1; i >= 0; i--) {
      const d = list[i], r = Math.min(d.anchor.r, sh.rows - 1), c = Math.min(d.anchor.c, sh.cols - 1);
      if ((inFrozenR && r >= sh.freezeR) || (inFrozenC && c >= sh.freezeC)) continue;   // hidden under the frozen panes
      const x0 = g.X(c) + d.dx, y0 = g.Y(r) + d.dy;
      if (x >= x0 && x <= x0 + d.w && y >= y0 && y <= y0 + d.h) return d.id;
    }
    return null;
  }

  select(id, focus = true) {
    this.sel = id || null;
    for (const [k, b] of this.boxes) b.el.classList.toggle('selected', k === this.sel);
    if (this.sel && focus) this.boxes.get(this.sel)?.el.focus({ preventScroll: true });
  }

  // Tekan dari grid (gambar belum dipilih): pilih; mouse langsung bisa menggeser.
  press(e, id, touch) {
    this.app.closePopups?.();
    if (this.app.grid.edit && !this.app.grid.commitEdit(0, 0)) return;
    this.select(id);
    const b = this.boxes.get(id);
    if (!touch && b && !this.app.readOnly) this.drag(e, b, null, true);
  }

  // Geser (tanpa pegangan) atau ubah ukuran dari sudut (rasio terjaga; Shift = bebas).
  drag(e, b, handle, fromGrid = false) {
    e.preventDefault();
    e.stopPropagation();
    this.app.closePopups?.();
    if (this.sel !== b.def.id) this.select(b.def.id);
    else b.el.focus({ preventScroll: true });
    if (this.app.readOnly) return;
    const g = this.app.grid, p0 = g.pos(e), d0 = { ...b.def, anchor: { ...b.def.anchor } }, a0 = this.abs(d0);
    const ratio = d0.w / d0.h;
    let cur = d0, moved = false;
    const move = ev => {
      const p = g.pos(ev), dx = p.x - p0.x, dy = p.y - p0.y;
      if (!moved && Math.hypot(dx, dy) * g.z < 3) return;
      moved = true;
      let x = a0.x, y = a0.y, w = d0.w, hh = d0.h;
      if (!handle) { x += dx; y += dy; }
      else {
        w = d0.w + (handle.includes('e') ? dx : -dx);
        hh = d0.h + (handle.includes('s') ? dy : -dy);
        if (!ev.shiftKey) { if (Math.abs(w - d0.w) * d0.h >= Math.abs(hh - d0.h) * d0.w) hh = w / ratio; else w = hh * ratio; }
        if (w < MIN) { w = MIN; if (!ev.shiftKey) hh = w / ratio; }
        if (hh < MIN) { hh = MIN; if (!ev.shiftKey) w = hh * ratio; }
        if (handle.includes('w')) x = a0.x + d0.w - w;
        if (handle.includes('n')) y = a0.y + d0.h - hh;
      }
      cur = this.anchorFor(x, y, { ...d0, w: Math.round(w * 10) / 10, h: Math.round(hh * 10) / 10 });
      this.place(b, cur);
    };
    const up = () => {
      document.removeEventListener('pointermove', move);
      document.removeEventListener('pointerup', up);
      document.removeEventListener('pointercancel', up);
      if (!moved) return;
      this.app.cmd('imgUpdate', cur, (handle ? 'Ubah ukuran ' : 'Pindahkan ') + what(d0));
    };
    document.addEventListener('pointermove', move);
    document.addEventListener('pointerup', up);
    document.addEventListener('pointercancel', up);
    if (fromGrid) this.app.grid.drag = null;
  }

  onKey(e, b) {
    const mod = e.ctrlKey || e.metaKey, k = e.key;
    const stop = () => { e.preventDefault(); e.stopPropagation(); };
    if (k === 'Escape') { stop(); this.select(null); this.app.grid.focus(); return; }
    if (mod && (k === 'z' || k === 'Z')) { stop(); e.shiftKey ? this.app.redo() : this.app.undo(); return; }
    if (mod && (k === 'y' || k === 'Y')) { stop(); this.app.redo(); return; }
    if (this.app.readOnly) { if (!mod && k.length === 1) stop(); return; }
    if (k === 'Delete' || k === 'Backspace') { stop(); this.remove(b.def.id); return; }
    if (mod && (k === 'd' || k === 'D')) { stop(); this.duplicate(b.def.id); return; }
    if (k === 'Enter') { stop(); this.replace(b.def.id); return; }
    const nav = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[k];
    if (nav) {
      stop();
      const step = e.shiftKey ? 10 : 1, a = this.abs(b.def);
      this.app.cmd('imgUpdate', this.anchorFor(a.x + nav[0] * step, a.y + nav[1] * step, b.def), 'Pindahkan ' + what(b.def));
      return;
    }
    if (k === 'Tab') { stop(); this.select(null); this.app.grid.focus(); return; }
    // typing while a signature is selected must not land in the cell under it
    if (!mod && !e.altKey && k.length === 1) stop();
  }

  menu(id, at) {
    const d = this.sh.images.find(x => x.id === id);
    if (!d) return;
    this.select(id, false);
    const ro = this.app.readOnly, many = this.sh.images.length > 1;
    openMenu([
      { label: `Ganti ${what(d)}…`, icon: 'draw', disabled: ro, run: () => this.replace(id) },
      { label: 'Duplikat', icon: 'content_copy', kbd: `${this.app.MOD}+D`, disabled: ro, run: () => this.duplicate(id) },
      { label: 'Ukuran awal', icon: 'aspect_ratio', disabled: ro, run: () => { const s = defaultSize(d); this.app.cmd('imgUpdate', { ...d, ...s }, 'Ubah ukuran ' + what(d)); } },
      many ? { label: 'Bawa ke depan', icon: 'flip_to_front', disabled: ro, run: () => this.app.cmd('imgOrder', id, 1) } : null,
      many ? { label: 'Kirim ke belakang', icon: 'flip_to_back', disabled: ro, run: () => this.app.cmd('imgOrder', id, -1) } : null,
      { label: 'Download gambar', icon: 'download', run: () => fetch(d.src).then(r => r.blob()).then(blob => download((d.name || 'Tanda tangan') + '.' + ((blob.type.split('/')[1] || 'png').replace('jpeg', 'jpg')), blob)) },
      { sep: true },
      { label: 'Hapus', icon: 'delete', kbd: 'Delete', disabled: ro, run: () => this.remove(id) },
    ], at, { alignRight: true });
  }

  // ---------- tindakan ----------
  remove(id) { this.app.cmd('imgDelete', id); this.select(null, false); this.app.grid.focus(); }
  duplicate(id) {
    const d = this.sh.images.find(x => x.id === id);
    if (!d) return;
    const a = this.abs(d), c = { ...this.anchorFor(a.x + 16, a.y + 16, d), id: uid('im') };
    this.app.cmd('imgAdd', c, 'Duplikat ' + what(d));
    this.select(c.id);
  }

  // Dialog tanda tangan: tersimpan / gambar / ketik / unggah. onUse menerima {url, w, h}.
  dialog(title, onUse) {
    let dlg = null, btn = null, used = false;
    const use = sig => { if (used || !sig) return; used = true; dlg?.close(); onUse(sig); };
    const panel = signaturePanel({ toast, onReady: ok => { if (btn) btn.disabled = !ok; }, onPick: use });
    dlg = dialog({
      title, width: 'min(640px, calc(100vw - 32px))', body: panel.el, onClose: () => panel.destroy(),
      buttons: [{ label: 'Batal' }, { label: title.startsWith('Ganti') ? 'Ganti' : 'Sisipkan', primary: true, run: async () => {
        const sig = await panel.result();
        if (!sig) { toast('Buat tanda tangan dulu'); return false; }
        used = true;
        setTimeout(() => onUse(sig), 0);   // after the dialog closed and handed focus back to the grid
      } }],
    });
    btn = dlg.el.querySelector('.dlg-foot .btn.primary');
    if (btn) btn.disabled = true;
    setTimeout(() => { if (!dlg.el.contains(document.activeElement)) panel.focus(); }, 60);
  }

  insertDialog() {
    if (this.app.readOnly) { toast('Anda hanya dapat melihat spreadsheet ini.'); return; }
    if (this.app.grid.edit && !this.app.grid.commitEdit(0, 0)) return;
    const sh = this.sh, rg = { ...this.app.grid.range };
    this.dialog('Sisipkan tanda tangan', sig => { if (this.app.sheet === sh) this.insert(sig, rg); });
  }

  // Satu sel: tanda tangan ±3 baris tingginya, mulai dari sel itu. Rentang (mis. B80:C83): dimuat & ditengahkan di dalamnya.
  insert(sig, rg) {
    const sh = this.sh, g = this.app.grid.geo();
    let { r1, c1, r2, c2 } = rg;
    const m = sh.mergeAt(r1, c1);
    if (m && r1 === r2 && c1 === c2) ({ r1, c1, r2, c2 } = m);
    r2 = Math.min(r2, sh.rows - 1); c2 = Math.min(c2, sh.cols - 1);
    let { w, h } = defaultSize(sig);
    let x = g.colX[c1] + 4, y = g.rowY[r1] + 2;
    const bw = g.colX[c2 + 1] - g.colX[c1], bh = g.rowY[r2 + 1] - g.rowY[r1];
    const block = (r2 > r1 || c2 > c1) && bw > 24 && bh > 16 && bw < 2000 && bh < 1200;
    if (block) {
      const k = Math.min((bw - 8) / w, (bh - 6) / h, 2.5);
      w *= k; h *= k;
      if (w < MIN || h < MIN) { const k2 = Math.max(MIN / w, MIN / h); w *= k2; h *= k2; }
      x = g.colX[c1] + (bw - w) / 2; y = g.rowY[r1] + (bh - h) / 2;
    }
    const d = this.anchorFor(x, y, { id: uid('im'), name: 'Tanda tangan', src: sig.url, w: Math.round(w), h: Math.round(h), sig: true });
    if (this.app.cmd('imgAdd', d, 'Sisipkan tanda tangan') === false) return;
    if (!sh.images.some(i => i.id === d.id)) return;   // ditolak (mis. tautan lihat saja)
    this.select(d.id);
    toast('Tanda tangan disisipkan. Geser untuk memindahkan, tarik sudutnya untuk mengubah ukuran.');
  }

  replace(id) {
    if (this.app.readOnly) return;
    const sh = this.sh, d0 = sh.images.find(x => x.id === id);
    if (!d0) return;
    this.dialog('Ganti tanda tangan', sig => {
      const d = sh.images.find(x => x.id === id);
      if (!d || this.app.sheet !== sh) return;
      const h = d.h, w = Math.round(h * (sig.w / sig.h));
      this.app.cmd('imgUpdate', { ...d, src: sig.url, w, h, sig: true }, 'Ganti tanda tangan');
      this.select(id);
    });
  }
}
