// Right-hand side panels: conditional formatting, validation, charts, pivot, names, versions, banding.
import { el, icon, rangeAddr, parseRange, colName, uid, addr, quoteSheet } from '../util.js';
import { select, checkbox, field, segmented, pickColor, toast, confirmDlg, promptDlg } from './ui.js';
import { CHIP_COLORS } from '../cf.js';
import * as store from '../store.js';
import { Workbook } from '../model.js';
import { CHART_TYPES } from '../charts.js';
import * as io from '../io.js';

const rangesText = rs => rs.map(g => rangeAddr(g)).join(', ');
function parseRanges(app, text) {
  const out = [];
  for (const part of text.split(/[,;]/)) {
    const g = parseRange(part.trim());
    if (!g) return null;
    const sh = app.sheet;
    out.push({ r1: g.r1, c1: g.c1, r2: g.r2 === Infinity ? sh.rows - 1 : g.r2, c2: g.c2 === Infinity ? sh.cols - 1 : g.c2 });
  }
  return out.length ? out : null;
}

function rangeField(app, value, onChange) {
  const inp = el('input.inp', { value, 'aria-label': 'Rentang' });
  inp.addEventListener('change', () => onChange(inp.value));
  const pick = el('button.icon-btn', { type: 'button', 'data-tip': 'Gunakan seleksi saat ini', onclick: () => { inp.value = rangesText(app.grid.sel.ranges); onChange(inp.value); } }, icon('grid_on'));
  return el('div.range-field', inp, pick);
}

function styleChip(st) {
  return el('span.style-chip', { style: { background: st.bg || 'transparent', color: st.color || 'inherit', fontWeight: st.b ? 700 : 400, fontStyle: st.i ? 'italic' : 'normal', textDecoration: [st.u ? 'underline' : '', st.s ? 'line-through' : ''].join(' ') } }, '123');
}

const CF_CONDS = [
  { group: 'Sel', items: [{ value: 'empty', label: 'Sel kosong' }, { value: 'notEmpty', label: 'Sel tidak kosong' }, { value: 'error', label: 'Berisi error' }, { value: 'noError', label: 'Tidak berisi error' }] },
  { group: 'Teks', items: [{ value: 'text:contains', label: 'Teks berisi' }, { value: 'text:notContains', label: 'Teks tidak berisi' }, { value: 'text:starts', label: 'Teks diawali dengan' }, { value: 'text:ends', label: 'Teks diakhiri dengan' }, { value: 'text:exact', label: 'Teks sama persis' }] },
  { group: 'Tanggal', items: [{ value: 'date:is', label: 'Tanggal adalah' }, { value: 'date:before', label: 'Tanggal sebelum' }, { value: 'date:after', label: 'Tanggal setelah' }] },
  { group: 'Angka', items: [{ value: 'num:gt', label: 'Lebih besar dari' }, { value: 'num:ge', label: 'Lebih besar dari atau sama dengan' }, { value: 'num:lt', label: 'Kurang dari' }, { value: 'num:le', label: 'Kurang dari atau sama dengan' }, { value: 'num:eq', label: 'Sama dengan' }, { value: 'num:ne', label: 'Tidak sama dengan' }, { value: 'num:between', label: 'Di antara' }, { value: 'num:notBetween', label: 'Tidak di antara' }] },
  { group: 'Excel', items: [{ value: 'dup', label: 'Nilai duplikat' }, { value: 'unique', label: 'Nilai unik' }, { value: 'top', label: 'N teratas' }, { value: 'bottom', label: 'N terbawah' }, { value: 'above', label: 'Di atas rata-rata' }, { value: 'below', label: 'Di bawah rata-rata' }] },
  { group: 'Lanjutan', items: [{ value: 'formula', label: 'Rumus kustom adalah' }] },
];
const STYLE_PRESETS = [
  { bg: '#b7e1cd', color: '#000000' }, { bg: '#fce8b2', color: '#000000' }, { bg: '#f4c7c3', color: '#000000' },
  { bg: '#c6efce', color: '#006100' }, { bg: '#ffeb9c', color: '#9c5700' }, { bg: '#ffc7ce', color: '#9c0006' },
  { color: '#d93025', b: true }, { color: '#188038', b: true }, { bg: '#d2e3fc', color: '#174ea6' },
];
const SCALE_PRESETS = [
  ['#57bb8a', null, '#ffffff'], ['#ffffff', null, '#57bb8a'], ['#e67c73', '#ffd666', '#57bb8a'], ['#57bb8a', '#ffd666', '#e67c73'],
  ['#ffffff', null, '#e67c73'], ['#ffffff', null, '#4285f4'], ['#f8696b', '#ffeb84', '#63be7b'], ['#5a8ac6', '#fcfcff', '#f8696b'],
];
const BAND_PRESETS = [
  { h: '#5b95f9', a: '#ffffff', b: '#e8f0fe', ht: '#ffffff' }, { h: '#bdbdbd', a: '#ffffff', b: '#f3f3f3' }, { h: '#4dd0e1', a: '#ffffff', b: '#e0f7fa' },
  { h: '#63d297', a: '#ffffff', b: '#e7f9ef' }, { h: '#f7cb4d', a: '#ffffff', b: '#fef8e3' }, { h: '#f46524', a: '#ffffff', b: '#ffe6dd', ht: '#ffffff' },
  { h: '#5b3f86', a: '#ffffff', b: '#ece7f4', ht: '#ffffff' }, { h: '#e94b3c', a: '#ffffff', b: '#fde8e6', ht: '#ffffff' },
  { h: '#4472c4', a: '#d9e1f2', b: '#ffffff', ht: '#ffffff' }, { h: '#70ad47', a: '#e2efda', b: '#ffffff', ht: '#ffffff' }, { h: '#ed7d31', a: '#fce4d6', b: '#ffffff', ht: '#ffffff' },
  { h: '#262626', a: '#f2f2f2', b: '#ffffff', ht: '#ffffff' },
];

export class Sidebar {
  constructor(app, host) {
    this.app = app;
    this.host = host;
    this.kind = null;
  }
  close() {
    this.kind = null;
    this.host.innerHTML = '';
    document.body.classList.remove('side-open');
    this.app.grid?.resize();
    this.refresh = null; this.onSel = null; this.onSheet = null;
  }
  frame(title, body, foot, back) {
    this.host.innerHTML = '';
    this.host.append(el('div.side-head',
      back ? el('button.icon-btn', { type: 'button', 'aria-label': 'Kembali', onclick: back }, icon('arrow_back')) : null,
      el('h3', title),
      el('button.icon-btn', { type: 'button', 'aria-label': 'Tutup', onclick: () => this.close() }, icon('close'))),
    el('div.side-body', body), ...(foot ? [el('div.side-foot', foot)] : []));
    if (!document.body.classList.contains('side-open')) { document.body.classList.add('side-open'); this.app.grid.resize(); }
  }
  open(kind, opts = {}) {
    if (this.app.grid.edit) this.app.grid.commitEdit(0, 0);
    this.kind = kind;
    this.refresh = null; this.onSel = null; this.onSheet = null;
    ({ cf: () => this.cfList(), dv: () => (opts.newRule ? this.dvEdit(null) : this.dvList()), chart: () => this.chartEdit(opts.chart || null, opts.create), pivot: () => this.pivot(opts.sheet), names: () => this.names(opts.define), versions: () => this.versions(), banding: () => this.banding() })[kind]?.();
  }

  // ---------- conditional formatting ----------
  cfList() {
    const app = this.app, sh = app.sheet;
    const list = el('div.rule-list');
    const describe = r => {
      const all = CF_CONDS.flatMap(g => g.items);
      if (r.type === 'scale') return 'Skala warna';
      if (r.type === 'bar') return 'Bilah data';
      if (r.type === 'icons') return 'Set ikon';
      const k = ['text', 'date', 'num'].includes(r.type) ? `${r.type}:${r.op}` : r.type;
      const lb = all.find(x => x.value === k)?.label || r.type;
      return `${lb}${r.v1 != null && r.v1 !== '' ? ' ' + r.v1 : ''}${r.v2 != null && r.v2 !== '' ? ' dan ' + r.v2 : ''}`;
    };
    const draw = () => {
      list.innerHTML = '';
      const rules = app.sheet.cf;
      if (!rules.length) list.append(el('div.empty.sm', icon('format_color_fill'), el('span', 'Belum ada aturan format bersyarat di sheet ini')));
      rules.forEach((r, i) => {
        const prev = r.type === 'scale' ? el('span.style-chip.scale', { style: { background: `linear-gradient(90deg, ${r.scale.min.color}, ${r.scale.mid ? r.scale.mid.color + ',' : ''} ${r.scale.max.color})` } }) :
          r.type === 'bar' ? el('span.style-chip.scale', { style: { background: `linear-gradient(90deg, ${r.bar.color} 60%, transparent 60%)` } }) :
          r.type === 'icons' ? el('span.style-chip', icon({ arrows: 'trending_up', traffic: 'circle', flags: 'flag', stars: 'star' }[r.icons.set] || 'trending_up')) : styleChip(r.style || {});
        list.append(el('div.rule', { onclick: () => this.cfEdit(r) },
          prev,
          el('div.rule-main', el('div.rule-title', describe(r)), el('div.muted.small', rangesText(r.ranges))),
          el('div.rule-act',
            el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Naikkan', disabled: i === 0, onclick: e => { e.stopPropagation(); app.cmd('cfMove', r.id, -1); draw(); } }, icon('arrow_upward')),
            el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Hapus aturan', onclick: e => { e.stopPropagation(); app.cmd('cfDelete', r.id); draw(); } }, icon('delete')))));
      });
    };
    draw();
    this.frame('Aturan format bersyarat', el('div.stack', list,
      el('button.btn.add-rule', { type: 'button', onclick: () => this.cfEdit(null) }, icon('add'), 'Tambahkan aturan lain')));
    this.onSheet = () => this.cfList();
  }

  cfEdit(rule) {
    const app = this.app;
    const isNew = !rule;
    const r = rule ? JSON.parse(JSON.stringify(rule)) : { id: uid('cf'), ranges: app.ranges(), type: 'notEmpty', style: { ...STYLE_PRESETS[0] } };
    if (!rule && app.grid.range.r1 === app.grid.range.r2 && app.grid.range.c1 === app.grid.range.c2) r.ranges = [io.currentRegion(app, app.grid.sel.r, app.grid.sel.c)];
    let mode = ['scale', 'bar', 'icons'].includes(r.type) ? r.type : 'single';
    const body = el('div.stack');
    const save = () => { if (!r.ranges?.length) return; app.cmd('cfSave', r); };
    const rangeEl = field('Terapkan ke rentang', rangeField(app, rangesText(r.ranges), t => { const rs = parseRanges(app, t); if (rs) { r.ranges = rs; save(); } else toast('Rentang tidak valid'); }));
    const tabs = segmented([{ value: 'single', label: 'Satu warna' }, { value: 'scale', label: 'Skala warna' }, { value: 'bar', label: 'Bilah data' }, { value: 'icons', label: 'Set ikon' }], mode, v => { mode = v; setType(); draw(); save(); });
    tabs.classList.add('full');
    const content = el('div.stack');
    const setType = () => {
      if (mode === 'scale') { r.type = 'scale'; r.scale = r.scale || { min: { type: 'min', color: '#ffffff' }, max: { type: 'max', color: '#57bb8a' } }; }
      else if (mode === 'bar') { r.type = 'bar'; r.bar = r.bar || { color: '#638ec6' }; }
      else if (mode === 'icons') { r.type = 'icons'; r.icons = r.icons || { set: 'arrows' }; }
      else if (['scale', 'bar', 'icons'].includes(r.type)) { r.type = 'notEmpty'; r.style = r.style || { ...STYLE_PRESETS[0] }; }
    };
    const draw = () => {
      content.innerHTML = '';
      if (mode === 'single') {
        const k = ['text', 'date', 'num'].includes(r.type) ? `${r.type}:${r.op}` : r.type;
        const cond = select(CF_CONDS, k, v => { const [t, op] = v.split(':'); r.type = t; r.op = op; if (t === 'top' || t === 'bottom') r.v1 = r.v1 || '10'; draw(); save(); });
        content.append(field('Format sel jika…', cond));
        const needs1 = ['text', 'num', 'formula', 'top', 'bottom', 'date'].includes(r.type);
        const needs2 = r.type === 'num' && (r.op === 'between' || r.op === 'notBetween');
        if (r.type === 'date') {
          const dsel = select([{ value: 'today', label: 'Hari ini' }, { value: 'yesterday', label: 'Kemarin' }, { value: 'tomorrow', label: 'Besok' }, ...(r.op === 'is' ? [{ value: 'week', label: 'Dalam seminggu terakhir' }, { value: 'month', label: 'Dalam sebulan terakhir' }, { value: 'year', label: 'Dalam setahun terakhir' }] : []), { value: '', label: 'Tanggal tertentu…' }], ['today', 'yesterday', 'tomorrow', 'week', 'month', 'year'].includes(r.v1) ? r.v1 : '', v => { r.v1 = v || ''; draw(); save(); });
          content.append(dsel);
          if (!['today', 'yesterday', 'tomorrow', 'week', 'month', 'year'].includes(r.v1)) { const i = el('input.inp', { value: r.v1 || '', placeholder: 'mis. 31/12/2026' }); i.addEventListener('change', () => { r.v1 = i.value; save(); }); content.append(i); }
        } else if (needs1) {
          const i1 = el('input.inp', { value: r.v1 ?? '', placeholder: r.type === 'formula' ? 'mis. =$B2>100' : r.type === 'top' || r.type === 'bottom' ? 'Jumlah (N)' : 'Nilai atau rumus' });
          i1.addEventListener('change', () => { r.v1 = i1.value; save(); });
          content.append(i1);
          if (r.type === 'top' || r.type === 'bottom') content.append(checkbox('Sebagai persentase', !!r.pct, v => { r.pct = v; save(); }));
          if (r.type === 'formula') content.append(el('div.muted.small', 'Rumus relatif terhadap sel kiri atas rentang. Gunakan $ untuk mengunci kolom atau baris.'));
        }
        if (needs2) { const i2 = el('input.inp', { value: r.v2 ?? '', placeholder: 'Nilai kedua' }); i2.addEventListener('change', () => { r.v2 = i2.value; save(); }); content.append(i2); }
        // style
        const st = r.style || (r.style = {});
        const prev = styleChip(st);
        const tog = (k, ic, tip) => { const b = el('button.tb-btn' + (st[k] ? '.on' : ''), { type: 'button', 'data-tip': tip, onclick: () => { st[k] = !st[k] || undefined; b.classList.toggle('on', !!st[k]); refreshPrev(); save(); } }, icon(ic)); return b; };
        const colorBtn = (k, ic, tip) => el('button.tb-btn', { type: 'button', 'data-tip': tip, onclick: e => pickColor(e.currentTarget.getBoundingClientRect(), c => { st[k] = c || undefined; refreshPrev(); save(); }) }, icon(ic));
        const refreshPrev = () => prev.replaceWith(Object.assign(styleChip(st), {}));
        content.append(el('div.side-label', 'Gaya pemformatan'),
          el('div.presets', ...STYLE_PRESETS.map(p => el('button.preset', { type: 'button', style: { background: p.bg || 'transparent', color: p.color, fontWeight: p.b ? 700 : 400 }, onclick: () => { r.style = { ...p }; draw(); save(); } }, '123'))),
          el('div.row.tight', prev, tog('b', 'format_bold', 'Tebal'), tog('i', 'format_italic', 'Miring'), tog('u', 'format_underlined', 'Garis bawah'), tog('s', 'strikethrough_s', 'Coret'), colorBtn('color', 'format_color_text', 'Warna teks'), colorBtn('bg', 'format_color_fill', 'Warna isian')));
        content.append(checkbox('Hentikan jika benar (Excel)', !!r.stop, v => { r.stop = v; save(); }));
      } else if (mode === 'scale') {
        const s = r.scale;
        content.append(el('div.side-label', 'Pratinjau'), el('div.presets', ...SCALE_PRESETS.map(([a, m, b]) => el('button.preset.wide', { type: 'button', style: { background: `linear-gradient(90deg, ${a}, ${m ? m + ',' : ''} ${b})` }, onclick: () => { s.min.color = a; s.max.color = b; if (m) s.mid = { type: 'pctl', v: 50, ...(s.mid || {}), color: m }; else delete s.mid; draw(); save(); } }))));
        const point = (key, label, types) => {
          const p = s[key] || { type: types[0].value, v: '', color: '#ffd666' };
          const enabled = !!s[key];
          const typeSel = select(types, enabled ? p.type : 'none', v => { if (v === 'none') delete s[key]; else s[key] = { ...p, type: v }; draw(); save(); });
          const val = el('input.inp.sm', { value: p.v ?? '', placeholder: 'Nilai', disabled: !enabled || ['min', 'max'].includes(p.type) });
          val.addEventListener('change', () => { s[key].v = val.value; save(); });
          const col = el('input.inp.color', { type: 'color', value: p.color || '#ffffff', disabled: !enabled });
          col.addEventListener('input', () => { s[key].color = col.value; save(); });
          return el('div.scale-pt', el('span.side-label', label), el('div.row.tight', typeSel, val, col));
        };
        const base = [{ value: 'num', label: 'Angka' }, { value: 'pct', label: 'Persen' }, { value: 'pctl', label: 'Persentil' }];
        content.append(point('min', 'Titik minimum', [{ value: 'min', label: 'Nilai min' }, ...base]), point('mid', 'Titik tengah', [{ value: 'none', label: 'Tidak ada' }, ...base]), point('max', 'Titik maksimum', [{ value: 'max', label: 'Nilai maks' }, ...base]));
      } else if (mode === 'bar') {
        const col = el('input.inp.color', { type: 'color', value: r.bar.color });
        col.addEventListener('input', () => { r.bar.color = col.value; save(); });
        content.append(field('Warna bilah', col), checkbox('Isian solid (tanpa gradasi)', !!r.bar.solid, v => { r.bar.solid = v; save(); }), el('div.muted.small', 'Nilai negatif ditampilkan dengan warna merah.'));
      } else {
        content.append(field('Gaya ikon', select([{ value: 'arrows', label: '3 Panah' }, { value: 'traffic', label: '3 Lampu lalu lintas' }, { value: 'flags', label: '3 Bendera' }, { value: 'stars', label: '3 Bintang' }], r.icons.set, v => { r.icons.set = v; save(); })),
          checkbox('Balik urutan ikon', !!r.icons.reverse, v => { r.icons.reverse = v; save(); }),
          el('div.muted.small', 'Ikon dibagi menjadi tiga kelompok: ≥ 67%, ≥ 33%, dan sisanya.'));
      }
    };
    setType();
    draw();
    body.append(rangeEl, tabs, content);
    if (isNew) save();
    this.frame('Format bersyarat', body,
      el('div.row', el('button.btn', { type: 'button', onclick: () => { app.cmd('cfDelete', r.id); this.cfList(); } }, isNew ? 'Batal' : 'Hapus'), el('span.flex1'), el('button.btn.primary', { type: 'button', onclick: () => { save(); this.cfList(); } }, 'Selesai')),
      () => this.cfList());
  }

  // ---------- data validation ----------
  dvList() {
    const app = this.app, sh = app.sheet;
    const names = { list: 'Dropdown', range: 'Dropdown (dari rentang)', checkbox: 'Kotak centang', number: 'Angka', date: 'Tanggal', textlen: 'Panjang teks', text: 'Teks', custom: 'Rumus kustom' };
    const list = el('div.rule-list');
    if (!sh.dv.length) list.append(el('div.empty.sm', icon('rule'), el('span', 'Belum ada aturan validasi data')));
    for (const r of sh.dv) list.append(el('div.rule', { onclick: () => this.dvEdit(r) },
      el('span.style-chip', icon(r.type === 'checkbox' ? 'check_box' : r.type === 'list' || r.type === 'range' ? 'arrow_drop_down_circle' : 'rule')),
      el('div.rule-main', el('div.rule-title', names[r.type] + (r.type === 'list' ? ': ' + (r.items || []).slice(0, 4).join(', ') : r.type === 'range' ? ': ' + r.src : '')), el('div.muted.small', rangesText(r.ranges))),
      el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Hapus aturan', onclick: e => { e.stopPropagation(); app.cmd('dvDelete', r.id); this.dvList(); } }, icon('delete'))));
    this.frame('Aturan validasi data', el('div.stack', list, el('button.btn.add-rule', { type: 'button', onclick: () => this.dvEdit(null) }, icon('add'), 'Tambahkan aturan')));
    this.onSheet = () => this.dvList();
  }

  dvEdit(rule) {
    const app = this.app;
    const r = rule ? JSON.parse(JSON.stringify(rule)) : { id: uid('dv'), ranges: app.ranges(), type: 'list', items: ['Opsi 1', 'Opsi 2'], colors: [CHIP_COLORS[1], CHIP_COLORS[2]], strict: true, chip: 'chip' };
    if (!rule) {
      // seed items from existing values in the selection
      const vals = new Set();
      const sh = app.sheet;
      for (const g of r.ranges) for (let rr = g.r1; rr <= Math.min(g.r2, sh.dataRows() - 1); rr++) for (let c = g.c1; c <= Math.min(g.c2, sh.cols - 1); c++) { const v = app.wb.eng.value(sh, rr, c); if (v != null && v !== '' && vals.size < 30) vals.add(String(v)); }
      if (vals.size) { r.items = [...vals]; r.colors = r.items.map((_, i) => CHIP_COLORS[(i % (CHIP_COLORS.length - 1)) + 1]); }
    }
    const content = el('div.stack');
    const draw = () => {
      content.innerHTML = '';
      const crit = select([
        { value: 'list', label: 'Dropdown' }, { value: 'range', label: 'Dropdown (dari rentang)' }, { value: 'checkbox', label: 'Kotak centang' },
        { value: 'number', label: 'Angka' }, { value: 'date', label: 'Tanggal' }, { value: 'textlen', label: 'Panjang teks' }, { value: 'text', label: 'Teks' }, { value: 'custom', label: 'Rumus kustom' },
      ], r.type, v => { r.type = v; if (['number', 'date', 'textlen'].includes(v)) r.op = r.op || 'between'; if (v === 'text') r.op = 'contains'; draw(); });
      content.append(field('Kriteria', crit));
      if (r.type === 'list') {
        const items = el('div.stack.sm');
        const drawItems = () => {
          items.innerHTML = '';
          (r.items || []).forEach((it, i) => {
            const inp = el('input.inp.sm', { value: it });
            inp.addEventListener('input', () => (r.items[i] = inp.value));
            const sw = el('button.swatch-btn', { type: 'button', style: { background: r.colors?.[i] || CHIP_COLORS[0] }, onclick: e => pickColor(e.currentTarget.getBoundingClientRect(), c => { r.colors = r.colors || []; r.colors[i] = c || CHIP_COLORS[0]; sw.style.background = r.colors[i]; }, { resetLabel: 'Default' }) });
            items.append(el('div.row.tight', sw, inp, el('button.icon-btn.sm', { type: 'button', onclick: () => { r.items.splice(i, 1); r.colors?.splice(i, 1); drawItems(); } }, icon('delete'))));
          });
          items.append(el('button.btn.link', { type: 'button', onclick: () => { r.items.push(''); (r.colors = r.colors || []).push(CHIP_COLORS[(r.items.length % (CHIP_COLORS.length - 1)) + 1]); drawItems(); items.querySelectorAll('input')[r.items.length - 1]?.focus(); } }, icon('add'), 'Tambahkan item lain'));
        };
        drawItems();
        content.append(items);
      }
      if (r.type === 'range') {
        const inp = el('input.inp', { value: r.src || '', placeholder: 'mis. Sheet2!A2:A20' });
        inp.addEventListener('change', () => (r.src = inp.value.replace(/^=/, '')));
        content.append(inp);
      }
      if (r.type === 'list' || r.type === 'range') {
        content.append(checkbox('Izinkan beberapa pilihan', !!r.multi, v => (r.multi = v)),
          field('Gaya tampilan', select([{ value: 'chip', label: 'Chip' }, { value: 'arrow', label: 'Panah' }, { value: 'plain', label: 'Teks biasa' }], r.chip || 'chip', v => (r.chip = v))));
      }
      if (r.type === 'checkbox') {
        let custom = r.checked != null;
        const on = el('input.inp.sm', { value: r.checked ?? 'Ya', placeholder: 'Nilai tercentang' }), off = el('input.inp.sm', { value: r.unchecked ?? 'Tidak', placeholder: 'Nilai tidak tercentang' });
        const wrap = el('div.grid2', on, off);
        wrap.style.display = custom ? '' : 'none';
        content.append(checkbox('Gunakan nilai sel kustom', custom, v => { custom = v; wrap.style.display = v ? '' : 'none'; if (!v) { delete r.checked; delete r.unchecked; } }), wrap);
        on.addEventListener('input', () => custom && (r.checked = on.value));
        off.addEventListener('input', () => custom && (r.unchecked = off.value));
        r._custom = () => { if (custom) { r.checked = on.value; r.unchecked = off.value; } };
      }
      if (['number', 'date', 'textlen'].includes(r.type)) {
        const op = select([{ value: 'between', label: 'di antara' }, { value: 'notBetween', label: 'tidak di antara' }, { value: 'gt', label: 'lebih besar dari' }, { value: 'ge', label: 'lebih besar dari atau sama dengan' }, { value: 'lt', label: 'kurang dari' }, { value: 'le', label: 'kurang dari atau sama dengan' }, { value: 'eq', label: 'sama dengan' }, { value: 'ne', label: 'tidak sama dengan' }], r.op, v => { r.op = v; draw(); });
        const a = el('input.inp.sm', { value: r.v1 ?? '', placeholder: r.type === 'date' ? 'mis. 1/1/2026' : 'Nilai' }), b = el('input.inp.sm', { value: r.v2 ?? '', placeholder: 'Nilai' });
        a.addEventListener('input', () => (r.v1 = a.value)); b.addEventListener('input', () => (r.v2 = b.value));
        content.append(op, el('div.row.tight', a, r.op === 'between' || r.op === 'notBetween' ? b : null));
      }
      if (r.type === 'text') {
        content.append(select([{ value: 'contains', label: 'berisi' }, { value: 'notContains', label: 'tidak berisi' }, { value: 'exact', label: 'sama persis' }, { value: 'email', label: 'adalah email yang valid' }, { value: 'url', label: 'adalah URL yang valid' }], r.op, v => { r.op = v; draw(); }));
        if (!['email', 'url'].includes(r.op)) { const a = el('input.inp.sm', { value: r.v1 ?? '' }); a.addEventListener('input', () => (r.v1 = a.value)); content.append(a); }
      }
      if (r.type === 'custom') { const a = el('input.inp', { value: r.v1 ?? '', placeholder: 'mis. =ISNUMBER(A1)' }); a.addEventListener('input', () => (r.v1 = a.value)); content.append(a); }
      const help = el('input.inp.sm', { value: r.help || '', placeholder: 'Teks bantuan (opsional)' });
      help.addEventListener('input', () => (r.help = help.value));
      content.append(el('details.adv', { open: !!r.help || !r.strict }, el('summary', 'Opsi lanjutan'),
        el('div.stack.sm', help, el('div.side-label', 'Jika data tidak valid'),
          el('label.radio', el('input', { type: 'radio', name: 'dvs', checked: !r.strict, onchange: () => (r.strict = false) }), el('span', 'Tampilkan peringatan')),
          el('label.radio', el('input', { type: 'radio', name: 'dvs', checked: !!r.strict, onchange: () => (r.strict = true) }), el('span', 'Tolak input')))));
    };
    draw();
    const rangeEl = field('Terapkan ke rentang', rangeField(app, rangesText(r.ranges), t => { const rs = parseRanges(app, t); if (rs) r.ranges = rs; else toast('Rentang tidak valid'); }));
    this.frame('Validasi data', el('div.stack', rangeEl, content),
      el('div.row', rule ? el('button.btn', { type: 'button', onclick: () => { app.cmd('dvDelete', r.id); this.dvList(); } }, 'Hapus aturan') : el('button.btn', { type: 'button', onclick: () => this.dvList() }, 'Batal'), el('span.flex1'),
        el('button.btn.primary', { type: 'button', onclick: () => { r._custom?.(); delete r._custom; if (r.type === 'list') { const keep = r.items.map((x, i) => [x.trim(), r.colors?.[i]]).filter(([x]) => x); r.items = keep.map(k => k[0]); r.colors = keep.map(k => k[1]); } app.cmd('dvSave', r); this.dvList(); } }, 'Selesai')),
      () => this.dvList());
  }

  // ---------- charts ----------
  chartEdit(chart, create) {
    const app = this.app, sh = app.sheet;
    let ch = chart ? JSON.parse(JSON.stringify(chart)) : null;
    if (!ch) {
      let g = { ...app.grid.range };
      if (g.r1 === g.r2 && g.c1 === g.c2) g = io.currentRegion(app, g.r1, g.c1);
      if (g.r2 >= sh.rows - 1) g.r2 = Math.max(g.r1, sh.dataRows() - 1);
      const geo = app.grid.geo();
      const anchor = { r: g.r1, c: Math.min(sh.cols - 1, g.c2 + 2) };
      ch = { id: uid('ch'), type: io.guessChart(app, g), range: rangeAddr(g), anchor, dx: 0, dy: 0, w: 520, h: 320, opts: { title: '', legend: 'bottom', headers: 'auto', labels: 'auto' } };
      if (create) { app.cmd('chartAdd', ch); app.charts.select(ch.id); }
    }
    const save = () => app.cmd('chartUpdate', ch);
    const tabs = segmented([{ value: 'setup', label: 'Penyiapan' }, { value: 'custom', label: 'Sesuaikan' }], 'setup', v => draw(v));
    tabs.classList.add('full');
    const content = el('div.stack');
    const draw = (tab = 'setup') => {
      content.innerHTML = '';
      const o = ch.opts;
      if (tab === 'setup') {
        content.append(el('div.side-label', 'Jenis diagram'), el('div.chart-types', ...CHART_TYPES.map(t => el('button.ct' + (ch.type === t.id ? '.on' : ''), { type: 'button', 'data-tip': t.label, onclick: () => { ch.type = t.id; save(); draw(tab); } }, icon(t.icon), el('span', t.label)))));
        content.append(field('Rentang data', rangeField(app, ch.range, t => { if (parseRange(t.split(',')[0])) { ch.range = t; save(); } else toast('Rentang tidak valid'); })));
        content.append(el('div.checks',
          checkbox('Gunakan baris pertama sebagai header', o.headers !== 'off', v => { o.headers = v ? 'on' : 'off'; save(); }),
          checkbox('Gunakan kolom pertama sebagai label', o.labels !== 'off', v => { o.labels = v ? 'on' : 'off'; save(); }),
          checkbox('Tukar baris/kolom', !!o.swap, v => { o.swap = v; save(); }),
          checkbox('Bertumpuk', !!o.stacked, v => { o.stacked = v; save(); })));
      } else {
        const title = el('input.inp', { value: o.title || '', placeholder: 'Judul diagram' });
        title.addEventListener('change', () => { o.title = title.value; save(); });
        const sub = el('input.inp', { value: o.subtitle || '', placeholder: 'Subjudul' });
        sub.addEventListener('change', () => { o.subtitle = sub.value; save(); });
        const xt = el('input.inp.sm', { value: o.xTitle || '', placeholder: 'Judul sumbu horizontal' }), yt = el('input.inp.sm', { value: o.yTitle || '', placeholder: 'Judul sumbu vertikal' });
        xt.addEventListener('change', () => { o.xTitle = xt.value; save(); }); yt.addEventListener('change', () => { o.yTitle = yt.value; save(); });
        const bg = el('input.inp.color', { type: 'color', value: o.bg || '#ffffff' });
        bg.addEventListener('input', () => { o.bg = bg.value; save(); });
        const colorsBox = el('div.stack.sm');
        const series = app.charts.seriesNames(ch);
        series.forEach((nm, i) => {
          const c = el('input.inp.color', { type: 'color', value: (o.colors && o.colors[i]) || app.charts.palette()[i % 10] });
          c.addEventListener('input', () => { o.colors = o.colors || []; o.colors[i] = c.value; save(); });
          colorsBox.append(el('div.row.tight', c, el('span', nm)));
        });
        content.append(field('Judul', title), field('Subjudul', sub),
          field('Legenda', select([{ value: 'top', label: 'Atas' }, { value: 'bottom', label: 'Bawah' }, { value: 'right', label: 'Kanan' }, { value: 'left', label: 'Kiri' }, { value: 'none', label: 'Tidak ada' }], o.legend || 'bottom', v => { o.legend = v; save(); })),
          el('div.grid2', xt, yt),
          el('div.checks',
            checkbox('Label data', !!o.dataLabels, v => { o.dataLabels = v; save(); }),
            checkbox('Garis kisi', o.grid !== false, v => { o.grid = v; save(); }),
            checkbox('Garis halus', !!o.smooth, v => { o.smooth = v; save(); }),
            checkbox('Sumbu Y mulai dari nol', o.zero !== false, v => { o.zero = v; save(); })),
          field('Warna latar', bg),
          el('div.side-label', 'Warna seri'), colorsBox);
      }
    };
    draw();
    this.frame('Editor diagram', el('div.stack', tabs, content),
      el('div.row', el('button.btn', { type: 'button', onclick: () => { app.cmd('chartDelete', ch.id); this.close(); } }, icon('delete'), 'Hapus'), el('span.flex1'),
        el('button.btn', { type: 'button', onclick: () => app.charts.download(ch.id) }, icon('download'), 'PNG'),
        el('button.btn.primary', { type: 'button', onclick: () => this.close() }, 'Selesai')));
  }

  // ---------- pivot ----------
  pivot(sheet) {
    const app = this.app;
    const sh = sheet || app.sheet;
    if (!sh.pivot) { this.frame('Editor tabel pivot', el('div.empty', icon('pivot_table_chart'), el('span', 'Pilih sheet tabel pivot untuk mengedit.'))); return; }
    const def = JSON.parse(JSON.stringify(sh.pivot));
    const src = app.wb.sheetByName(def.srcSheet) || app.sheet;
    const head = [];
    for (let c = def.range.c1; c <= def.range.c2; c++) { const v = app.wb.eng.value(src, def.range.r1, c); head.push({ value: c, label: v == null || v === '' ? 'Kolom ' + colName(c) : String(v) }); }
    const apply = () => { app.withSheet(sh, () => {}); const i = app.wb.sheets.indexOf(sh); if (i !== app.wb.active) app.switchSheet(i); sh.pivot = def; app.cmd('pivotRefresh'); };
    const section = (title, list, kind) => {
      const box = el('div.pv-sec', el('div.pv-head', el('span', title), el('button.btn.sm', { type: 'button', onclick: e => {
        const used = new Set([...def.rows, ...def.cols]);
        (async () => {
          const { openMenu } = await import('./ui.js');
          openMenu(head.filter(h => kind === 'values' || kind === 'filters' || !used.has(h.value)).map(h => ({ label: h.label, run: () => {
            if (kind === 'values') def.values.push({ col: h.value, agg: isNumericCol(h.value) ? 'SUM' : 'COUNTA' });
            else if (kind === 'filters') (def.filters = def.filters || []).push({ col: h.value, hidden: [] });
            else def[kind].push(h.value);
            apply(); this.pivot(sh);
          } })), e.currentTarget.getBoundingClientRect());
        })();
      } }, icon('add'), 'Tambah')));
      list.forEach((it, i) => {
        const col = kind === 'values' || kind === 'filters' ? it.col : it;
        const label = head.find(h => h.value === col)?.label || colName(col);
        const extra = kind === 'values' ? select(['SUM', 'COUNT', 'COUNTA', 'COUNTUNIQUE', 'AVERAGE', 'MEDIAN', 'MAX', 'MIN'].map(a => ({ value: a, label: a })), it.agg, v => { it.agg = v; apply(); }) : kind === 'rows' && i === 0 ? select([{ value: 'asc', label: 'Naik' }, { value: 'desc', label: 'Turun' }], def.rowDesc ? 'desc' : 'asc', v => { def.rowDesc = v === 'desc'; apply(); }) : null;
        box.append(el('div.pv-item', el('b', label), extra, el('button.icon-btn.sm', { type: 'button', onclick: () => { list.splice(i, 1); apply(); this.pivot(sh); } }, icon('close'))));
      });
      return box;
    };
    const isNumericCol = c => { for (let r = def.range.r1 + 1; r <= Math.min(def.range.r2, def.range.r1 + 20); r++) if (typeof app.wb.eng.value(src, r, c) === 'number') return true; return false; };
    const rangeInp = el('input.inp', { value: rangeAddr(def.range, def.srcSheet) });
    rangeInp.addEventListener('change', () => { const g = parseRange(rangeInp.value); if (!g) { toast('Rentang tidak valid'); return; } def.srcSheet = g.sheet || def.srcSheet; def.range = { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }; apply(); this.pivot(sh); });
    this.frame('Editor tabel pivot', el('div.stack',
      field('Rentang sumber', rangeInp),
      section('Baris', def.rows, 'rows'), section('Kolom', def.cols, 'cols'), section('Nilai', def.values, 'values'),
      checkbox('Tampilkan total keseluruhan', def.totals !== false, v => { def.totals = v; apply(); })),
    el('div.row', el('button.btn', { type: 'button', onclick: () => apply() }, icon('refresh'), 'Perbarui'), el('span.flex1'), el('button.btn.primary', { type: 'button', onclick: () => this.close() }, 'Selesai')));
  }

  // ---------- named ranges ----------
  names(define) {
    const app = this.app;
    const list = el('div.rule-list');
    for (const n of app.wb.names.values()) list.append(el('div.rule', { onclick: () => app.fx.gotoName(n.name) },
      el('span.style-chip', icon('label')), el('div.rule-main', el('div.rule-title', n.name), el('div.muted.small', String(n.ref))),
      el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Hapus', onclick: e => { e.stopPropagation(); app.cmd('deleteNamedRange', n.name); this.names(); } }, icon('delete'))));
    if (!app.wb.names.size) list.append(el('div.empty.sm', icon('label'), el('span', 'Belum ada rentang bernama')));
    const name = el('input.inp', { placeholder: 'mis. Penjualan_2026' });
    const absRef = () => quoteSheet(app.sheet.name) + '!' + rangeAddr(app.grid.range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
    const ref = el('input.inp', { value: absRef() });
    const add = () => {
      const n = name.value.trim();
      if (!/^[A-Za-z_\\][\w.]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^(TRUE|FALSE)$/i.test(n)) { toast('Nama harus diawali huruf/garis bawah, tanpa spasi, dan bukan alamat sel'); return; }
      app.cmd('addNamedRange', n, ref.value.replace(/^=/, ''));
      this.names();
    };
    this.frame('Rentang bernama', el('div.stack', list, el('div.card', el('div.side-label', 'Tambahkan rentang'), name, ref, el('button.btn.primary', { type: 'button', onclick: add }, 'Selesai'))));
    this.onSel = () => { if (document.activeElement !== ref) ref.value = absRef(); };
    if (define) setTimeout(() => name.focus(), 30);
  }

  // ---------- versions ----------
  async versions() {
    const app = this.app;
    const list = el('div.ver-list');
    this.frame('Histori versi', el('div.stack', el('div.muted.small', 'Versi disimpan otomatis di perangkat ini setiap 10 menit saat Anda mengedit, dan setiap kali Anda memberi nama versi.'), list),
      el('div.row', el('button.btn', { type: 'button', onclick: async () => { const n = await promptDlg('Beri nama versi', 'Nama versi', ''); if (n) { await store.saveVersion(app.wb, n); this.versions(); } } }, icon('bookmark_add'), 'Beri nama versi saat ini')));
    const vs = await store.listVersions(app.wb.id);
    if (!vs.length) list.append(el('div.empty.sm', icon('history'), el('span', 'Belum ada versi tersimpan')));
    let day = '';
    for (const v of vs) {
      const d = new Date(v.at);
      const dl = d.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      if (dl !== day) { day = dl; list.append(el('div.side-label', dl)); }
      list.append(el('div.ver', el('div.rule-main', el('div.rule-title', v.label || d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })), el('div.muted.small', v.label ? d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }) : 'Versi otomatis')),
        el('button.btn.sm', { type: 'button', onclick: async () => {
          if (!(await confirmDlg('Pulihkan versi', 'Pulihkan versi ini? Versi saat ini akan disimpan terlebih dahulu ke histori.', 'Pulihkan'))) return;
          await store.saveVersion(app.wb, 'Sebelum pemulihan');
          const j = await store.getVersion(v.vid);
          const wb = Workbook.fromJSON(j); wb.id = app.wb.id; wb.lock = app.wb.lock; if (app.wb.share) wb.share = { ...app.wb.share, clean: false };
          app.openWorkbook(wb); toast('Versi dipulihkan');
        } }, 'Pulihkan'),
        el('button.icon-btn.sm', { type: 'button', 'data-tip': 'Buat salinan', onclick: async () => { const j = await store.getVersion(v.vid); const wb = Workbook.fromJSON(j); wb.id = uid('wb'); wb.lock = app.wb.lock; wb.name = 'Salinan dari ' + (v.label || app.wb.name); app.openWorkbook(wb); } }, icon('file_copy'))));
    }
  }

  // ---------- alternating colors / format as table ----------
  banding() {
    const app = this.app, sh = app.sheet;
    let g = { ...app.grid.range };
    if (g.r1 === g.r2 && g.c1 === g.c2) g = io.currentRegion(app, g.r1, g.c1);
    if (g.r2 >= sh.rows - 1) g.r2 = Math.max(g.r1, sh.dataRows() - 1);
    const existing = sh.banding.find(b => b.r1 <= g.r2 && b.r2 >= g.r1 && b.c1 <= g.c2 && b.c2 >= g.c1);
    if (existing) g = { r1: existing.r1, c1: existing.c1, r2: existing.r2, c2: existing.c2 };
    const cur = existing ? { h: existing.h, a: existing.a, b: existing.b, f: existing.f, ht: existing.ht } : { ...BAND_PRESETS[0] };
    let header = existing ? existing.header : true, footer = existing ? existing.footer : false;
    const apply = () => { app.grid.selectRange(g, g.r1, g.c1); app.cmd('banding', { ...cur }, { header, footer, headerStyle: !existing }); };
    const preview = p => el('div.band-prev', el('i', { style: { background: p.h } }), el('i', { style: { background: p.a } }), el('i', { style: { background: p.b } }), el('i', { style: { background: p.a } }));
    const colorRow = (k, label) => { const c = el('input.inp.color', { type: 'color', value: cur[k] || '#ffffff' }); c.addEventListener('input', () => { cur[k] = c.value; apply(); }); return el('div.row.tight', c, el('span', label)); };
    const body = el('div.stack',
      field('Terapkan ke rentang', rangeField(app, rangeAddr(g), t => { const rs = parseRanges(app, t); if (rs) { g = rs[0]; apply(); } })),
      el('div.checks', checkbox('Header', header, v => { header = v; apply(); }), checkbox('Footer', footer, v => { footer = v; apply(); })),
      el('div.side-label', 'Gaya default'),
      el('div.band-grid', ...BAND_PRESETS.map(p => el('button.band', { type: 'button', onclick: () => { Object.assign(cur, p); delete cur.f; apply(); this.banding(); } }, preview(p)))),
      el('div.side-label', 'Gaya kustom'),
      colorRow('h', 'Header'), colorRow('a', 'Warna 1'), colorRow('b', 'Warna 2'), footer ? colorRow('f', 'Footer') : null);
    this.frame('Warna bergantian', body,
      el('div.row', el('button.btn', { type: 'button', onclick: () => { app.grid.selectRange(g, g.r1, g.c1); app.cmd('banding', null); this.close(); } }, 'Hapus warna bergantian'), el('span.flex1'), el('button.btn.primary', { type: 'button', onclick: () => { apply(); this.close(); } }, 'Selesai')));
  }
}
