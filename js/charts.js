// Floating charts (Chart.js) anchored to cells.
import { el, icon, parseRange, loadScript, LIBS, download, keyR } from './util.js';
import { format, fmtSmart } from './format.js';
import { FErr } from './formula/core.js';
import { HH } from './grid.js';
import { openMenu } from './ui/ui.js';

export const CHART_TYPES = [
  { id: 'column', label: 'Kolom', icon: 'bar_chart' },
  { id: 'column-stacked', label: 'Kolom bertumpuk', icon: 'stacked_bar_chart' },
  { id: 'bar', label: 'Batang', icon: 'align_horizontal_left' },
  { id: 'line', label: 'Garis', icon: 'show_chart' },
  { id: 'area', label: 'Area', icon: 'area_chart' },
  { id: 'combo', label: 'Kombo', icon: 'insights' },
  { id: 'pie', label: 'Lingkaran', icon: 'pie_chart' },
  { id: 'doughnut', label: 'Donat', icon: 'donut_large' },
  { id: 'scatter', label: 'Sebar', icon: 'scatter_plot' },
  { id: 'radar', label: 'Radar', icon: 'radar' },
  { id: 'polar', label: 'Area kutub', icon: 'track_changes' },
];

const PALETTE = ['#4285f4', '#ea4335', '#fbbc04', '#34a853', '#ff6d01', '#46bdc6', '#7baaf7', '#f07b72', '#fcd04f', '#71c287'];

export class Charts {
  constructor(app, layer) {
    this.app = app;
    this.layer = layer;
    this.boxes = new Map();
    this.sel = null;
    this.ready = false;
    app.grid.scroller.addEventListener('pointerdown', () => this.select(null));
  }
  palette() { return PALETTE; }

  async lib() {
    if (window.Chart) return window.Chart;
    await loadScript(LIBS.chart);
    return window.Chart;
  }

  reset() {
    for (const b of this.boxes.values()) { b.chart?.destroy(); b.el.remove(); }
    this.boxes.clear();
    this.update();
  }

  update(force = false) {
    const sh = this.app.sheet;
    if (!sh) return;
    const ids = new Set(sh.charts.map(c => c.id));
    for (const [id, b] of this.boxes) if (!ids.has(id) || b.sheet !== sh) { b.chart?.destroy(); b.el.remove(); this.boxes.delete(id); }
    if (!sh.charts.length) return;
    for (const ch of sh.charts) {
      let b = this.boxes.get(ch.id);
      if (!b) { b = this.makeBox(ch); this.boxes.set(ch.id, b); }
      b.def = ch;
      this.renderChart(b, force);
    }
    this.position();
  }

  makeBox(ch) {
    const canvas = el('canvas');
    const menuBtn = el('button.icon-btn.sm.chart-menu', { type: 'button', 'aria-label': 'Opsi diagram', onclick: e => { e.stopPropagation(); this.menu(ch.id, e.currentTarget.getBoundingClientRect()); } }, icon('more_vert'));
    const box = el('div.chart-box', { tabIndex: 0 }, el('div.chart-inner', canvas), menuBtn,
      ...['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(h => el('span.ch-h.h-' + h, { dataset: { h } })));
    const b = { el: box, canvas, chart: null, sheet: this.app.sheet, def: ch };
    box.addEventListener('pointerdown', e => this.onDown(e, b));
    box.addEventListener('dblclick', () => this.app.sidebar.open('chart', { chart: this.app.sheet.charts.find(c => c.id === b.def.id) }));
    box.addEventListener('keydown', e => {
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.app.cmd('chartDelete', b.def.id); }
      if (e.key === 'Escape') { this.select(null); this.app.grid.focus(); }
    });
    this.layer.append(box);
    return b;
  }

  menu(id, at) {
    const ch = this.app.sheet.charts.find(c => c.id === id);
    if (!ch) return;
    this.select(id);
    openMenu([
      { label: 'Edit diagram', icon: 'edit', run: () => this.app.sidebar.open('chart', { chart: ch }) },
      { label: 'Download PNG', icon: 'download', run: () => this.download(id) },
      { label: 'Salin diagram', icon: 'content_copy', run: () => { const c = JSON.parse(JSON.stringify(ch)); c.id = 'ch' + Date.now().toString(36); c.dx += 24; c.dy += 24; this.app.cmd('chartAdd', c); } },
      { sep: true },
      { label: 'Hapus diagram', icon: 'delete', run: () => this.app.cmd('chartDelete', id) },
    ], at, { alignRight: true });
  }

  select(id) {
    this.sel = id;
    for (const [k, b] of this.boxes) b.el.classList.toggle('selected', k === id);
    if (id) this.boxes.get(id)?.el.focus({ preventScroll: true });
  }

  onDown(e, b) {
    if (e.target.closest('.chart-menu')) return;
    e.preventDefault();
    e.stopPropagation();
    this.select(b.def.id);
    const h = e.target.dataset?.h;
    const z = this.app.grid.z;
    const start = { x: e.clientX, y: e.clientY, dx: b.def.dx, dy: b.def.dy, w: b.def.w, h: b.def.h };
    const tmp = { ...b.def };
    const move = ev => {
      const dx = (ev.clientX - start.x) / z, dy = (ev.clientY - start.y) / z;
      if (!h) { tmp.dx = start.dx + dx; tmp.dy = start.dy + dy; }
      else {
        if (h.includes('e')) tmp.w = Math.max(160, start.w + dx);
        if (h.includes('s')) tmp.h = Math.max(120, start.h + dy);
        if (h.includes('w')) { tmp.w = Math.max(160, start.w - dx); tmp.dx = start.dx + (start.w - tmp.w); }
        if (h.includes('n')) { tmp.h = Math.max(120, start.h - dy); tmp.dy = start.dy + (start.h - tmp.h); }
      }
      this.place(b, tmp);
      if (h) b.chart?.resize();
    };
    const up = ev => {
      document.removeEventListener('pointermove', move);
      document.removeEventListener('pointerup', up);
      if (Math.abs(ev.clientX - start.x) < 2 && Math.abs(ev.clientY - start.y) < 2) return;
      // re-anchor to the cell under the top-left corner
      const g = this.app.grid;
      const x = g.X(tmp.anchor.c) + tmp.dx, y = g.Y(tmp.anchor.r) + tmp.dy;
      const c = g.colAt(Math.max(g.geo().hw, x)), r = g.rowAt(Math.max(HH, y));
      tmp.anchor = { r, c };
      tmp.dx = x - g.X(c); tmp.dy = y - g.Y(r);
      this.app.cmd('chartUpdate', { ...tmp });
    };
    document.addEventListener('pointermove', move);
    document.addEventListener('pointerup', up);
  }

  place(b, d) {
    const g = this.app.grid, z = g.z;
    const p = g.boxAt(b.el, Math.min(d.anchor.r, this.app.sheet.rows - 1), Math.min(d.anchor.c, this.app.sheet.cols - 1), d.dx, d.dy);
    Object.assign(b.el.style, { left: p.x + 'px', top: p.y + 'px', width: d.w * z + 'px', height: d.h * z + 'px' });
  }

  position() {
    if (!this.boxes.size) return;
    const g = this.app.grid, geo = g.geo();
    this.layer.style.clipPath = `inset(${HH * g.z}px 0 0 ${geo.hw * g.z}px)`;
    for (const b of this.boxes.values()) this.place(b, b.def);
  }

  values(ch) {
    const wb = this.app.wb;
    const parts = String(ch.range).split(',').map(s => parseRange(s.trim())).filter(Boolean);
    if (!parts.length) return null;
    const sh = (parts[0].sheet && wb.sheetByName(parts[0].sheet)) || this.app.sheet;
    let cols = [];
    let R = 0;
    for (const p of parts) {
      const r2 = p.r2 === Infinity ? sh.dataRows() - 1 : Math.min(p.r2, sh.dataRows() - 1);
      const c2 = p.c2 === Infinity ? sh.dataCols() - 1 : p.c2;
      for (let c = p.c1; c <= c2; c++) {
        const col = [];
        for (let r = p.r1; r <= r2; r++) col.push({ v: wb.eng.value(sh, r, c), r, c });
        cols.push(col);
        R = Math.max(R, col.length);
      }
    }
    let rows = Array.from({ length: R }, (_, i) => cols.map(col => col[i] || { v: null }));
    if (ch.opts.swap) rows = rows[0].map((_, j) => rows.map(r => r[j]));
    return { rows, sh };
  }

  text(sh, cell) {
    if (!cell || cell.v == null) return '';
    if (cell.v instanceof FErr) return cell.v.code;
    const c = sh.get(cell.r, cell.c);
    const st = this.app.wb.styles[c ? c.s : 0] || {};
    return format(cell.v, st.fmt || c?.af).t;
  }

  dataset(ch) {
    const vals = this.values(ch);
    if (!vals || !vals.rows.length) return { labels: [], series: [] };
    const { rows, sh } = vals;
    const o = ch.opts;
    const isNum = x => typeof x.v === 'number';
    const header = o.headers === 'on' || (o.headers !== 'off' && rows.length > 1 && rows[0].slice(1).some(x => typeof x.v === 'string') && rows.slice(1).some(r => r.slice(1).some(isNum)));
    const body = header ? rows.slice(1) : rows;
    const labelsCol = o.labels === 'on' || (o.labels !== 'off' && rows[0].length > 1 && body.some(r => !isNum(r[0]) && r[0].v != null)) || (o.labels !== 'off' && rows[0].length > 1 && body.length && body.every(r => isNum(r[0])) && (() => { const c = sh.get(body[0][0].r, body[0][0].c); return /[dmy]/.test(this.app.wb.styles[c?.s || 0]?.fmt || c?.af || ''); })());
    const labels = labelsCol ? body.map(r => this.text(sh, r[0])) : body.map((_, i) => String(i + 1));
    const series = [];
    for (let j = labelsCol ? 1 : 0; j < rows[0].length; j++) {
      series.push({ name: header ? this.text(sh, rows[0][j]) || `Seri ${j}` : `Seri ${j + (labelsCol ? 0 : 1)}`, data: body.map(r => (isNum(r[j]) ? r[j].v : null)), xs: body.map(r => (isNum(r[0]) ? r[0].v : null)) });
    }
    return { labels, series, labelsCol };
  }
  seriesNames(ch) { return this.dataset(ch).series.map(s => s.name); }

  async renderChart(b, force) {
    const Chart = await this.lib().catch(() => null);
    if (!Chart) { b.el.querySelector('.chart-inner').textContent = 'Diagram memerlukan koneksi internet untuk memuat pustaka.'; return; }
    const ch = b.def, o = ch.opts;
    const ds = this.dataset(ch);
    const css = getComputedStyle(document.documentElement);
    const textColor = o.bg && o.bg !== '#ffffff' ? autoText(o.bg) : css.getPropertyValue('--chart-text').trim() || '#3c4043';
    const gridColor = css.getPropertyValue('--chart-grid').trim() || 'rgba(0,0,0,.08)';
    const colors = i => (o.colors && o.colors[i]) || PALETTE[i % PALETTE.length];
    const type = ch.type;
    const circular = type === 'pie' || type === 'doughnut' || type === 'polar';
    let cfg;
    const base = {
      responsive: true, maintainAspectRatio: false, animation: { duration: force ? 0 : 350 },
      layout: { padding: { top: 6, right: 14, bottom: 6, left: 6 } },
      plugins: {
        legend: { display: o.legend !== 'none' && (circular || ds.series.length > 1 || type === 'combo'), position: o.legend && o.legend !== 'none' ? o.legend : 'bottom', labels: { color: textColor, boxWidth: 12, boxHeight: 12, usePointStyle: true, font: { family: 'Inter', size: 11 } } },
        title: { display: !!o.title, text: o.title, color: textColor, font: { family: 'Inter', size: 14, weight: '600' }, padding: { bottom: o.subtitle ? 2 : 10 } },
        subtitle: { display: !!o.subtitle, text: o.subtitle, color: textColor, font: { family: 'Inter', size: 11 }, padding: { bottom: 10 } },
        tooltip: {
          backgroundColor: 'rgba(32,33,36,.92)', titleFont: { family: 'Inter' }, bodyFont: { family: 'Inter' }, padding: 8, cornerRadius: 6,
          callbacks: { label: c => { const y = c.parsed && typeof c.parsed === 'object' ? (c.chart.options.indexAxis === 'y' ? c.parsed.x : c.parsed.y ?? c.parsed.r) : c.parsed; return `${c.dataset.label || c.label}: ${typeof y === 'number' ? fmtSmart(y) : y}`; } },
        },
      },
    };
    const axis = (title, stacked, horizontal) => ({
      stacked, grid: { color: gridColor, display: o.grid !== false, drawTicks: false }, border: { display: false },
      ticks: { color: textColor, font: { family: 'Inter', size: 11 }, padding: 6, maxRotation: 0, autoSkip: true, callback(v) { return this.type === 'linear' ? fmtSmart(v) : this.getLabelForValue(v); } },
      title: { display: !!title, text: title, color: textColor, font: { family: 'Inter', size: 11, weight: '600' } },
      beginAtZero: o.zero !== false,
    });
    if (circular) {
      const s = ds.series[0] || { data: [] };
      cfg = { type: type === 'polar' ? 'polarArea' : type, data: { labels: ds.labels, datasets: [{ label: s.name, data: s.data, backgroundColor: ds.labels.map((_, i) => colors(i)), borderColor: o.bg || '#fff', borderWidth: 2 }] }, options: { ...base, cutout: type === 'doughnut' ? '58%' : undefined } };
      if (type === 'polar') cfg.options.scales = { r: { grid: { color: gridColor }, ticks: { display: false } } };
    } else if (type === 'scatter') {
      cfg = { type: 'scatter', data: { datasets: ds.series.map((s, i) => ({ label: s.name, data: s.data.map((y, k) => ({ x: ds.labelsCol ? s.xs[k] ?? k + 1 : k + 1, y })).filter(p => p.y != null), backgroundColor: colors(i), pointRadius: 4 })) }, options: { ...base, scales: { x: axis(o.xTitle, false), y: axis(o.yTitle, false) } } };
    } else if (type === 'radar') {
      cfg = { type: 'radar', data: { labels: ds.labels, datasets: ds.series.map((s, i) => ({ label: s.name, data: s.data, borderColor: colors(i), backgroundColor: colors(i) + '33', pointRadius: 2 })) }, options: { ...base, scales: { r: { grid: { color: gridColor }, angleLines: { color: gridColor }, pointLabels: { color: textColor, font: { family: 'Inter', size: 11 } }, ticks: { display: false } } } } };
    } else {
      const stacked = type === 'column-stacked' || !!o.stacked;
      const horizontal = type === 'bar';
      const line = type === 'line' || type === 'area';
      cfg = {
        type: line ? 'line' : 'bar',
        data: {
          labels: ds.labels,
          datasets: ds.series.map((s, i) => {
            const asLine = line || (type === 'combo' && i > 0);
            return {
              type: asLine ? 'line' : 'bar', label: s.name, data: s.data,
              backgroundColor: asLine ? (type === 'area' ? colors(i) + '40' : colors(i)) : colors(i),
              borderColor: colors(i), borderWidth: asLine ? 2 : 0, borderRadius: asLine ? 0 : 3,
              fill: type === 'area' ? (stacked ? (i ? '-1' : 'origin') : 'origin') : false,
              tension: o.smooth ? 0.38 : 0, pointRadius: asLine ? (ds.labels.length > 40 ? 0 : 3) : 0, pointHoverRadius: 5,
              maxBarThickness: 48, order: asLine ? 0 : 1, spanGaps: true,
            };
          }),
        },
        options: { ...base, indexAxis: horizontal ? 'y' : 'x', scales: { x: axis(horizontal ? o.yTitle : o.xTitle, stacked), y: axis(horizontal ? o.xTitle : o.yTitle, stacked) }, interaction: { mode: 'index', intersect: false } },
      };
      if (horizontal) { cfg.options.scales.y.grid.display = false; } else cfg.options.scales.x.grid.display = false;
    }
    cfg.plugins = [bgPlugin(o.bg || css.getPropertyValue('--chart-bg').trim() || '#ffffff'), labelsPlugin(o.dataLabels, textColor)];
    b.el.style.setProperty('--chart-bg-color', o.bg || '');
    const sig = JSON.stringify([cfg.type, ch.type, o, ds.labels.length, ds.series.length, document.documentElement.dataset.theme]);
    if (b.chart && b.sig === sig && !force) {
      b.chart.data.labels = cfg.data.labels;
      cfg.data.datasets.forEach((d, i) => { if (b.chart.data.datasets[i]) b.chart.data.datasets[i].data = d.data; });
      b.chart.update('none');
      return;
    }
    b.chart?.destroy();
    b.sig = sig;
    b.chart = new Chart(b.canvas, cfg);
  }

  download(id) {
    const b = this.boxes.get(id);
    if (!b?.chart) return;
    const url = b.chart.toBase64Image('image/png', 1);
    fetch(url).then(r => r.blob()).then(blob => download((b.def.opts.title || 'Diagram') + '.png', blob));
  }
}

function autoText(bg) {
  const m = /^#?([\da-f]{2})([\da-f]{2})([\da-f]{2})/i.exec(bg);
  if (!m) return '#3c4043';
  const [r, g, b] = [m[1], m[2], m[3]].map(x => parseInt(x, 16));
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? '#3c4043' : '#f1f3f4';
}

function bgPlugin(color) {
  return { id: 'bg', beforeDraw: c => { const ctx = c.ctx; ctx.save(); ctx.fillStyle = color; ctx.fillRect(0, 0, c.width, c.height); ctx.restore(); } };
}
function labelsPlugin(on, color) {
  return {
    id: 'labels',
    afterDatasetsDraw: c => {
      if (!on) return;
      const ctx = c.ctx;
      ctx.save();
      ctx.font = '600 10px Inter, Arial';
      ctx.fillStyle = color;
      ctx.textAlign = 'center';
      c.data.datasets.forEach((d, i) => {
        const meta = c.getDatasetMeta(i);
        if (meta.hidden) return;
        meta.data.forEach((pt, k) => {
          const v = d.data[k];
          const n = typeof v === 'object' && v ? v.y : v;
          if (n == null) return;
          const txt = fmtSmart(n);
          const p = pt.tooltipPosition();
          if (c.config.type === 'pie' || c.config.type === 'doughnut') { ctx.fillStyle = '#fff'; ctx.fillText(txt, p.x, p.y + 3); ctx.fillStyle = color; }
          else if (c.options.indexAxis === 'y') { ctx.textAlign = 'left'; ctx.fillText(txt, p.x + 4, p.y + 3); ctx.textAlign = 'center'; }
          else ctx.fillText(txt, p.x, p.y - 6);
        });
      });
      ctx.restore();
    },
  };
}
