// Formula bar, name box, function autocomplete and signature help.
import { el, icon, parseRange, rangeAddr, addr } from '../util.js';
import { lex, toggleAbsAt } from '../formula/parse.js';
import { DOCS } from '../docs.js';
import { highlightFormula, refsOf } from '../grid.js';

export class FormulaBar {
  constructor(app, host) {
    this.app = app;
    this.nameBox = el('input.namebox', { 'aria-label': 'Kotak nama', spellcheck: false });
    this.nameBtn = el('button.icon-btn.nb-drop', { type: 'button', 'aria-label': 'Rentang bernama', 'data-tip': 'Rentang bernama' }, icon('arrow_drop_down'));
    this.input = el('textarea.fx-input', { spellcheck: false, rows: 1, 'aria-label': 'Bilah formula' });
    this.hl = el('div.fx-hl');
    this.expand = el('button.icon-btn.fx-expand', { type: 'button', 'data-tip': 'Perluas bilah formula' }, icon('expand_more'));
    this.box = el('div.fx-box', this.hl, this.input);
    host.append(
      el('div.nb-wrap', this.nameBox, this.nameBtn),
      el('div.fx-sep'),
      el('span.fx-icon', { 'data-tip': 'Fungsi' }, el('i', 'fx')),
      this.box, this.expand,
    );
    this.ac = el('div.ac.glass', { role: 'listbox' });
    this.sig = el('div.sig.glass');
    document.getElementById('popups').append(this.ac, this.sig);
    this.acItems = []; this.acIdx = 0;
    this.bind();
  }

  bind() {
    const app = this.app, inp = this.input;
    this.nameBox.addEventListener('focus', () => this.nameBox.select());
    this.nameBox.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); this.gotoName(this.nameBox.value); app.grid.focus(); }
      if (e.key === 'Escape') { this.sync(); app.grid.focus(); }
    });
    this.nameBtn.addEventListener('click', () => app.namedRangesMenu(this.nameBtn.getBoundingClientRect()));
    this.expand.addEventListener('click', () => { document.body.classList.toggle('fx-tall'); this.expand.querySelector('.ms').textContent = document.body.classList.contains('fx-tall') ? 'expand_less' : 'expand_more'; app.grid.resize(); });
    inp.addEventListener('focus', () => {
      const g = app.grid;
      if (!g.edit) {
        if (!app.canEdit(g.sel.r, g.sel.c)) { inp.blur(); return; }
        g.startEdit(null, 'edit', 'bar');
      }
      if (g.edit) g.edit.src = 'bar';
      this.focused = true;
    });
    inp.addEventListener('blur', () => { this.focused = false; setTimeout(() => this.hideAssist(), 120); });
    inp.addEventListener('input', () => app.grid.onInput(inp));
    inp.addEventListener('keydown', e => app.onKey(e, 'bar'));
    inp.addEventListener('click', () => this.refreshAssist());
    inp.addEventListener('keyup', e => { if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) this.refreshAssist(); });
    inp.addEventListener('scroll', () => { this.hl.scrollTop = inp.scrollTop; });
    this.ac.addEventListener('pointerdown', e => {
      const it = e.target.closest('.ac-item');
      if (!it) return;
      e.preventDefault();
      this.acIdx = +it.dataset.i;
      this.acceptAc();
    });
  }

  activeInput() { return this.focused ? this.input : this.app.grid.input; }

  gotoName(text) {
    const app = this.app, t = text.trim();
    if (!t) return;
    const nm = app.wb.names.get(t.toUpperCase());
    const g = parseRange(nm ? String(nm.ref).replace(/^=/, '') : t);
    if (!g) { app.toast(`Rentang "${t}" tidak valid`); this.sync(); return; }
    if (g.sheet) { const i = app.wb.sheets.findIndex(s => s.name.toLowerCase() === g.sheet.toLowerCase()); if (i < 0) { app.toast('Sheet tidak ditemukan'); return; } app.switchSheet(i); }
    const sh = app.sheet;
    const r2 = g.r2 === Infinity ? sh.rows - 1 : g.r2, c2 = g.c2 === Infinity ? sh.cols - 1 : g.c2;
    if (r2 >= sh.rows || c2 >= sh.cols) sh.growTo(r2 + 1, c2 + 1);
    app.grid.selectRange({ r1: g.r1, c1: g.c1, r2, c2 }, g.r1, g.c1, { scroll: true });
  }

  sync() {
    const app = this.app, g = app.grid, sh = app.sheet;
    const rg = g.range;
    let label;
    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    const named = [...app.wb.names.values()].find(n => {
      const p = parseRange(String(n.ref).replace(/^=/, ''));
      return p && (!p.sheet || p.sheet.toLowerCase() === sh.name.toLowerCase()) && p.r1 === rg.r1 && p.c1 === rg.c1 && (p.r2 === rg.r2 || (p.r2 === Infinity && rg.r2 === sh.rows - 1)) && (p.c2 === rg.c2 || (p.c2 === Infinity && rg.c2 === sh.cols - 1));
    });
    if (named) label = named.name;
    else if (multi) {
      if (rg.r1 === 0 && rg.r2 >= sh.rows - 1) label = rangeAddr(rg).replace(/\d+/g, '');
      else if (rg.c1 === 0 && rg.c2 >= sh.cols - 1) label = `${rg.r1 + 1}:${rg.r2 + 1}`;
      else label = rangeAddr(rg);
    } else label = addr(g.sel.r, g.sel.c);
    if (document.activeElement !== this.nameBox) this.nameBox.value = label;
    if (!g.edit) {
      const t = app.editTextOf(sh, g.sel.r, g.sel.c);
      this.input.value = t;
      this.paint(t);
    }
    this.input.readOnly = !app.canEdit(g.sel.r, g.sel.c, true);
  }

  mirror(text) { if (!this.focused) { this.input.value = text; } this.paint(text); }
  paint(text) {
    if (text.startsWith('=')) { this.hl.innerHTML = highlightFormula(text, refsOf(text)) + '​'; this.box.classList.add('formula'); }
    else { this.hl.textContent = ''; this.box.classList.remove('formula'); }
  }
  endEdit() { this.hideAssist(); if (this.focused) this.input.blur(); this.sync(); }

  // ---------- assist ----------
  hideAssist() { this.ac.classList.remove('show'); this.sig.classList.remove('show'); this.acItems = []; }

  anchorRect() {
    const g = this.app.grid;
    if (this.focused) { const r = this.input.getBoundingClientRect(); return { x: r.left + 6, y: r.bottom + 4, top: r.top }; }
    const r = g.input.getBoundingClientRect();
    return { x: r.left, y: r.bottom + 2, top: r.top };
  }

  refreshAssist() {
    const g = this.app.grid;
    if (!g.edit) { this.hideAssist(); return; }
    const inp = this.activeInput();
    const text = inp.value, caret = inp.selectionStart;
    if (!text.startsWith('=') || inp.selectionStart !== inp.selectionEnd) { this.hideAssist(); this.textComplete(inp); return; }
    const before = text.slice(1, caret);
    const toks = lex(before);
    const last = toks[toks.length - 1];
    const at = this.anchorRect();
    // autocomplete for identifiers being typed
    let showAc = false;
    if (last && (last.t === 'id' || (last.t === 'ref' && /^[A-Za-z]+\d*$/.test(last.s))) && last.p + last.s.length === before.length) {
      const prev = toks.filter(t => t.t !== 'ws');
      const pv = prev[prev.length - 2];
      if (!pv || pv.t === 'op' || pv.t === '(' || pv.t === ',' || pv.t === ';' || pv.t === '{') {
        const q = last.s.toUpperCase();
        const fns = [...DOCS.values()].filter(d => d.name.startsWith(q)).sort((a, b) => a.name.length - b.name.length || a.name.localeCompare(b.name)).slice(0, 9);
        const names = [...this.app.wb.names.values()].filter(n => n.name.toUpperCase().startsWith(q)).slice(0, 3).map(n => ({ name: n.name, named: true, desc: String(n.ref) }));
        this.acItems = [...names, ...fns];
        this.acToken = { start: last.p + 1, end: caret };
        if (this.acItems.length && !(this.acItems.length === 1 && this.acItems[0].name === q && !names.length)) {
          showAc = true;
          this.acIdx = Math.min(this.acIdx, this.acItems.length - 1);
          this.renderAc(at);
        }
      }
    }
    if (!showAc) { this.ac.classList.remove('show'); this.acItems = []; }
    // signature help: innermost open call
    const stack = [];
    toks.forEach((t, i) => {
      if (t.t === '(') { const prev = toks.slice(0, i).filter(x => x.t !== 'ws').pop(); stack.push({ name: prev && prev.t === 'id' ? prev.v.toUpperCase() : null, arg: 0 }); }
      else if (t.t === ')') stack.pop();
      else if ((t.t === ',' || t.t === ';') && stack.length) stack[stack.length - 1].arg++;
      else if (t.t === '{') stack.push({ name: null, arg: 0, brace: true });
      else if (t.t === '}') stack.pop();
    });
    const top = [...stack].reverse().find(s => s.name && !s.brace);
    const doc = top && DOCS.get(top.name);
    if (doc && !showAc) {
      const args = doc.args ? doc.args.split(/,\s*/) : [];
      let idx = top.arg;
      const rep = args.findIndex(a => a.includes('...'));
      if (rep >= 0 && idx >= rep) idx = rep;
      this.sig.innerHTML = '';
      const sigLine = el('div.sig-line', el('b', doc.name), '(');
      args.forEach((a, i) => { if (i) sigLine.append(', '); sigLine.append(i === Math.min(idx, args.length - 1) ? el('span.cur', a) : a); });
      sigLine.append(')');
      this.sig.append(sigLine, el('div.sig-desc', doc.desc), el('div.sig-cat', doc.cat));
      this.sig.classList.add('show');
      this.sig.style.left = Math.max(6, Math.min(at.x, window.innerWidth - this.sig.offsetWidth - 8)) + 'px';
      const h = this.sig.getBoundingClientRect().height;
      this.sig.style.top = (at.top - h - 6 < 4 ? at.y : at.top - h - 6) + 'px';
    } else this.sig.classList.remove('show');
  }

  renderAc(at) {
    this.ac.innerHTML = '';
    this.acItems.forEach((d, i) => {
      this.ac.append(el('div.ac-item' + (i === this.acIdx ? '.on' : ''), { 'data-i': i, role: 'option' },
        el('span.ac-name', d.named ? icon('label', 'sm') : null, d.name),
        el('span.ac-desc', d.named ? d.desc : d.desc)));
    });
    const cur = this.acItems[this.acIdx];
    if (cur && !cur.named) this.ac.append(el('div.ac-foot', el('code', `${cur.name}(${cur.args})`)));
    this.ac.style.top = at.y + 'px';
    this.ac.classList.add('show');
    this.ac.style.left = Math.max(6, Math.min(at.x, window.innerWidth - this.ac.offsetWidth - 8)) + 'px';
    const r = this.ac.getBoundingClientRect();
    if (r.bottom > window.innerHeight - 4) this.ac.style.top = (at.top - r.height - 4) + 'px';
  }

  // returns true if the key was consumed
  handleKey(e) {
    if (this.acItems.length && this.ac.classList.contains('show')) {
      if (e.key === 'ArrowDown') { this.acIdx = (this.acIdx + 1) % this.acItems.length; this.renderAc(this.anchorRect()); return true; }
      if (e.key === 'ArrowUp') { this.acIdx = (this.acIdx - 1 + this.acItems.length) % this.acItems.length; this.renderAc(this.anchorRect()); return true; }
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey && !e.ctrlKey && !e.metaKey)) { this.acceptAc(); return true; }
      if (e.key === 'Escape') { this.hideAssist(); return true; }
    }
    if (e.key === 'F4') {
      const inp = this.activeInput();
      const res = toggleAbsAt(inp.value, inp.selectionStart);
      if (res) { inp.value = res.text; inp.setSelectionRange(res.caret, res.caret); this.app.grid.onInput(inp); }
      return true;
    }
    return false;
  }

  acceptAc() {
    const d = this.acItems[this.acIdx];
    if (!d) return;
    const inp = this.activeInput();
    const t = inp.value, { start, end } = this.acToken;
    const ins = d.named ? d.name : d.name + '(';
    inp.value = t.slice(0, start) + ins + t.slice(end);
    const pos = start + ins.length;
    inp.setSelectionRange(pos, pos);
    this.acItems = [];
    this.ac.classList.remove('show');
    this.app.grid.onInput(inp);
  }

  // Excel-style AutoComplete for plain text using values in the same column
  textComplete(inp) {
    const g = this.app.grid, ed = g.edit;
    if (!ed || ed.mode !== 'enter' || inp !== g.input || this.app.settings.autocomplete === false) return;
    if (this.lastLen !== undefined && inp.value.length < this.lastLen) { this.lastLen = inp.value.length; return; }
    this.lastLen = inp.value.length;
    const typed = inp.value;
    if (!typed || typed.length < 2 || /^[\d=+\-.,]/.test(typed)) return;
    const sh = ed.sh, low = typed.toLowerCase();
    const found = new Set();
    for (const dir of [-1, 1]) {
      for (let r = ed.r + dir, n = 0; r >= 0 && r < sh.rows && n < 2000; r += dir, n++) {
        const c = sh.get(r, ed.c);
        if (!c || c.v == null) { if (n > 0) break; continue; }
        if (typeof c.v === 'string' && c.v.toLowerCase().startsWith(low) && c.v.length > typed.length) found.add(c.v);
      }
    }
    if (found.size !== 1) return;
    const m = [...found][0];
    inp.value = typed + m.slice(typed.length);
    inp.setSelectionRange(typed.length, m.length);
    this.mirror(inp.value);
  }
}
