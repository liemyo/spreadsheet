// File lock: a password (PBKDF2 hash in wb.lock, stored beside the workbook like wb.share) asked before the
// spreadsheet is shown. Unlocks are remembered for this browser tab only; share links are not affected.
import { el, icon } from './util.js';
import { dialog, field, toast } from './ui/ui.js';

const IT = 150000;
const hex = b => [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join('');
async function derive(pw, salt, it = IT) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveBits']);
  return hex(await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: new TextEncoder().encode(salt), iterations: it }, key, 256));
}
async function makeLock(pw, hint) {
  const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
  return { v: 1, salt, it: IT, hash: await derive(pw, salt), hint: hint || '', at: Date.now() };
}
const check = async (lock, pw) => !!lock && (await derive(pw, lock.salt, lock.it || IT)) === lock.hash;

const SK = 'as.unlocked';
const tag = wb => wb.id + ':' + wb.lock.hash.slice(0, 16);
const opened = () => { try { return JSON.parse(sessionStorage.getItem(SK) || '[]'); } catch { return []; } };
const keep = list => { try { sessionStorage.setItem(SK, JSON.stringify(list.slice(-200))); } catch { /* ask again next time */ } };
export const isOpen = wb => !wb.lock || opened().includes(tag(wb));
const markOpen = wb => { if (wb.lock && !isOpen(wb)) keep([...opened(), tag(wb)]); };
const forget = wb => keep(opened().filter(x => !x.startsWith(wb.id + ':')));

const pwInput = (ph, auto) => el('input.inp', { type: 'password', placeholder: ph, autocomplete: 'off', autofocus: auto || null });

export function clearGate() { document.getElementById('lockscreen')?.remove(); document.body.classList.remove('file-locked'); }

// Cover the app until the password is entered; the workbook is already loaded behind but not shown.
export function gate(app, wb) {
  clearGate();
  if (isOpen(wb)) return;
  const inp = pwInput('Kata sandi', true);
  const err = el('div.lk-err', { role: 'alert' });
  const go = el('button.btn.primary', { type: 'submit' }, icon('lock_open'), 'Buka');
  let tries = 0, wait = 0;
  const form = el('form.lk-card.glass',
    el('div.lk-ic', icon('lock')),
    el('h2', wb.name || 'Spreadsheet terkunci'),
    el('p.muted', 'Spreadsheet ini dikunci dengan kata sandi.'),
    inp, wb.lock.hint ? el('div.small.muted.lk-hint', icon('lightbulb'), 'Petunjuk: ' + wb.lock.hint) : null, err,
    el('div.lk-btns', el('button.btn', { type: 'button', onclick: () => app.openOther() }, icon('folder_open'), 'Buka file lain'), go));
  form.addEventListener('submit', async e => {
    e.preventDefault();
    if (Date.now() < wait) return;
    go.disabled = true;
    const ok = await check(wb.lock, inp.value);
    go.disabled = false;
    if (app.wb !== wb) return;
    if (ok) { markOpen(wb); clearGate(); app.grid.render(); app.grid.focus(); return; }
    tries++;
    wait = Date.now() + (tries >= 5 ? 30000 : 0);
    err.textContent = tries >= 5 ? 'Terlalu banyak percobaan. Coba lagi dalam 30 detik.' : 'Kata sandi salah.';
    form.classList.remove('shake'); void form.offsetWidth; form.classList.add('shake');
    inp.select();
  });
  document.body.classList.add('file-locked');
  document.body.append(el('div#lockscreen.lk-screen', form));
  setTimeout(() => inp.focus(), 30);
}

export function menuItems(app) {
  const wb = app.wb, owner = !(wb.share && !wb.share.owner);
  return [
    { label: wb.lock ? 'Ubah kata sandi…' : 'Kunci file dengan kata sandi…', icon: 'lock', disabled: () => !owner, run: () => setDialog(app) },
    { label: 'Kunci sekarang', icon: 'lock_clock', disabled: () => !wb.lock, run: () => lockNow(app) },
    { label: 'Hapus kata sandi…', icon: 'lock_open', disabled: () => !wb.lock || !owner, run: () => removeDialog(app) },
  ];
}

function setDialog(app) {
  const wb = app.wb;
  const cur = wb.lock ? pwInput('Kata sandi saat ini', true) : null;
  const a = pwInput('Minimal 4 karakter', !cur), b = pwInput('Ketik ulang kata sandi');
  const hint = el('input.inp', { placeholder: 'Opsional — jangan tulis kata sandinya', value: wb.lock?.hint || '' });
  const err = el('div.lk-err');
  dialog({
    title: wb.lock ? 'Ubah kata sandi file' : 'Kunci file dengan kata sandi', width: 440,
    body: el('div.stack',
      cur ? field('Kata sandi saat ini', cur) : null,
      field('Kata sandi baru', a), field('Konfirmasi kata sandi', b), field('Petunjuk kata sandi', hint), err,
      el('p.dlg-text.small.muted', 'Kata sandi diminta setiap kali spreadsheet dibuka di tab baru, termasuk di perangkat lain akun Anda. Kata sandi tidak dapat dipulihkan — simpan baik-baik. Tautan Bagikan yang aktif tetap dapat dibuka tanpa kata sandi.')),
    buttons: [{ label: 'Batal' }, { label: wb.lock ? 'Simpan' : 'Kunci', primary: true, run: async () => {
      if (cur && !(await check(wb.lock, cur.value))) { err.textContent = 'Kata sandi saat ini salah.'; return false; }
      if (a.value.length < 4) { err.textContent = 'Kata sandi minimal 4 karakter.'; return false; }
      if (a.value !== b.value) { err.textContent = 'Konfirmasi kata sandi tidak sama.'; return false; }
      if (hint.value && hint.value.includes(a.value)) { err.textContent = 'Petunjuk tidak boleh berisi kata sandi.'; return false; }
      if (app.wb !== wb) return true;
      forget(wb);
      wb.lock = await makeLock(a.value, hint.value.trim());
      markOpen(wb);
      await app.save();
      app.updateLockBadge();
      toast(cur ? 'Kata sandi diubah' : 'File dikunci. Kata sandi diminta saat dibuka berikutnya.');
    } }],
  });
}

function removeDialog(app) {
  const wb = app.wb, cur = pwInput('Kata sandi saat ini', true), err = el('div.lk-err');
  dialog({
    title: 'Hapus kata sandi', width: 400,
    body: el('div.stack', el('p.dlg-text', 'Spreadsheet akan dapat dibuka tanpa kata sandi.'), field('Kata sandi saat ini', cur), err),
    buttons: [{ label: 'Batal' }, { label: 'Hapus kata sandi', danger: true, run: async () => {
      if (!(await check(wb.lock, cur.value))) { err.textContent = 'Kata sandi salah.'; return false; }
      forget(wb);
      wb.lock = null;
      await app.save();
      app.updateLockBadge();
      toast('Kata sandi dihapus');
    } }],
  });
}

async function lockNow(app) {
  const wb = app.wb;
  if (!wb.lock) return;
  if (app.grid.edit) app.grid.commitEdit(0, 0);
  await app.save();
  forget(wb);
  app.closePopups();
  gate(app, wb);
}
