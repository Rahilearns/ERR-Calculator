// Authentication client + screens (login / register / pending) and the admin approval portal.
// Talks to the /server API at API_BASE_URL. Only loaded/used when AUTH_ENABLED is true.
import { el, toast } from './components.js?v=20260603zzn';
import { API_BASE_URL } from './config.js?v=20260603zzn';

const TOKEN_KEY = 'err_auth_token';
export const getToken = () => localStorage.getItem(TOKEN_KEY) || '';
const setToken = (t) => localStorage.setItem(TOKEN_KEY, t);
export const clearToken = () => localStorage.removeItem(TOKEN_KEY);

async function api(path, { method = 'GET', body = null, token = getToken() } = {}) {
  let res;
  try {
    res = await fetch(API_BASE_URL + path, {
      method,
      headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: 'Bearer ' + token } : {}) },
      body: body ? JSON.stringify(body) : undefined,
    });
  } catch {
    throw new Error('Cannot reach the server. Check your connection and try again.');
  }
  let data = null;
  try { data = await res.json(); } catch { /* no body */ }
  if (!res.ok) { const e = new Error((data && data.error) || `Request failed (${res.status})`); e.status = res.status; throw e; }
  return data;
}

export async function fetchMe() {
  if (!getToken()) return null;
  try { const { user } = await api('/api/me'); return user; }
  catch { clearToken(); return null; }
}
export async function login(cif, password) {
  const { token, user } = await api('/api/login', { method: 'POST', body: { cif, password }, token: null });
  setToken(token); return user;
}
export async function register(cif, email, password) {
  return api('/api/register', { method: 'POST', body: { cif, email, password }, token: null });
}
export function logout() { clearToken(); location.reload(); }

// ---- small form helpers ----
function field(labelText, attrs) {
  const id = 'a_' + Math.random().toString(36).slice(2, 8);
  const input = el('input', { id, autocomplete: 'off', ...attrs });
  const wrap = el('div', { class: 'auth-field' }, el('label', { for: id }, labelText), input);
  return { wrap, input };
}
function setMsg(box, text, kind) {
  box.textContent = text || '';
  box.className = 'auth-msg' + (text ? ' ' + (kind || 'error') : '');
}

// ---- login / register screen ----
// Renders into `root`; calls onAuthed(user) after a successful sign-in.
export function renderAuthScreen(root, onAuthed) {
  root.innerHTML = '';
  let mode = 'login';

  const card = el('div', { class: 'auth-card' });
  const tabs = el('div', { class: 'auth-tabs' });
  const loginTab = el('button', { class: 'auth-tab active', type: 'button' }, 'Sign in');
  const regTab = el('button', { class: 'auth-tab', type: 'button' }, 'Register');
  tabs.append(loginTab, regTab);

  const body = el('div', { class: 'auth-body' });
  const card2 = el('div', { class: 'auth-wrap' },
    el('div', { class: 'auth-brand' },
      el('div', { class: 'auth-title' }, 'Effective Rate of Return Calculator'),
      el('div', { class: 'auth-subtitle' }, 'Sign in to continue'),
    ),
    card,
  );
  card.append(tabs, body);
  root.appendChild(card2);

  function showLogin() {
    mode = 'login'; loginTab.classList.add('active'); regTab.classList.remove('active');
    body.innerHTML = '';
    const cif = field('CIF', { type: 'text', placeholder: 'Your CIF', name: 'cif' });
    const pw = field('Password', { type: 'password', placeholder: 'Your password', name: 'password' });
    const msg = el('div', { class: 'auth-msg' });
    const btn = el('button', { class: 'primary-btn auth-submit', type: 'submit' }, 'Sign in');
    const form = el('form', { class: 'auth-form' }, cif.wrap, pw.wrap, msg, btn);
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      const c = cif.input.value.trim(), p = pw.input.value;
      if (!c || !p) return setMsg(msg, 'Enter your CIF and password.');
      btn.disabled = true; btn.textContent = 'Signing in…'; setMsg(msg, '');
      try { const user = await login(c, p); onAuthed(user); }
      catch (err) { setMsg(msg, err.message); btn.disabled = false; btn.textContent = 'Sign in'; }
    });
    body.appendChild(form);
    cif.input.focus();
  }

  function showRegister() {
    mode = 'register'; regTab.classList.add('active'); loginTab.classList.remove('active');
    body.innerHTML = '';
    const cif = field('CIF', { type: 'text', placeholder: 'Your CIF', name: 'cif' });
    const email = field('IDLC Email', { type: 'email', placeholder: 'name@idlc.com', name: 'email' });
    const pw = field('Password', { type: 'password', placeholder: 'At least 8 characters', name: 'password' });
    const pw2 = field('Confirm Password', { type: 'password', placeholder: 'Re-type your password', name: 'password2' });
    const msg = el('div', { class: 'auth-msg' });
    const btn = el('button', { class: 'primary-btn auth-submit', type: 'submit' }, 'Create account');
    const form = el('form', { class: 'auth-form' }, cif.wrap, email.wrap, pw.wrap, pw2.wrap, msg, btn);
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      const c = cif.input.value.trim(), em = email.input.value.trim(), p = pw.input.value, p2 = pw2.input.value;
      if (!c || !em || !p) return setMsg(msg, 'Fill in every field.');
      if (p.length < 8) return setMsg(msg, 'Password must be at least 8 characters.');
      if (p !== p2) return setMsg(msg, 'The two passwords do not match.');
      btn.disabled = true; btn.textContent = 'Submitting…'; setMsg(msg, '');
      try {
        const r = await register(c, em, p);
        renderPending(body, r.message);
      } catch (err) { setMsg(msg, err.message); btn.disabled = false; btn.textContent = 'Create account'; }
    });
    body.appendChild(form);
    cif.input.focus();
  }

  function renderPending(container, message) {
    container.innerHTML = '';
    container.appendChild(el('div', { class: 'auth-pending' },
      el('div', { class: 'auth-pending-icon' }, '⏳'),
      el('h3', {}, 'Registration received'),
      el('p', {}, message || 'You can sign in once an administrator approves your account.'),
      (() => { const b = el('button', { class: 'ghost-btn', type: 'button' }, '← Back to sign in'); b.addEventListener('click', showLogin); return b; })(),
    ));
  }

  loginTab.addEventListener('click', showLogin);
  regTab.addEventListener('click', showRegister);
  showLogin();
}

// ---- admin portal ----
export function renderAdminPanel(root, onBack) {
  root.innerHTML = '';
  let filter = 'pending';
  const card = el('div', { class: 'section-card' });
  const head = el('div', { class: 'admin-head' },
    el('div', {}, el('h2', { style: 'margin:0' }, 'Admin Portal'), el('div', { class: 'help' }, 'Approve or reject registrations')),
    (() => { const b = el('button', { class: 'ghost-btn', type: 'button' }, '← Back to calculator'); b.addEventListener('click', onBack); return b; })(),
  );
  card.appendChild(head);

  const chips = el('div', { class: 'admin-filters' });
  const FILTERS = [['pending', 'Pending'], ['approved', 'Approved'], ['rejected', 'Rejected'], ['', 'All']];
  const tableWrap = el('div', { class: 'admin-table-wrap' });
  card.append(chips, tableWrap);
  root.appendChild(card);

  function renderChips() {
    chips.innerHTML = '';
    FILTERS.forEach(([val, label]) => {
      const b = el('button', { class: 'admin-chip' + (val === filter ? ' active' : ''), type: 'button' }, label);
      b.addEventListener('click', () => { filter = val; renderChips(); load(); });
      chips.appendChild(b);
    });
  }

  async function load() {
    tableWrap.innerHTML = '';
    tableWrap.appendChild(el('div', { class: 'help' }, 'Loading…'));
    let users;
    try { const r = await api('/api/admin/users' + (filter ? '?status=' + filter : '')); users = r.users; }
    catch (err) { tableWrap.innerHTML = ''; tableWrap.appendChild(el('div', { class: 'auth-msg error' }, err.message)); return; }
    tableWrap.innerHTML = '';
    if (!users.length) { tableWrap.appendChild(el('div', { class: 'help' }, 'No ' + (filter || '') + ' registrations.')); return; }

    const rows = users.map((u) => {
      const tds = [
        el('td', {}, u.cif),
        el('td', {}, u.email),
        el('td', {}, el('span', { class: 'status-pill status-' + u.status }, u.status)),
        el('td', {}, new Date(u.createdAt).toLocaleString()),
      ];
      const actions = el('td', { class: 'admin-actions' });
      if (u.role !== 'admin' && u.status === 'pending') {
        const ap = el('button', { class: 'mini-btn approve', type: 'button' }, 'Approve');
        const rj = el('button', { class: 'mini-btn reject', type: 'button' }, 'Reject');
        ap.addEventListener('click', () => act('/api/admin/approve', u, ap, rj));
        rj.addEventListener('click', () => act('/api/admin/reject', u, ap, rj));
        actions.append(ap, rj);
      } else if (u.role === 'admin') {
        actions.appendChild(el('span', { class: 'help' }, 'admin'));
      } else {
        actions.appendChild(el('span', { class: 'help' }, '—'));
      }
      tds.push(actions);
      return el('tr', {}, ...tds);
    });
    const table = el('table', { class: 'admin-table' },
      el('thead', {}, el('tr', {}, ...['CIF', 'Email', 'Status', 'Registered', 'Action'].map((h) => el('th', {}, h)))),
      el('tbody', {}, ...rows),
    );
    tableWrap.appendChild(table);
  }

  async function act(path, user, ap, rj) {
    if (ap) ap.disabled = true; if (rj) rj.disabled = true;
    try { await api(path, { method: 'POST', body: { id: user.id } }); toast('Done — ' + user.cif + ' updated.', 'success'); load(); }
    catch (err) { toast(err.message, 'error'); if (ap) ap.disabled = false; if (rj) rj.disabled = false; }
  }

  renderChips();
  load();
}

// ---- header chip ----
export function userChip(user, { onAdmin, onLogout }) {
  const chip = el('div', { class: 'user-chip' });
  chip.appendChild(el('span', { class: 'user-chip-name', title: user.email }, user.cif));
  if (user.role === 'admin') {
    const a = el('button', { class: 'ghost-btn user-chip-btn', type: 'button' }, 'Admin');
    a.addEventListener('click', onAdmin);
    chip.appendChild(a);
  }
  const out = el('button', { class: 'ghost-btn user-chip-btn', type: 'button' }, 'Log out');
  out.addEventListener('click', onLogout);
  chip.appendChild(out);
  return chip;
}
