// Standalone web verzia Archív Ravafol (bez Outlooku/Office) — funguje v ľubovoľnom prehliadači.
// Token v localStorage, API = vlastný origin (Caddy proxuje /v1).
import PostalMime from './vendor/postal-mime.js';

let API = localStorage.getItem('oa_api') || window.location.origin;
let TOKEN = localStorage.getItem('oa_token') || '';
let lastResults = [];

const $ = (id) => document.getElementById(id);
const show = (id) => ['loginView', 'searchView', 'viewerView'].forEach(v => $(v).classList.toggle('hidden', v !== id));

window.addEventListener('DOMContentLoaded', () => {
  $('apiBase').value = API;
  $('loginBtn').onclick = doLogin;
  $('logoutBtn').onclick = doLogout;
  $('searchForm').onsubmit = (e) => { e.preventDefault(); doSearch(); };
  $('backBtn').onclick = () => show('searchView');
  if (TOKEN) { show('searchView'); $('logoutBtn').classList.remove('hidden'); $('query').focus(); }
  else show('loginView');
});

function save() { localStorage.setItem('oa_api', API); localStorage.setItem('oa_token', TOKEN); }

async function api(path, opts = {}) {
  const res = await fetch(`${API}${path}`, {
    ...opts,
    headers: { 'Content-Type': 'application/json', ...(TOKEN ? { Authorization: `Bearer ${TOKEN}` } : {}), ...(opts.headers || {}) },
  });
  if (res.status === 401) { doLogout(); throw new Error('Relácia vypršala – prihlás sa znova.'); }
  return res;
}

async function doLogin() {
  $('loginError').textContent = '';
  API = ($('apiBase').value || window.location.origin).trim().replace(/\/+$/, '');
  const email = $('email').value.trim().toLowerCase();
  const password = $('password').value;
  if (!email || !password) { $('loginError').textContent = 'Vyplň email aj heslo.'; return; }
  try {
    const res = await fetch(`${API}/v1/auth/login`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password }),
    });
    if (!res.ok) { $('loginError').textContent = 'Nesprávny email alebo heslo.'; return; }
    const data = await res.json();
    TOKEN = data.accessToken || data.token;
    if (!TOKEN) { $('loginError').textContent = 'Server nevrátil token.'; return; }
    save();
    $('password').value = '';
    $('logoutBtn').classList.remove('hidden');
    show('searchView'); $('query').focus();
  } catch (e) { $('loginError').textContent = 'Nedá sa pripojiť k serveru. ' + e.message; }
}

function doLogout() { TOKEN = ''; save(); $('logoutBtn').classList.add('hidden'); show('loginView'); }

async function doSearch() {
  const q = $('query').value.trim();
  if (!q) return;
  $('status').textContent = 'Hľadám…'; $('results').innerHTML = '';
  try {
    const res = await api(`/v1/search?keywords=${encodeURIComponent(q)}&page=1&limit=25`);
    if (!res.ok) { $('status').textContent = 'Chyba vyhľadávania.'; return; }
    const data = await res.json();
    lastResults = data.hits || [];
    $('status').textContent = `Nájdené: ${data.total ?? lastResults.length}`;
    renderResults();
  } catch (e) { $('status').textContent = e.message; }
}

const esc = (s) => (s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function renderResults() {
  const ul = $('results'); ul.innerHTML = '';
  if (!lastResults.length) { ul.innerHTML = '<li class="muted">Žiadne výsledky.</li>'; return; }
  lastResults.forEach((hit) => {
    const li = document.createElement('li');
    const date = hit.timestamp ? new Date(hit.timestamp).toLocaleString('sk-SK') : '';
    li.innerHTML = `<div class="r-subject">${esc(hit.subject) || '(bez predmetu)'}</div>
      <div class="r-meta">od: ${esc(hit.from)} • ${esc(date)}</div>`;
    li.onclick = () => openEmail(hit.id);
    ul.appendChild(li);
  });
}

async function openEmail(id) {
  show('viewerView');
  $('emailMeta').innerHTML = '<span class="muted">Načítavam e-mail…</span>';
  $('emailBody').innerHTML = ''; $('attachments').innerHTML = '';
  try {
    const r1 = await api(`/v1/archived-emails/${encodeURIComponent(id)}`);
    if (!r1.ok) throw new Error('E-mail sa nenašiel.');
    const meta = await r1.json();
    const path = meta.storagePath || meta.storage_path || meta.email?.storagePath;
    const r2 = await api(`/v1/storage/download?path=${encodeURIComponent(path)}`);
    if (!r2.ok) throw new Error('Obsah e-mailu sa nepodarilo stiahnuť.');
    const email = await new PostalMime().parse(await r2.arrayBuffer());
    renderEmail(email);
  } catch (e) { $('emailMeta').innerHTML = `<span class="error">${esc(e.message)}</span>`; }
}

function renderEmail(email) {
  const addr = (a) => Array.isArray(a) ? a.map(x => x.address).join(', ') : (a?.address || '');
  const date = email.date ? new Date(email.date).toLocaleString('sk-SK') : '';
  $('emailMeta').innerHTML = `
    <div class="subj">${esc(email.subject) || '(bez predmetu)'}</div>
    <div class="row">Od: ${esc(addr(email.from))}</div>
    <div class="row">Komu: ${esc(addr(email.to))}</div>
    ${email.cc?.length ? `<div class="row">Kópia: ${esc(addr(email.cc))}</div>` : ''}
    <div class="row">${esc(date)}</div>`;
  const body = $('emailBody');
  if (email.html) {
    const iframe = document.createElement('iframe');
    iframe.setAttribute('sandbox', '');
    iframe.srcdoc = email.html;
    body.appendChild(iframe);
  } else {
    const pre = document.createElement('pre');
    pre.style.whiteSpace = 'pre-wrap';
    pre.textContent = email.text || '(prázdne telo)';
    body.appendChild(pre);
  }
  if (email.attachments?.length) {
    const wrap = $('attachments');
    wrap.innerHTML = '<div class="muted" style="margin-top:8px">Prílohy:</div>';
    email.attachments.forEach((att) => {
      const blob = new Blob([att.content], { type: att.mimeType || 'application/octet-stream' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = att.filename || 'priloha';
      a.textContent = `📎 ${att.filename || 'priloha'}`;
      wrap.appendChild(a);
    });
  }
}
