/* =============================================================
   THE DRIVE — script.js
   4/20/2026 — Sidebar + Row-based browse view
   ============================================================= */

// ─── CONFIG ───────────────────────────────────────────────────
const SHEET_CSV_URL    = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-pTR5TVQn64f0w1o8Z4JeJ9rj9GtOPDoAA1R5cDeg7YYrgscYwPVxJIqgdP9Bn9ywCnDjCjm7nsTR/pub?output=csv';
const DRIVE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbykhH4OK9zXWh9o4dhMRpdMccBpX7LUfM2fyAB-rMkRJrpgONtkVTz82XY48pgLaAT_Tw/exec';

// ─── ACCESS KEY GATE ──────────────────────────────────────────
const LOCAL_KEY_STORE = 'thedrive_access_key_v1';
const LOCAL_DEVICE_ID = 'thedrive_device_id_v1';

function getSavedKey() {
  try { return localStorage.getItem(LOCAL_KEY_STORE) || null; } catch(e) { return null; }
}
function saveKey(key) {
  try { localStorage.setItem(LOCAL_KEY_STORE, key); } catch(e) {}
}

function getDeviceId() {
  try {
    let did = localStorage.getItem(LOCAL_DEVICE_ID);
    if (!did) {
      if (crypto && crypto.randomUUID) {
        did = 'did-' + crypto.randomUUID().replace(/-/g, '').slice(0, 12).toUpperCase();
      } else {
        did = 'did-' + Array.from({ length: 12 }, () => Math.floor(Math.random() * 16).toString(16)).join('').toUpperCase();
      }
      localStorage.setItem(LOCAL_DEVICE_ID, did);
    }
    return did;
  } catch(e) { return 'did-UNKNOWN'; }
}

function callKeyAction(action, keyStr, existing) {
  return new Promise((resolve) => {
    const cbName = '__keyCallback_' + Date.now();
    const script = document.createElement('script');
    const timer  = setTimeout(() => { cleanup(); resolve({ error: 'timeout' }); }, 12000);
    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }
    window[cbName] = function(data) { cleanup(); resolve(data); };
    script.src = DRIVE_SCRIPT_URL
      + '?action=' + action
      + '&key='    + encodeURIComponent(keyStr)
      + (existing ? '&existing=1' : '')
      + '&did='    + encodeURIComponent(getDeviceId())
      + '&callback=' + cbName
      + '&_cb=' + Date.now();
    script.onerror = () => { cleanup(); resolve({ error: 'network' }); };
    document.head.appendChild(script);
  });
}

let gateResolveFn = null;

function showDenied() {
  const overlay   = document.getElementById('gate-overlay');
  const titleEl   = overlay && overlay.querySelector('.gate-title');
  const fieldEl   = overlay && overlay.querySelector('.modal-field');
  const submitBtn = document.getElementById('gate-submit');
  if (overlay) overlay.classList.remove('gate-overlay-hidden');
  if (titleEl) {
    titleEl.textContent = 'ACCESS DENIED';
    if (!overlay.querySelector('.gate-denied-msg')) {
      const msg = document.createElement('p');
      msg.className = 'gate-denied-msg';
      msg.textContent = 'Your device has been blocked';
      titleEl.insertAdjacentElement('afterend', msg);
    }
  }
  if (fieldEl) fieldEl.style.display = 'none';
  if (submitBtn) {
    submitBtn.classList.remove('loading');
    submitBtn.classList.add('denied');
    submitBtn.disabled = true;
    submitBtn.innerHTML = `<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>`;
  }
}

function showGate() {
  return new Promise(resolve => {
    gateResolveFn = resolve;
    const overlay   = document.getElementById('gate-overlay');
    const input     = document.getElementById('gate-key-input');
    const submitBtn = document.getElementById('gate-submit');
    const errorEl   = document.getElementById('gate-error');
    if (!overlay) { resolve(); return; }
    overlay.classList.remove('gate-overlay-hidden');

    function showError(msg) {
      errorEl.textContent = msg;
      errorEl.hidden = false;
      input.style.borderColor = 'var(--red)';
      submitBtn.classList.remove('loading');
      submitBtn.textContent = 'ENTER THE DRIVE';
    }
    function clearError() { errorEl.hidden = true; input.style.borderColor = ''; }

    input.addEventListener('input', () => {
      const start = input.selectionStart, end = input.selectionEnd;
      input.value = input.value.toUpperCase();
      input.setSelectionRange(start, end);
      clearError();
    });

    async function attempt() {
      const keyStr = input.value.trim().toUpperCase();
      if (!keyStr) { showError('Please enter your access key.'); return; }
      submitBtn.classList.add('loading');
      submitBtn.textContent = 'CHECKING…';
      clearError();
      const validation = await callKeyAction('validateKey', keyStr);
      if (validation.error) { showError('Could not reach the server. Check your connection and try again.'); return; }
      if (!validation.valid) {
        if (validation.reason === 'device_blocked') showDenied();
        else if (validation.reason === 'expired') showError('This key has reached its device limit. Please request a new key.');
        else showError('Invalid key. Please check and try again.');
        return;
      }
      const consume = await callKeyAction('useKey', keyStr);
      if (!consume.success && consume.reason === 'device_blocked') { showDenied(); return; }
      if (!consume.success && consume.reason === 'expired') { showError('This key just hit its device limit. Please request a new key.'); return; }
      saveKey(keyStr);
      overlay.classList.add('gate-overlay-hidden');
      setTimeout(() => { overlay.style.display = 'none'; }, 350);
      resolve();
    }

    submitBtn.addEventListener('click', attempt);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') attempt(); });
    setTimeout(() => input.focus(), 100);
  });
}

async function initWithGate() {
  const overlay = document.getElementById('gate-overlay');
  const earlySavedKey = getSavedKey();
  if (earlySavedKey && overlay) {
    const earlyInput = document.getElementById('gate-key-input');
    const earlyBtn   = document.getElementById('gate-submit');
    if (earlyInput) { earlyInput.value = earlySavedKey; earlyInput.disabled = true; }
    if (earlyBtn)   { earlyBtn.textContent = 'CHECKING PERMISSIONS…'; earlyBtn.classList.add('loading'); }
  }

  if (DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
    const deviceCheck = await new Promise(resolve => {
      const cbName = '__deviceCheckCallback_' + Date.now();
      const script  = document.createElement('script');
      const timer   = setTimeout(() => { cleanup(); resolve({ allowed: true }); }, 8000);
      function cleanup() { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }
      window[cbName] = function(data) { cleanup(); resolve(data); };
      script.src = DRIVE_SCRIPT_URL
        + '?action=checkDevice'
        + '&did='      + encodeURIComponent(getDeviceId())
        + '&key='      + encodeURIComponent(getSavedKey() || '')
        + '&callback=' + cbName
        + '&_cb='      + Date.now();
      script.onerror = () => { cleanup(); resolve({ allowed: true }); };
      document.head.appendChild(script);
    });
    if (deviceCheck.allowed === false) { showDenied(); return; }
    if (deviceCheck.keyCleared) { try { localStorage.removeItem(LOCAL_KEY_STORE); } catch(e) {} }
  }

  const savedKey = getSavedKey();
  if (savedKey) {
    if (overlay) {
      overlay.classList.remove('gate-overlay-hidden');
      const submitBtn = document.getElementById('gate-submit');
      const input     = document.getElementById('gate-key-input');
      if (submitBtn) { submitBtn.textContent = 'CHECKING KEY…'; submitBtn.classList.add('loading'); }
      if (input)     { input.value = savedKey; input.disabled = true; }
    }
    const validation = await callKeyAction('validateKey', savedKey, true);
    if (validation.reason === 'device_blocked') {
      showDenied();
    } else if (validation.valid || validation.error) {
      if (overlay) { overlay.classList.add('gate-overlay-hidden'); overlay.style.display = 'none'; }
    } else {
      try { localStorage.removeItem(LOCAL_KEY_STORE); } catch(e) {}
      const submitBtn = document.getElementById('gate-submit');
      const input     = document.getElementById('gate-key-input');
      const errorEl   = document.getElementById('gate-error');
      if (submitBtn) { submitBtn.textContent = 'ENTER THE DRIVE'; submitBtn.classList.remove('loading'); }
      if (input)     { input.value = ''; input.disabled = false; }
      if (errorEl)   { errorEl.textContent = 'Your access key is no longer valid. Please enter a new one.'; errorEl.hidden = false; }
      await showGate();
    }
  } else {
    await showGate();
  }
}

// ─── DEMO DATA ────────────────────────────────────────────────
const DEMO_MOVIES = [
  { title: "Inception", resolution: "4K", maturityRating: "PG-13", releaseDate: "2010-07-16", fileSize: "58 GB", imdbRating: "8.8" },
  { title: "The Dark Knight", resolution: "4K", maturityRating: "PG-13", releaseDate: "2008-07-18", fileSize: "62 GB", imdbRating: "9.0" },
  { title: "Interstellar", resolution: "4K", maturityRating: "PG-13", releaseDate: "2014-11-07", fileSize: "55 GB", imdbRating: "8.6" },
  { title: "Parasite", resolution: "1080p", maturityRating: "R", releaseDate: "2019-11-08", fileSize: "14 GB", imdbRating: "8.5" },
  { title: "Dune", resolution: "4K", maturityRating: "PG-13", releaseDate: "2021-10-22", fileSize: "47 GB", imdbRating: "8.0" },
  { title: "The Godfather", resolution: "1080p", maturityRating: "R", releaseDate: "1972-03-24", fileSize: "20 GB", imdbRating: "9.2" },
  { title: "Pulp Fiction", resolution: "1080p", maturityRating: "R", releaseDate: "1994-10-14", fileSize: "18 GB", imdbRating: "8.9" },
  { title: "Mad Max Fury Road", resolution: "4K", maturityRating: "R", releaseDate: "2015-05-15", fileSize: "50 GB", imdbRating: "8.1" },
  { title: "Oppenheimer", resolution: "4K", maturityRating: "R", releaseDate: "2023-07-21", fileSize: "72 GB", imdbRating: "8.3" },
  { title: "Everything Everywhere All at Once", resolution: "1080p", maturityRating: "R", releaseDate: "2022-03-25", fileSize: "12 GB", imdbRating: "7.8" },
  { title: "Arrival", resolution: "1080p", maturityRating: "PG-13", releaseDate: "2016-11-11", fileSize: "16 GB", imdbRating: "7.9" },
  { title: "The Revenant", resolution: "4K", maturityRating: "R", releaseDate: "2015-12-25", fileSize: "53 GB", imdbRating: "8.0" },
  { title: "Blade Runner 2049", resolution: "4K", maturityRating: "R", releaseDate: "2017-10-06", fileSize: "60 GB", imdbRating: "8.0" },
  { title: "1917", resolution: "1080p", maturityRating: "R", releaseDate: "2019-12-25", fileSize: "22 GB", imdbRating: "8.3" },
  { title: "Whiplash", resolution: "1080p", maturityRating: "R", releaseDate: "2014-10-10", fileSize: "11 GB", imdbRating: "8.5" },
  { title: "The Grand Budapest Hotel", resolution: "1080p", maturityRating: "R", releaseDate: "2014-03-28", fileSize: "9 GB", imdbRating: "8.1" },
  { title: "Tenet", resolution: "4K", maturityRating: "PG-13", releaseDate: "2020-09-03", fileSize: "48 GB", imdbRating: "7.3" },
  { title: "Joker", resolution: "4K", maturityRating: "R", releaseDate: "2019-10-04", fileSize: "44 GB", imdbRating: "8.4" },
  { title: "The Shawshank Redemption", resolution: "1080p", maturityRating: "R", releaseDate: "1994-09-23", fileSize: "15 GB", imdbRating: "9.3" },
  { title: "Schindler's List", resolution: "1080p", maturityRating: "R", releaseDate: "1993-12-15", fileSize: "19 GB", imdbRating: "8.9" },
];

const DEMO_DRIVE = {
  "inception": { id: "demo1", link: "#" },
  "thedarkknight": { id: "demo2", link: "#" },
  "interstellar": { id: "demo3", link: "#" },
  "parasite": { id: "demo4", link: "#" },
  "dune": { id: "demo5", link: "#" },
  "oppenheimer": { id: "demo6", link: "#" },
  "pulpfiction": { id: "demo7", link: "#" },
  "whiplash": { id: "demo8", link: "#" },
  "theshawshankredemption": { id: "demo9", link: "#" },
};

// ─── STATE ────────────────────────────────────────────────────
let allMovies  = [];
let filtered   = [];
let currentSort = 'title';
let currentDir  = 'asc';
let isDemoMode  = false;
let posterMap   = {};

// Active sidebar filters
let activeFilters = {
  maturity:   new Set(),
  status:     new Set(),
  resolution: new Set(),
};

function hasActiveFilters() {
  const search = searchInput ? searchInput.value.trim() : '';
  return search.length > 0
    || activeFilters.maturity.size   > 0
    || activeFilters.status.size     > 0
    || activeFilters.resolution.size > 0;
}

// ─── REQUEST COUNTS ───────────────────────────────────────────
let requestCounts = {};
const LOCAL_USER_REQ_KEY = 'thedrive_user_reqs_v1';

function loadUserRequested() {
  try { return new Set(JSON.parse(localStorage.getItem(LOCAL_USER_REQ_KEY) || '[]')); } catch(e) { return new Set(); }
}
function saveUserRequested() {
  try { localStorage.setItem(LOCAL_USER_REQ_KEY, JSON.stringify([...userRequested])); } catch(e) {}
}
let userRequested = loadUserRequested();

function hasUserRequested(title) { return userRequested.has(normalize(title)); }

function getRatingScore(title) {
  const r = ratingCounts[normalize(title)];
  if (!r) return 0;
  return (r.up || 0) - (r.down || 0);
}

function getRequestCount(title) { return requestCounts[normalize(title)] || 0; }

async function postRequest(title) {
  const key = normalize(title);
  userRequested.add(key);
  saveUserRequested();
  requestCounts[key] = (requestCounts[key] || 0) + 1;
  const localCount = requestCounts[key];
  if (!DRIVE_SCRIPT_URL || DRIVE_SCRIPT_URL === 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') return localCount;
  return new Promise(resolve => {
    const cbName = '__requestCallback_' + Date.now();
    const script = document.createElement('script');
    const timer = setTimeout(() => { cleanup(); resolve(localCount); }, 10000);
    function cleanup() { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }
    window[cbName] = function(data) {
      cleanup();
      if (data && data.count !== undefined) { requestCounts[key] = data.count; resolve(data.count); }
      else resolve(localCount);
    };
    const url = DRIVE_SCRIPT_URL
      + '?action=request'
      + '&title=' + encodeURIComponent(title)
      + '&key='   + encodeURIComponent(getSavedKey() || '')
      + '&did='   + encodeURIComponent(getDeviceId())
      + '&callback=' + cbName;
    script.src = url;
    script.onerror = () => { cleanup(); resolve(localCount); };
    document.head.appendChild(script);
  });
}

requestCounts = {};

// ─── RATINGS ─────────────────────────────────────────────────
const LOCAL_RATINGS_KEY = 'thedrive_ratings_v1';
let ratingCounts = {};

function loadUserRatings() {
  try { return JSON.parse(localStorage.getItem(LOCAL_RATINGS_KEY) || '{}'); } catch(e) { return {}; }
}
function saveUserRatings() {
  try { localStorage.setItem(LOCAL_RATINGS_KEY, JSON.stringify(userRatings)); } catch(e) {}
}
let userRatings = loadUserRatings();

function getUserRating(title) { return userRatings[normalize(title)] || null; }
function getRatingCount(title, type) { return (ratingCounts[normalize(title)] || {})[type] || 0; }

async function postRating(title, type) {
  const key = normalize(title);
  const prev = userRatings[key];
  if (prev === type) delete userRatings[key];
  else userRatings[key] = type;
  saveUserRatings();
  if (!DRIVE_SCRIPT_URL || DRIVE_SCRIPT_URL === 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') return;
  return new Promise(resolve => {
    const cbName = '__ratingCallback_' + Date.now();
    const script = document.createElement('script');
    const timer = setTimeout(() => { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); resolve(); }, 10000);
    window[cbName] = function(data) {
      clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script);
      if (data && typeof data.up === 'number' && typeof data.down === 'number') ratingCounts[key] = { up: data.up, down: data.down };
      resolve(data);
    };
    script.src = DRIVE_SCRIPT_URL
      + '?action=rateMovie'
      + '&title=' + encodeURIComponent(title)
      + '&type='  + encodeURIComponent(type)
      + '&prev='  + encodeURIComponent(prev || '')
      + '&key='   + encodeURIComponent(getSavedKey() || '')
      + '&did='   + encodeURIComponent(getDeviceId())
      + '&callback=' + cbName;
    script.onerror = () => { if (script.parentNode) script.parentNode.removeChild(script); resolve(); };
    document.head.appendChild(script);
  });
}

// ─── SETTINGS PERSISTENCE ─────────────────────────────────────
const LOCAL_SETTINGS_KEY = 'thedrive_settings_v2';

function saveSettings() {
  try {
    localStorage.setItem(LOCAL_SETTINGS_KEY, JSON.stringify({
      search: searchInput ? searchInput.value : '',
      sort:   currentSort,
      dir:    currentDir,
      maturity:   [...activeFilters.maturity],
      status:     [...activeFilters.status],
      resolution: [...activeFilters.resolution],
    }));
  } catch(e) {}
}

function loadSettings() {
  try { return JSON.parse(localStorage.getItem(LOCAL_SETTINGS_KEY) || 'null'); } catch(e) { return null; }
}

function applySettings(s) {
  if (!s) return;
  if (s.search && searchInput) {
    searchInput.value = s.search;
    if (clearSearch) clearSearch.classList.toggle('visible', s.search.length > 0);
  }
  if (s.sort) { currentSort = s.sort; if (sortBy) sortBy.value = s.sort; }
  if (s.dir)  { currentDir = s.dir; if (sortDirBtn) sortDirBtn.textContent = currentDir === 'desc' ? '↓' : '↑'; }
  if (s.maturity)   s.maturity.forEach(v => activeFilters.maturity.add(v));
  if (s.status)     s.status.forEach(v => activeFilters.status.add(v));
  if (s.resolution) s.resolution.forEach(v => activeFilters.resolution.add(v));
}

// ─── DOM REFS ─────────────────────────────────────────────────
const $ = id => document.getElementById(id);

const searchInput  = $('search-input');
const clearSearch  = $('clear-search');
const sortBy       = $('sort-by');
const sortDirBtn   = $('sort-dir-btn');
const movieCount   = $('movie-count');
const availCount   = $('available-count');
const resultsSummary = $('results-summary');
const scanBar      = $('scan-bar');
const lastUpdatedEl  = $('last-updated');
const refreshBtn   = $('refresh-btn');
const scanFill     = $('scan-fill');
const toast        = $('toast');
const rowView      = $('row-view');
const gridView     = $('grid-view');
const movieGrid    = $('movie-grid');
const gridEmpty    = $('grid-empty');
const sidebarClearBtn = $('sidebar-clear-btn');

// ─── UTILITIES ────────────────────────────────────────────────
function normalize(str) { return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

function levenshtein(a, b) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  if (!m) return n; if (!n) return m;
  const dp = Array.from({length: m + 1}, (_, i) =>
    Array.from({length: n + 1}, (_, j) => i === 0 ? j : j === 0 ? i : 0));
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
  return dp[m][n];
}

function normalizeFilename(str) {
  return normalize(String(str || '').replace(/\(\d{4}\)/g, '').replace(/\[.*?\]/g, ''));
}

function findDriveMatch(title, driveMap) {
  const key = normalizeFilename(title);
  for (const [driveKey, val] of Object.entries(driveMap)) {
    if (normalizeFilename(driveKey) === key) return val;
  }
  return null;
}

function parseCSV(text) {
  const lines = text.trim().split('\n');
  if (lines.length < 2) return [];
  const headers = splitCSVLine(lines[0]).map(h => h.trim().toLowerCase().replace(/\s+/g, '_'));
  return lines.slice(1).map(line => {
    const vals = splitCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] || '').trim(); });
    return obj;
  }).filter(r => r[headers[0]]);
}

function splitCSVLine(line) {
  const result = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { if (inQ && line[i+1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
    else if (c === ',' && !inQ) { result.push(cur); cur = ''; }
    else cur += c;
  }
  result.push(cur);
  return result;
}

function extractYear(dateStr) {
  if (!dateStr) return '—';
  const m = String(dateStr).match(/\d{4}/);
  return m ? m[0] : '—';
}

function parseSizeGB(sizeStr) {
  if (!sizeStr) return 0;
  const n = parseFloat(sizeStr), s = sizeStr.toUpperCase();
  if (s.includes('TB')) return n * 1024;
  if (s.includes('GB')) return n;
  if (s.includes('MB')) return n / 1024;
  return n;
}

function parseRuntimeMinutes(str) {
  if (!str) return 0;
  const hm = str.match(/(\d+)\s*h(?:r|ours?)?\s*(\d+)?\s*m?/i);
  if (hm) return parseInt(hm[1]) * 60 + (parseInt(hm[2]) || 0);
  const m = str.match(/(\d+)/);
  return m ? parseInt(m[1]) : 0;
}

const MATURITY_ORDER = { 'G': 1, 'PG': 2, 'PG-13': 3, 'PG13': 3, 'R': 4, 'NC-17': 5, 'NR': 6 };

function parseResolutionScore(res) {
  if (!res) return 0;
  const s = String(res).toUpperCase().trim();
  if (s === '4K' || s === 'UHD' || s.includes('2160')) return 2160;
  const m = s.match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

function imdbClass(rating) {
  const r = parseFloat(rating);
  if (r >= 8) return 'imdb-high';
  if (r >= 6.5) return 'imdb-mid';
  return 'imdb-low';
}

function resClass(res) {
  const r = String(res).toUpperCase();
  if (r.includes('4K') || r.includes('2160')) return 'res-4k';
  if (r.includes('1080')) return 'res-1080';
  if (r.includes('720') || r.includes('576')) return 'res-720';
  return 'res-other';
}

function ratingClass(rating) {
  const r = String(rating || '').toUpperCase().replace(/[\s-]/g, '');
  if (r === 'G')    return 'rating-g';
  if (r === 'PG')   return 'rating-pg';
  if (r === 'PG13') return 'rating-pg13';
  if (r === 'R')    return 'rating-r';
  return '';
}

let toastTimer;

function logClientEvent(event, detail) {
  if (!DRIVE_SCRIPT_URL) return;
  const cbName = '__logCallback_' + Date.now();
  const script = document.createElement('script');
  window[cbName] = function() { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); };
  script.src = DRIVE_SCRIPT_URL
    + '?action=logEvent'
    + '&event='  + encodeURIComponent(event)
    + '&detail=' + encodeURIComponent(detail || '')
    + '&key='    + encodeURIComponent(getSavedKey() || '')
    + '&did='    + encodeURIComponent(getDeviceId())
    + '&callback=' + cbName;
  script.onerror = () => { if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

function showToast(msg, duration = 3000) {
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), duration);
}

function updateLastUpdated(date) {
  const d = (date instanceof Date && !isNaN(date)) ? date : new Date();
  let h = d.getHours(), m = d.getMinutes();
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  const mm = String(m).padStart(2, '0');
  if (lastUpdatedEl) lastUpdatedEl.textContent = h + ':' + mm + ' ' + ampm;
}

function setProgress(pct) { scanFill.style.width = pct + '%'; }

// ─── FETCH & MERGE ────────────────────────────────────────────
async function fetchURL(url, bustCache = false) {
  if (bustCache) url += (url.includes('?') ? '&' : '?') + '_cb=' + Date.now();
  return fetch(url, { redirect: 'follow', cache: bustCache ? 'no-store' : 'default' });
}

function fetchScriptJSON(url, bustCache = false) {
  return new Promise((resolve, reject) => {
    fetchURL(url, bustCache)
      .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(resolve)
      .catch(() => {
        const cbName = '__driveCallback_' + Date.now();
        const script = document.createElement('script');
        const timer = setTimeout(() => { cleanup(); reject(new Error('JSONP timeout')); }, 15000);
        function cleanup() { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }
        window[cbName] = function(data) { cleanup(); resolve(data); };
        const jsonpUrl = bustCache ? url + (url.includes('?') ? '&' : '?') + '_cb=' + Date.now() : url;
        script.src = jsonpUrl + (jsonpUrl.includes('?') ? '&' : '?') + 'callback=' + cbName;
        script.onerror = () => { cleanup(); reject(new Error('JSONP script error')); };
        document.head.appendChild(script);
      });
  });
}

function saveCache(_data) {}
function loadCache() { return null; }

function applyDriveData(rawData, csvRows) {
  const rawMovies = rawData.movies || rawData;
  posterMap = rawData.posters || {};
  if (rawData.requests) {
    requestCounts = {};
    for (const [k, v] of Object.entries(rawData.requests)) requestCounts[normalize(k)] = v;
  }
  if (rawData.ratings) {
    ratingCounts = {};
    try { localStorage.removeItem('thedrive_rating_counts_v1'); } catch(e) {}
    for (const [k, v] of Object.entries(rawData.ratings)) ratingCounts[normalize(k)] = v;
    try { localStorage.setItem('thedrive_rating_counts_v1', JSON.stringify(ratingCounts)); } catch(e) {}
  } else {
    try { const stored = JSON.parse(localStorage.getItem('thedrive_rating_counts_v1') || '{}'); ratingCounts = stored; } catch(e) {}
  }
  const videoMimeTypes = ['video/', 'application/octet-stream'];
  const driveMap = Object.fromEntries(
    Object.entries(rawMovies).filter(([, val]) =>
      typeof val === 'object' && val !== null &&
      (!val.mimeType || videoMimeTypes.some(t => val.mimeType.startsWith(t)))
    )
  );
  allMovies = mergeData(csvRows, driveMap, posterMap);
  render();
  populateFilterCheckboxes();
  updateCounts();
  updateLastUpdated();
}

function fetchRatings(scriptURL, isRefresh = false) {
  return new Promise(resolve => {
    const cbName = '__ratingsCallback_' + Date.now();
    const script = document.createElement('script');
    const timer = setTimeout(() => { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); resolve(); }, 10000);
    window[cbName] = function(data) {
      clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script);
      if (data && data.ratings) {
        ratingCounts = {};
        for (const [k, v] of Object.entries(data.ratings)) ratingCounts[normalize(k)] = v;
        document.querySelectorAll('.rating-btn').forEach(b => {
          const title = b.dataset.ratingTitle, type = b.dataset.ratingType;
          if (!title || !type) return;
          const countEl = b.querySelector('.rating-count');
          if (countEl) countEl.textContent = getRatingCount(title, type) || 0;
        });
        applySort();
      }
      resolve();
    };
    script.src = scriptURL + '?action=getRatings&key=' + encodeURIComponent(getSavedKey() || '') + '&did=' + encodeURIComponent(getDeviceId()) + '&refresh=' + (isRefresh ? '1' : '0') + '&callback=' + cbName + '&_cb=' + Date.now();
    script.onerror = () => { if (script.parentNode) script.parentNode.removeChild(script); resolve(); };
    document.head.appendChild(script);
  });
}

function jsonpAction(url) {
  return new Promise((resolve, reject) => {
    const cbName = '__cb_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    const timer  = setTimeout(() => { cleanup(); reject(new Error('timeout')); }, 15000);
    function cleanup() { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }
    window[cbName] = data => { cleanup(); resolve(data); };
    script.src     = url + (url.includes('?') ? '&' : '?') + 'callback=' + cbName + '&_cb=' + Date.now();
    script.onerror = () => { cleanup(); reject(new Error('script error')); };
    document.head.appendChild(script);
  });
}

const SEQUENTIAL_BATCH_SIZE = 10;

async function loadData(sheetURL, scriptURL, forceRefresh = false) {
  setProgress(5);
  let csvRows = [];
  try {
    const r = await fetchURL(sheetURL, forceRefresh);
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const text = await r.text();
    csvRows = parseCSV(text);
    setProgress(15);
  } catch (e) {
    showToast('⚠ Could not load Sheet CSV. Check the URL & sharing settings.');
    console.error('CSV fetch error:', e);
  }

  allMovies = mergeData(csvRows, {}, {});
  render();
  populateFilterCheckboxes();
  updateCounts();
  setProgress(20);

  const driveURL = scriptURL && scriptURL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE' ? scriptURL : null;
  if (!driveURL) { setProgress(100); setTimeout(() => scanBar.classList.add('hidden'), 300); return; }

  if (forceRefresh) { await loadDataBulkFallback(driveURL, csvRows, true, false); return; }

  setProgress(25);
  let serverPayload = null, cacheAgeS = null;
  try {
    const cacheResult = await jsonpAction(driveURL + '?action=getScanCache&_cb=' + Date.now());
    if (cacheResult && cacheResult.ok && cacheResult.payload && cacheResult.payload.movies) {
      serverPayload = cacheResult.payload;
      cacheAgeS     = typeof cacheResult.age_s === 'number' ? cacheResult.age_s : null;
    }
  } catch (e) { console.warn('getScanCache failed:', e); }

  if (serverPayload) {
    applyDriveData(serverPayload, csvRows);
    setProgress(100);
    setTimeout(() => scanBar.classList.add('hidden'), 300);
    fetchRatings(driveURL, false);
    const cacheDate = cacheAgeS !== null ? new Date(Date.now() - cacheAgeS * 1000) : new Date();
    updateLastUpdated(cacheDate);
    return;
  }

  await loadDataBulkFallback(driveURL, csvRows, false, false);
}

async function loadDataBulkFallback(driveURL, csvRows, forceRefresh, background = false) {
  let files = [];
  try {
    const listData = await jsonpAction(driveURL + '?action=getFileList');
    if (listData && listData.ok && Array.isArray(listData.files)) files = listData.files;
    else throw new Error(listData && listData.error ? listData.error : 'getFileList failed');
  } catch (e) {
    if (!background) { showToast('⚠ Could not load Drive file list. Check the Script URL & deployment.'); setProgress(100); setTimeout(() => scanBar.classList.add('hidden'), 300); }
    console.error('getFileList error:', e);
    fetchRatings(driveURL, forceRefresh);
    return;
  }

  if (files.length === 0) {
    if (!background) { setProgress(100); setTimeout(() => scanBar.classList.add('hidden'), 300); }
    fetchRatings(driveURL, forceRefresh);
    return;
  }

  const accumMovies = {}, accumPosters = {}, accumRequests = {}, accumRatings = {};
  const SCAN_BATCH_SIZE = 10, CONCURRENCY = 6;
  const total = files.length;
  const progressStart = 25, progressEnd = 95;
  const batches = [];
  for (let i = 0; i < total; i += SCAN_BATCH_SIZE) batches.push(files.slice(i, i + SCAN_BATCH_SIZE));
  let completedFiles = 0;

  for (let i = 0; i < batches.length; i += CONCURRENCY) {
    const concurrentBatches = batches.slice(i, i + CONCURRENCY);
    const results = await Promise.all(concurrentBatches.map(batch => {
      const fileIds   = batch.map(f => f.id).join(',');
      const isPosters = batch.map(f => f.isPosters ? '1' : '0').join(',');
      return jsonpAction(driveURL + '?action=scanFiles&fileIds=' + encodeURIComponent(fileIds) + '&isPosters=' + encodeURIComponent(isPosters) + '&key=' + encodeURIComponent(getSavedKey() || '') + '&did=' + encodeURIComponent(getDeviceId())).catch(err => { console.warn('scanFiles batch failed:', err); return null; });
    }));
    for (const result of results) {
      if (result && result.ok) { Object.assign(accumMovies, result.movies || {}); Object.assign(accumPosters, result.posters || {}); }
    }
    completedFiles += concurrentBatches.reduce((s, b) => s + b.length, 0);
    if (!background) setProgress(progressStart + ((completedFiles / total) * (progressEnd - progressStart)));
    applyDriveData({ movies: accumMovies, posters: accumPosters, requests: accumRequests, ratings: accumRatings }, csvRows);
  }

  let liveRequests = {}, liveRatings = {};
  try {
    const ratingsData = await new Promise((resolve) => {
      const cbName = '__postScanRatings_' + Date.now();
      const script = document.createElement('script');
      const timer  = setTimeout(() => { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); resolve({}); }, 10000);
      window[cbName] = function(data) { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); resolve(data || {}); };
      script.src = driveURL + '?action=getRatings&key=' + encodeURIComponent(getSavedKey() || '') + '&did=' + encodeURIComponent(getDeviceId()) + '&callback=' + cbName + '&_cb=' + Date.now();
      script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); resolve({}); };
      document.head.appendChild(script);
    });
    if (ratingsData.ratings) {
      for (const [k, v] of Object.entries(ratingsData.ratings)) liveRatings[normalize(k)] = v;
      ratingCounts = { ...liveRatings };
    }
  } catch(e) {}

  try {
    // Fetch request counts live from the sheet so the full-scan payload
    // always contains up-to-date data (getScanCache would be stale here
    // since the new cache hasn't been written yet).
    const reqResult = await jsonpAction(driveURL + '?action=getRequests&key=' + encodeURIComponent(getSavedKey() || '') + '&did=' + encodeURIComponent(getDeviceId()) + '&_cb=' + Date.now());
    if (reqResult && reqResult.requests) {
      for (const [k, v] of Object.entries(reqResult.requests)) liveRequests[normalize(k)] = v;
      requestCounts = { ...liveRequests };
    }
  } catch(e) {}

  const finalPayload = { movies: accumMovies, posters: accumPosters, requests: liveRequests, ratings: liveRatings };
  applyDriveData(finalPayload, csvRows);

  { const total = allMovies.length, available = allMovies.filter(m => m.available).length; pushSnapshot(total, available); }

  try {
    fetch(driveURL, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ action: 'writeCache', payload: finalPayload, key: getSavedKey() || '', did: getDeviceId() }), mode: 'no-cors' }).catch(() => {});
  } catch(e) {}

  if (!background) { setProgress(100); setTimeout(() => scanBar.classList.add('hidden'), 300); }
  updateLastUpdated();
}

function mergeData(rows, driveMap, posterMap = {}) {
  const mapped = rows.map(row => {
    const title          = row.title || row.movie_title || row['movie title'] || '';
    const runtime        = row.runtime || row.run_time || row.duration || '';
    const resolution     = row.resolution || row.res || '';
    const maturityRating = row.maturity_rating || row.rating || row.maturityrating || '';
    const releaseDate    = row.release_date || row.releasedate || row.date || '';
    const fileSize       = row.file_size || row.filesize || row.size || '';
    const imdbRating     = row.imdb_rating || row.imdbrating || row.imdb || '';
    const match          = findDriveMatch(title, driveMap);
    const posterKey      = normalize(title);
    const poster         = posterMap[posterKey] || null;
    return { title, runtime, resolution, maturityRating, releaseDate, year: extractYear(releaseDate), fileSize, imdbRating, available: !!match, driveLink: match ? match.link : null, driveResolution: match ? (match.name || '') : '', poster };
  }).filter(m => m.title);

  const seen = new Map();
  for (const m of mapped) {
    const key = normalize(m.title);
    if (!seen.has(key)) { seen.set(key, m); }
    else {
      const existing = seen.get(key);
      if (m.available && m.driveResolution) {
        const driveRes = m.driveResolution.toUpperCase(), mRes = m.resolution.toUpperCase(), existingRes = existing.resolution.toUpperCase();
        const mMatches = driveRes.includes(mRes) || mRes.split('P')[0] === driveRes.split('P')[0];
        const existingMatches = driveRes.includes(existingRes) || existingRes.split('P')[0] === driveRes.split('P')[0];
        if (mMatches && !existingMatches) seen.set(key, m);
      }
    }
  }
  return [...seen.values()];
}

// ─── SIDEBAR FILTER POPULATION ────────────────────────────────
function populateFilterCheckboxes() {
  // Maturity
  const maturityEl = $('filter-maturity-checks');
  if (maturityEl) {
    const ratings = [...new Set(allMovies.map(m => m.maturityRating).filter(Boolean))].sort((a, b) => {
      const oa = MATURITY_ORDER[a.toUpperCase().replace(/[\s-]/g,'')] || 99;
      const ob = MATURITY_ORDER[b.toUpperCase().replace(/[\s-]/g,'')] || 99;
      return oa - ob;
    });
    maturityEl.innerHTML = ratings.map(r => `
      <label class="check-row">
        <input type="checkbox" value="${escHtml(r)}" data-filter="maturity" ${activeFilters.maturity.has(r) ? 'checked' : ''} />
        <span class="check-label ${ratingClass(r)}">${escHtml(r)}</span>
      </label>`).join('');
  }

  // Resolution
  const resEl = $('filter-resolution-checks');
  if (resEl) {
    const resolutions = [...new Set(allMovies.map(m => m.resolution).filter(Boolean))].sort((a, b) => parseResolutionScore(b) - parseResolutionScore(a));
    resEl.innerHTML = resolutions.map(r => `
      <label class="check-row">
        <input type="checkbox" value="${escHtml(r)}" data-filter="resolution" ${activeFilters.resolution.has(r) ? 'checked' : ''} />
        <span class="check-label ${resClass(r)}">${escHtml(r)}</span>
      </label>`).join('');
  }

  // Re-bind checkbox listeners
  bindSidebarCheckboxes();
  updateClearBtn();
}

function bindSidebarCheckboxes() {
  document.querySelectorAll('.sidebar-checks input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      const filterType = cb.dataset.filter;
      const value      = cb.value;
      if (cb.checked) activeFilters[filterType].add(value);
      else            activeFilters[filterType].delete(value);
      updateClearBtn();
      render();
      saveSettings();
    });
  });
}

function updateClearBtn() {
  if (!sidebarClearBtn) return;
  const anyActive = activeFilters.maturity.size > 0 || activeFilters.status.size > 0 || activeFilters.resolution.size > 0;
  sidebarClearBtn.hidden = !anyActive;
}

function clearAllFilters() {
  activeFilters.maturity.clear();
  activeFilters.status.clear();
  activeFilters.resolution.clear();
  if (searchInput) { searchInput.value = ''; clearSearch && clearSearch.classList.remove('visible'); }
  document.querySelectorAll('.sidebar-checks input[type="checkbox"]').forEach(cb => cb.checked = false);
  updateClearBtn();
  render();
  saveSettings();
}

function updateCounts() {
  if (movieCount) movieCount.textContent = allMovies.length + ' films';
  if (availCount) availCount.textContent = allMovies.filter(m => m.available).length + ' available';
}

// ─── SORT & FILTER ────────────────────────────────────────────
function applyFilters() {
  const q = normalize(searchInput ? searchInput.value : '');

  filtered = allMovies.filter(m => {
    if (q && !normalize(m.title).includes(q)) return false;
    if (activeFilters.maturity.size   > 0 && !activeFilters.maturity.has(m.maturityRating))     return false;
    if (activeFilters.resolution.size > 0 && !activeFilters.resolution.has(m.resolution))       return false;
    if (activeFilters.status.size > 0) {
      const status = m.available ? 'Available' : 'Not Uploaded';
      if (!activeFilters.status.has(status)) return false;
    }
    return true;
  });

  applySort();
}

function applySort() {
  const key = currentSort, dir = currentDir;
  filtered.sort((a, b) => {
    let va, vb;
    if (key === 'title')    { va = a.title.toLowerCase(); vb = b.title.toLowerCase(); }
    else if (key === 'imdb')     { va = parseFloat(a.imdbRating) || 0; vb = parseFloat(b.imdbRating) || 0; }
    else if (key === 'year')     { va = parseInt(a.year) || 0; vb = parseInt(b.year) || 0; }
    else if (key === 'size')     { va = parseSizeGB(a.fileSize); vb = parseSizeGB(b.fileSize); }
    else if (key === 'requests') { va = getRequestCount(a.title); vb = getRequestCount(b.title); }
    else if (key === 'rating')   { va = getRatingScore(a.title); vb = getRatingScore(b.title); }
    else if (key === 'runtime')  { va = parseRuntimeMinutes(a.runtime); vb = parseRuntimeMinutes(b.runtime); }
    else if (key === 'maturity') {
      va = MATURITY_ORDER[a.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99;
      vb = MATURITY_ORDER[b.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99;
    }
    else if (key === 'status')   { va = a.available ? 0 : 1; vb = b.available ? 0 : 1; }
    else if (key === 'res')      { va = parseResolutionScore(a.resolution); vb = parseResolutionScore(b.resolution); }
    if (va < vb) return dir === 'asc' ? -1 : 1;
    if (va > vb) return dir === 'asc' ? 1 : -1;
    return 0;
  });

  renderCurrentView();
}

// ─── RENDER ───────────────────────────────────────────────────
function render() { applyFilters(); }

function renderCurrentView() {
  if (hasActiveFilters()) {
    // Switch to grid view when any filter/search is active
    rowView.classList.remove('active');
    gridView.classList.add('active');
    renderGrid();
    const total = allMovies.length, shown = filtered.length;
    if (resultsSummary) resultsSummary.textContent = shown === total ? `Showing all ${total} films` : `Showing ${shown} of ${total} films`;
  } else {
    // Default: show row-based browse view
    gridView.classList.remove('active');
    rowView.classList.add('active');
    renderRows();
    if (resultsSummary) resultsSummary.textContent = `${allMovies.length} films in the library`;
  }
}

// ── Row View ──
function renderRows() {
  const availableMovies = [...allMovies]
    .filter(m => m.available)
    .sort((a, b) => getRatingScore(b.title) - getRatingScore(a.title));

  const requestedMovies = [...allMovies]
    .filter(m => getRequestCount(m.title) > 0)
    .sort((a, b) => getRequestCount(b.title) - getRequestCount(a.title));

  const imdbMovies = [...allMovies]
    .filter(m => parseFloat(m.imdbRating) > 0)
    .sort((a, b) => (parseFloat(b.imdbRating) || 0) - (parseFloat(a.imdbRating) || 0));

  const rowAvailableEl   = $('row-available-cards');
  const rowRequestedEl   = $('row-requested-cards');
  const rowImdbEl        = $('row-imdb-cards');
  const rowRequestedSec  = $('row-requested');

  if (rowAvailableEl)  renderRowCards(rowAvailableEl, availableMovies.slice(0, 30));

  // Hide Most Requested row if no requests yet
  if (rowRequestedEl && rowRequestedSec) {
    if (requestedMovies.length > 0) {
      rowRequestedSec.style.display = '';
      renderRowCards(rowRequestedEl, requestedMovies.slice(0, 30));
    } else {
      rowRequestedSec.style.display = 'none';
    }
  }

  if (rowImdbEl) renderRowCards(rowImdbEl, imdbMovies.slice(0, 30));
}

function renderRowCards(container, movies) {
  container.innerHTML = '';
  const frag = document.createDocumentFragment();
  movies.forEach((m, i) => {
    const card = buildCard(m, i, true);
    frag.appendChild(card);
  });
  container.appendChild(frag);
  // Update button visibility after cards are rendered
  const scroller = container.closest('.movie-row-scroll');
  if (scroller) updateRowScrollBtns(scroller);
}

// ── Row scroll buttons ──
(function initRowScrollBtns() {
  document.addEventListener('click', function(e) {
    const btn = e.target.closest('.row-scroll-btn');
    if (!btn) return;
    const dir      = parseInt(btn.dataset.dir, 10);
    const targetId = btn.dataset.target;
    const track    = document.getElementById(targetId);
    if (!track) return;
    const scroller = track.closest('.movie-row-scroll');
    if (!scroller) return;
    // Scroll by ~3 card widths (200px card + 14px gap)
    scroller.scrollBy({ left: dir * (214 * 3), behavior: 'smooth' });
  });

  // Update button visibility on scroll
  document.addEventListener('scroll', handleRowScroll, true);

  function handleRowScroll(e) {
    if (!e.target.classList || !e.target.classList.contains('movie-row-scroll')) return;
    updateRowScrollBtns(e.target);
  }
})();

function updateRowScrollBtns(scroller) {
  const wrapper = scroller.closest('.row-scroll-wrapper');
  if (!wrapper) return;
  const leftBtn  = wrapper.querySelector('.row-scroll-btn--left');
  const rightBtn = wrapper.querySelector('.row-scroll-btn--right');
  if (!leftBtn || !rightBtn) return;
  const atStart = scroller.scrollLeft <= 4;
  const atEnd   = scroller.scrollLeft >= scroller.scrollWidth - scroller.clientWidth - 4;
  leftBtn.dataset.hidden  = atStart ? '1' : '0';
  rightBtn.dataset.hidden = atEnd   ? '1' : '0';
}

// ── Grid View ──
function renderGrid() {
  if (!movieGrid) return;
  movieGrid.innerHTML = '';

  if (filtered.length === 0) {
    if (gridEmpty) gridEmpty.hidden = false;
    return;
  }
  if (gridEmpty) gridEmpty.hidden = true;

  const frag = document.createDocumentFragment();
  filtered.forEach((m, i) => {
    const card = buildCard(m, i, false);
    frag.appendChild(card);
  });
  movieGrid.appendChild(frag);
}

// ── Shared card builder ──
function buildCard(m, i, isRowCard) {
  const card = document.createElement('div');
  const key  = normalize(m.title);
  card.className = isRowCard ? 'movie-card row-card' : 'movie-card';
  card.dataset.key = key;
  card.style.animationDelay = Math.min(i * 30, 400) + 'ms';

  const cardReqCount   = getRequestCount(m.title);
  const cardIRequested = hasUserRequested(m.title);
  const posterClasses  = ['card-poster'];
  if (m.driveLink) posterClasses.push('card-poster--playable');
  else posterClasses.push('card-poster--requestable');

  card.innerHTML = `
    <div class="${posterClasses.join(' ')}">
      ${m.poster ? `<img src="${m.poster}" alt="${escHtml(m.title)}" loading="lazy" onload="this.classList.add('loaded')" />` : ''}
      ${m.driveLink
        ? `<a class="card-play-overlay drive-link" href="${m.driveLink}" target="_blank" rel="noopener" data-title="${escHtml(m.title)}" aria-label="Watch ${escHtml(m.title)}"><div class="card-play-btn"><span class="card-play-icon">&#9654;</span></div></a>`
        : cardIRequested
          ? `<div class="card-play-overlay card-request-overlay card-request-overlay--done"><div class="card-request-btn card-request-btn--done"><span class="card-request-icon">&#10003;</span><span class="card-request-label">REQUESTED${cardReqCount ? ' <span class="request-count">' + cardReqCount + '</span>' : ''}</span></div></div>`
          : `<button class="card-play-overlay card-request-overlay request-btn" data-title="${escHtml(m.title)}" aria-label="Request ${escHtml(m.title)}"><div class="card-request-btn"><span class="card-request-icon">&#65291;</span><span class="card-request-label">REQUEST${cardReqCount ? ' <span class="request-count">' + cardReqCount + '</span>' : ''}</span></div></button>`}
    </div>
    <div class="card-title">${escHtml(m.title)}</div>
    <div class="card-meta">
      <span class="card-year">${escHtml(m.year)}</span>
      <span class="card-sep">·</span>
      <span class="card-rating ${ratingClass(m.maturityRating)}">${escHtml(m.maturityRating) || '—'}</span>
      ${m.runtime ? `<span class="card-sep">·</span><span class="card-runtime">${escHtml(m.runtime)}</span>` : ''}
    </div>
    <div class="card-row">
      <span class="card-imdb ${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span>
      <span class="card-res ${resClass(m.resolution)}">${escHtml(m.resolution) || '—'}</span>
    </div>
    <div class="card-footer">
      <span class="status-pill ${m.available ? 'status-available' : 'status-missing'}">
        ${m.available ? 'AVAILABLE' : 'NOT UPLOADED'}
      </span>
      ${m.driveLink ? `<div class="card-rating-row">${ratingHTML(m.title)}</div>` : ''}
    </div>
  `;
  return card;
}

function ratingHTML(title) {
  const userVote = getUserRating(title);
  const ups = getRatingCount(title, 'up'), downs = getRatingCount(title, 'down');
  return `<div class="rating-wrap">
    <button class="rating-btn rating-btn--up ${userVote === 'up' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="up" title="Liked it">
      ▲<span class="rating-count">${ups || 0}</span>
    </button>
    <button class="rating-btn rating-btn--down ${userVote === 'down' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="down" title="Didn't like it">
      ▼<span class="rating-count">${downs || 0}</span>
    </button>
  </div>`;
}

function escHtml(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─── REQUEST COUNT TOOLTIP ────────────────────────────────────
(function initRequestTooltip() {
  const tooltip = document.getElementById('req-tooltip');
  if (!tooltip) return;
  let hideTimer;

  function showTooltipForTitle(title, anchorEl) {
    clearTimeout(hideTimer);
    const count = getRequestCount(title);
    tooltip.innerHTML = count <= 0
      ? '<span class="tt-count">0</span> requests'
      : '<span class="tt-count">' + count + '</span> request' + (count === 1 ? '' : 's');
    tooltip.removeAttribute('hidden');
    tooltip.getBoundingClientRect();
    tooltip.classList.add('visible');
    positionTooltip(anchorEl);
  }

  function positionTooltip(el) {
    const rect = el.getBoundingClientRect(), ttW = tooltip.offsetWidth, ttH = tooltip.offsetHeight;
    let left = rect.left + rect.width / 2 - ttW / 2, top = rect.top - ttH - 8;
    if (left < 8) left = 8;
    if (left + ttW > window.innerWidth - 8) left = window.innerWidth - ttW - 8;
    if (top < 8) top = rect.bottom + 8;
    tooltip.style.left = left + 'px'; tooltip.style.top = top + 'px';
  }

  function hideTooltip() {
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => { tooltip.classList.remove('visible'); tooltip.setAttribute('hidden', ''); }, 80);
  }

  const mainContent = document.getElementById('main-content');
  if (!mainContent) return;
  mainContent.addEventListener('mouseover', e => {
    const cardTitle = e.target.closest('.card-title');
    if (cardTitle) { showTooltipForTitle(cardTitle.textContent.trim(), cardTitle); }
  });
  mainContent.addEventListener('mouseout', e => {
    if (e.target.closest('.card-title')) hideTooltip();
  });
  document.addEventListener('scroll', () => { tooltip.classList.remove('visible'); }, { passive: true });
})();

// ─── EVENTS ───────────────────────────────────────────────────

// Search
let searchTimer, searchLogTimer;
if (searchInput) {
  searchInput.addEventListener('input', () => {
    const query = searchInput.value;
    if (clearSearch) clearSearch.classList.toggle('visible', query.length > 0);
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => { render(); saveSettings(); }, 200);
    clearTimeout(searchLogTimer);
    if (query.trim().length > 0) searchLogTimer = setTimeout(() => logClientEvent('Search', query.trim()), 1500);
  });
}

if (clearSearch) {
  clearSearch.addEventListener('click', () => {
    searchInput.value = '';
    clearSearch.classList.remove('visible');
    render();
    saveSettings();
    searchInput.focus();
  });
}

// Sidebar sort
if (sortBy) {
  sortBy.addEventListener('change', () => {
    currentSort = sortBy.value;
    if (hasActiveFilters()) applySort();
    saveSettings();
    logClientEvent('Sort', currentSort + '-' + currentDir);
  });
}

if (sortDirBtn) {
  sortDirBtn.addEventListener('click', () => {
    currentDir = currentDir === 'desc' ? 'asc' : 'desc';
    sortDirBtn.textContent = currentDir === 'desc' ? '↓' : '↑';
    sortDirBtn.title = currentDir === 'desc' ? 'Descending' : 'Ascending';
    if (hasActiveFilters()) applySort();
    saveSettings();
    logClientEvent('Sort Direction', currentDir);
  });
}

// Clear all filters button
if (sidebarClearBtn) {
  sidebarClearBtn.addEventListener('click', clearAllFilters);
}

// Watch link clicks
const mainContent = document.getElementById('main-content');
if (mainContent) {
  mainContent.addEventListener('click', e => {
    const link = e.target.closest('.drive-link');
    if (!link) return;
    const title = link.dataset.title || '', key = getSavedKey() || '';
    if (!DRIVE_SCRIPT_URL || !key) return;
    const cbName = '__openLinkCallback_' + Date.now();
    const script = document.createElement('script');
    window[cbName] = function() { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); };
    script.src = DRIVE_SCRIPT_URL + '?action=openLink&title=' + encodeURIComponent(title) + '&key=' + encodeURIComponent(key) + '&did=' + encodeURIComponent(getDeviceId()) + '&callback=' + cbName;
    script.onerror = () => { if (script.parentNode) script.parentNode.removeChild(script); };
    document.head.appendChild(script);
  });

  // Rating buttons
  mainContent.addEventListener('click', async e => {
    const btn = e.target.closest('.rating-btn');
    if (!btn) return;
    const title = btn.dataset.ratingTitle, type = btn.dataset.ratingType;
    if (!title || !type) return;
    await postRating(title, type);
    document.querySelectorAll(`[data-rating-title="${CSS.escape(title)}"]`).forEach(b => {
      const userVote = getUserRating(title);
      b.classList.toggle('active', userVote === b.dataset.ratingType);
      const countEl = b.querySelector('.rating-count');
      if (countEl) countEl.textContent = getRatingCount(title, b.dataset.ratingType) || 0;
    });
    const userVote = getUserRating(title);
    if (userVote === 'up') showToast('▲ You liked ' + title);
    else if (userVote === 'down') showToast('▼ You disliked ' + title);
    else showToast('Rating removed for ' + title);
  });

  // Request buttons
  mainContent.addEventListener('click', async e => {
    const btn = e.target.closest('.request-btn');
    if (!btn || btn.classList.contains('request-btn--done')) return;
    btn.disabled = true;
    const title = btn.dataset.title;
    const optimisticCount = (requestCounts[normalize(title)] || 0) + 1;
    setRequestedState(title, optimisticCount);
    const serverCount = await postRequest(title);
    if (serverCount !== optimisticCount) setRequestedState(title, serverCount);
    showToast('✓ Requested: ' + title);
  });
}

function setRequestedState(title, count) {
  const countHtml = count ? ' <span class="request-count">' + count + '</span>' : '';
  document.querySelectorAll('.request-btn[data-title="' + CSS.escape(title) + '"]:not(.card-request-overlay)').forEach(b => {
    b.disabled = false; b.classList.add('request-btn--done');
    b.innerHTML = '<span class="request-icon">&#10003;</span> REQUESTED' + countHtml;
    b.dataset.title = title;
  });
  document.querySelectorAll('.card-request-overlay.request-btn[data-title="' + CSS.escape(title) + '"]').forEach(b => {
    b.disabled = true; b.classList.add('card-request-overlay--done');
    const inner = b.querySelector('.card-request-btn');
    if (inner) { inner.classList.add('card-request-btn--done'); inner.innerHTML = '<span class="card-request-icon">&#10003;</span><span class="card-request-label">REQUESTED' + countHtml + '</span>'; }
    b.dataset.title = title;
  });
}

// ─── FOOTER COMMENT FORM ─────────────────────────────────────
(function() {
  const submitBtn = document.getElementById('footer-submit');
  const nameInput = document.getElementById('footer-name');
  const msgInput  = document.getElementById('footer-message');
  const statusEl  = document.getElementById('footer-form-status');
  if (!submitBtn) return;

  function setStatus(msg, type) { statusEl.textContent = msg; statusEl.className = 'footer-form-status ' + type; statusEl.hidden = false; }

  submitBtn.addEventListener('click', async () => {
    const message = (msgInput.value || '').trim();
    if (!message) { setStatus('Please enter a message before sending.', 'error'); msgInput.focus(); return; }
    submitBtn.disabled = true; submitBtn.textContent = 'SENDING…'; statusEl.hidden = true;
    const name = (nameInput.value || '').trim() || 'Anonymous', key = getSavedKey() || '';
    try {
      await new Promise((resolve, reject) => {
        const cbName = '__formCallback_' + Date.now();
        const script = document.createElement('script');
        const timer  = setTimeout(() => { cleanup(); reject(new Error('timeout')); }, 12000);
        function cleanup() { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }
        window[cbName] = function(data) { cleanup(); if (data && data.ok) resolve(); else reject(new Error(data && data.error ? data.error : 'unknown')); };
        script.src = DRIVE_SCRIPT_URL + '?action=submitForm&name=' + encodeURIComponent(name) + '&message=' + encodeURIComponent(message) + '&key=' + encodeURIComponent(key) + '&did=' + encodeURIComponent(getDeviceId()) + '&callback=' + cbName + '&_cb=' + Date.now();
        script.onerror = () => { cleanup(); reject(new Error('network')); };
        document.head.appendChild(script);
      });
      setStatus('✓ Message sent — thanks!', 'success');
      nameInput.value = ''; msgInput.value = '';
    } catch(err) {
      setStatus('Something went wrong. Please try again.', 'error');
    } finally {
      submitBtn.disabled = false; submitBtn.textContent = 'SEND';
    }
  });
})();

// ─── REFRESH BUTTON ───────────────────────────────────────────
if (refreshBtn) {
  refreshBtn.addEventListener('click', async () => {
    if (refreshBtn.classList.contains('spinning')) return;
    refreshBtn.classList.add('spinning');
    scanBar.classList.remove('hidden');
    setProgress(0);
    requestCounts = {}; ratingCounts = {};
    try { localStorage.removeItem('thedrive_cache_v3'); } catch(e) {}
    try { localStorage.removeItem('thedrive_requests_v1'); } catch(e) {}
    try { localStorage.removeItem('thedrive_rating_counts_v1'); } catch(e) {}
    if (DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
      try { await jsonpAction(DRIVE_SCRIPT_URL + '?action=bustCache&key=' + encodeURIComponent(getSavedKey() || '') + '&did=' + encodeURIComponent(getDeviceId())); } catch(e) {}
    }
    await loadData(SHEET_CSV_URL, DRIVE_SCRIPT_URL, true);
    updateLastUpdated();
    refreshBtn.classList.remove('spinning');
  });
}

// ─── INIT ─────────────────────────────────────────────────────
(async function init() {
  const savedSettings = loadSettings();
  if (savedSettings) applySettings(savedSettings);
  else {
    currentSort = 'title'; currentDir = 'asc';
    if (sortBy) sortBy.value = 'title';
    if (sortDirBtn) sortDirBtn.textContent = '↓';
  }

  await initWithGate();
  loadData(SHEET_CSV_URL, DRIVE_SCRIPT_URL);

  if (DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
    fetchOnlineCount();
    setInterval(fetchOnlineCount, 60 * 1000);
    setInterval(pingHeartbeat, 4 * 60 * 1000);
    setInterval(pushPresencePing, 10 * 1000);
  }

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + tab).classList.add('active');
      if (tab === 'stats') initStatsTab();
    });
  });
})();

// ─── ONLINE COUNT ─────────────────────────────────────────────
function fetchOnlineCount() {
  const cbName = '__onlineCallback_' + Date.now();
  const script = document.createElement('script');
  const timer  = setTimeout(() => { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }, 10000);
  window[cbName] = function(data) {
    clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script);
    const el = document.getElementById('online-count');
    if (el && data && typeof data.online === 'number') el.textContent = data.online;
  };
  script.src = DRIVE_SCRIPT_URL + '?action=getOnlineCount&callback=' + cbName + '&_cb=' + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

function pingHeartbeat() {
  const cbName = '__heartbeatCallback_' + Date.now();
  const script = document.createElement('script');
  const timer  = setTimeout(() => { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }, 10000);
  window[cbName] = function() { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); };
  script.src = DRIVE_SCRIPT_URL + '?action=checkDevice&did=' + encodeURIComponent(getDeviceId()) + '&key=' + encodeURIComponent(getSavedKey() || '') + '&callback=' + cbName + '&_cb=' + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

// ─── STATS TAB ────────────────────────────────────────────────
let statsLoaded = false, statsLoadedAt = 0;
let chartLibrary = null, chartUsers = null, chartPresence = null;
let lastPresenceAppendAt = 0;

function initStatsTab() {
  renderLocalStats();
  if ((!statsLoaded || Date.now() - statsLoadedAt > 60000) && DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
    fetchStatsData(); statsLoaded = true; statsLoadedAt = Date.now();
  }
}

function renderLocalStats() {
  if (!allMovies.length) return;
  const total = allMovies.length, available = allMovies.filter(m => m.available).length;
  const pct = total > 0 ? ((available / total) * 100).toFixed(2) : '0.00';
  const fracEl = $('upload-fraction'), pctEl = $('upload-pct'), fillEl = $('upload-fill');
  if (fracEl)  fracEl.textContent  = available + ' / ' + total + ' films uploaded';
  if (pctEl)   pctEl.textContent   = pct + '%';
  if (fillEl)  fillEl.style.width  = parseFloat(pct) + '%';
  setText('stat-total-films', total);
  setText('stat-available', available);
  let totalGB = 0; allMovies.forEach(m => { totalGB += parseSizeGB(m.fileSize); });
  setText('stat-total-size', totalGB > 0 ? totalGB.toFixed(1) + ' GB' : '—');
  let totalMins = 0; allMovies.forEach(m => { totalMins += parseRuntimeMinutes(m.runtime); });
  if (totalMins > 0) setText('stat-total-runtime', Math.floor(totalMins / 60) + 'h ' + (totalMins % 60) + 'm');
  const rated = allMovies.filter(m => parseFloat(m.imdbRating) > 0);
  if (rated.length) setText('stat-avg-imdb', '★ ' + (rated.reduce((s, m) => s + parseFloat(m.imdbRating), 0) / rated.length).toFixed(1));
  const matNorm = r => String(r || '').toUpperCase().replace(/[\s-]/g, '');
  setText('stat-g',    allMovies.filter(m => matNorm(m.maturityRating) === 'G').length);
  setText('stat-pg',   allMovies.filter(m => matNorm(m.maturityRating) === 'PG').length);
  setText('stat-pg13', allMovies.filter(m => matNorm(m.maturityRating) === 'PG13').length);
  setText('stat-r',    allMovies.filter(m => matNorm(m.maturityRating) === 'R').length);
  setText('stat-4k',   allMovies.filter(m => /4k|2160/i.test(m.resolution)).length);
  setText('stat-1080', allMovies.filter(m => /1080/i.test(m.resolution)).length);
}

function setText(id, val) { const el = $(id); if (el) el.textContent = String(val); }

function fetchStatsData() {
  const cbName = '__statsCallback_' + Date.now();
  const script = document.createElement('script');
  const timer  = setTimeout(() => { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }, 15000);
  window[cbName] = function(data) {
    clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script);
    if (!data) return;
    if (typeof data.uniqueDevices === 'number') setText('stat-total-users', data.uniqueDevices);
    if (data.snapshots   && data.snapshots.length)   renderLibraryChart(data.snapshots);
    if (data.userHistory && data.userHistory.length)  renderUserChart(data.userHistory);
    if (data.presence    && data.presence.length)     renderPresenceChart(data.presence);
    else showPresencePlaceholder();
  };
  script.src = DRIVE_SCRIPT_URL + '?action=getStatsData&callback=' + cbName + '&_cb=' + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

function renderLibraryChart(snapshots) {
  const canvas = $('chart-library'); if (!canvas) return;
  const cfg = { type: 'line', data: { labels: snapshots.map(s => s.date), datasets: [{ label: 'Total Films', data: snapshots.map(s => s.total), borderColor: '#9090a8', backgroundColor: 'rgba(144,144,168,0.08)', borderWidth: 2, pointRadius: 3, pointBackgroundColor: '#9090a8', tension: 0.3, fill: true }, { label: 'Available', data: snapshots.map(s => s.available), borderColor: '#e8c547', backgroundColor: 'rgba(232,197,71,0.10)', borderWidth: 2, pointRadius: 3, pointBackgroundColor: '#e8c547', tension: 0.3, fill: true }] }, options: chartOptions('Films') };
  if (chartLibrary) chartLibrary.destroy();
  chartLibrary = new Chart(canvas, cfg);
}

function renderUserChart(userHistory) {
  const canvas = $('chart-users'); if (!canvas) return;
  const cfg = { type: 'line', data: { labels: userHistory.map(u => u.date), datasets: [{ label: 'Unique Users', data: userHistory.map(u => u.users), borderColor: '#e8c547', backgroundColor: 'rgba(232,197,71,0.10)', borderWidth: 2, pointRadius: 3, pointBackgroundColor: '#e8c547', tension: 0.3, fill: true }] }, options: chartOptions('Users') };
  if (chartUsers) chartUsers.destroy();
  chartUsers = new Chart(canvas, cfg);
}

function showPresencePlaceholder() {
  const canvas = $('chart-presence'); if (!canvas) return;
  const wrap = canvas.closest('.chart-wrap'); if (!wrap) return;
  canvas.style.display = 'none';
  if (!wrap.querySelector('.presence-placeholder')) {
    const msg = document.createElement('div');
    msg.className = 'presence-placeholder';
    msg.innerHTML = `<span class="presence-placeholder-icon">◎</span><p>No history yet — data will appear here as users come online.</p>`;
    wrap.appendChild(msg);
  }
}

function showPresenceCanvas() {
  const canvas = $('chart-presence'); if (!canvas) return;
  canvas.style.display = '';
  const wrap = canvas.closest('.chart-wrap');
  if (wrap) { const ph = wrap.querySelector('.presence-placeholder'); if (ph) ph.remove(); }
}

function renderPresenceChart(presence) {
  showPresenceCanvas();
  const canvas = $('chart-presence'); if (!canvas) return;
  const INTERVAL_MS = 10 * 1000, GAP_THRESH = INTERVAL_MS * 2;
  function tsToMs(ts) { return new Date(ts.replace(' ', 'T')).getTime(); }
  const filled = [];
  for (let i = 0; i < presence.length; i++) {
    filled.push(presence[i]);
    if (i < presence.length - 1) {
      const gap = tsToMs(presence[i + 1].ts) - tsToMs(presence[i].ts);
      if (gap > GAP_THRESH) {
        const afterTs = new Date(tsToMs(presence[i].ts) + INTERVAL_MS);
        const pad = n => String(n).padStart(2, '0');
        filled.push({ ts: presence[i].ts.slice(0, 11) + pad(afterTs.getHours()) + ':' + pad(afterTs.getMinutes()) + ':' + pad(afterTs.getSeconds()), online: 0 });
      }
    }
  }
  const step = Math.max(1, Math.floor(filled.length / 500));
  const sampled = filled.filter((_, i) => i % step === 0);
  const times = sampled.map(p => { const m = String(p.ts).match(/(\d{1,2}:\d{2})(?::\d{2})?/); return m ? m[1] : ''; });
  const rawValues = sampled.map(p => p.online);
  const values = rawValues.map((v, i, arr) => {
    const p2 = arr[i-2] !== undefined ? arr[i-2] : v, p1 = arr[i-1] !== undefined ? arr[i-1] : v;
    const n1 = arr[i+1] !== undefined ? arr[i+1] : v, n2 = arr[i+2] !== undefined ? arr[i+2] : v;
    return Math.round((p2 + p1 + v + n1 + n2) / 5 * 100) / 100;
  });
  const cfg = { type: 'line', data: { labels: sampled.map((_, i) => i), datasets: [{ label: 'Online', data: values, borderColor: '#3ecf74', backgroundColor: 'rgba(62,207,116,0.10)', borderWidth: 2, pointRadius: 0, tension: 0, fill: true }] }, options: presenceChartOptions(times) };
  if (chartPresence) chartPresence.destroy();
  chartPresence = new Chart(canvas, cfg);
  chartPresence._times = times;
  lastPresenceAppendAt = Date.now();
}

function chartOptions(yLabel) {
  return {
    responsive: true, maintainAspectRatio: true,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { labels: { color: '#9090a8', font: { family: 'DM Mono', size: 11 }, boxWidth: 12 } },
      tooltip: { backgroundColor: '#18181f', borderColor: '#252530', borderWidth: 1, titleColor: '#e8e8f0', bodyColor: '#9090a8', titleFont: { family: 'DM Mono', size: 11 }, bodyFont: { family: 'DM Mono', size: 11 } }
    },
    scales: {
      x: { ticks: { color: '#78788f', font: { family: 'DM Mono', size: 10 }, maxTicksLimit: 10 }, grid: { color: 'rgba(37,37,48,0.6)' } },
      y: { title: { display: true, text: yLabel, color: '#78788f', font: { family: 'DM Mono', size: 10 } }, ticks: { color: '#78788f', font: { family: 'DM Mono', size: 10 }, precision: 0 }, grid: { color: 'rgba(37,37,48,0.6)' }, beginAtZero: true }
    }
  };
}

function presenceChartOptions(times) {
  const base = chartOptions('Users');
  base.scales.x.type = 'category';
  base.scales.x.ticks = { color: '#78788f', font: { family: 'DM Mono', size: 10 }, maxTicksLimit: 10, callback: function(val) { return times[val] || ''; } };
  base.plugins.tooltip.callbacks = { title: function(items) { return times[items[0].dataIndex] || ''; } };
  return base;
}

function pushPresencePing() {
  const onlineEl = $('online-count');
  const count = onlineEl ? (parseInt(onlineEl.textContent, 10) || 0) : 0;
  const cbName = '__presencePingCallback_' + Date.now();
  const script  = document.createElement('script');
  const timer   = setTimeout(() => { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }, 8000);
  window[cbName] = function(resp) {
    clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script);
    const serverCount = (resp && typeof resp.online === 'number') ? resp.online : 0;
    const onlineEl = $('online-count');
    if (onlineEl) onlineEl.textContent = serverCount;
    if (chartPresence) {
      const now = new Date(), pad = n => String(n).padStart(2, '0');
      const hhmm = pad(now.getHours()) + ':' + pad(now.getMinutes());
      const labels = chartPresence.data.labels, vals = chartPresence.data.datasets[0].data, times = chartPresence._times || [];
      if (lastPresenceAppendAt > 0 && (now.getTime() - lastPresenceAppendAt) > 20000) {
        const zeroTime = new Date(lastPresenceAppendAt + 10000);
        times.push(pad(zeroTime.getHours()) + ':' + pad(zeroTime.getMinutes())); labels.push(labels.length); vals.push(0);
      }
      times.push(hhmm); labels.push(labels.length); vals.push(serverCount);
      lastPresenceAppendAt = now.getTime();
      if (labels.length > 500) { labels.shift(); vals.shift(); times.shift(); }
      chartPresence.update('none');
    }
  };
  script.src = DRIVE_SCRIPT_URL + '?action=recordPresence&count=' + encodeURIComponent(count) + '&callback=' + cbName + '&_cb=' + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

function pushSnapshot(total, available) {
  if (!DRIVE_SCRIPT_URL || DRIVE_SCRIPT_URL === 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') return;
  const cbName = '__snapshotCallback_' + Date.now();
  const script = document.createElement('script');
  const timer  = setTimeout(() => { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); }, 8000);
  window[cbName] = function() { clearTimeout(timer); delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); };
  script.src = DRIVE_SCRIPT_URL + '?action=pushSnapshot&total=' + encodeURIComponent(total) + '&available=' + encodeURIComponent(available) + '&callback=' + cbName + '&_cb=' + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}
