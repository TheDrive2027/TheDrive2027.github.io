/* =============================================================
   THE DRIVE — script.js
   Fetches Sheet CSV + Drive JSON, merges them, renders the UI.
   No external dependencies except Google Fonts (CSS only).
   4/20/2026 6:08 PM
   ============================================================= */

// ─── CONFIG ───────────────────────────────────────────────────
// Sheet published as CSV
const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRRk-WuFbb7q-_ZNbCjC6AaeV5yR6cGDuVCBJp0-wQI3zRQmdSaw87uzsUwI3dFgXTvsO_qBs6ach1C/pub?output=csv';
// ↓↓ PASTE YOUR APPS SCRIPT /exec URL HERE ↓↓
const DRIVE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzeCbThT62-e0fz7G8dcyufpYmoGWFq2awOyk8T3fShM9c__zqsCL82CrQ8njwmYM7V9Q/exec';


// ─── ACCESS KEY GATE ──────────────────────────────────────────
// Keys are validated against the "Keys" sheet via the Apps Script.
// Once validated, the key is saved to localStorage permanently.
// "Uses" = number of unique devices/browsers that have authenticated
// with this key. Re-visits from the same device never increment the
// count — the gate is skipped entirely if the key is already saved locally.

const LOCAL_KEY_STORE    = 'thedrive_access_key_v1';
const LOCAL_DEVICE_ID    = 'thedrive_device_id_v1';

function getSavedKey() {
  try { return localStorage.getItem(LOCAL_KEY_STORE) || null; } catch(e) { return null; }
}
function saveKey(key) {
  try { localStorage.setItem(LOCAL_KEY_STORE, key); } catch(e) {}
}

/**
 * Returns the persistent device ID for this browser.
 * Generated once using crypto.randomUUID (or a fallback) and stored forever.
 * Format: did-<8 hex chars>-<4 hex chars>-<4 hex chars>-<12 hex chars>
 */
function getDeviceId() {
  try {
    let did = localStorage.getItem(LOCAL_DEVICE_ID);
    if (!did) {
      if (crypto && crypto.randomUUID) {
        did = 'did-' + crypto.randomUUID().replace(/-/g, '').slice(0, 12).toUpperCase();
      } else {
        // Fallback for older browsers
        did = 'did-' + Array.from(
          { length: 12 },
          () => Math.floor(Math.random() * 16).toString(16)
        ).join('').toUpperCase();
      }
      localStorage.setItem(LOCAL_DEVICE_ID, did);
    }
    return did;
  } catch(e) {
    return 'did-UNKNOWN';
  }
}

/**
 * Call the Apps Script via JSONP to validate or consume a key.
 * action = 'validateKey' | 'useKey'
 */
function callKeyAction(action, keyStr, existing) {
  return new Promise((resolve) => {
    const cbName = '__keyCallback_' + Date.now();
    const script = document.createElement('script');
    const timer  = setTimeout(() => {
      cleanup();
      resolve({ error: 'timeout' });
    }, 12000);

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

let gateResolveFn = null; // resolved when the gate is passed

/**
 * Switches the gate modal into "Access Denied" mode:
 *  - Title changes to ACCESS DENIED
 *  - A subtitle "Your device has been blocked" appears beneath
 *  - Input field group is hidden
 *  - Button becomes a non-interactive SVG lock icon
 */
function showDenied() {
  const overlay   = document.getElementById('gate-overlay');
  const titleEl   = overlay && overlay.querySelector('.gate-title');
  const fieldEl   = overlay && overlay.querySelector('.modal-field');
  const submitBtn = document.getElementById('gate-submit');

  if (overlay) overlay.classList.remove('gate-overlay-hidden');

  if (titleEl) {
    titleEl.textContent = 'ACCESS DENIED';

    // Insert subtitle beneath the title if not already there
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
    // SVG lock icon (Lucide-style outline)
    submitBtn.innerHTML = `<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
      <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
      <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
    </svg>`;
  }
}


function showGate() {
  return new Promise(resolve => {
    gateResolveFn = resolve;
    const overlay   = document.getElementById('gate-overlay');
    const input     = document.getElementById('gate-key-input');
    const submitBtn = document.getElementById('gate-submit');
    const errorEl   = document.getElementById('gate-error');

    if (!overlay) { resolve(); return; } // no gate in HTML, skip

    overlay.classList.remove('gate-overlay-hidden');

    function showError(msg) {
      errorEl.textContent = msg;
      errorEl.hidden = false;
      input.style.borderColor = 'var(--red)';
      submitBtn.classList.remove('loading');
      submitBtn.textContent = 'ENTER THE DRIVE';
    }

    function clearError() {
      errorEl.hidden = true;
      input.style.borderColor = '';
    }

    input.addEventListener('input', () => {
      // Force uppercase as the user types, preserving cursor position
      const start = input.selectionStart;
      const end   = input.selectionEnd;
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

      // First validate (read-only check)
      const validation = await callKeyAction('validateKey', keyStr);

      if (validation.error) {
        showError('Could not reach the server. Check your connection and try again.');
        return;
      }

      if (!validation.valid) {
        if (validation.reason === 'device_blocked') {
          showDenied();
        } else if (validation.reason === 'expired') {
          showError('This key has reached its device limit. Please request a new key.');
        } else {
          showError('Invalid key. Please check and try again.');
        }
        return;
      }

      // Key is valid — consume it (counts this as one new device)
      const consume = await callKeyAction('useKey', keyStr);
      if (!consume.success && consume.reason === 'device_blocked') {
        showDenied();
        return;
      }
      if (!consume.success && consume.reason === 'expired') {
        // Race condition — another device used the last slot between validate and consume
        showError('This key just hit its device limit. Please request a new key.');
        return;
      }

      // Success — save locally and dismiss gate
      saveKey(keyStr);
      overlay.classList.add('gate-overlay-hidden');
      setTimeout(() => { overlay.style.display = 'none'; }, 350);
      resolve();
    }

    submitBtn.addEventListener('click', attempt);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') attempt(); });

    // Auto-focus the input
    setTimeout(() => input.focus(), 100);
  });
}

async function initWithGate() {
  const overlay = document.getElementById('gate-overlay');

  // ── Pre-populate UI immediately if a key is already saved ──
  // This runs before any network calls so the user sees their cached key
  // and a "Checking permissions…" loading state instead of a blank gate.
  const earlySavedKey = getSavedKey();
  if (earlySavedKey && overlay) {
    const earlyInput = document.getElementById('gate-key-input');
    const earlyBtn   = document.getElementById('gate-submit');
    if (earlyInput) { earlyInput.value = earlySavedKey; earlyInput.disabled = true; }
    if (earlyBtn)   { earlyBtn.textContent = 'CHECKING PERMISSIONS…'; earlyBtn.classList.add('loading'); }
  }

  // ── Step 1: silently check if this device is blocked ──
  // Do this before showing anything — blocked devices skip the gate entirely
  // and go straight to the denied screen.
  if (DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
    const deviceCheck = await new Promise(resolve => {
      const cbName = '__deviceCheckCallback_' + Date.now();
      const script  = document.createElement('script');
      const timer   = setTimeout(() => {
        cleanup();
        resolve({ allowed: true }); // fail open on timeout
      }, 8000);
      function cleanup() {
        clearTimeout(timer);
        delete window[cbName];
        if (script.parentNode) script.parentNode.removeChild(script);
      }
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

    if (deviceCheck.allowed === false) {
      showDenied();
      return; // stop here — never show the gate or load data
    }

    // Key column was cleared in the Devices sheet — wipe locally and
    // force them to re-enter a key, but don't block them entirely.
    if (deviceCheck.keyCleared) {
      try { localStorage.removeItem(LOCAL_KEY_STORE); } catch(e) {}
    }
  }

  // ── Step 2: normal gate / saved-key flow ──
  const savedKey = getSavedKey();

  if (savedKey) {
    // Key is saved locally — revalidate against the server in case it was deleted
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
      // Valid (or server unreachable — give benefit of the doubt) — proceed silently
      if (overlay) {
        overlay.classList.add('gate-overlay-hidden');
        overlay.style.display = 'none';
      }
    } else {
      // Key was deleted or expired — clear it and make them re-enter
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
    // No saved key — show gate and wait for it to be passed
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

// Simulate some being "available" in demo mode
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

// ─── STATE ───────────────────────────────────────────────────
let allMovies = [];     // merged final array
let filtered  = [];     // after search/filter
let currentView = 'grid';
let currentSort = 'imdb';
let currentDir  = 'desc';
let isDemoMode  = false;
let posterMap   = {};   // normalized title → poster URL

// ─── REQUEST COUNTS ──────────────────────────────────────────
// TWO separate stores:
//   requestCounts  — server totals (refreshed from server on every load, shown to all users)
//   userRequested  — Set of normalized titles THIS user has clicked (persisted forever in localStorage)
//
// This way:
//   • A refresh always pulls fresh totals from the server.
//   • Whether *you* requested something is remembered locally, independent of the count.
//   • Movies you haven't requested still show their total request count.

let requestCounts = {}; // { normalizedTitle: number } — from server

const LOCAL_USER_REQ_KEY = 'thedrive_user_reqs_v1';  // which titles THIS user requested (never wiped)

// ── User-requested set ──
function loadUserRequested() {
  try { return new Set(JSON.parse(localStorage.getItem(LOCAL_USER_REQ_KEY) || '[]')); } catch(e) { return new Set(); }
}
function saveUserRequested() {
  try { localStorage.setItem(LOCAL_USER_REQ_KEY, JSON.stringify([...userRequested])); } catch(e) {}
}
let userRequested = loadUserRequested(); // persists across refreshes

function hasUserRequested(title) {
  return userRequested.has(normalize(title));
}

// ── Server counts ──
function getRatingScore(title) {
  const r = ratingCounts[normalize(title)];
  if (!r) return 0;
  return (r.up || 0) - (r.down || 0);
}

function getRequestCount(title) {
  return requestCounts[normalize(title)] || 0;
}

async function postRequest(title) {
  const key = normalize(title);

  // Mark this user as having requested this title — persisted locally forever
  userRequested.add(key);
  saveUserRequested();

  // Optimistically bump the count in memory so the UI updates instantly
  requestCounts[key] = (requestCounts[key] || 0) + 1;
  const localCount = requestCounts[key];

  if (!DRIVE_SCRIPT_URL || DRIVE_SCRIPT_URL === 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
    return localCount;
  }

  // Use JSONP-style GET instead of POST to avoid Apps Script CORS issues.
  return new Promise(resolve => {
    const cbName = '__requestCallback_' + Date.now();
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      cleanup();
      console.warn('Request timed out, using local count');
      resolve(localCount);
    }, 10000);

    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }

    window[cbName] = function(data) {
      cleanup();
      if (data && data.count !== undefined) {
        requestCounts[key] = data.count;
        resolve(data.count);
      } else {
        resolve(localCount);
      }
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

// requestCounts starts empty — server data fills it in applyDriveData()
requestCounts = {};

// ─── RATINGS (thumbs up / down for available movies) ─────────────
// userRatings — persisted locally: { normalizedTitle: 'up' | 'down' }
// ratingCounts — from server: { normalizedTitle: { up: N, down: N } }

const LOCAL_RATINGS_KEY = 'thedrive_ratings_v1';
let ratingCounts = {}; // filled by applyDriveData

function loadUserRatings() {
  try { return JSON.parse(localStorage.getItem(LOCAL_RATINGS_KEY) || '{}'); } catch(e) { return {}; }
}
function saveUserRatings() {
  try { localStorage.setItem(LOCAL_RATINGS_KEY, JSON.stringify(userRatings)); } catch(e) {}
}
let userRatings = loadUserRatings();

function getUserRating(title) {
  return userRatings[normalize(title)] || null;
}
function getRatingCount(title, type) {
  return (ratingCounts[normalize(title)] || {})[type] || 0;
}

async function postRating(title, type) {
  const key = normalize(title);
  const prev = userRatings[key];

  // Toggle off if clicking the same type again
  if (prev === type) {
    delete userRatings[key];
  } else {
    userRatings[key] = type;
  }
  saveUserRatings();

  if (!DRIVE_SCRIPT_URL || DRIVE_SCRIPT_URL === 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') return;

  // Call Apps Script and update counts from the server's authoritative response
  return new Promise(resolve => {
    const cbName = '__ratingCallback_' + Date.now();
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
      resolve();
    }, 10000);
    window[cbName] = function(data) {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
      // Always overwrite local counts with what the server says
      if (data && typeof data.up === 'number' && typeof data.down === 'number') {
        ratingCounts[key] = { up: data.up, down: data.down };
      }
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

// ─── SETTINGS PERSISTENCE ────────────────────────────────────
const LOCAL_SETTINGS_KEY = 'thedrive_settings_v1';

function saveSettings() {
  try {
    localStorage.setItem(LOCAL_SETTINGS_KEY, JSON.stringify({
      search:     searchInput.value,
      resolution: filterRes.value,
      maturity:   filterMat.value,
      status:     filterStat.value,
      sort:       sortBy.value,
        dir:        currentDir,
      view:       currentView,
    }));
  } catch(e) {}
}

function loadSettings() {
  try { return JSON.parse(localStorage.getItem(LOCAL_SETTINGS_KEY) || 'null'); } catch(e) { return null; }
}

function applySettings(s) {
  if (!s) return;
  if (s.search)     { searchInput.value = s.search; clearSearch.classList.toggle('visible', s.search.length > 0); }
  if (s.resolution) filterRes.value  = s.resolution;
  if (s.maturity)   filterMat.value  = s.maturity;
  if (s.status)     filterStat.value = s.status;
  if (s.sort)       { sortBy.value = s.sort; currentSort = s.sort; }
  if (s.dir)        { currentDir = s.dir; sortDirBtn.textContent = currentDir === 'desc' ? '↓' : '↑'; }
  if (s.view && s.view !== currentView) {
    currentView = s.view;
    document.querySelectorAll('.toggle-btn').forEach(b => b.classList.toggle('active', b.dataset.view === s.view));
    tableView.classList.toggle('active', s.view === 'table');
    gridView.classList.toggle('active',  s.view === 'grid');
  }
}

// ─── DOM REFS ─────────────────────────────────────────────────
const $  = id => document.getElementById(id);
const tableBody   = $('table-body');
const movieGrid   = $('movie-grid');
const searchInput = $('search-input');
const clearSearch = $('clear-search');
const filterRes   = $('filter-resolution');
const filterMat   = $('filter-maturity');
const filterStat  = $('filter-status');
const sortBy      = $('sort-by');
const sortDirBtn  = $('sort-dir-btn');
const movieCount  = $('movie-count');
const availCount  = $('available-count');
const resultsSummary = $('results-summary');
const scanBar     = $('scan-bar');
const lastUpdatedEl = $('last-updated');
const refreshBtn  = $('refresh-btn');
const scanFill    = $('scan-fill');
const configModal = $('config-modal');
const tableView   = $('table-view');
const gridView    = $('grid-view');
const tableEmpty  = $('table-empty');
const gridEmpty   = $('grid-empty');
const toast       = $('toast');

// ─── UTILITIES ────────────────────────────────────────────────

/** Normalize for matching: lowercase, strip non-alphanumeric */
function normalize(str) {
  return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

/** Simple Levenshtein for fuzzy matching */
function levenshtein(a, b) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  if (!m) return n; if (!n) return m;
  const dp = Array.from({length: m + 1}, (_, i) =>
    Array.from({length: n + 1}, (_, j) => i === 0 ? j : j === 0 ? i : 0)
  );
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i-1] === b[j-1]
        ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[m][n];
}

/** Strip (year) and [tags] from a filename, then normalize */
function normalizeFilename(str) {
  return normalize(
    String(str || '')
      .replace(/\(\d{4}\)/g, '')   // remove (2014)
      .replace(/\[.*?\]/g, '')      // remove [1080p], [BluRay], etc.
  );
}

/** Find best Drive match for a movie title */
function findDriveMatch(title, driveMap) {
  const key = normalizeFilename(title);
  // Build a stripped version of the drive map on the fly and do an exact match
  for (const [driveKey, val] of Object.entries(driveMap)) {
    if (normalizeFilename(driveKey) === key) return val;
  }
  return null;
}

/** Parse CSV text → array of objects */
function parseCSV(text) {
  const lines = text.trim().split('\n');
  if (lines.length < 2) return [];
  const headers = splitCSVLine(lines[0]).map(h => h.trim().toLowerCase().replace(/\s+/g, '_'));
  return lines.slice(1).map(line => {
    const vals = splitCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] || '').trim(); });
    return obj;
  }).filter(r => r[headers[0]]); // skip empty rows
}

/** Handle quoted CSV fields */
function splitCSVLine(line) {
  const result = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      if (inQ && line[i+1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (c === ',' && !inQ) {
      result.push(cur); cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur);
  return result;
}

/** Extract year from a date string */
function extractYear(dateStr) {
  if (!dateStr) return '—';
  const m = String(dateStr).match(/\d{4}/);
  return m ? m[0] : '—';
}

/** Parse file size to GB float for sorting */
function parseSizeGB(sizeStr) {
  if (!sizeStr) return 0;
  const n = parseFloat(sizeStr);
  const s = sizeStr.toUpperCase();
  if (s.includes('TB')) return n * 1024;
  if (s.includes('GB')) return n;
  if (s.includes('MB')) return n / 1024;
  return n;
}

/** Parse runtime string to total minutes for sorting (e.g. "2h 18m", "138 min", "138") */
function parseRuntimeMinutes(str) {
  if (!str) return 0;
  const hm = str.match(/(\d+)\s*h(?:r|ours?)?\s*(\d+)?\s*m?/i);
  if (hm) return parseInt(hm[1]) * 60 + (parseInt(hm[2]) || 0);
  const m = str.match(/(\d+)/);
  return m ? parseInt(m[1]) : 0;
}

/** Maturity rating sort order */
const MATURITY_ORDER = { 'G': 1, 'PG': 2, 'PG-13': 3, 'PG13': 3, 'R': 4, 'NC-17': 5, 'NR': 6 };
function parseResolutionScore(res) {
  if (!res) return 0;
  const s = String(res).toUpperCase().trim();
  // 4K / UHD / 2160 → 2160
  if (s === '4K' || s === 'UHD' || s.includes('2160')) return 2160;
  // Extract raw number like "1080p", "720P", "576p", etc.
  const m = s.match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

/** IMDb rating CSS class */
function imdbClass(rating) {
  const r = parseFloat(rating);
  if (r >= 8) return 'imdb-high';
  if (r >= 6.5) return 'imdb-mid';
  return 'imdb-low';
}

/** Resolution CSS class */
function resClass(res) {
  const r = String(res).toUpperCase();
  if (r.includes('4K') || r.includes('2160')) return 'res-4k';
  if (r.includes('1080')) return 'res-1080';
  if (r.includes('720') || r.includes('576')) return 'res-720';
  return 'res-other';
}

/** Maturity rating CSS class */
function ratingClass(rating) {
  const r = String(rating || '').toUpperCase().replace(/[\s-]/g, '');
  if (r === 'G')    return 'rating-g';
  if (r === 'PG')   return 'rating-pg';
  if (r === 'PG13') return 'rating-pg13';
  if (r === 'R')    return 'rating-r';
  return '';
}

/** Show toast notification */
let toastTimer;

/** Fire-and-forget log of a client-side event to the server */
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

/** Format and display the last-updated timestamp.
 *  @param {Date} [date] — when the cache was written; defaults to now (fresh scan). */
function updateLastUpdated(date) {
  const d = (date instanceof Date && !isNaN(date)) ? date : new Date();
  let h = d.getHours(), m = d.getMinutes();
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  const mm = String(m).padStart(2, '0');
  if (lastUpdatedEl) lastUpdatedEl.textContent = h + ':' + mm + ' ' + ampm;
}

/** Set progress bar */
function setProgress(pct) {
  scanFill.style.width = pct + '%';
}

// ─── FETCH & MERGE ────────────────────────────────────────────

/**
 * Fetch a URL that may redirect (like Google Apps Script /exec endpoints).
 * Apps Script redirects to a Google-CDN URL that has proper CORS headers,
 * so we must follow redirects without locking the mode to 'cors'.
 */
async function fetchURL(url, bustCache = false) {
  if (bustCache) {
    url += (url.includes('?') ? '&' : '?') + '_cb=' + Date.now();
  }
  return fetch(url, { redirect: 'follow', cache: bustCache ? 'no-store' : 'default' });
}

/**
 * Fetch the Apps Script JSON via a JSONP-style callback to sidestep
 * any remaining CORS preflight issues. Falls back to direct fetch.
 */
function fetchScriptJSON(url, bustCache = false) {
  return new Promise((resolve, reject) => {
    // Primary: plain fetch with redirect following (works in most cases)
    fetchURL(url, bustCache)
      .then(r => {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      })
      .then(resolve)
      .catch(() => {
        // Fallback: JSONP via <script> tag — bypasses CORS entirely
        const cbName = '__driveCallback_' + Date.now();
        const script = document.createElement('script');
        const timer = setTimeout(() => {
          cleanup();
          reject(new Error('JSONP timeout'));
        }, 15000);

        function cleanup() {
          clearTimeout(timer);
          delete window[cbName];
          if (script.parentNode) script.parentNode.removeChild(script);
        }

        window[cbName] = function(data) {
          cleanup();
          resolve(data);
        };

        // Apps Script supports ?callback= for JSONP
        const jsonpUrl = bustCache ? url + (url.includes('?') ? '&' : '?') + '_cb=' + Date.now() : url;
        script.src = jsonpUrl + (jsonpUrl.includes('?') ? '&' : '?') + 'callback=' + cbName;
        script.onerror = () => { cleanup(); reject(new Error('JSONP script error')); };
        document.head.appendChild(script);
      });
  });
}

// ── Server-side cache only ──
// Drive data is cached exclusively on the Apps Script server (5-min TTL via
// CacheService). localStorage is NOT used for Drive data — every load asks
// the server, which either returns its cached payload instantly or triggers
// a fresh Drive scan if the cache is cold (> 5 min old or first load).

function saveCache(_data) {
  // No-op: Drive cache lives on the server, not in localStorage.
  // The writeCache POST in loadDataBulkFallback handles server-side persistence.
}

function loadCache() {
  // Always returns null — we never use a local Drive cache.
  return null;
}

function applyDriveData(rawData, csvRows) {
  const rawMovies = rawData.movies || rawData;
  posterMap = rawData.posters || {};
  // Always replace counts with the authoritative server totals on load.
  // The user's personal requested-set (userRequested) is stored separately
  // and never wiped, so their "✓ REQUESTED" state survives refreshes.
  if (rawData.requests) {
    // Normalize every key from the server so they match getRequestCount's normalize() lookup.
    // Server may store keys as "Inception", "inception", "the dark knight", etc — all get
    // collapsed to the same lowercase alphanumeric form used everywhere else.
    requestCounts = {};
    for (const [k, v] of Object.entries(rawData.requests)) {
      requestCounts[normalize(k)] = v;
    }
  }
  if (rawData.ratings) {
    ratingCounts = {};
    try { localStorage.removeItem('thedrive_rating_counts_v1'); } catch(e) {}
    for (const [k, v] of Object.entries(rawData.ratings)) {
      ratingCounts[normalize(k)] = v; // v = { up: N, down: N }
    }
    // Store separately from the Drive cache so ratings are always fresh on load
    try { localStorage.setItem('thedrive_rating_counts_v1', JSON.stringify(ratingCounts)); } catch(e) {}
  } else {
    // No ratings in this response — load from our separate store as fallback
    try {
      const stored = JSON.parse(localStorage.getItem('thedrive_rating_counts_v1') || '{}');
      ratingCounts = stored;
    } catch(e) {}
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
  populateResFilter();
  updateCounts();
  updateLastUpdated();
}

/** Update only the availability, link, and poster on already-rendered grid cards */
function patchGridCards() {
  applyFilters();
}


/** Fetch fresh ratings from the server and re-render any visible rating widgets */
function fetchRatings(scriptURL, isRefresh = false) {
  return new Promise(resolve => {
    const cbName = '__ratingsCallback_' + Date.now();
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
      resolve();
    }, 10000);
    window[cbName] = function(data) {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
      if (data && data.ratings) {
        ratingCounts = {};
        for (const [k, v] of Object.entries(data.ratings)) {
          ratingCounts[normalize(k)] = v;
        }
        // Patch all visible rating count elements without a full re-render
        document.querySelectorAll('.rating-btn').forEach(b => {
          const title = b.dataset.ratingTitle;
          const type  = b.dataset.ratingType;
          if (!title || !type) return;
          const countEl = b.querySelector('.rating-count');
          if (countEl) countEl.textContent = getRatingCount(title, type) || 0;
        });
        // Re-sort so rating sort order reflects fresh data
        applySort();
      }
      resolve();
    };
    script.src = scriptURL + '?action=getRatings&key=' + encodeURIComponent(getSavedKey() || '') + '&did=' + encodeURIComponent(getDeviceId()) + '&refresh=' + (isRefresh ? '1' : '0') + '&callback=' + cbName + '&_cb=' + Date.now();
    script.onerror = () => { if (script.parentNode) script.parentNode.removeChild(script); resolve(); };
    document.head.appendChild(script);
  });
}

/**
 * Fetch fresh request counts from the server and patch all visible
 * request buttons and tooltips without a full re-render.
 */
// ── JSONP helper for a single Apps Script action call ──
function jsonpAction(url) {
  return new Promise((resolve, reject) => {
    const cbName = '__cb_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    const timer  = setTimeout(() => { cleanup(); reject(new Error('timeout')); }, 15000);
    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }
    window[cbName] = data => { cleanup(); resolve(data); };
    script.src     = url + (url.includes('?') ? '&' : '?') + 'callback=' + cbName + '&_cb=' + Date.now();
    script.onerror = () => { cleanup(); reject(new Error('script error')); };
    document.head.appendChild(script);
  });
}

// Number of movie keys fetched per sequential request
const SEQUENTIAL_BATCH_SIZE = 10;

async function loadData(sheetURL, scriptURL, forceRefresh = false) {
  setProgress(5);
  let csvRows = [];

  // ── 1. Fetch CSV ──
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

  // Show CSV titles immediately so the page isn't blank while we wait for Drive
  allMovies = mergeData(csvRows, {}, {});
  render();
  populateResFilter();
  updateCounts();
  setProgress(20);

  const driveURL = scriptURL && scriptURL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE' ? scriptURL : null;
  if (!driveURL) {
    setProgress(100);
    setTimeout(() => scanBar.classList.add('hidden'), 300);
    return;
  }

  // ── 2. Try the Apps Script server cache (5-min TTL via CacheService) ──
  //
  //  • forceRefresh=false (normal page load):
  //      Call the plain doGet with no extra params. If the server cache is warm
  //      (someone scanned Drive in the last 5 minutes), it returns instantly.
  //      If the cache is cold, the plain doGet attempts a folder walk which can
  //      time out for large libraries — in that case we catch the error and fall
  //      through to the batched scan below.
  //
  //  • forceRefresh=true (refresh button):
  //      Add bust=1 to force the server to skip its cache and do a fresh scan,
  //      but the plain doGet folder walk can time out for large libraries.
  //      We therefore go straight to the batched bulk-fallback for refreshes.

  if (forceRefresh) {
    // Force a full batched rescan — plain doGet isn't reliable for large libraries
    await loadDataBulkFallback(driveURL, csvRows, true, false);
    return;
  }

  // Normal load — ask the server for its shared scan cache.
  // The getScanCache action checks both CacheService (L1, fast) and
  // the ScanCache Google Sheet (L2, durable, cross-device). If the
  // data is < 5 minutes old it returns it instantly; otherwise {stale:true}.
  setProgress(25);
  let serverPayload = null;
  try {
    const cacheResult = await jsonpAction(
      driveURL + '?action=getScanCache&_cb=' + Date.now()
    );
    if (cacheResult && cacheResult.ok && cacheResult.payload && cacheResult.payload.movies) {
      serverPayload = cacheResult.payload;
      console.log('Scan cache hit (' + (cacheResult.source || '?') + ', ' + Math.round(cacheResult.age_s) + 's old)');
    } else {
      console.log('Scan cache stale or missing — running full Drive scan');
    }
  } catch (e) {
    console.warn('getScanCache failed:', e);
  }

  if (serverPayload) {
    // ── Cache hit — apply data immediately, done ──
    applyDriveData(serverPayload, csvRows);
    setProgress(100);
    setTimeout(() => scanBar.classList.add('hidden'), 300);
    fetchRatings(driveURL, false);
    // Show when the cache was written, not the current time
    const cacheDate = (typeof cacheResult.age_s === 'number')
      ? new Date(Date.now() - cacheResult.age_s * 1000)
      : new Date();
    updateLastUpdated(cacheDate);
    return;
  }

  // ── Cache miss — run the full batched Drive scan ──
  // loadDataBulkFallback scans Drive in parallel batches, then POSTs
  // the result to writeCache which saves it to both CacheService and
  // the ScanCache sheet so every other device benefits immediately.
  await loadDataBulkFallback(driveURL, csvRows, false, false);
}

// Scans Drive file-by-file using batched scanFiles calls.
// background=true means the scan runs silently without touching the scan bar
// (used when stale cache was already shown and we\'re refreshing in the background).
async function loadDataBulkFallback(driveURL, csvRows, forceRefresh, background = false) {
  // ── Step 1: get the flat list of every file in the Drive tree ──
  let files = [];
  try {
    const listData = await jsonpAction(driveURL + '?action=getFileList');
    if (listData && listData.ok && Array.isArray(listData.files)) {
      files = listData.files; // [{ id, name, isPosters }, ...]
    } else {
      throw new Error(listData && listData.error ? listData.error : 'getFileList failed');
    }
  } catch (e) {
    if (!background) {
      showToast('⚠ Could not load Drive file list. Check the Script URL & deployment.');
      setProgress(100);
      setTimeout(() => scanBar.classList.add('hidden'), 300);
    }
    console.error('getFileList error:', e);
    fetchRatings(driveURL, forceRefresh);
    return;
  }

  if (files.length === 0) {
    if (!background) {
      setProgress(100);
      setTimeout(() => scanBar.classList.add('hidden'), 300);
    }
    fetchRatings(driveURL, forceRefresh);
    return;
  }

  // ── Step 2: scan files in batches — each request processes SCAN_BATCH_SIZE
  //    files server-side in a single Apps Script execution, dramatically
  //    cutting the number of round trips vs. one file per call.
  //    Multiple batches run in parallel (CONCURRENCY) for further speed.
  const accumMovies   = {};
  const accumPosters  = {};
  const accumRequests = {};
  const accumRatings  = {};

  const SCAN_BATCH_SIZE = 10;
  const CONCURRENCY     = 6;
  const total           = files.length;
  const progressStart   = 25;
  const progressEnd     = 95;

  const batches = [];
  for (let i = 0; i < total; i += SCAN_BATCH_SIZE) {
    batches.push(files.slice(i, i + SCAN_BATCH_SIZE));
  }

  let completedFiles = 0;

  for (let i = 0; i < batches.length; i += CONCURRENCY) {
    const concurrentBatches = batches.slice(i, i + CONCURRENCY);

    const results = await Promise.all(concurrentBatches.map(batch => {
      const fileIds   = batch.map(f => f.id).join(',');
      const isPosters = batch.map(f => f.isPosters ? '1' : '0').join(',');
      return jsonpAction(
        driveURL
        + '?action=scanFiles'
        + '&fileIds='   + encodeURIComponent(fileIds)
        + '&isPosters=' + encodeURIComponent(isPosters)
        + '&key='       + encodeURIComponent(getSavedKey() || '')
        + '&did='       + encodeURIComponent(getDeviceId())
      ).catch(err => {
        console.warn('scanFiles batch failed:', err);
        return null;
      });
    }));

    for (const result of results) {
      if (result && result.ok) {
        Object.assign(accumMovies,  result.movies  || {});
        Object.assign(accumPosters, result.posters || {});
      }
    }

    completedFiles += concurrentBatches.reduce((s, b) => s + b.length, 0);
    if (!background) {
      setProgress(progressStart + ((completedFiles / total) * (progressEnd - progressStart)));
    }

    applyDriveData({
      movies:   accumMovies,
      posters:  accumPosters,
      requests: accumRequests,
      ratings:  accumRatings,
    }, csvRows);
  }

  // ── Step 3: fetch live request + rating counts from the sheet ──
  // scanFiles only collects Drive file metadata; requests/ratings live in
  // the Google Sheet and must be fetched separately after the scan completes.
  let liveRequests = {};
  let liveRatings  = {};
  try {
    const ratingsData = await new Promise((resolve) => {
      const cbName = '__postScanRatings_' + Date.now();
      const script = document.createElement('script');
      const timer  = setTimeout(() => {
        delete window[cbName];
        if (script.parentNode) script.parentNode.removeChild(script);
        resolve({});
      }, 10000);
      window[cbName] = function(data) {
        clearTimeout(timer);
        delete window[cbName];
        if (script.parentNode) script.parentNode.removeChild(script);
        resolve(data || {});
      };
      script.src = driveURL
        + '?action=getRatings'
        + '&key='      + encodeURIComponent(getSavedKey() || '')
        + '&did='      + encodeURIComponent(getDeviceId())
        + '&callback=' + cbName
        + '&_cb='      + Date.now();
      script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); resolve({}); };
      document.head.appendChild(script);
    });
    if (ratingsData.ratings) {
      for (const [k, v] of Object.entries(ratingsData.ratings)) {
        liveRatings[normalize(k)] = v;
      }
      ratingCounts = { ...liveRatings };
    }
  } catch(e) {}

  // Pull request counts from the live sheet via getScanCache
  // (at this point we just POSTed writeCache so L1/L2 are fresh)
  try {
    const reqResult = await jsonpAction(driveURL + '?action=getScanCache&_cb=' + Date.now());
    const reqPayload = (reqResult && reqResult.ok) ? reqResult.payload : null;
    if (reqPayload && reqPayload.requests) {
      for (const [k, v] of Object.entries(reqPayload.requests)) {
        liveRequests[normalize(k)] = v;
      }
      requestCounts = { ...liveRequests };
    }
  } catch(e) {}

  // ── Step 4: assemble final payload with live counts ──
  const finalPayload = {
    movies:   accumMovies,
    posters:  accumPosters,
    requests: liveRequests,
    ratings:  liveRatings,
  };

  // Apply the complete data (movies + fresh counts) to the UI
  applyDriveData(finalPayload, csvRows);

  // ── Push Library Growth snapshot — only on a full Drive scan ──
  // The server's upsertSnapshot replaces today's entry if it already exists,
  // so this is safe to call every scan and will never create duplicate rows.
  {
    // Use allMovies (populated by applyDriveData above) so counts match the stats panel:
    // total = full CSV library size, available = those with a Drive file attached.
    const total     = allMovies.length;
    const available = allMovies.filter(m => m.available).length;
    pushSnapshot(total, available);
  }

  // ── Step 5: write the assembled payload to the Apps Script cache (5-min TTL)
  //    so the next person to load the site gets it instantly.
  try {
    // Apps Script web apps return a 302 redirect before executing doPost.
    // Using redirect:'follow' causes the browser to convert the POST to a GET
    // on the redirect, dropping the body — so doPost never runs and the cache
    // is never written. mode:'no-cors' sends the POST directly without
    // following redirects, which is the correct pattern for Apps Script POSTs.
    fetch(driveURL, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain' }, // no-cors requires simple headers
      body:    JSON.stringify({
        action:  'writeCache',
        payload: finalPayload,
        key:     getSavedKey() || '',
        did:     getDeviceId(),
      }),
      mode:     'no-cors',
    }).catch(() => {}); // fire-and-forget — failure is non-critical
  } catch(e) {}

  if (!background) {
    setProgress(100);
    setTimeout(() => scanBar.classList.add('hidden'), 300);
  }
  updateLastUpdated();
}

function mergeData(rows, driveMap, posterMap = {}) {
  // Normalize column names: handles "Title", "title", "Movie Title", etc.
  const mapped = rows.map(row => {
    const title       = row.title || row.movie_title || row['movie title'] || '';
    const runtime     = row.runtime || row.run_time || row.duration || '';
    const resolution  = row.resolution || row.res || '';
    const maturityRating = row.maturity_rating || row.rating || row.maturityrating || '';
    const releaseDate = row.release_date || row.releasedate || row.date || '';
    const fileSize    = row.file_size || row.filesize || row.size || '';
    const imdbRating  = row.imdb_rating || row.imdbrating || row.imdb || '';

    const match = findDriveMatch(title, driveMap);
    const posterKey = normalize(title);
    const poster = posterMap[posterKey] || null;

    return {
      title,
      runtime,
      resolution,
      maturityRating,
      releaseDate,
      year: extractYear(releaseDate),
      fileSize,
      imdbRating,
      available: !!match,
      driveLink: match ? match.link : null,
      driveResolution: match ? (match.name || '') : '',
      poster,
    };
  }).filter(m => m.title);

  // Deduplicate rows with the same normalized title.
  // If a Drive file exists, keep the row whose resolution matches the drive file.
  // Otherwise keep the first row encountered.
  const seen = new Map();
  for (const m of mapped) {
    const key = normalize(m.title);
    if (!seen.has(key)) {
      seen.set(key, m);
    } else {
      const existing = seen.get(key);
      // Prefer whichever row's resolution matches the actual drive file
      if (m.available && m.driveResolution) {
        const driveRes = m.driveResolution.toUpperCase();
        const mRes = m.resolution.toUpperCase();
        const existingRes = existing.resolution.toUpperCase();
        const mMatches = driveRes.includes(mRes) || mRes.split('P')[0] === driveRes.split('P')[0];
        const existingMatches = driveRes.includes(existingRes) || existingRes.split('P')[0] === driveRes.split('P')[0];
        if (mMatches && !existingMatches) {
          seen.set(key, m);
        }
      }
      // If neither or both match, keep existing (first row wins)
    }
  }

  return [...seen.values()];
}

function populateResFilter() {
  const resolutions = [...new Set(allMovies.map(m => m.resolution).filter(Boolean))].sort();
  filterRes.innerHTML = '<option value="">All</option>';
  resolutions.forEach(r => {
    const opt = document.createElement('option');
    opt.value = r; opt.textContent = r;
    filterRes.appendChild(opt);
  });

  const ratings = [...new Set(allMovies.map(m => m.maturityRating).filter(Boolean))].sort();
  filterMat.innerHTML = '<option value="">All</option>';
  ratings.forEach(r => {
    const opt = document.createElement('option');
    opt.value = r; opt.textContent = r;
    filterMat.appendChild(opt);
  });
}

function updateCounts() {
  movieCount.textContent   = allMovies.length + ' films';
  availCount.textContent   = allMovies.filter(m => m.available).length + ' available';
}

// ─── SORT & FILTER ────────────────────────────────────────────

function applyFilters() {
  const q   = normalize(searchInput.value);
  const res = filterRes.value;
  const mat = filterMat.value;
  const st  = filterStat.value;

  filtered = allMovies.filter(m => {
    if (q && !normalize(m.title).includes(q)) return false;
    if (res && m.resolution !== res) return false;
    if (mat && m.maturityRating !== mat) return false;
    if (st === 'Available' && !m.available) return false;
    if (st === 'Not Uploaded' && m.available) return false;
    return true;
  });

  applySort();
}

function applySort() {
  const key = currentSort;
  const dir = currentDir;
  filtered.sort((a, b) => {
    let va, vb;
    if (key === 'title') { va = a.title.toLowerCase(); vb = b.title.toLowerCase(); }
    else if (key === 'imdb') { va = parseFloat(a.imdbRating) || 0; vb = parseFloat(b.imdbRating) || 0; }
    else if (key === 'year') { va = parseInt(a.year) || 0; vb = parseInt(b.year) || 0; }
    else if (key === 'size') { va = parseSizeGB(a.fileSize); vb = parseSizeGB(b.fileSize); }
    else if (key === 'requests') { va = getRequestCount(a.title); vb = getRequestCount(b.title); }
    else if (key === 'rating') { va = getRatingScore(a.title); vb = getRatingScore(b.title); }
    else if (key === 'runtime') { va = parseRuntimeMinutes(a.runtime); vb = parseRuntimeMinutes(b.runtime); }
    else if (key === 'maturity') { va = MATURITY_ORDER[a.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99; vb = MATURITY_ORDER[b.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99; }
    else if (key === 'status') { va = a.available ? 0 : 1; vb = b.available ? 0 : 1; }
    else if (key === 'res') {
      va = parseResolutionScore(a.resolution);
      vb = parseResolutionScore(b.resolution);
    }
    else if (key === 'link') {
      // Tier 1: available (has link) sorted by upvotes desc
      // Tier 2: requested (no link) sorted by request count desc
      // Tier 3: unrequested (no link, no requests)
      const availA = a.available ? 1 : 0, availB = b.available ? 1 : 0;
      if (availA !== availB) return availB - availA; // available first
      if (availA === 1) {
        // Both available: sort by rating score desc
        return getRatingScore(b.title) - getRatingScore(a.title);
      }
      // Both unavailable: requested before unrequested, then by request count desc
      const rqA = getRequestCount(a.title), rqB = getRequestCount(b.title);
      const hasReqA = rqA > 0 ? 1 : 0, hasReqB = rqB > 0 ? 1 : 0;
      if (hasReqA !== hasReqB) return hasReqB - hasReqA; // requested first
      return rqB - rqA; // more requests first
    }

    if (va < vb) return dir === 'asc' ? -1 : 1;
    if (va > vb) return dir === 'asc' ? 1 : -1;
    return 0;
  });

  renderCurrentView();
}

// ─── RENDER ───────────────────────────────────────────────────

function render() {
  applyFilters();
}

function renderCurrentView() {
  if (currentView === 'table') renderTable();
  else renderGrid();

  // Summary
  const total = allMovies.length;
  const shown = filtered.length;
  if (shown === total) {
    resultsSummary.textContent = `Showing all ${total} films`;
  } else {
    resultsSummary.textContent = `Showing ${shown} of ${total} films`;
  }
}


/** Build the rating widget HTML for a movie */
function ratingHTML(title) {
  const userVote = getUserRating(title);
  const ups   = getRatingCount(title, 'up');
  const downs = getRatingCount(title, 'down');
  return `<div class="rating-wrap">
    <button class="rating-btn rating-btn--up ${userVote === 'up' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="up" title="Liked it">
      ▲<span class="rating-count">${ups || 0}</span>
    </button>
    <button class="rating-btn rating-btn--down ${userVote === 'down' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="down" title="Didn't like it">
      ▼<span class="rating-count">${downs || 0}</span>
    </button>
  </div>`;
}

function renderTable() {
  tableBody.innerHTML = '';

  if (filtered.length === 0) {
    tableEmpty.hidden = false;
    return;
  }
  tableEmpty.hidden = true;

  const frag = document.createDocumentFragment();
  filtered.forEach((m, i) => {
    const tr = document.createElement('tr');
    tr.style.animationDelay = Math.min(i * 18, 300) + 'ms';
    const reqCount = getRequestCount(m.title);
    const iRequested = hasUserRequested(m.title);
    tr.innerHTML = `
      <td class="td-num">${i + 1}</td>
      <td class="td-title">${escHtml(m.title)}</td>
      <td class="td-res"><span class="${resClass(m.resolution)}">${escHtml(m.resolution) || '—'}</span></td>
      <td class="td-rating"><span class="${ratingClass(m.maturityRating)}">${escHtml(m.maturityRating) || '—'}</span></td>
      <td class="td-year">${escHtml(m.year)}</td>
      <td class="td-runtime">${escHtml(m.runtime) || '—'}</td>
      <td class="td-size">${escHtml(m.fileSize) || '—'}</td>
      <td class="td-imdb"><span class="${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span></td>
      <td>
        <span class="status-pill ${m.available ? 'status-available' : 'status-missing'}">
          ${m.available ? 'AVAILABLE' : 'NOT UPLOADED'}
        </span>
      </td>
      <td class="td-link">
        ${m.driveLink
          ? `<div class="td-link-inner"><a class="drive-link" href="${m.driveLink}" target="_blank" rel="noopener" data-title="${escHtml(m.title)}">▶ WATCH</a>${ratingHTML(m.title)}</div>`
          : iRequested
            ? `<button class="request-btn request-btn--done" data-title="${escHtml(m.title)}"><span class="request-icon">✓</span> REQUESTED${reqCount ? ' <span class="request-count">' + reqCount + '</span>' : ''}</button>`
            : `<button class="request-btn" data-title="${escHtml(m.title)}"><span class="request-icon">＋</span> REQUEST${reqCount ? ' <span class="request-count">' + reqCount + '</span>' : ''}</button>`}
      </td>
    `;
    frag.appendChild(tr);
  });
  tableBody.appendChild(frag);
}

function renderGrid() {
  // Track which cards are already rendered so we only animate new arrivals.
  const existingKeys = new Set(
    Array.from(movieGrid.querySelectorAll('.movie-card[data-key]'))
      .map(el => el.dataset.key)
  );
  movieGrid.innerHTML = '';

  if (filtered.length === 0) {
    gridEmpty.hidden = false;
    return;
  }
  gridEmpty.hidden = true;

  const frag = document.createDocumentFragment();
  filtered.forEach((m, i) => {
    const card = document.createElement('div');
    const key  = normalize(m.title);
    const isNew = !existingKeys.has(key);
    card.className = 'movie-card';
    card.dataset.key = key;
    if (isNew) {
      card.style.animationDelay = Math.min(i * 30, 400) + 'ms';
    } else {
      card.style.animation = 'none';
    }
    const cardReqCount = getRequestCount(m.title);
    const cardIRequested = hasUserRequested(m.title);
    const posterClasses = ['card-poster'];
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
        ${m.fileSize ? `<span class="card-sep">·</span><span class="card-size">${escHtml(m.fileSize)}</span>` : ''}
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
    frag.appendChild(card);
  });
  movieGrid.appendChild(frag);
}

function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ─── REQUEST COUNT TOOLTIP ────────────────────────────────────
// Shows a floating tooltip with the server request count when
// hovering over any movie title cell (table) or card title (grid).

(function initRequestTooltip() {
  const tooltip = document.getElementById('req-tooltip');
  if (!tooltip) return;

  let hideTimer;

  function showTooltipForTitle(title, anchorEl) {
    clearTimeout(hideTimer);
    const count = getRequestCount(title);
    if (count <= 0) {
      // Still show "No requests yet" so users know the feature exists
      tooltip.innerHTML = '<span class="tt-count">0</span> requests';
    } else {
      tooltip.innerHTML = '<span class="tt-count">' + count + '</span> request' + (count === 1 ? '' : 's');
    }
    tooltip.removeAttribute('hidden');
    // Force reflow before adding visible so the CSS transition fires
    tooltip.getBoundingClientRect();
    tooltip.classList.add('visible');
    positionTooltip(anchorEl);
  }

  function positionTooltip(el) {
    const rect = el.getBoundingClientRect();
    const ttW  = tooltip.offsetWidth;
    const ttH  = tooltip.offsetHeight;
    let left   = rect.left + rect.width / 2 - ttW / 2;
    let top    = rect.top - ttH - 8;
    // Keep within viewport
    if (left < 8) left = 8;
    if (left + ttW > window.innerWidth - 8) left = window.innerWidth - ttW - 8;
    if (top < 8) top = rect.bottom + 8; // flip below if no room above
    tooltip.style.left = left + 'px';
    tooltip.style.top  = top  + 'px';
  }

  function hideTooltip() {
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => {
      tooltip.classList.remove('visible');
      tooltip.setAttribute('hidden', '');
    }, 80);
  }

  // Delegate from main content — works for both table and grid views
  document.getElementById('main-content').addEventListener('mouseover', e => {
    // Table: hovering the title cell td.td-title or its text
    const tdTitle = e.target.closest('td.td-title');
    if (tdTitle) {
      const titleText = tdTitle.textContent.trim();
      showTooltipForTitle(titleText, tdTitle);
      return;
    }
    // Grid: hovering the .card-title div
    const cardTitle = e.target.closest('.card-title');
    if (cardTitle) {
      const titleText = cardTitle.textContent.trim();
      showTooltipForTitle(titleText, cardTitle);
      return;
    }
  });

  document.getElementById('main-content').addEventListener('mouseout', e => {
    const leaving = e.target.closest('td.td-title, .card-title');
    if (leaving) hideTooltip();
  });

  // Keep tooltip positioned if the element scrolls
  document.addEventListener('scroll', () => { tooltip.classList.remove('visible'); }, { passive: true });
})();

// ─── EVENTS ───────────────────────────────────────────────────

// Search
let searchTimer;
let searchLogTimer;
searchInput.addEventListener('input', () => {
  const query = searchInput.value;
  clearSearch.classList.toggle('visible', query.length > 0);

  // Render immediately (debounced 200ms)
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { render(); saveSettings(); }, 200);

  // Log the search query after the user pauses for 1.5s (non-empty only)
  clearTimeout(searchLogTimer);
  if (query.trim().length > 0) {
    searchLogTimer = setTimeout(() => {
      logClientEvent('Search', query.trim());
    }, 1500);
  }
});

clearSearch.addEventListener('click', () => {
  searchInput.value = '';
  clearSearch.classList.remove('visible');
  render();
  saveSettings();
  searchInput.focus();
});

// Filters / sort
filterRes.addEventListener('change',  () => { render(); saveSettings(); logClientEvent('Filter Resolution', filterRes.value || 'All'); });
filterMat.addEventListener('change',  () => { render(); saveSettings(); logClientEvent('Filter Maturity', filterMat.value || 'All'); });
filterStat.addEventListener('change', () => { render(); saveSettings(); logClientEvent('Filter Status', filterStat.value || 'All'); });
sortBy.addEventListener('change', () => {
  currentSort = sortBy.value;
  applySort();
  saveSettings();
  logClientEvent('Sort', currentSort + '-' + currentDir);
});

sortDirBtn.addEventListener('click', () => {
  currentDir = currentDir === 'desc' ? 'asc' : 'desc';
  sortDirBtn.textContent = currentDir === 'desc' ? '↓' : '↑';
  sortDirBtn.title = currentDir === 'desc' ? 'Descending' : 'Ascending';
  applySort();
  saveSettings();
  logClientEvent('Sort Direction', currentDir);
});

// Column header sort
document.querySelectorAll('th.sortable').forEach(th => {
  th.addEventListener('click', () => {
    const [key, defaultDir] = th.dataset.sort.split('-');
    if (currentSort === key) {
      currentDir = currentDir === 'asc' ? 'desc' : 'asc';
    } else {
      currentSort = key;
      currentDir = defaultDir || 'desc';
    }
    sortDirBtn.textContent = currentDir === 'desc' ? '↓' : '↑';
    if (sortBy.querySelector(`option[value="${key}"]`)) sortBy.value = key;
    document.querySelectorAll('th.sortable').forEach(t => t.classList.remove('sort-active'));
    th.classList.add('sort-active');
    applySort();
    saveSettings();
    logClientEvent('Sort', currentSort + '-' + currentDir);
  });
});

// View toggle
document.querySelectorAll('.toggle-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const v = btn.dataset.view;
    if (v === currentView) return;
    currentView = v;
    document.querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    tableView.classList.toggle('active', v === 'table');
    gridView.classList.toggle('active', v === 'grid');
    renderCurrentView();
    saveSettings();
    logClientEvent('Switch View', v);
  });
});

// Watch link clicks — fire-and-forget log to server
$('main-content').addEventListener('click', e => {
  const link = e.target.closest('.drive-link');
  if (!link) return;
  const title = link.dataset.title || '';
  const key = getSavedKey() || '';
  if (!DRIVE_SCRIPT_URL || !key) return;
  const cbName = '__openLinkCallback_' + Date.now();
  const script = document.createElement('script');
  window[cbName] = function() { delete window[cbName]; if (script.parentNode) script.parentNode.removeChild(script); };
  script.src = DRIVE_SCRIPT_URL
    + '?action=openLink'
    + '&title=' + encodeURIComponent(title)
    + '&key='   + encodeURIComponent(key)
    + '&did='   + encodeURIComponent(getDeviceId())
    + '&callback=' + cbName;
  script.onerror = () => { if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
});

// Rating buttons (event delegation on main content)
$('main-content').addEventListener('click', async e => {
  const btn = e.target.closest('.rating-btn');
  if (!btn) return;

  const title = btn.dataset.ratingTitle;
  const type  = btn.dataset.ratingType;
  if (!title || !type) return;

  await postRating(title, type);

  // Re-render all matching rating widgets in the DOM
  document.querySelectorAll(`[data-rating-title="${CSS.escape(title)}"]`).forEach(b => {
    const userVote = getUserRating(title);
    const isUp = b.dataset.ratingType === 'up';
    b.classList.toggle('active', userVote === b.dataset.ratingType);
    const countEl = b.querySelector('.rating-count');
    if (countEl) countEl.textContent = getRatingCount(title, b.dataset.ratingType) || 0;
  });

  const userVote = getUserRating(title);
  if (userVote === 'up')   showToast('▲ You liked ' + title);
  else if (userVote === 'down') showToast('▼ You disliked ' + title);
  else showToast('Rating removed for ' + title);
});

// Request button (event delegation on main content)
$('main-content').addEventListener('click', async e => {
  const btn = e.target.closest('.request-btn');
  if (!btn || btn.classList.contains('request-btn--done')) return;

  btn.disabled = true;
  const title = btn.dataset.title;

  // Optimistically update all matching buttons right away
  const optimisticCount = (requestCounts[normalize(title)] || 0) + 1;
  setRequestedState(title, optimisticCount);

  // Fire POST and correct count if server differs
  const serverCount = await postRequest(title);
  if (serverCount !== optimisticCount) {
    setRequestedState(title, serverCount);
  }

  showToast('✓ Requested: ' + title);
});

function setRequestedState(title, count) {
  const countHtml = count ? ' <span class="request-count">' + count + '</span>' : '';

  // Table view: plain .request-btn elements (not inside an overlay)
  document.querySelectorAll('.request-btn[data-title="' + CSS.escape(title) + '"]:not(.card-request-overlay)').forEach(b => {
    b.disabled = false;
    b.classList.add('request-btn--done');
    b.innerHTML = '<span class="request-icon">&#10003;</span> REQUESTED' + countHtml;
    b.dataset.title = title;
  });

  // Grid view: .card-request-overlay buttons — swap to "done" appearance
  document.querySelectorAll('.card-request-overlay.request-btn[data-title="' + CSS.escape(title) + '"]').forEach(b => {
    b.disabled = true;
    b.classList.add('card-request-overlay--done');
    const inner = b.querySelector('.card-request-btn');
    if (inner) {
      inner.classList.add('card-request-btn--done');
      inner.innerHTML = '<span class="card-request-icon">&#10003;</span><span class="card-request-label">REQUESTED' + countHtml + '</span>';
    }
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

  function setStatus(msg, type) {
    statusEl.textContent = msg;
    statusEl.className = 'footer-form-status ' + type;
    statusEl.hidden = false;
  }

  submitBtn.addEventListener('click', async () => {
    const message = (msgInput.value || '').trim();
    if (!message) { setStatus('Please enter a message before sending.', 'error'); msgInput.focus(); return; }

    submitBtn.disabled = true;
    submitBtn.textContent = 'SENDING…';
    statusEl.hidden = true;

    const name = (nameInput.value || '').trim() || 'Anonymous';
    const key  = getSavedKey() || '';

    try {
      await new Promise((resolve, reject) => {
        const cbName = '__formCallback_' + Date.now();
        const script = document.createElement('script');
        const timer  = setTimeout(() => { cleanup(); reject(new Error('timeout')); }, 12000);

        function cleanup() {
          clearTimeout(timer);
          delete window[cbName];
          if (script.parentNode) script.parentNode.removeChild(script);
        }

        window[cbName] = function(data) {
          cleanup();
          if (data && data.ok) resolve();
          else reject(new Error(data && data.error ? data.error : 'unknown'));
        };

        script.src = DRIVE_SCRIPT_URL
          + '?action=submitForm'
          + '&name='    + encodeURIComponent(name)
          + '&message=' + encodeURIComponent(message)
          + '&key='     + encodeURIComponent(key)
          + '&did='     + encodeURIComponent(getDeviceId())
          + '&callback=' + cbName
          + '&_cb='     + Date.now();
        script.onerror = () => { cleanup(); reject(new Error('network')); };
        document.head.appendChild(script);
      });

      setStatus('✓ Message sent — thanks!', 'success');
      nameInput.value = '';
      msgInput.value  = '';
    } catch(err) {
      setStatus('Something went wrong. Please try again.', 'error');
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = 'SEND';
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

    // Wipe in-memory state and any stale localStorage left from older versions
    requestCounts = {};
    ratingCounts  = {};
    try { localStorage.removeItem('thedrive_cache_v3'); } catch(e) {}
    try { localStorage.removeItem('thedrive_requests_v1'); } catch(e) {}
    try { localStorage.removeItem('thedrive_rating_counts_v1'); } catch(e) {}

    // Bust the server-side cache first so the upcoming scan starts clean
    if (DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
      try {
        await jsonpAction(
          DRIVE_SCRIPT_URL + '?action=bustCache'
          + '&key=' + encodeURIComponent(getSavedKey() || '')
          + '&did=' + encodeURIComponent(getDeviceId())
        );
      } catch(e) {
        console.warn('bustCache failed (non-critical):', e);
      }
    }

    // loadData(forceRefresh=true) skips the server cache and goes straight to
    // the batched Drive scan, then writes the fresh result back to the server.
    await loadData(SHEET_CSV_URL, DRIVE_SCRIPT_URL, true);

    updateLastUpdated();
    refreshBtn.classList.remove('spinning');
  });
}

// ─── INIT ─────────────────────────────────────────────────────

(async function init() {
  const modal = $('config-modal');
  if (modal) modal.classList.add('hidden');

  // Set default view to grid
  tableView.classList.remove('active');
  gridView.classList.add('active');
  $('btn-table').classList.remove('active');
  $('btn-grid').classList.add('active');

  // Restore saved settings, or fall back to defaults
  const savedSettings = loadSettings();
  if (savedSettings) {
    applySettings(savedSettings);
  } else {
    sortBy.value = 'imdb';
    currentSort = 'imdb';
    currentDir  = 'desc';
    sortDirBtn.textContent = '↓';
  }

  // Show the access gate if no key is saved locally
  await initWithGate();

  // Only load data after the gate is passed
  loadData(SHEET_CSV_URL, DRIVE_SCRIPT_URL);

  // ── Online presence + stats pings ──
  if (DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
    fetchOnlineCount();
    setInterval(fetchOnlineCount, 60 * 1000);
    setInterval(pingHeartbeat,   4 * 60 * 1000);
    // Record presence every 10 seconds
    setInterval(pushPresencePing, 10 * 1000);
  }

  // ── Tab switching ──
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
  const timer  = setTimeout(() => {
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
  }, 10000);
  window[cbName] = function(data) {
    clearTimeout(timer);
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
    const el = document.getElementById('online-count');
    if (el && data && typeof data.online === 'number') {
      el.textContent = data.online;
    }
  };
  script.src = DRIVE_SCRIPT_URL
    + '?action=getOnlineCount'
    + '&callback=' + cbName
    + '&_cb=' + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

/** Re-ping checkDevice so our Last Seen stays current between page loads. */
function pingHeartbeat() {
  const cbName = '__heartbeatCallback_' + Date.now();
  const script = document.createElement('script');
  const timer  = setTimeout(() => {
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
  }, 10000);
  window[cbName] = function() {
    clearTimeout(timer);
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
  };
  script.src = DRIVE_SCRIPT_URL
    + '?action=checkDevice'
    + '&did='      + encodeURIComponent(getDeviceId())
    + '&key='      + encodeURIComponent(getSavedKey() || '')
    + '&callback=' + cbName
    + '&_cb='      + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

// ─── STATS TAB ────────────────────────────────────────────────

let statsLoaded    = false;
let statsLoadedAt  = 0;
let chartLibrary        = null;
let chartUsers          = null;
let chartPresence       = null;
let lastPresenceAppendAt = 0; // ms timestamp of the last live-appended point

function initStatsTab() {
  // Render local stats immediately from allMovies (no network needed)
  renderLocalStats();
  // Then fetch server stats (snapshots, presence, device count)
  if ((!statsLoaded || Date.now() - statsLoadedAt > 60000) && DRIVE_SCRIPT_URL && DRIVE_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') {
    fetchStatsData();
    statsLoaded   = true;
    statsLoadedAt = Date.now();
  }
}

function renderLocalStats() {
  if (!allMovies.length) return;

  const total     = allMovies.length;
  const available = allMovies.filter(m => m.available).length;

  // Upload progress bar
  const pct = total > 0 ? ((available / total) * 100).toFixed(2) : '0.00';
  const fracEl  = document.getElementById('upload-fraction');
  const pctEl   = document.getElementById('upload-pct');
  const fillEl  = document.getElementById('upload-fill');
  if (fracEl)  fracEl.textContent  = available + ' / ' + total + ' films uploaded';
  if (pctEl)   pctEl.textContent   = pct + '%';
  if (fillEl)  fillEl.style.width  = parseFloat(pct) + '%';

  // Stat cards from local data
  setText('stat-total-films', total);
  setText('stat-available', available);

  // Total size
  let totalGB = 0;
  allMovies.forEach(m => { totalGB += parseSizeGB(m.fileSize); });
  setText('stat-total-size', totalGB > 0 ? totalGB.toFixed(1) + ' GB' : '—');

  // Total runtime
  let totalMins = 0;
  allMovies.forEach(m => { totalMins += parseRuntimeMinutes(m.runtime); });
  if (totalMins > 0) {
    const h = Math.floor(totalMins / 60);
    const m = totalMins % 60;
    setText('stat-total-runtime', h + 'h ' + m + 'm');
  }

  // Avg IMDb
  const rated = allMovies.filter(m => parseFloat(m.imdbRating) > 0);
  if (rated.length) {
    const avg = rated.reduce((s, m) => s + parseFloat(m.imdbRating), 0) / rated.length;
    setText('stat-avg-imdb', '★ ' + avg.toFixed(1));
  }

  // Maturity counts
  const matNorm = r => String(r || '').toUpperCase().replace(/[\s-]/g, '');
  setText('stat-g',    allMovies.filter(m => matNorm(m.maturityRating) === 'G').length);
  setText('stat-pg',   allMovies.filter(m => matNorm(m.maturityRating) === 'PG').length);
  setText('stat-pg13', allMovies.filter(m => matNorm(m.maturityRating) === 'PG13').length);
  setText('stat-r',    allMovies.filter(m => matNorm(m.maturityRating) === 'R').length);

  // Resolution counts
  setText('stat-4k',   allMovies.filter(m => /4k|2160/i.test(m.resolution)).length);
  setText('stat-1080', allMovies.filter(m => /1080/i.test(m.resolution)).length);

  // Snapshot is pushed only after a full Drive scan completes (in loadDataBulkFallback),
  // not here, so the Library Growth chart only reflects actual scan results.
}

function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = String(val);
}

function fetchStatsData() {
  const cbName = '__statsCallback_' + Date.now();
  const script = document.createElement('script');
  const timer  = setTimeout(() => {
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
  }, 15000);
  window[cbName] = function(data) {
    clearTimeout(timer);
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
    if (!data) return;
    if (typeof data.uniqueDevices === 'number') setText('stat-total-users', data.uniqueDevices);
    if (data.snapshots    && data.snapshots.length)    renderLibraryChart(data.snapshots);
    if (data.userHistory  && data.userHistory.length)  renderUserChart(data.userHistory);
    if (data.presence     && data.presence.length)     renderPresenceChart(data.presence);
    else showPresencePlaceholder();
  };
  script.src = DRIVE_SCRIPT_URL
    + '?action=getStatsData'
    + '&callback=' + cbName
    + '&_cb=' + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

// ── Library growth chart ──
function renderLibraryChart(snapshots) {
  const canvas = document.getElementById('chart-library');
  if (!canvas) return;

  const labels    = snapshots.map(s => s.date);
  const totals    = snapshots.map(s => s.total);
  const available = snapshots.map(s => s.available);

  const cfg = {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: 'Total Films',
          data: totals,
          borderColor: '#9090a8',
          backgroundColor: 'rgba(144,144,168,0.08)',
          borderWidth: 2,
          pointRadius: 3,
          pointBackgroundColor: '#9090a8',
          tension: 0.3,
          fill: true,
        },
        {
          label: 'Available',
          data: available,
          borderColor: '#e8c547',
          backgroundColor: 'rgba(232,197,71,0.10)',
          borderWidth: 2,
          pointRadius: 3,
          pointBackgroundColor: '#e8c547',
          tension: 0.3,
          fill: true,
        }
      ]
    },
    options: chartOptions('Films')
  };

  if (chartLibrary) chartLibrary.destroy();
  chartLibrary = new Chart(canvas, cfg);
}

// ── Unique user growth chart ──
function renderUserChart(userHistory) {
  const canvas = document.getElementById('chart-users');
  if (!canvas) return;

  const labels = userHistory.map(u => u.date);
  const data   = userHistory.map(u => u.users);

  const cfg = {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Unique Users',
        data,
        borderColor: '#e8c547',
        backgroundColor: 'rgba(232,197,71,0.10)',
        borderWidth: 2,
        pointRadius: 3,
        pointBackgroundColor: '#e8c547',
        tension: 0.3,
        fill: true,
      }]
    },
    options: chartOptions('Users')
  };

  if (chartUsers) chartUsers.destroy();
  chartUsers = new Chart(canvas, cfg);
}

// ── Presence history chart ──
function showPresencePlaceholder() {
  const canvas = document.getElementById('chart-presence');
  if (!canvas) return;
  const wrap = canvas.closest('.chart-wrap');
  if (!wrap) return;
  // Hide canvas, show a styled placeholder message
  canvas.style.display = 'none';
  if (!wrap.querySelector('.presence-placeholder')) {
    const msg = document.createElement('div');
    msg.className = 'presence-placeholder';
    msg.innerHTML = `<span class="presence-placeholder-icon">◎</span><p>No history yet — data will appear here as users come online.</p>`;
    wrap.appendChild(msg);
  }
}

function showPresenceCanvas() {
  const canvas = document.getElementById('chart-presence');
  if (!canvas) return;
  canvas.style.display = '';
  const wrap = canvas.closest('.chart-wrap');
  if (wrap) {
    const ph = wrap.querySelector('.presence-placeholder');
    if (ph) ph.remove();
  }
}

function renderPresenceChart(presence) {
  showPresenceCanvas();
  const canvas = document.getElementById('chart-presence');
  if (!canvas) return;

  const INTERVAL_MS = 10 * 1000;
  const GAP_THRESH  = INTERVAL_MS * 2;

  function tsToMs(ts) {
    return new Date(ts.replace(' ', 'T')).getTime();
  }

  const filled = [];
  for (let i = 0; i < presence.length; i++) {
    filled.push(presence[i]);
    if (i < presence.length - 1) {
      const gap = tsToMs(presence[i + 1].ts) - tsToMs(presence[i].ts);
      if (gap > GAP_THRESH) {
        const afterTs = new Date(tsToMs(presence[i].ts) + INTERVAL_MS);
        const pad = n => String(n).padStart(2, '0');
        const fakeTs = presence[i].ts.slice(0, 11)
          + pad(afterTs.getHours()) + ':' + pad(afterTs.getMinutes()) + ':' + pad(afterTs.getSeconds());
        filled.push({ ts: fakeTs, online: 0 });
      }
    }
  }

  const step    = Math.max(1, Math.floor(filled.length / 500));
  const sampled = filled.filter((_, i) => i % step === 0);

  // Use numeric indices as labels so Chart.js never tries to parse them as dates.
  // The actual HH:MM strings live in _times and are surfaced via the tick callback.
  const times  = sampled.map(p => {
    const m = String(p.ts).match(/(\d{1,2}:\d{2})(?::\d{2})?/);
    return m ? m[1] : '';
  });
  const rawValues = sampled.map(p => p.online);
  // 5-point rolling average to smooth out single-sample spikes
  const values = rawValues.map((v, i, arr) => {
    const p2 = arr[i - 2] !== undefined ? arr[i - 2] : v;
    const p1 = arr[i - 1] !== undefined ? arr[i - 1] : v;
    const n1 = arr[i + 1] !== undefined ? arr[i + 1] : v;
    const n2 = arr[i + 2] !== undefined ? arr[i + 2] : v;
    return Math.round((p2 + p1 + v + n1 + n2) / 5 * 100) / 100;
  });
  const labels = sampled.map((_, i) => i);

  const cfg = {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Online',
        data: values,
        borderColor: '#3ecf74',
        backgroundColor: 'rgba(62,207,116,0.10)',
        borderWidth: 2,
        pointRadius: 0,
        tension: 0,
        fill: true,
      }]
    },
    options: presenceChartOptions(times)
  };

  if (chartPresence) chartPresence.destroy();
  chartPresence = new Chart(canvas, cfg);
  chartPresence._times = times;
  lastPresenceAppendAt = Date.now();
}

function chartOptions(yLabel) {
  return {
    responsive: true,
    maintainAspectRatio: true,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: {
        labels: {
          color: '#9090a8',
          font: { family: 'DM Mono', size: 11 },
          boxWidth: 12,
        }
      },
      tooltip: {
        backgroundColor: '#18181f',
        borderColor: '#252530',
        borderWidth: 1,
        titleColor: '#e8e8f0',
        bodyColor: '#9090a8',
        titleFont: { family: 'DM Mono', size: 11 },
        bodyFont:  { family: 'DM Mono', size: 11 },
      }
    },
    scales: {
      x: {
        ticks: { color: '#78788f', font: { family: 'DM Mono', size: 10 }, maxTicksLimit: 10 },
        grid:  { color: 'rgba(37,37,48,0.6)' },
      },
      y: {
        title: { display: true, text: yLabel, color: '#78788f', font: { family: 'DM Mono', size: 10 } },
        ticks: { color: '#78788f', font: { family: 'DM Mono', size: 10 }, precision: 0 },
        grid:  { color: 'rgba(37,37,48,0.6)' },
        beginAtZero: true,
      }
    }
  };
}

function presenceChartOptions(times) {
  const base = chartOptions('Users');
  base.scales.x.type = 'category';
  base.scales.x.ticks = {
    color: '#78788f',
    font: { family: 'DM Mono', size: 10 },
    maxTicksLimit: 10,
    callback: function(val) {
      return times[val] || '';
    }
  };
  // Override tooltip so the title shows HH:MM instead of the numeric index
  base.plugins.tooltip.callbacks = {
    title: function(tooltipItems) {
      const idx = tooltipItems[0].dataIndex;
      return times[idx] || '';
    }
  };
  return base;
}

// ── Push presence ping (called every 10s) ──
function pushPresencePing() {
  const onlineEl = document.getElementById('online-count');
  const count    = onlineEl ? (parseInt(onlineEl.textContent, 10) || 0) : 0;
  const cbName   = '__presencePingCallback_' + Date.now();
  const script   = document.createElement('script');
  const timer    = setTimeout(() => {
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
  }, 8000);
  window[cbName] = function(resp) {
    clearTimeout(timer);
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
    // Use the authoritative count from the server response
    const serverCount = (resp && typeof resp.online === 'number') ? resp.online : 0;
    // Update the header online count display
    const onlineEl = document.getElementById('online-count');
    if (onlineEl) onlineEl.textContent = serverCount;
    // Live-append to the presence chart if it's already rendered
    if (chartPresence) {
      const now    = new Date();
      const pad    = n => String(n).padStart(2, '0');
      const hhmm   = pad(now.getHours()) + ':' + pad(now.getMinutes());
      const labels = chartPresence.data.labels;
      const vals   = chartPresence.data.datasets[0].data;
      const times  = chartPresence._times || [];
      if (lastPresenceAppendAt > 0 && (now.getTime() - lastPresenceAppendAt) > 20000) {
        const zeroTime = new Date(lastPresenceAppendAt + 10000);
        times.push(pad(zeroTime.getHours()) + ':' + pad(zeroTime.getMinutes()));
        labels.push(labels.length);
        vals.push(0);
      }
      times.push(hhmm);
      labels.push(labels.length);
      vals.push(serverCount);
      lastPresenceAppendAt = now.getTime();
      if (labels.length > 500) { labels.shift(); vals.shift(); times.shift(); }
      chartPresence.update('none');
    }
  };
  script.src = DRIVE_SCRIPT_URL
    + '?action=recordPresence'
    + '&count='    + encodeURIComponent(count)
    + '&callback=' + cbName
    + '&_cb='      + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}

// ── Push today's snapshot counts ──
function pushSnapshot(total, available) {
  if (!DRIVE_SCRIPT_URL || DRIVE_SCRIPT_URL === 'YOUR_APPS_SCRIPT_EXEC_URL_HERE') return;
  const cbName = '__snapshotCallback_' + Date.now();
  const script = document.createElement('script');
  const timer  = setTimeout(() => {
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
  }, 8000);
  window[cbName] = function() {
    clearTimeout(timer);
    delete window[cbName];
    if (script.parentNode) script.parentNode.removeChild(script);
  };
  script.src = DRIVE_SCRIPT_URL
    + '?action=pushSnapshot'
    + '&total='     + encodeURIComponent(total)
    + '&available=' + encodeURIComponent(available)
    + '&callback='  + cbName
    + '&_cb='       + Date.now();
  script.onerror = () => { clearTimeout(timer); if (script.parentNode) script.parentNode.removeChild(script); };
  document.head.appendChild(script);
}
