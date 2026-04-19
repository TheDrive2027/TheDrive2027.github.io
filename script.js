/* =============================================================
   THE DRIVE — script.js
   Fetches Sheet CSV + Drive JSON, merges them, renders the UI.
   No external dependencies except Google Fonts (CSS only).
   ============================================================= */

// ─── CONFIG ───────────────────────────────────────────────────
// Sheet published as CSV
const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRRk-WuFbb7q-_ZNbCjC6AaeV5yR6cGDuVCBJp0-wQI3zRQmdSaw87uzsUwI3dFgXTvsO_qBs6ach1C/pub?output=csv';
// ↓↓ PASTE YOUR APPS SCRIPT /exec URL HERE ↓↓
const DRIVE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwQ8GKmR5MalfoYgp9Uuz5eBSqe0A-s_GM61NQD1-uQyyAAK6rcZGbtl128N_Cog9eUag/exec';

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
let currentSort = 'imdb-desc';
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

const LOCAL_REQUEST_KEY  = 'thedrive_requests_v1';   // server counts cache (wiped on refresh)
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
function getRequestCount(title) {
  return requestCounts[normalize(title)] || 0;
}

function loadLocalCounts() {
  try { return JSON.parse(localStorage.getItem(LOCAL_REQUEST_KEY) || '{}'); } catch(e) { return {}; }
}
function saveLocalCounts() {
  try { localStorage.setItem(LOCAL_REQUEST_KEY, JSON.stringify(requestCounts)); } catch(e) {}
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
      + '&callback=' + cbName;
    script.src = url;
    script.onerror = () => { cleanup(); resolve(localCount); };
    document.head.appendChild(script);
  });
}

// requestCounts starts empty — server data fills it in applyDriveData()
requestCounts = {};

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

/** Find best Drive match for a movie title */
function findDriveMatch(title, driveMap) {
  const key = normalize(title);
  // 1. Exact match
  if (driveMap[key]) return driveMap[key];
  // 2. Prefix match — sheet title is a prefix of the drive key (e.g. "f1" vs "f1themovie")
  //    or drive key is a prefix of the sheet title
  for (const [driveKey, val] of Object.entries(driveMap)) {
    if (driveKey.startsWith(key) || key.startsWith(driveKey)) return val;
  }
  // 3. Fuzzy: allow up to 2 edits for close typos
  let best = null, bestDist = Infinity;
  for (const [driveKey, val] of Object.entries(driveMap)) {
    if (Math.abs(driveKey.length - key.length) > key.length * 0.5) continue;
    const dist = levenshtein(key, driveKey);
    if (dist <= 2 && dist < bestDist) {
      bestDist = dist;
      best = val;
    }
  }
  return best;
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
function showToast(msg, duration = 3000) {
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), duration);
}

/** Format and display the last-updated timestamp */
function updateLastUpdated() {
  const now = new Date();
  let h = now.getHours(), m = now.getMinutes();
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

const CACHE_KEY   = 'thedrive_cache_v2';
const CACHE_MAX_AGE = 5 * 60 * 1000; // 5 minutes

function saveCache(data) {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), data }));
  } catch(e) {}
}

function loadCache() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return null;
    const { ts, data } = JSON.parse(raw);
    if (Date.now() - ts > CACHE_MAX_AGE) return null;
    return data;
  } catch(e) { return null; }
}

function applyDriveData(rawData, csvRows) {
  const rawMovies = rawData.movies || rawData;
  posterMap = rawData.posters || {};
  // Always replace counts with the authoritative server totals on load.
  // The user's personal requested-set (userRequested) is stored separately
  // and never wiped, so their "✓ REQUESTED" state survives refreshes.
  if (rawData.requests) {
    requestCounts = { ...rawData.requests };
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

async function loadData(sheetURL, scriptURL, forceRefresh = false) {
  setProgress(10);
  let csvRows = [];

  // ── 1. Fetch CSV ──
  try {
    const r = await fetchURL(sheetURL, forceRefresh);
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const text = await r.text();
    csvRows = parseCSV(text);
    setProgress(30);
  } catch (e) {
    showToast('⚠ Could not load Sheet CSV. Check the URL & sharing settings.');
    console.error('CSV fetch error:', e);
  }

  // ── 2. Render immediately from cache if available (skipped on force refresh) ──
  const cached = forceRefresh ? null : loadCache();
  if (cached) {
    applyDriveData(cached, csvRows);
    setProgress(100);
    setTimeout(() => scanBar.classList.add('hidden'), 300);
  } else {
    // No cache — show movies without availability while Drive loads
    allMovies = mergeData(csvRows, {}, {});
    render();
    populateResFilter();
    updateCounts();
  }

  // ── 3. Fetch Drive JSON in background (with one silent retry) ──
  const driveURL = scriptURL && scriptURL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE' ? scriptURL : null;
  if (driveURL) {
    let driveData = null;
    let lastError = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        if (attempt > 0) await new Promise(r => setTimeout(r, 3000)); // wait 3s before retry
        driveData = await fetchScriptJSON(driveURL, forceRefresh);
        if (driveData && driveData.error) throw new Error(driveData.error);
        break; // success
      } catch (e) {
        lastError = e;
        console.warn('Drive fetch attempt ' + (attempt + 1) + ' failed:', e);
      }
    }
    if (driveData) {
      saveCache(driveData);
      applyDriveData(driveData, csvRows);
      setProgress(100);
      setTimeout(() => scanBar.classList.add('hidden'), 300);
    } else {
      // Only show the toast if both attempts failed
      if (!cached) showToast('⚠ Could not load Drive data. Check the Script URL & deployment.');
      console.error('Drive JSON error after retry:', lastError);
      setProgress(100);
      setTimeout(() => scanBar.classList.add('hidden'), 300);
    }
  } else {
    setProgress(100);
    setTimeout(() => scanBar.classList.add('hidden'), 300);
  }
}

function mergeData(rows, driveMap, posterMap = {}) {
  // Normalize column names: handles "Title", "title", "Movie Title", etc.
  const mapped = rows.map(row => {
    const title       = row.title || row.movie_title || row['movie title'] || '';
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
  const [key, dir] = currentSort.split('-');
  filtered.sort((a, b) => {
    let va, vb;
    if (key === 'title') { va = a.title.toLowerCase(); vb = b.title.toLowerCase(); }
    else if (key === 'imdb') { va = parseFloat(a.imdbRating) || 0; vb = parseFloat(b.imdbRating) || 0; }
    else if (key === 'year') { va = parseInt(a.year) || 0; vb = parseInt(b.year) || 0; }
    else if (key === 'size') { va = parseSizeGB(a.fileSize); vb = parseSizeGB(b.fileSize); }
    else if (key === 'requests') { va = getRequestCount(a.title); vb = getRequestCount(b.title); }

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
      <td class="td-size">${escHtml(m.fileSize) || '—'}</td>
      <td class="td-imdb"><span class="${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span></td>
      <td>
        <span class="status-pill ${m.available ? 'status-available' : 'status-missing'}">
          ${m.available ? 'AVAILABLE' : 'NOT UPLOADED'}
        </span>
      </td>
      <td class="td-link">
        ${m.driveLink
          ? `<a class="drive-link" href="${m.driveLink}" target="_blank" rel="noopener">▶ WATCH</a>`
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
  movieGrid.innerHTML = '';

  if (filtered.length === 0) {
    gridEmpty.hidden = false;
    return;
  }
  gridEmpty.hidden = true;

  const frag = document.createDocumentFragment();
  filtered.forEach((m, i) => {
    const card = document.createElement('div');
    card.className = 'movie-card';
    card.style.animationDelay = Math.min(i * 30, 400) + 'ms';
    const cardReqCount = getRequestCount(m.title);
    const cardIRequested = hasUserRequested(m.title);
    card.innerHTML = `
      ${m.poster ? `<div class="card-poster"><img src="${m.poster}" alt="${escHtml(m.title)}" loading="lazy" onload="this.classList.add('loaded')" /></div>` : ''}
      <div class="card-title">${escHtml(m.title)}</div>
      <div class="card-meta">
        <span class="card-year">${escHtml(m.year)}</span>
        <span class="card-sep">·</span>
        <span class="card-rating ${ratingClass(m.maturityRating)}">${escHtml(m.maturityRating) || '—'}</span>
      </div>
      <div class="card-row">
        <span class="card-imdb ${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span>
        <span class="card-res ${resClass(m.resolution)}">${escHtml(m.resolution) || '—'}</span>
      </div>
      <div class="card-size">${escHtml(m.fileSize) || '—'}</div>
      <div class="card-footer">
        <span class="status-pill ${m.available ? 'status-available' : 'status-missing'}">
          ${m.available ? 'AVAILABLE' : 'NOT UPLOADED'}
        </span>
        ${m.driveLink
          ? `<a class="drive-link" href="${m.driveLink}" target="_blank" rel="noopener">▶</a>`
          : cardIRequested
            ? `<button class="request-btn request-btn--done" data-title="${escHtml(m.title)}"><span class="request-icon">✓</span> REQUESTED${cardReqCount ? ' <span class="request-count">' + cardReqCount + '</span>' : ''}</button>`
            : `<button class="request-btn" data-title="${escHtml(m.title)}"><span class="request-icon">＋</span> REQUEST${cardReqCount ? ' <span class="request-count">' + cardReqCount + '</span>' : ''}</button>`}
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

// ─── EVENTS ───────────────────────────────────────────────────

// Search
let searchTimer;
searchInput.addEventListener('input', () => {
  clearSearch.classList.toggle('visible', searchInput.value.length > 0);
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { render(); saveSettings(); }, 200);
});

clearSearch.addEventListener('click', () => {
  searchInput.value = '';
  clearSearch.classList.remove('visible');
  render();
  saveSettings();
  searchInput.focus();
});

// Filters / sort
filterRes.addEventListener('change',  () => { render(); saveSettings(); });
filterMat.addEventListener('change',  () => { render(); saveSettings(); });
filterStat.addEventListener('change', () => { render(); saveSettings(); });
sortBy.addEventListener('change', () => {
  currentSort = sortBy.value;
  applySort();
  saveSettings();
});

// Column header sort
document.querySelectorAll('th.sortable').forEach(th => {
  th.addEventListener('click', () => {
    const s = th.dataset.sort;
    // Toggle direction if same key
    const [key] = s.split('-');
    const [curKey, curDir] = currentSort.split('-');
    if (curKey === key) {
      currentSort = key + '-' + (curDir === 'asc' ? 'desc' : 'asc');
    } else {
      currentSort = s;
    }
    sortBy.value = currentSort;
    document.querySelectorAll('th.sortable').forEach(t => t.classList.remove('sort-active'));
    th.classList.add('sort-active');
    applySort();
    saveSettings();
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
  });
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
  document.querySelectorAll('.request-btn[data-title="' + CSS.escape(title) + '"]').forEach(b => {
    b.disabled = false;
    b.classList.add('request-btn--done');
    const countHtml = count ? ' <span class="request-count">' + count + '</span>' : '';
    b.innerHTML = '<span class="request-icon">✓</span> REQUESTED' + countHtml;
    b.dataset.title = title;
  });
}

// ─── REFRESH BUTTON ───────────────────────────────────────────
if (refreshBtn) {
  refreshBtn.addEventListener('click', async () => {
    if (refreshBtn.classList.contains('spinning')) return;

    // Spin the icon
    refreshBtn.classList.add('spinning');
    scanBar.classList.remove('hidden');
    setProgress(0);

    // Clear the drive cache and server counts — but KEEP userRequested
    try { localStorage.removeItem(CACHE_KEY); } catch(e) {}
    try { localStorage.removeItem(LOCAL_REQUEST_KEY); } catch(e) {}
    requestCounts = {};

    // Re-fetch everything fresh — bypass cache and bust browser/CDN caches
    await loadData(SHEET_CSV_URL, DRIVE_SCRIPT_URL, true);

    refreshBtn.classList.remove('spinning');
  });
}

// ─── INIT ─────────────────────────────────────────────────────

(function init() {
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
    sortBy.value = 'imdb-desc';
    currentSort = 'imdb-desc';
  }

  loadData(SHEET_CSV_URL, DRIVE_SCRIPT_URL);
})();
