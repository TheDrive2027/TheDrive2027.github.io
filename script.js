/* =============================================================
   THE DRIVE — script.js
   Fetches Sheet CSV + Drive JSON, merges them, renders the UI.
   No external dependencies except Google Fonts (CSS only).
   ============================================================= */

// ─── CONFIG ───────────────────────────────────────────────────
// Sheet published as CSV
const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRRk-WuFbb7q-_ZNbCjC6AaeV5yR6cGDuVCBJp0-wQI3zRQmdSaw87uzsUwI3dFgXTvsO_qBs6ach1C/pub?output=csv';
// ↓↓ PASTE YOUR APPS SCRIPT /exec URL HERE ↓↓
const DRIVE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbypnRQ-tawdumrXDO02ht_1zYeV9piH48IDaftdV6DLpssK9NmuXnWodyEQZG8gOaOG/exec';

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
let currentView = 'table';
let currentSort = 'title-asc';
let isDemoMode  = false;
let posterMap   = {};   // normalized title → poster URL

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
async function fetchURL(url) {
  return fetch(url, { redirect: 'follow' });
}

/**
 * Fetch the Apps Script JSON via a JSONP-style callback to sidestep
 * any remaining CORS preflight issues. Falls back to direct fetch.
 */
function fetchScriptJSON(url) {
  return new Promise((resolve, reject) => {
    // Primary: plain fetch with redirect following (works in most cases)
    fetchURL(url)
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
        script.src = url + (url.includes('?') ? '&' : '?') + 'callback=' + cbName;
        script.onerror = () => { cleanup(); reject(new Error('JSONP script error')); };
        document.head.appendChild(script);
      });
  });
}

const CACHE_KEY   = 'thedrive_cache_v1';
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
  const videoMimeTypes = ['video/', 'application/octet-stream'];
  const driveMap = Object.fromEntries(
    Object.entries(rawMovies).filter(([, val]) =>
      !val.mimeType || videoMimeTypes.some(t => val.mimeType.startsWith(t))
    )
  );
  allMovies = mergeData(csvRows, driveMap, posterMap);
  render();
  populateResFilter();
  updateCounts();
}

async function loadData(sheetURL, scriptURL) {
  setProgress(10);
  let csvRows = [];

  // ── 1. Fetch CSV ──
  try {
    const r = await fetchURL(sheetURL);
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const text = await r.text();
    csvRows = parseCSV(text);
    setProgress(30);
  } catch (e) {
    showToast('⚠ Could not load Sheet CSV. Check the URL & sharing settings.');
    console.error('CSV fetch error:', e);
  }

  // ── 2. Render immediately from cache if available ──
  const cached = loadCache();
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

  // ── 3. Fetch Drive JSON in background ──
  const driveURL = scriptURL && scriptURL !== 'YOUR_APPS_SCRIPT_EXEC_URL_HERE' ? scriptURL : null;
  if (driveURL) {
    try {
      const driveData = await fetchScriptJSON(driveURL);
      if (driveData && driveData.error) throw new Error(driveData.error);
      saveCache(driveData);
      applyDriveData(driveData, csvRows);
      setProgress(100);
      setTimeout(() => scanBar.classList.add('hidden'), 300);
    } catch (e) {
      showToast('⚠ Could not load Drive data. Check the Script URL & deployment.');
      console.error('Drive JSON error:', e);
      if (!cached) {
        setProgress(100);
        setTimeout(() => scanBar.classList.add('hidden'), 300);
      }
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
      <td>
        ${m.driveLink
          ? `<a class="drive-link" href="${m.driveLink}" target="_blank" rel="noopener">▶ WATCH</a>`
          : `<span class="no-link">—</span>`}
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
          : ''}
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
  searchTimer = setTimeout(() => render(), 200);
});

clearSearch.addEventListener('click', () => {
  searchInput.value = '';
  clearSearch.classList.remove('visible');
  render();
  searchInput.focus();
});

// Filters / sort
filterRes.addEventListener('change', render);
filterMat.addEventListener('change', render);
filterStat.addEventListener('change', render);
sortBy.addEventListener('change', () => {
  currentSort = sortBy.value;
  applySort();
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
  });
});

// ─── INIT ─────────────────────────────────────────────────────

(function init() {
  const modal = $('config-modal');
  if (modal) modal.classList.add('hidden');
  loadData(SHEET_CSV_URL, DRIVE_SCRIPT_URL);
})();
