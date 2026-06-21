/* =============================================================
   THE DRIVE — script.js
   ============================================================= */

// ─── CONFIG ───────────────────────────────────────────────────
let API_BASE = ''; // Populated by config.json in init()
const SHEET_CSV_URL    = '';
const SHOWS_CSV_URL    = '';
const DRIVE_SCRIPT_URL = '';

// Auto-reload the full tab every 30 minutes
const AUTO_RELOAD_MS = 30 * 60 * 1000;
setTimeout(() => location.reload(), AUTO_RELOAD_MS);

// ─── ACCESS KEY GATE ──────────────────────────────────────────
const LOCAL_KEY_STORE = 'thedrive_access_key_v1';
const LOCAL_DEVICE_ID = 'thedrive_device_id_v1';

function getSavedKey() { try { return localStorage.getItem(LOCAL_KEY_STORE) || null; } catch(e) { return null; } }
function saveKey(key) { try { localStorage.setItem(LOCAL_KEY_STORE, key); } catch(e) {} }

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

async function initWithGate() {
  const overlay = document.getElementById('gate-overlay');
  const input = document.getElementById('gate-key-input');
  const submitBtn = document.getElementById('gate-submit');
  const errorEl = document.getElementById('gate-error');
  const did = getDeviceId();
  const savedKey = getSavedKey();

  // Helper to trigger auto-retry when server is down
  function triggerServerRetry() {
    if (overlay) overlay.classList.remove('gate-overlay-hidden');
    if (submitBtn) { submitBtn.classList.add('loading'); submitBtn.textContent = 'RETRYING...'; }
    if (input) input.disabled = true;
    if (errorEl) { 
      errorEl.textContent = 'Cannot reach server. Retrying in 5 seconds...'; 
      errorEl.hidden = false; 
    }
    setTimeout(() => location.reload(), 5000);
  }

  // 1. Check if device is blocked globally
  try {
    const blockRes = await fetch(`${API_BASE}/api/keys/check-device?did=${did}`);
    const blockData = await blockRes.json();
    if (blockData.blocked) {
      if (overlay) overlay.classList.remove('gate-overlay-hidden');
      const titleEl = overlay.querySelector('.gate-title');
      if (titleEl) titleEl.textContent = 'ACCESS DENIED';
      return;
    }
  } catch(e) { 
    console.warn("Device check failed", e);
    triggerServerRetry();
    return;
  }

  // 2. If saved key exists, validate it silently
  if (savedKey) {
    try {
      const valRes = await fetch(`${API_BASE}/api/keys/validate?code=${savedKey}&did=${did}`);
      const valData = await valRes.json();
      if (valData.valid) {
        if (overlay) { overlay.classList.add('gate-overlay-hidden'); overlay.style.display = 'none'; }
        return; // Gate passed!
      } else {
        // Server reached, but key is invalid. Clear it.
        localStorage.removeItem(LOCAL_KEY_STORE);
      }
    } catch(e) { 
      console.warn("Key validation failed", e);
      triggerServerRetry();
      return;
    }
  }

  // 3. Show gate and wait for user input
  if (overlay) overlay.classList.remove('gate-overlay-hidden');
  if (submitBtn) submitBtn.classList.remove('loading');
  if (input) input.disabled = false;
  
  return new Promise(resolve => {
    if (!submitBtn || !input) { resolve(); return; }

    async function attempt() {
      const keyStr = input.value.trim().toUpperCase();
      if (!keyStr) return;
      
      submitBtn.classList.add('loading');
      submitBtn.textContent = 'CHECKING…';
      errorEl.hidden = true;

      try {
        const valRes = await fetch(`${API_BASE}/api/keys/validate?code=${keyStr}&did=${did}`);
        const valData = await valRes.json();
        
        if (!valData.valid) {
          errorEl.textContent = 'Invalid key.';
          errorEl.hidden = false;
          submitBtn.classList.remove('loading');
          submitBtn.textContent = 'ENTER THE DRIVE';
          return;
        }

        const useRes = await fetch(`${API_BASE}/api/keys/use?code=${keyStr}&did=${did}`);
        const useData = await useRes.json();
        
        if (useData.success) {
          saveKey(keyStr);
          overlay.classList.add('gate-overlay-hidden');
          setTimeout(() => { overlay.style.display = 'none'; }, 350);
          resolve();
        } else {
          errorEl.textContent = 'Failed to activate key.';
          errorEl.hidden = false;
          submitBtn.classList.remove('loading');
          submitBtn.textContent = 'ENTER THE DRIVE';
        }
      } catch(e) {
        // If manual entry fails due to network, save the key they typed and trigger retry
        saveKey(keyStr);
        triggerServerRetry();
        resolve();
      }
    }

    submitBtn.addEventListener('click', attempt, { once: true });
    input.addEventListener('keydown', e => { if (e.key === 'Enter') attempt(); }, { once: true });
    setTimeout(() => input.focus(), 100);
  });
}

// ─── STATE ────────────────────────────────────────────────────
let allMovies   = [];
let allShows    = []; 
let filtered    = [];
let currentSort = 'title';
let currentDir  = 'asc';
let activeTab   = 'movies'; 
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

// ─── RATINGS ─────────────────────────────────────────────────
const ratingInflight = new Set();
const LOCAL_RATINGS_KEY = 'thedrive_ratings_v1';
let ratingCounts = {};

function loadUserRatings() { try { return JSON.parse(localStorage.getItem(LOCAL_RATINGS_KEY) || '{}'); } catch(e) { return {}; } }
function saveUserRatings() { try { localStorage.setItem(LOCAL_RATINGS_KEY, JSON.stringify(userRatings)); } catch(e) {} }
let userRatings = loadUserRatings();

function getUserRating(title) { return userRatings[normalize(title)] || null; }
function getRatingCount(title, type) { return (ratingCounts[normalize(title)] || {})[type] || 0; }
function getRatingScore(title) {
  const r = ratingCounts[normalize(title)];
  if (!r) return 0;
  return (r.up || 0) - (r.down || 0);
}

function applyRatingDOM(title, nextVote, upCount, downCount, clickedBtn) {
  document.querySelectorAll(`[data-rating-title="${CSS.escape(title)}"]`).forEach(b => {
    const bType    = b.dataset.ratingType;
    const isActive = nextVote === bType;
    b.classList.toggle('active', isActive);

    if (b === clickedBtn && isActive) {
      b.classList.remove('just-voted');
      void b.offsetWidth;
      b.classList.add('just-voted');
    }

    const countEl = b.querySelector('.rating-count');
    if (countEl) countEl.textContent = bType === 'up' ? upCount : downCount;
  });
}

// ─── REQUEST COUNTS (Local) ───────────────────────────────────
let requestCounts = {};
const LOCAL_USER_REQ_KEY = 'thedrive_user_reqs_v1';
function loadUserRequested() { try { return new Set(JSON.parse(localStorage.getItem(LOCAL_USER_REQ_KEY) || '[]')); } catch(e) { return new Set(); } }
function saveUserRequested() { try { localStorage.setItem(LOCAL_USER_REQ_KEY, JSON.stringify([...userRequested])); } catch(e) {} }
let userRequested = loadUserRequested();
function hasUserRequested(title) { return userRequested.has(normalize(title)); }
function getRequestCount(title) { return requestCounts[normalize(title)] || 0; }

// ─── SETTINGS PERSISTENCE ─────────────────────────────────────
const LOCAL_SETTINGS_KEY = 'thedrive_settings_v2';
function saveSettings() { }
function loadSettings() { return null; }
function applySettings(s) { if (!s) return; }

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
function extractYear(dateStr) { if (!dateStr) return '—'; const m = String(dateStr).match(/\d{4}/); return m ? m[0] : '—'; }
function parseSizeGB(sizeStr) { if (!sizeStr) return 0; const n = parseFloat(sizeStr), s = sizeStr.toUpperCase(); if (s.includes('TB')) return n * 1024; if (s.includes('GB')) return n; if (s.includes('MB')) return n / 1024; return n; }
function parseRuntimeMinutes(str) { if (!str) return 0; const hm = str.match(/(\d+)\s*h(?:r|ours?)?\s*(\d+)?\s*m?/i); if (hm) return parseInt(hm[1]) * 60 + (parseInt(hm[2]) || 0); const m = str.match(/(\d+)/); return m ? parseInt(m[1]) : 0; }
function formatEpRuntime(str) { const mins = parseRuntimeMinutes(str); if (!mins) return ''; const h = Math.floor(mins / 60); const m = mins % 60; if (h > 0 && m > 0) return h + 'h ' + m + 'm'; if (h > 0) return h + 'h'; return m + 'm'; }

const MATURITY_ORDER = { 'G': 1, 'PG': 2, 'PG-13': 3, 'PG13': 3, 'R': 4, 'NC-17': 5, 'NR': 6 };
function parseResolutionScore(res) { if (!res) return 0; const s = String(res).toUpperCase().trim(); if (s === '4K' || s === 'UHD' || s.includes('2160')) return 2160; const m = s.match(/(\d+)/); return m ? parseInt(m[1], 10) : 0; }
function imdbClass(rating) { const r = parseFloat(rating); if (r >= 8) return 'imdb-high'; if (r >= 6.5) return 'imdb-mid'; return 'imdb-low'; }
function resClass(res) { const r = String(res).toUpperCase(); if (r.includes('4K') || r.includes('2160')) return 'res-4k'; if (r.includes('1080')) return 'res-1080'; if (r.includes('720') || r.includes('576')) return 'res-720'; return 'res-other'; }
function ratingClass(rating) { const r = String(rating || '').toUpperCase().replace(/[\s-]/g, ''); if (r === 'G') return 'rating-g'; if (r === 'PG') return 'rating-pg'; if (r === 'PG13') return 'rating-pg13'; if (r === 'R') return 'rating-r'; return ''; }

let toastTimer;
function showToast(msg, duration = 3000) { toast.textContent = msg; toast.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove('show'), duration); }
function updateLastUpdated(date) { const d = (date instanceof Date && !isNaN(date)) ? date : new Date(); let h = d.getHours(), m = d.getMinutes(); const ampm = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12; const mm = String(m).padStart(2, '0'); if (lastUpdatedEl) lastUpdatedEl.textContent = h + ':' + mm + ' ' + ampm; }
function setProgress(pct) { scanFill.style.width = pct + '%'; }

// ─── SHOWS (Local Stubs) ──────────────────────────────────────
function showAvailableCount(show) { return show.seasons.reduce((t, s) => t + s.episodes.filter(e => e.available).length, 0); }
function showTotalCount(show) { return show.seasons.reduce((t, s) => t + s.episodes.length, 0); }

async function loadShowsData(forceRefresh = false) {
  allShows = [];
  renderShows();
}

function renderShows() {
  const container = document.getElementById('shows-grid');
  if (!container) return;
  container.innerHTML = '';
  if (!allShows.length) {
    container.innerHTML = '<div class="empty-state"><span class="empty-icon">◻</span><p>No shows found.</p></div>';
    return;
  }
  updateCounts();
}

function filterAndRenderShows() { renderShows(); }

// ─── SIDEBAR FILTER POPULATION ────────────────────────────────
function populateFilterCheckboxes() {
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

  const resEl = $('filter-resolution-checks');
  if (resEl) {
    const resolutions = [...new Set(allMovies.map(m => m.resolution).filter(Boolean))].sort((a, b) => parseResolutionScore(b) - parseResolutionScore(a));
    resEl.innerHTML = resolutions.map(r => `
      <label class="check-row">
        <input type="checkbox" value="${escHtml(r)}" data-filter="resolution" ${activeFilters.resolution.has(r) ? 'checked' : ''} />
        <span class="check-label ${resClass(r)}">${escHtml(r)}</span>
      </label>`).join('');
  }

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
  const totalMovies = allMovies.length;
  const totalEps    = allShows.reduce((t, s) => t + showTotalCount(s), 0);
  let totalText, availText;
  if (activeTab === 'shows') { totalText = totalEps + ' Episodes'; availText = '0 Available'; } 
  else if (activeTab === 'stats') { totalText = (totalMovies + totalEps) + ' Files'; availText = totalMovies + ' Available'; } 
  else { totalText = totalMovies + ' Movies'; availText = totalMovies + ' Available'; }
  if (movieCount) movieCount.textContent = totalText;
  if (availCount) availCount.textContent = availText;
}

// ─── SORT & FILTER ────────────────────────────────────────────
function applyFilters() {
  const q = normalize(searchInput ? searchInput.value : '');
  filtered = allMovies.filter(m => {
    if (q && !normalize(m.title).includes(q)) return false;
    if (activeFilters.maturity.size   > 0 && !activeFilters.maturity.has(m.maturityRating))     return false;
    if (activeFilters.resolution.size > 0 && !activeFilters.resolution.has(m.resolution))       return false;
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
    else if (key === 'maturity') { va = MATURITY_ORDER[a.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99; vb = MATURITY_ORDER[b.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99; }
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
    rowView.classList.remove('active');
    gridView.classList.add('active');
    renderGrid();
    const total = allMovies.length, shown = filtered.length;
    if (resultsSummary) resultsSummary.textContent = shown === total ? `Showing all ${total} movies` : `Showing ${shown} of ${total} movies`;
  } else {
    gridView.classList.remove('active');
    rowView.classList.add('active');
    renderRows();
    if (resultsSummary) resultsSummary.textContent = `${allMovies.length} movies in the library`;
  }
}

function renderRows() {
  const availableMovies = [...allMovies].sort((a, b) => getRatingScore(b.title) - getRatingScore(a.title));
  const requestedMovies = [...allMovies].filter(m => getRequestCount(m.title) > 0).sort((a, b) => getRequestCount(b.title) - getRequestCount(a.title));
  const imdbMovies = [...allMovies].filter(m => parseFloat(m.imdbRating) > 0).sort((a, b) => (parseFloat(b.imdbRating) || 0) - (parseFloat(a.imdbRating) || 0));

  const rowAvailableEl   = $('row-available-cards');
  const rowRequestedEl   = $('row-requested-cards');
  const rowImdbEl        = $('row-imdb-cards');
  const rowRequestedSec  = $('row-requested');

  if (rowAvailableEl)  renderRowCards(rowAvailableEl, availableMovies.slice(0, 30));
  if (rowRequestedEl && rowRequestedSec) {
    if (requestedMovies.length > 0) { rowRequestedSec.style.display = ''; renderRowCards(rowRequestedEl, requestedMovies.slice(0, 30)); } 
    else { rowRequestedSec.style.display = 'none'; }
  }
  if (rowImdbEl) renderRowCards(rowImdbEl, imdbMovies.slice(0, 30));
}

function renderRowCards(container, movies) {
  container.innerHTML = '';
  const frag = document.createDocumentFragment();
  movies.forEach((m, i) => frag.appendChild(buildCard(m, i, true)));
  container.appendChild(frag);
  const scroller = container.closest('.movie-row-scroll');
  if (scroller) updateRowScrollBtns(scroller);
}

(function initRowScrollBtns() {
  document.addEventListener('click', function(e) {
    const btn = e.target.closest('.row-scroll-btn');
    if (!btn) return;
    const dir = parseInt(btn.dataset.dir, 10);
    const track = document.getElementById(btn.dataset.target);
    if (!track) return;
    const scroller = track.closest('.movie-row-scroll');
    if (!scroller) return;
    scroller.scrollBy({ left: dir * (214 * 3), behavior: 'smooth' });
  });
  document.addEventListener('scroll', (e) => {
    if (!e.target.classList || !e.target.classList.contains('movie-row-scroll')) return;
    updateRowScrollBtns(e.target);
  }, true);
})();

function updateRowScrollBtns(scroller) {
  const wrapper = scroller.closest('.row-scroll-wrapper');
  if (!wrapper) return;
  const leftBtn  = wrapper.querySelector('.row-scroll-btn--left');
  const rightBtn = wrapper.querySelector('.row-scroll-btn--right');
  if (!leftBtn || !rightBtn) return;
  leftBtn.dataset.hidden  = scroller.scrollLeft <= 4 ? '1' : '0';
  rightBtn.dataset.hidden = scroller.scrollLeft >= scroller.scrollWidth - scroller.clientWidth - 4 ? '1' : '0';
}

function renderGrid() {
  if (!movieGrid) return;
  movieGrid.innerHTML = '';
  if (filtered.length === 0) { if (gridEmpty) gridEmpty.hidden = false; return; }
  if (gridEmpty) gridEmpty.hidden = true;
  const frag = document.createDocumentFragment();
  filtered.forEach((m, i) => frag.appendChild(buildCard(m, i, false)));
  movieGrid.appendChild(frag);
}

function buildCard(m, i, isRowCard) {
  const card = document.createElement('div');
  card.className = isRowCard ? 'movie-card row-card' : 'movie-card';
  card.dataset.key = normalize(m.title);
  card.style.animationDelay = Math.min(i * 30, 400) + 'ms';

  const posterClasses  = ['card-poster'];
  if (m.driveLink) posterClasses.push('card-poster--playable');
  else posterClasses.push('card-poster--requestable');

  card.innerHTML = `
    <div class="${posterClasses.join(' ')}">
      ${m.poster ? `<img src="${m.poster}" alt="${escHtml(m.title)}" loading="lazy" onload="this.classList.add('loaded')" />` : ''}
      ${m.driveLink
        ? `<a class="card-play-overlay drive-link" href="${m.driveLink}" target="_blank" rel="noopener" data-title="${escHtml(m.title)}" aria-label="Watch ${escHtml(m.title)}"><div class="card-play-btn"><span class="card-play-icon">&#9654;</span></div></a>`
        : `<div class="card-play-overlay card-request-overlay card-request-overlay--done"><div class="card-request-btn card-request-btn--done"><span class="card-request-icon">&#10003;</span><span class="card-request-label">REQUESTED</span></div></div>`}
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
      ${m.driveLink ? ratingHTML(m.title) : ''}
    </div>
  `;
  return card;
}

function ratingHTML(title) {
  const userVote = getUserRating(title);
  const ups      = getRatingCount(title, 'up');
  const downs    = getRatingCount(title, 'down');

  return `<div class="rating-wrap">
    <button class="rating-btn rating-btn--up ${userVote === 'up' ? 'active' : ''}"
            data-rating-title="${escHtml(title)}" data-rating-type="up" title="Liked it">
      <span class="rating-icon">👍</span><span class="rating-count">${ups || 0}</span>
    </button>
    <button class="rating-btn rating-btn--down ${userVote === 'down' ? 'active' : ''}"
            data-rating-title="${escHtml(title)}" data-rating-type="down" title="Didn't like it">
      <span class="rating-icon">👎</span><span class="rating-count">${downs || 0}</span>
    </button>
  </div>`;
}

function escHtml(str) { return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ─── EVENTS ───────────────────────────────────────────────────
let searchTimer, searchLogTimer;
if (searchInput) {
  searchInput.addEventListener('input', () => {
    const query = searchInput.value;
    if (clearSearch) clearSearch.classList.toggle('visible', query.length > 0);
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => { if (activeTab === 'shows') filterAndRenderShows(); else { render(); saveSettings(); } }, 200);
  });
}
if (clearSearch) {
  clearSearch.addEventListener('click', () => {
    searchInput.value = '';
    clearSearch.classList.remove('visible');
    if (activeTab === 'shows') filterAndRenderShows(); else { render(); saveSettings(); }
    searchInput.focus();
  });
}
if (sortBy) {
  sortBy.addEventListener('change', () => {
    currentSort = sortBy.value;
    if (hasActiveFilters()) applySort();
    saveSettings();
  });
}
if (sortDirBtn) {
  sortDirBtn.addEventListener('click', () => {
    currentDir = currentDir === 'desc' ? 'asc' : 'desc';
    sortDirBtn.textContent = currentDir === 'desc' ? '↓' : '↑';
    sortDirBtn.title = currentDir === 'desc' ? 'Descending' : 'Ascending';
    if (hasActiveFilters()) applySort();
    saveSettings();
  });
}
if (sidebarClearBtn) sidebarClearBtn.addEventListener('click', clearAllFilters);

const mainContent = document.getElementById('main-content');
if (mainContent) {
  mainContent.addEventListener('click', e => {
    const btn = e.target.closest('.rating-btn');
    if (!btn) return;

    const title = btn.dataset.ratingTitle;
    const type  = btn.dataset.ratingType;
    if (!title || !type) return;

    const key = normalize(title);
    if (ratingInflight.has(key)) return;
    ratingInflight.add(key);

    const prevVote = getUserRating(title);
    const prevUp   = getRatingCount(title, 'up');
    const prevDown = getRatingCount(title, 'down');
    const nextVote = prevVote === type ? null : type; // toggle off if same

    // Optimistically update counts
    const delta = { up: prevUp, down: prevDown };
    if (prevVote === 'up')   delta.up   = Math.max(0, delta.up   - 1);
    if (prevVote === 'down') delta.down = Math.max(0, delta.down - 1);
    if (nextVote === 'up')   delta.up   += 1;
    if (nextVote === 'down') delta.down += 1;
    ratingCounts[key] = { up: delta.up, down: delta.down };

    if (nextVote) userRatings[key] = nextVote;
    else          delete userRatings[key];
    saveUserRatings();

    applyRatingDOM(title, nextVote, delta.up, delta.down, btn);

    if (nextVote === 'up')        showToast('👍 Liked ' + title);
    else if (nextVote === 'down') showToast('👎 Disliked ' + title);
    else                          showToast('Rating removed for ' + title);

    // Send to server
    fetch(`${API_BASE}/api/rate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ title, type: nextVote, prev: prevVote, did: getDeviceId() })
    })
    .then(res => res.json())
    .then(data => {
      if (data.ratings) {
        ratingCounts[key] = data.ratings;
        applyRatingDOM(title, getUserRating(title), data.ratings.up, data.ratings.down, null);
      }
    })
    .catch(e => console.error('Rating failed:', e))
    .finally(() => ratingInflight.delete(key));
  });
}

// ─── NOTIFICATIONS (Local Stubs) ─────────────────────────────
let serverNotifications = []; let clientNotifications = [];
function notifId(n) { return (n.name || '') + '||' + (n.message || '') + '||' + (n.time || ''); }
function getReadNotifIds() { try { return new Set(JSON.parse(localStorage.getItem('thedrive_readNotifIds_v1') || '[]')); } catch(e) { return new Set(); } }
function saveReadNotifIds(set) { try { localStorage.setItem('thedrive_readNotifIds_v1', JSON.stringify([...set])); } catch(e) {} }
function fetchServerNotifications() { return Promise.resolve(); }
function checkAndNotifyRequestFulfillments() { }
function hasUnreadNotifications() { return false; }
function updateNotificationBell() { const dot = refreshBtn ? refreshBtn.querySelector('.notif-dot') : null; if (dot) dot.style.display = 'none'; }
function markAllAsRead() { }
function renderNotificationPanel() { const panel = document.getElementById('notif-panel'); if (panel) { const listEl = panel.querySelector('.notif-list'); if (listEl) listEl.innerHTML = '<div class="notif-empty">No notifications</div>'; } }
function toggleNotificationPanel() { const panel = document.getElementById('notif-panel'); if (panel) panel.style.display = panel.style.display === 'block' ? 'none' : 'block'; }

(function initNotificationBell() {
  if (!refreshBtn) return;
  refreshBtn.id = 'notif-btn';   
  refreshBtn.classList.remove('spinning');
  refreshBtn.innerHTML = `
    <svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round">
      <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path>
      <path d="M13.73 21a2 2 0 0 1-3.46 0"></path>
    </svg>
    <span class="notif-dot" style="display:none; position:absolute; top:4px; right:4px; width:8px; height:8px; background:var(--red); border-radius:50%; border:1px solid var(--bg);"></span>
  `;
  refreshBtn.title = 'Notifications';
  const panel = document.createElement('div');
  panel.id = 'notif-panel';
  panel.style.cssText = ['display:none','position:absolute','top:calc(100% + 8px)','right:0','width:320px','max-height:400px','overflow-y:auto','background:var(--surface)','border:1px solid var(--border)','border-radius:8px','box-shadow:0 8px 24px rgba(0,0,0,0.4)','z-index:1000','padding:12px','font-size:13px','color:var(--text)'].join(';');
  panel.innerHTML = '<div class="notif-list"></div>';
  refreshBtn.parentNode.style.position = 'relative';
  refreshBtn.parentNode.appendChild(panel);
  refreshBtn.addEventListener('click', (e) => { e.stopPropagation(); toggleNotificationPanel(); });
  document.addEventListener('click', (e) => {
    const panel = document.getElementById('notif-panel');
    if (!panel || panel.style.display !== 'block') return;
    if (!refreshBtn.contains(e.target) && !panel.contains(e.target)) panel.style.display = 'none';
  });
})();

// ─── FOOTER FORM ──────────────────────────────────────────────
(function() {
  const submitBtn = document.getElementById('footer-submit');
  const msgInput  = document.getElementById('footer-message');
  const statusEl  = document.getElementById('footer-form-status');
  if (!submitBtn) return;
  function setStatus(msg, type) { statusEl.textContent = msg; statusEl.className = 'footer-form-status ' + type; statusEl.hidden = false; }
  submitBtn.addEventListener('click', async () => {
    const message = (msgInput.value || '').trim();
    if (!message) { setStatus('Please enter a message before sending.', 'error'); msgInput.focus(); return; }
    setStatus('⚠ Messaging disabled in local mode.', 'error');
  });
})();

// ─── LOAD DATA ────────────────────────────────────────────────
async function loadData() {
  setProgress(5);
  try {
    // Fetch videos
    const url = API_BASE + '/videos' + ('?_cb=' + Date.now());
    const r = await fetch(url);
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const videos = await r.json();

    allMovies = videos.map(v => ({
      title:           v.title,
      runtime:         v.runtime || '',
      resolution:      '', 
      maturityRating:  v.maturityRating || '',
      releaseDate:     v.year || '',
      year:            v.year || '—',
      fileSize:        '',
      imdbRating:      v.imdbRating || '',
      available:       true,
      driveLink:       API_BASE + v.video,
      driveResolution: '',
      poster:          v.poster ? (API_BASE + v.poster) : null
    }));

    // Fetch ratings
    try {
      const ratingsRes = await fetch(`${API_BASE}/api/ratings`);
      if (ratingsRes.ok) ratingCounts = await ratingsRes.json();
    } catch(e) { console.warn("Failed to fetch ratings", e); }

    setProgress(100);
    render();
    populateFilterCheckboxes();
    updateCounts();
    updateLastUpdated();
    setTimeout(() => scanBar.classList.add('hidden'), 300);
  } catch(e) {
    console.error('Failed to load videos:', e);
    showToast('⚠ Could not reach backend at ' + (API_BASE || 'localhost') + '. Is the server running?');
    setProgress(100);
    setTimeout(() => scanBar.classList.add('hidden'), 300);
  }
}

// ─── DONATIONS BAR (Local Stub) ────────────────────────────────
function fetchDonationsData() { }
(function() {
  const overlay   = document.getElementById('donation-modal-overlay');
  const closeBtn  = document.getElementById('donation-modal-close');
  const readMore  = document.getElementById('donation-readmore-btn');
  if (readMore && overlay) readMore.addEventListener('click', () => { overlay.style.display = 'flex'; });
  if (closeBtn && overlay) closeBtn.addEventListener('click', () => { overlay.style.display = 'none'; });
  if (overlay) overlay.addEventListener('click', e => { if (e.target === overlay) overlay.style.display = 'none'; });
})();

// ─── HEARTBEAT & ONLINE COUNT ─────────────────────────────────
let heartbeatOnline = true;
function startHeartbeat() { }
function isClientOnline() { return heartbeatOnline; }

async function pushPresencePing() {
  try {
    const res = await fetch(`${API_BASE}/api/heartbeat?did=${getDeviceId()}`);
    const data = await res.json();
    const onlineEl = document.getElementById('online-count');
    if (onlineEl && data && typeof data.online === 'number') {
      onlineEl.textContent = data.online;
    }
  } catch(e) { /* ignore heartbeat errors */ }
}

// ─── STATS ────────────────────────────────────────────────────
let statsLoaded = false, statsLoadedAt = 0;
let chartLibrary = null, chartUsers = null, chartPresence = null;
let lastPresenceAppendAt = 0;

function initStatsTab() {
  renderLocalStats();
  if (!statsLoaded || Date.now() - statsLoadedAt > 60000) {
    fetchStatsData();
    statsLoaded = true;
    statsLoadedAt = Date.now();
  }
}

function renderLocalStats() {
  if (!allMovies.length && !allShows.length) return;
  const total = allMovies.length;
  const totalAll = total;
  const availAll = total;

  const pct = totalAll > 0 ? ((availAll / totalAll) * 100).toFixed(2) : '0.00';
  const fracEl = $('upload-fraction'), pctEl = $('upload-pct'), fillEl = $('upload-fill');
  if (fracEl)  fracEl.textContent  = availAll + ' / ' + totalAll + ' movies uploaded';
  if (pctEl)   pctEl.textContent   = pct + '%';
  if (fillEl)  fillEl.style.width  = parseFloat(pct) + '%';
  setText('stat-total-films', totalAll);
  setText('stat-available', availAll);
  setText('stat-total-size', '—');
  
  let totalMins = 0;
  allMovies.forEach(m => { totalMins += parseRuntimeMinutes(m.runtime); });
  if (totalMins > 0) setText('stat-total-runtime', Math.floor(totalMins / 60) + 'h ' + (totalMins % 60) + 'm');
  
  const rated = allMovies.filter(m => parseFloat(m.imdbRating) > 0);
  if (rated.length) setText('stat-avg-imdb', '★ ' + (rated.reduce((s, m) => s + parseFloat(m.imdbRating), 0) / rated.length).toFixed(1));
  
  const matNorm = r => String(r || '').toUpperCase().replace(/[\s-]/g, '');
  setText('stat-g',    allMovies.filter(m => matNorm(m.maturityRating) === 'G').length);
  setText('stat-pg',   allMovies.filter(m => matNorm(m.maturityRating) === 'PG').length);
  setText('stat-pg13', allMovies.filter(m => matNorm(m.maturityRating) === 'PG13').length);
  setText('stat-r',    allMovies.filter(m => matNorm(m.maturityRating) === 'R').length);
}

function setText(id, val) { const el = $(id); if (el) el.textContent = String(val); }

async function fetchStatsData() {
  try {
    const res = await fetch(`${API_BASE}/api/stats`);
    if (!res.ok) return;
    const data = await res.json();
    if (!data) return;
    
    if (typeof data.uniqueDevices === 'number') setText('stat-total-users', data.uniqueDevices);
    if (data.snapshots && data.snapshots.length) renderLibraryChart(data.snapshots);
    if (data.userHistory && data.userHistory.length) renderUserChart(data.userHistory);
    if (data.presence && data.presence.length) renderPresenceChart(data.presence);
    else showPresencePlaceholder();
    
    const onlineEl = document.getElementById('online-count');
    if (onlineEl && typeof data.online === 'number') onlineEl.textContent = data.online;
  } catch(e) {
    console.error('Stats fetch failed', e);
  }
}

function renderLibraryChart(snapshots) {
  if (!window.Chart) return;
  const canvas = $('chart-library'); if (!canvas) return;
  const cfg = { type: 'line', data: { labels: snapshots.map(s => s.date), datasets: [{ label: 'Total Files', data: snapshots.map(s => s.total), borderColor: '#9090a8', backgroundColor: 'rgba(144,144,168,0.08)', borderWidth: 2, pointRadius: 3, pointBackgroundColor: '#9090a8', tension: 0.3, fill: true }, { label: 'Available Files', data: snapshots.map(s => s.available), borderColor: '#e8c547', backgroundColor: 'rgba(232,197,71,0.10)', borderWidth: 2, pointRadius: 3, pointBackgroundColor: '#e8c547', tension: 0.3, fill: true }] }, options: chartOptions('Files') };
  if (chartLibrary) chartLibrary.destroy();
  chartLibrary = new Chart(canvas, cfg);
}

function renderUserChart(userHistory) {
  if (!window.Chart) return;
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
  if (!window.Chart) return;
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

// ─── DEVICE ID CORNER REVEAL ──────────────────────────────────
(function initDeviceIdReveal() {
  const el = document.createElement('div');
  el.id = 'device-id-corner';
  el.style.cssText = [
    'position:fixed','bottom:12px','right:14px','z-index:99999','font-family:monospace','font-size:10px','color:rgba(144,144,168,0.9)','background:rgba(18,18,26,0.85)','border:1px solid rgba(80,80,110,0.4)','border-radius:5px','padding:4px 9px','letter-spacing:0.08em','pointer-events:none','opacity:0','transition:opacity 0.2s ease'
  ].join(';');
  document.body.appendChild(el);

  const CORNER_PX = 12;
  let hideTimer = null;

  document.addEventListener('mousemove', function(e) {
    const nearRight  = window.innerWidth  - e.clientX < CORNER_PX;
    const nearBottom = window.innerHeight - e.clientY < CORNER_PX;

    if (nearRight && nearBottom) {
      if (!el.textContent) el.textContent = 'DID: ' + getDeviceId();
      el.style.opacity = '1';
      clearTimeout(hideTimer);
    } else if (el.style.opacity !== '0') {
      clearTimeout(hideTimer);
      hideTimer = setTimeout(function() { el.style.opacity = '0'; }, 300);
    }
  });
})();

// ─── MAIN INIT ────────────────────────────────────────────────
(async function init() {
  try { localStorage.removeItem(LOCAL_SETTINGS_KEY); } catch(e) {}
  currentSort = 'title'; currentDir = 'asc';
  if (sortBy) sortBy.value = 'title';
  if (sortDirBtn) sortDirBtn.textContent = '↓';

  // ── Fetch live backend URL from config.json ──
  try {
    const configResponse = await fetch('config.json?t=' + Date.now());
    if (configResponse.ok) {
        const config = await configResponse.json();
        API_BASE = config.API_BASE || '';
    }
  } catch (e) {
    console.log('No config.json found, defaulting to local backend.');
    API_BASE = ''; 
  }

  await initWithGate();
  loadShowsData();
  await loadData();

  // Start presence heartbeat
  pushPresencePing();
  setInterval(pushPresencePing, 10000);

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      activeTab = tab;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + tab).classList.add('active');
      updateCounts();
      if (tab === 'stats')  initStatsTab();
      if (tab === 'shows')  filterAndRenderShows();
    });
  });
})();
