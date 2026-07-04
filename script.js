/* =============================================================
   THE DRIVE — script.js
   ============================================================= */

// ─── CONFIG ───────────────────────────────────────────────────
let API_BASE = ''; 
const SHEET_CSV_URL    = '';
const SHOWS_CSV_URL    = '';
const DRIVE_SCRIPT_URL = '';

const AUTO_RELOAD_MS = 30 * 60 * 1000;
setTimeout(() => location.reload(), AUTO_RELOAD_MS);

// ─── ACCESS KEY GATE ──────────────────────────────────────────
const LOCAL_DEVICE_ID = 'thedrive_device_id_v1';
let cachedDeviceData = null;

async function getSavedKey() {
  try {
    if (!cachedDeviceData) {
      const did = getDeviceId();
      const res = await fetch(`${API_BASE}/api/device/data?did=${encodeURIComponent(did)}`);
      if (res.ok) cachedDeviceData = await res.json();
    }
    return cachedDeviceData?.access_key || null;
  } catch(e) { return null; }
}
function saveKey(key) {}

function getDeviceId() {
  try {
    let did = localStorage.getItem(LOCAL_DEVICE_ID);
    if (!did) {
      did = 'did-' + Array.from({ length: 12 }, () => Math.floor(Math.random() * 16).toString(16)).join('').toUpperCase();
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
  const savedKey = await getSavedKey();

  function triggerServerRetry() {
    if (overlay) overlay.classList.remove('gate-overlay-hidden');
    if (submitBtn) { submitBtn.classList.add('loading'); submitBtn.textContent = 'RETRYING...'; }
    if (input) input.disabled = true;
    if (errorEl) { errorEl.textContent = 'Cannot reach server. Retrying in 5 seconds...'; errorEl.hidden = false; }
    setTimeout(() => location.reload(), 5000);
  }

  try {
    const blockRes = await fetch(`${API_BASE}/api/keys/check-device?did=${did}`);
    const blockData = await blockRes.json();
    if (blockData.blocked) {
      if (overlay) overlay.classList.remove('gate-overlay-hidden');
      const titleEl = overlay.querySelector('.gate-title');
      if (titleEl) titleEl.textContent = 'ACCESS DENIED';
      return;
    }
  } catch(e) { triggerServerRetry(); return; }

  if (savedKey) {
    try {
      const valRes = await fetch(`${API_BASE}/api/keys/validate?code=${savedKey}&did=${did}`);
      const valData = await valRes.json();
      if (valData.valid) {
        if (overlay) { overlay.classList.add('gate-overlay-hidden'); overlay.style.display = 'none'; }
        return;
      }
    } catch(e) { triggerServerRetry(); return; }
  }

  if (overlay) overlay.classList.remove('gate-overlay-hidden');
  if (submitBtn) submitBtn.classList.remove('loading');
  if (input) input.disabled = false;
  
  return new Promise(resolve => {
    if (!submitBtn || !input) { resolve(); return; }
    async function attempt() {
      const keyStr = input.value.trim().toUpperCase();
      if (!keyStr) return;
      submitBtn.classList.add('loading'); submitBtn.textContent = 'CHECKING…'; errorEl.hidden = true;
      try {
        const valRes = await fetch(`${API_BASE}/api/keys/validate?code=${keyStr}&did=${did}`);
        const valData = await valRes.json();
        if (!valData.valid) {
          errorEl.textContent = 'Invalid key.'; errorEl.hidden = false;
          submitBtn.classList.remove('loading'); submitBtn.textContent = 'ENTER THE DRIVE'; return;
        }
        const useRes = await fetch(`${API_BASE}/api/keys/use?code=${keyStr}&did=${did}`);
        const useData = await useRes.json();
        if (useData.success) {
          saveKey(keyStr); overlay.classList.add('gate-overlay-hidden');
          setTimeout(() => { overlay.style.display = 'none'; }, 350); resolve();
        } else {
          errorEl.textContent = 'Failed to activate key.'; errorEl.hidden = false;
          submitBtn.classList.remove('loading'); submitBtn.textContent = 'ENTER THE DRIVE';
        }
      } catch(e) { saveKey(keyStr); triggerServerRetry(); resolve(); }
    }
    
    submitBtn.addEventListener('click', attempt);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') attempt(); });
    setTimeout(() => input.focus(), 100);
  });
}

// ─── STATE ────────────────────────────────────────────────────
let allMovies   = [], allShows = [], filtered = [];
let currentSort = 'title', currentDir = 'asc', activeTab = 'home';
let activeFilters = { maturity: new Set(), status: new Set(), resolution: new Set(), genre: new Set() };
let libraryData = { watched: [], unwatched: [] };

function hasActiveFilters() {
  const search = searchInput ? searchInput.value.trim() : '';
  return search.length > 0 || activeFilters.maturity.size > 0 || activeFilters.status.size > 0 || activeFilters.resolution.size > 0 || activeFilters.genre.size > 0;
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
function getRatingScore(title) { const r = ratingCounts[normalize(title)]; return r ? (r.up || 0) - (r.down || 0) : 0; }
function applyRatingDOM(title, nextVote, upCount, downCount, clickedBtn) {
  document.querySelectorAll(`[data-rating-title="${CSS.escape(title)}"]`).forEach(b => {
    const bType = b.dataset.ratingType, isActive = nextVote === bType;
    b.classList.toggle('active', isActive);
    if (b === clickedBtn && isActive) { b.classList.remove('just-voted'); void b.offsetWidth; b.classList.add('just-voted'); }
    const countEl = b.querySelector('.rating-count'); if (countEl) countEl.textContent = bType === 'up' ? upCount : downCount;
  });
}

// ─── DOM REFS ─────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const searchInput = $('search-input'), clearSearch = $('clear-search'), sortBy = $('sort-by'), sortDirBtn = $('sort-dir-btn');
const movieCount = $('movie-count'), availCount = $('available-count'), resultsSummary = $('results-summary');
const scanBar = $('scan-bar'), lastUpdatedEl = $('last-updated'), scanFill = $('scan-fill');
const toast = $('toast'), rowView = $('row-view'), movieGrid = $('movie-grid');
const gridEmpty = $('grid-empty'), sidebarClearBtn = $('sidebar-clear-btn');

// ─── UTILITIES ────────────────────────────────────────────────
function normalize(str) { return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function extractYear(dateStr) { if (!dateStr) return '—'; const m = String(dateStr).match(/\d{4}/); return m ? m[0] : '—'; }
function parseRuntimeMinutes(str) { if (!str) return 0; const hm = str.match(/(\d+)\s*h(?:r|ours?)?\s*(\d+)?\s*m?/i); if (hm) return parseInt(hm[1]) * 60 + (parseInt(hm[2]) || 0); const m = str.match(/(\d+)/); return m ? parseInt(m[1]) : 0; }

const MATURITY_ORDER = { 'G': 1, 'PG': 2, 'PG-13': 3, 'R': 4, 'NC-17': 5, 'NR': 6 };

function normalizeMaturity(str) {
  if (!str) return 'NR';
  // Remove "Rated " prefix, then remove spaces and hyphens for easy checking
  let s = String(str).toUpperCase().replace(/RATED\s*/, '').replace(/[\s-]/g, '').trim();
  if (s === 'NC17') return 'NC-17';
  if (s === 'PG13') return 'PG-13';
  if (s === 'R' || s === 'PG' || s === 'G') return s;
  return 'NR'; // Fallback for Unrated, Not Rated, etc.
}

function parseResolutionScore(res) { if (!res) return 0; const s = String(res).toUpperCase().trim(); if (s === '4K' || s === 'UHD' || s.includes('2160')) return 2160; const m = s.match(/(\d+)/); return m ? parseInt(m[1], 10) : 0; }
function imdbClass(rating) { const r = parseFloat(rating); if (r >= 8) return 'imdb-high'; if (r >= 6.5) return 'imdb-mid'; return 'imdb-low'; }
function resClass(res) { const r = String(res).toUpperCase(); if (r.includes('4K') || r.includes('2160')) return 'res-4k'; if (r.includes('1080')) return 'res-1080'; if (r.includes('720') || r.includes('576')) return 'res-720'; return 'res-other'; }
function ratingClass(rating) { const r = normalizeMaturity(rating); if (r === 'G') return 'rating-g'; if (r === 'PG') return 'rating-pg'; if (r === 'PG-13') return 'rating-pg13'; if (r === 'R') return 'rating-r'; if (r === 'NC-17') return 'rating-nc17'; return 'rating-nr'; }
let toastTimer;
function showToast(msg, duration = 3000) { toast.textContent = msg; toast.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove('show'), duration); }
function updateLastUpdated(date) { const d = (date instanceof Date && !isNaN(date)) ? date : new Date(); let h = d.getHours(), m = d.getMinutes(); const ampm = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12; const mm = String(m).padStart(2, '0'); if (lastUpdatedEl) lastUpdatedEl.textContent = h + ':' + mm + ' ' + ampm; }
function setProgress(pct) { scanFill.style.width = pct + '%'; }
function formatTime(sec) {
  if (isNaN(sec)) return '0:00';
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = Math.floor(sec % 60);
  return h > 0 ? `${h}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}` : `${m}:${String(s).padStart(2,'0')}`;
}

// ─── SHOWS DATA ───────────────────────────────────────────────
async function loadShowsData() {
  try {
    const r = await fetch(`${API_BASE}/shows?_cb=` + Date.now());
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const videos = await r.json();
    allShows = videos.map(v => ({
      title: v.title, runtime: v.runtime || '', resolution: v.resolution || '', maturityRating: v.maturityRating || '',
      year: v.year || '—', imdbRating: v.imdbRating || '', plot: v.plot || '', cast: v.cast || [], genres: v.genres || [],
      driveLink: API_BASE + v.video, poster: v.poster ? (API_BASE + v.poster) : null
    }));
  } catch(e) {
    allShows = [];
  }
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
  container.className = 'movie-grid';
  const frag = document.createDocumentFragment();
  allShows.forEach((m, i) => frag.appendChild(buildCard(m, i, false)));
  container.appendChild(frag);
  updateCounts();
}

// ─── UPDATES DATA ─────────────────────────────────────────────
async function loadUpdates() {
  const container = $('updates-container');
  if (!container) return;
  container.innerHTML = '<div class="empty-state"><span class="empty-icon">◻</span><p>Loading updates...</p></div>';
  try {
    const res = await fetch(`${API_BASE}/api/updates`);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const updates = await res.json();
    if (!updates || updates.length === 0) {
      container.innerHTML = '<div class="empty-state"><span class="empty-icon">◻</span><p>No updates yet.</p></div>';
      return;
    }
    container.innerHTML = updates.map(u => `
      <div class="update-panel">
        <div class="update-header">${escHtml(u.header)}</div>
        <span class="update-timestamp">${u.timestamp}</span>
        <div class="update-body">${escHtml(u.body)}</div>
      </div>
    `).join('');
  } catch(e) {
    container.innerHTML = '<div class="empty-state"><span class="empty-icon">◻</span><p>Failed to load updates.</p></div>';
  }
}

// ─── SIDEBAR FILTERS ──────────────────────────────────────────
function populateFilterCheckboxes() {
  const maturityEl = $('filter-maturity-checks');
  if (maturityEl) {
    // Normalize all ratings before putting them in the sidebar so they group properly
    const ratings = [...new Set(allMovies.map(m => normalizeMaturity(m.maturityRating)).filter(Boolean))].sort((a, b) => (MATURITY_ORDER[a] || 99) - (MATURITY_ORDER[b] || 99));
    maturityEl.innerHTML = ratings.map(r => `<label class="check-row"><input type="checkbox" value="${escHtml(r)}" data-filter="maturity" ${activeFilters.maturity.has(r) ? 'checked' : ''} /><span class="check-label ${ratingClass(r)}">${escHtml(r)}</span></label>`).join('');
  }
  const resEl = $('filter-resolution-checks');
  if (resEl) {
    const resolutions = [...new Set(allMovies.map(m => m.resolution).filter(Boolean))].sort((a, b) => parseResolutionScore(b) - parseResolutionScore(a));
    resEl.innerHTML = resolutions.map(r => `<label class="check-row"><input type="checkbox" value="${escHtml(r)}" data-filter="resolution" ${activeFilters.resolution.has(r) ? 'checked' : ''} /><span class="check-label ${resClass(r)}">${escHtml(r)}</span></label>`).join('');
  }
  const genreEl = $('filter-genre-checks');
  if (genreEl) {
    const genres = [...new Set(allMovies.flatMap(m => m.genres).filter(Boolean))].sort();
    genreEl.innerHTML = genres.map(r => `<label class="check-row"><input type="checkbox" value="${escHtml(r)}" data-filter="genre" ${activeFilters.genre.has(r) ? 'checked' : ''} /><span class="check-label">${escHtml(r)}</span></label>`).join('');
  }
  bindSidebarCheckboxes(); updateClearBtn();
}
function bindSidebarCheckboxes() {
  document.querySelectorAll('.sidebar-checks input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      const f = cb.dataset.filter, v = cb.value;
      if (cb.checked) activeFilters[f].add(v); else activeFilters[f].delete(v);
      updateClearBtn(); render(); saveSettings();
    });
  });
}
function updateClearBtn() { if (sidebarClearBtn) sidebarClearBtn.hidden = !(activeFilters.maturity.size > 0 || activeFilters.status.size > 0 || activeFilters.resolution.size > 0 || activeFilters.genre.size > 0); }
function clearAllFilters() { activeFilters.maturity.clear(); activeFilters.status.clear(); activeFilters.resolution.clear(); activeFilters.genre.clear(); if (searchInput) { searchInput.value = ''; clearSearch && clearSearch.classList.remove('visible'); } document.querySelectorAll('.sidebar-checks input[type="checkbox"]').forEach(cb => cb.checked = false); updateClearBtn(); render(); saveSettings(); }
function updateCounts() {
  const totalMovies = allMovies.length, totalEps = allShows.length;
  let totalText = activeTab === 'shows' ? totalEps + ' Shows' : totalMovies + ' Movies';
  if (movieCount) movieCount.textContent = totalText;
  if (availCount) availCount.style.display = 'none';
}

// ─── SORT & FILTER ────────────────────────────────────────────
function applyFilters() {
  const q = normalize(searchInput ? searchInput.value : '');
  filtered = allMovies.filter(m => {
    if (q && !normalize(m.title).includes(q)) return false;
    // Apply normalizeMaturity to the check so it matches the sidebar values
    if (activeFilters.maturity.size > 0 && !activeFilters.maturity.has(normalizeMaturity(m.maturityRating))) return false;
    if (activeFilters.resolution.size > 0 && !activeFilters.resolution.has(m.resolution)) return false;
    if (activeFilters.genre.size > 0) {
      if (!m.genres || !m.genres.some(g => activeFilters.genre.has(g))) return false;
    }
    return true;
  });
  applySort();
}
function applySort() {
  const key = currentSort, dir = currentDir;
  filtered.sort((a, b) => {
    let va, vb;
    if (key === 'title') { va = a.title.toLowerCase(); vb = b.title.toLowerCase(); }
    else if (key === 'imdb') { va = parseFloat(a.imdbRating) || 0; vb = parseFloat(b.imdbRating) || 0; }
    else if (key === 'year') { va = parseInt(a.year) || 0; vb = parseInt(b.year) || 0; }
    else if (key === 'rating') { va = getRatingScore(a.title); vb = getRatingScore(b.title); }
    else if (key === 'maturity') { va = MATURITY_ORDER[normalizeMaturity(a.maturityRating)] || 99; vb = MATURITY_ORDER[normalizeMaturity(b.maturityRating)] || 99; }
    else if (key === 'res') { va = parseResolutionScore(a.resolution); vb = parseResolutionScore(b.resolution); }
    if (va < vb) return dir === 'asc' ? -1 : 1; if (va > vb) return dir === 'asc' ? 1 : -1; return 0;
  });
  renderCurrentView();
}

// ─── RENDER ───────────────────────────────────────────────────
function render() {
  if (hasActiveFilters()) {
    if (activeTab !== 'movies') {
      const moviesTab = document.querySelector('.nav-btn[data-view="movies"]');
      if (moviesTab) moviesTab.click();
      return;
    }
    applyFilters();
  } else {
    if (activeTab === 'home') renderRows();
    else if (activeTab === 'movies') { filtered = [...allMovies]; applySort(); }
    else if (activeTab === 'shows') renderShows();
    else if (activeTab === 'library') renderLibrary();
  }
}

function renderCurrentView() {
  if (activeTab === 'movies') {
    renderGrid();
    if (resultsSummary) resultsSummary.textContent = hasActiveFilters() ? `Showing ${filtered.length} of ${allMovies.length} movies` : `${allMovies.length} movies in the library`;
  }
}

function createRowHtml(id, title, subtitle) {
  return `
    <section class="movie-row-section">
      <div class="row-header">
        <h2 class="row-title">${title}</h2>
        ${subtitle ? `<span class="row-subtitle">${subtitle}</span>` : ''}
      </div>
      <div class="row-scroll-wrapper">
        <button class="row-scroll-btn row-scroll-btn--left" data-dir="-1" data-target="${id}">‹</button>
        <button class="row-scroll-btn row-scroll-btn--right" data-dir="1" data-target="${id}">›</button>
        <div class="movie-row-scroll">
          <div id="${id}" class="movie-row"></div>
        </div>
      </div>
    </section>
  `;
}

function renderRows() {
  if (!rowView) return;
  let html = '';
  
  const continueMovies = allMovies.filter(m => {
    const p = getVideoProgress(m.title);
    return p.time > 10 && p.duration > 0;
  }).sort((a,b) => getVideoProgress(b.title).time - getVideoProgress(a.title).time);
  
  if (continueMovies.length > 0) {
    html += createRowHtml('row-continue-cards', 'CONTINUE WATCHING', 'PICK UP WHERE YOU LEFT OFF');
  }

  html += createRowHtml('row-available-cards', 'AVAILABLE FILMS', 'SORTED BY RATING');
  html += createRowHtml('row-imdb-cards', 'TOP RATED', 'IMDB');

  const genres = [...new Set(allMovies.flatMap(m => m.genres).filter(Boolean))].sort();
  genres.forEach(g => {
    html += createRowHtml('row-genre-' + normalize(g), g.toUpperCase(), 'GENRE');
  });

  rowView.innerHTML = html;

  if ($('row-continue-cards')) renderRowCards($('row-continue-cards'), continueMovies.slice(0, 20));
  if ($('row-available-cards')) {
    const availableMovies = [...allMovies].sort((a, b) => getRatingScore(b.title) - getRatingScore(a.title));
    renderRowCards($('row-available-cards'), availableMovies.slice(0, 30));
  }
  if ($('row-imdb-cards')) {
    const imdbMovies = [...allMovies].filter(m => parseFloat(m.imdbRating) > 0).sort((a, b) => (parseFloat(b.imdbRating) || 0) - (parseFloat(a.imdbRating) || 0));
    renderRowCards($('row-imdb-cards'), imdbMovies.slice(0, 30));
  }
  genres.forEach(g => {
    const id = 'row-genre-' + normalize(g);
    const cards = allMovies.filter(m => m.genres.includes(g)).sort((a,b) => getRatingScore(b.title) - getRatingScore(a.title)).slice(0, 30);
    if ($(id)) renderRowCards($(id), cards);
  });
}

function renderRowCards(container, movies) {
  container.innerHTML = '';
  const frag = document.createDocumentFragment();
  movies.forEach((m, i) => frag.appendChild(buildCard(m, i, true)));
  container.appendChild(frag);
  const scroller = container.closest('.movie-row-scroll'); if (scroller) updateRowScrollBtns(scroller);
}
(function initRowScrollBtns() {
  document.addEventListener('click', function(e) {
    const btn = e.target.closest('.row-scroll-btn'); if (!btn) return;
    const track = document.getElementById(btn.dataset.target); if (!track) return;
    const scroller = track.closest('.movie-row-scroll'); if (!scroller) return;
    scroller.scrollBy({ left: parseInt(btn.dataset.dir, 10) * (214 * 3), behavior: 'smooth' });
  });
  document.addEventListener('scroll', (e) => { if (e.target.classList && e.target.classList.contains('movie-row-scroll')) updateRowScrollBtns(e.target); }, true);
})();
function updateRowScrollBtns(scroller) {
  const wrapper = scroller.closest('.row-scroll-wrapper'); if (!wrapper) return;
  const lBtn = wrapper.querySelector('.row-scroll-btn--left'), rBtn = wrapper.querySelector('.row-scroll-btn--right');
  if (lBtn) lBtn.dataset.hidden = scroller.scrollLeft <= 4 ? '1' : '0';
  if (rBtn) rBtn.dataset.hidden = scroller.scrollLeft >= scroller.scrollWidth - scroller.clientWidth - 4 ? '1' : '0';
}
function renderGrid() {
  if (!movieGrid) return; movieGrid.innerHTML = '';
  if (filtered.length === 0) { if (gridEmpty) gridEmpty.hidden = false; return; }
  if (gridEmpty) gridEmpty.hidden = true;
  const frag = document.createDocumentFragment();
  filtered.forEach((m, i) => frag.appendChild(buildCard(m, i, false)));
  movieGrid.appendChild(frag);
}

function renderLibrary() {
  const watchedGrid = $('library-watched-grid');
  const unwatchedGrid = $('library-unwatched-grid');
  const watchedEmpty = $('library-watched-empty');
  const unwatchedEmpty = $('library-unwatched-empty');

  if (!watchedGrid || !unwatchedGrid) return;

  // WATCHING section (Top)
  watchedGrid.innerHTML = '';
  const watchingMovies = allMovies.filter(m => {
    const p = getVideoProgress(m.title);
    return p.time > 10 && p.duration > 0;
  }).sort((a,b) => getVideoProgress(b.title).time - getVideoProgress(a.title).time);

  if (watchingMovies.length === 0) {
    watchedEmpty.hidden = false;
  } else {
    watchedEmpty.hidden = true;
    const frag = document.createDocumentFragment();
    watchingMovies.forEach((m, i) => frag.appendChild(buildCard(m, i, false)));
    watchedGrid.appendChild(frag);
  }

  // UNWATCHED section (Bottom)
  unwatchedGrid.innerHTML = '';
  const unwatchedMovies = allMovies.filter(m => {
    const p = getVideoProgress(m.title);
    const hasProgress = p.time > 10 && p.duration > 0;
    const isInUnwatched = libraryData.unwatched.some(t => normalize(t) === normalize(m.title));
    return isInUnwatched && !hasProgress;
  });

  if (unwatchedMovies.length === 0) {
    unwatchedEmpty.hidden = false;
  } else {
    unwatchedEmpty.hidden = true;
    const frag = document.createDocumentFragment();
    unwatchedMovies.forEach((m, i) => frag.appendChild(buildCard(m, i, false)));
    unwatchedGrid.appendChild(frag);
  }
}

function updateLibraryButtons() {
  document.querySelectorAll('.library-btn').forEach(btn => {
    const title = btn.dataset.title;
    const inLib = isInLibrary(title);
    btn.textContent = inLib ? '−' : '+';
    btn.title = inLib ? 'Remove from library' : 'Add to library';
  });
}

function buildCard(m, i, isRowCard) {
  const card = document.createElement('div');
  card.className = isRowCard ? 'movie-card row-card' : 'movie-card';
  card.dataset.key = normalize(m.title);
  card.style.animationDelay = Math.min(i * 30, 400) + 'ms';
  card.style.cursor = 'pointer';

  const progObj = getVideoProgress(m.title);
  const progress = progObj.time || 0;
  const duration = progObj.duration || 0;
  const progressHtml = (progress > 10 && duration > 0) ? `<div class="card-progress-bar"><div class="card-progress-fill" style="width:${Math.min(100, (progress / duration) * 100)}%"></div></div>` : '';

  const maturity = normalizeMaturity(m.maturityRating);

  card.innerHTML = `
    <div class="card-poster card-poster--playable">
      ${m.poster ? `<img src="${m.poster}" alt="${escHtml(m.title)}" loading="lazy" onload="this.classList.add('loaded')" />` : ''}
      <div class="card-play-overlay"><div class="card-play-btn"><span class="card-play-icon">&#9654;</span></div></div>
      ${progressHtml}
    </div>
    <div class="card-title">${escHtml(m.title)}</div>
    <div class="card-meta">
      <span class="card-year">${escHtml(m.year)}</span><span class="card-sep">·</span>
      <span class="card-rating ${ratingClass(maturity)}">${escHtml(maturity)}</span>
    </div>
    <div class="card-row">
      <span class="card-imdb ${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span>
      <span class="card-res ${resClass(m.resolution)}">${escHtml(m.resolution) || '—'}</span>
    </div>
    <div class="card-footer">${ratingHTML(m.title)}${libraryButtonHTML(m.title)}</div>
  `;

  card.addEventListener('click', (e) => {
    if (e.target.closest('.rating-btn')) return;
    if (e.target.closest('.library-btn')) {
      e.stopPropagation();
      toggleLibrary(m.title);
      return;
    }
    openMovieViewer(m);
  });
  return card;
}

function libraryButtonHTML(title) {
  const inLib = isInLibrary(title);
  const icon = inLib ? '−' : '+';
  return `<button class="library-btn" data-title="${escHtml(title)}" title="${inLib ? 'Remove from library' : 'Add to library'}">${icon}</button>`;
}

function ratingHTML(title) {
  const userVote = getUserRating(title), ups = getRatingCount(title, 'up'), downs = getRatingCount(title, 'down');
  return `<div class="rating-wrap">
    <button class="rating-btn rating-btn--up ${userVote === 'up' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="up" title="Liked it"><span class="rating-icon">👍</span><span class="rating-count">${ups || 0}</span></button>
    <button class="rating-btn rating-btn--down ${userVote === 'down' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="down" title="Didn't like it"><span class="rating-icon">👎</span><span class="rating-count">${downs || 0}</span></button>
  </div>`;
}
function escHtml(str) { return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ─── MOVIE VIEWER, PLAYER & COMMENTS LOGIC ────────────────────
let progressCache = {};

async function hydrateProgress() {
  try {
    const res = await fetch(`${API_BASE}/api/device/data?did=${encodeURIComponent(getDeviceId())}`);
    if (res.ok) {
      const d = await res.json();
      progressCache = d.progress || {};
    }
  } catch(e) {}
}

function getVideoProgress(title) {
  return progressCache[normalize(title)] || {};
}

// ─── LIBRARY ───────────────────────────────────────────────────
function isInLibrary(title) {
  const norm = normalize(title);
  const p = getVideoProgress(title);
  const hasProgress = p.time > 10 && p.duration > 0;
  return libraryData.watched.some(t => normalize(t) === norm) ||
         libraryData.unwatched.some(t => normalize(t) === norm) ||
         hasProgress;
}

function getLibrarySection(title) {
  const norm = normalize(title);
  if (libraryData.watched.some(t => normalize(t) === norm)) return 'watched';
  if (libraryData.unwatched.some(t => normalize(t) === norm)) return 'unwatched';
  return null;
}

async function loadLibrary() {
  try {
    const res = await fetch(`${API_BASE}/api/library?did=${encodeURIComponent(getDeviceId())}`);
    if (res.ok) {
      const data = await res.json();
      libraryData = data.library || { watched: [], unwatched: [] };
    }
  } catch(e) {}
}

async function toggleLibrary(title, watched = false) {
  const section = getLibrarySection(title);
  const p = getVideoProgress(title);
  const hasProgress = p.time > 10 && p.duration > 0;

  if (section || hasProgress) {
    try {
      if (section) {
        await fetch(`${API_BASE}/api/library/remove`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ did: getDeviceId(), title })
        });
      }
      libraryData.watched = libraryData.watched.filter(t => normalize(t) !== normalize(title));
      libraryData.unwatched = libraryData.unwatched.filter(t => normalize(t) !== normalize(title));
      
      if (hasProgress) {
        await fetch(`${API_BASE}/api/device/progress/remove`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ did: getDeviceId(), title })
        });
        delete progressCache[normalize(title)];
      }
      
      showToast('Removed from library');
    } catch(e) {}
  } else {
    try {
      await fetch(`${API_BASE}/api/library/add`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ did: getDeviceId(), title, watched: false })
      });
      libraryData.unwatched.push(title);
      showToast('Added to library');
    } catch(e) {}
  }
  renderLibrary();
  updateLibraryButtons();
  if (currentViewerMovie) updateViewerLibraryButton();
}

function updateViewerLibraryButton() {
  if (!currentViewerMovie) return;
  const btn = $('viewer-library-btn');
  const icon = btn?.querySelector('.library-icon');
  const text = $('library-btn-text');
  if (!btn || !icon || !text) return;

  const inLib = isInLibrary(currentViewerMovie.title);
  if (inLib) {
    icon.textContent = '−';
    text.textContent = 'Remove from Library';
  } else {
    icon.textContent = '+';
    text.textContent = 'Add to Library';
  }
}

let progressTimers = {};
function saveVideoProgress(title, time, duration) {
  const key = normalize(title);
  progressCache[key] = { time, duration };
  clearTimeout(progressTimers[key]);
  progressTimers[key] = setTimeout(async () => {
    try {
      await fetch(`${API_BASE}/api/device/progress`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ did: getDeviceId(), title, time, duration })
      });

      if (duration > 0 && time / duration > 0.5) {
        const section = getLibrarySection(title);
        if (section !== 'watched') {
          libraryData.unwatched = libraryData.unwatched.filter(t => normalize(t) !== key);
          if (!libraryData.watched.some(t => normalize(t) === key)) {
            libraryData.watched.push(title);
            await fetch(`${API_BASE}/api/library/add`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ did: getDeviceId(), title, watched: true })
            });
          }
        }
      }
    } catch(e) {}
  }, 4000);
}

let currentViewerMovie = null;
let currentViewerComments = [];
const viewer = $('movie-viewer');
const viewerContent = $('viewer-content');
const videoEl = $('video-el');
let progressRaf = null;

async function openMovieViewer(m) {
  currentViewerMovie = m;
  viewerContent.classList.remove('player-active');
  $('viewer-details').style.display = 'flex';
  $('viewer-player').style.display = 'none';
  viewer.style.display = 'flex';
  document.body.style.overflow = 'hidden';

  $('viewer-poster').src = m.poster || '';
  $('viewer-title').textContent = m.title;
  
  const maturity = normalizeMaturity(m.maturityRating);
  $('viewer-meta').innerHTML = `
    <span>${m.year || '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-rating ${ratingClass(maturity)}">${maturity}</span>
    <span class="card-sep">·</span>
    <span>${m.runtime || '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-imdb ${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-res ${resClass(m.resolution)}">${m.resolution || '—'}</span>
  `;
  $('viewer-plot').textContent = m.plot || 'No plot available.';
  $('viewer-cast').innerHTML = m.cast.length ? `<b>CAST:</b> ${m.cast.join(', ')}` : '';
  
  const progObj = getVideoProgress(m.title);
  const progress = progObj.time || 0;
  const duration = progObj.duration || 0;
  const progContainer = $('viewer-progress-container');
  
  if (progress > 10 && duration > 0) {
    progContainer.style.display = 'block';
    $('viewer-progress-fill').style.width = Math.min(100, (progress / duration) * 100) + '%';
    $('viewer-progress-time').textContent = `${formatTime(progress)} / ${formatTime(duration)}`;
    $('play-btn-text').textContent = 'Resume';
  } else {
    progContainer.style.display = 'none';
    $('play-btn-text').textContent = 'Play';
  }

  currentViewerComments = await fetchComments(m.title);
  renderComments();
  updateViewerLibraryButton();
}

function closeMovieViewer() {
  viewer.style.display = 'none';
  viewerContent.classList.remove('player-active');
  if (!videoEl.paused) videoEl.pause();
  videoEl.removeAttribute('src'); videoEl.load();
  document.body.style.overflow = '';
  render();
}

function playVideo() {
  $('viewer-details').style.display = 'none';
  $('viewer-player').style.display = 'flex';
  viewerContent.classList.add('player-active');
  videoEl.src = currentViewerMovie.driveLink;
  
  const startTime = getVideoProgress(currentViewerMovie.title).time || 0;
  videoEl.addEventListener('loadedmetadata', () => {
    if (startTime > 10 && startTime < videoEl.duration - 10) {
      videoEl.currentTime = startTime;
    }
    videoEl.play();
  }, { once: true });
}

function updateProgressBar() {
  if (videoEl.duration) {
    const pct = (videoEl.currentTime / videoEl.duration) * 100;
    $('ctrl-progress-played').style.width = pct + '%';
    $('ctrl-time').textContent = `${formatTime(videoEl.currentTime)} / ${formatTime(videoEl.duration)}`;
  }
  progressRaf = requestAnimationFrame(updateProgressBar);
}

videoEl.addEventListener('timeupdate', () => {
  if (!videoEl.paused) saveVideoProgress(currentViewerMovie.title, videoEl.currentTime, videoEl.duration);
});

videoEl.addEventListener('click', (e) => {
  e.stopPropagation();
  if (videoEl.paused) videoEl.play(); 
  else videoEl.pause();
});

 $('ctrl-play-pause').addEventListener('click', (e) => {
  e.stopPropagation();
  if (videoEl.paused) videoEl.play(); 
  else videoEl.pause();
});
videoEl.addEventListener('play', () => {
  $('ctrl-play-pause').innerHTML = '&#10074;&#10074;';
  if (!progressRaf) progressRaf = requestAnimationFrame(updateProgressBar);
});
videoEl.addEventListener('pause', () => {
  $('ctrl-play-pause').innerHTML = '&#9654;';
  if (progressRaf) { cancelAnimationFrame(progressRaf); progressRaf = null; }
  updateProgressBar();
});
 $('ctrl-progress-track').addEventListener('click', (e) => {
  e.stopPropagation();
  const rect = e.currentTarget.getBoundingClientRect();
  const pct = (e.clientX - rect.left) / rect.width;
  if (videoEl.duration) videoEl.currentTime = videoEl.duration * pct;
});
 $('ctrl-fullscreen').addEventListener('click', (e) => {
  e.stopPropagation();
  if (!document.fullscreenElement) $('viewer-player').requestFullscreen();
  else document.exitFullscreen();
});
 $('viewer-play-btn').addEventListener('click', playVideo);
 $('viewer-library-btn').addEventListener('click', () => {
   if (currentViewerMovie) toggleLibrary(currentViewerMovie.title);
 });
 $('viewer-close').addEventListener('click', closeMovieViewer);
 $('viewer-backdrop').addEventListener('click', closeMovieViewer);

let seekInterval = null;
let seekTimeout = null;

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && viewer.style.display === 'flex') {
    closeMovieViewer();
    return;
  }
  if (viewer.style.display !== 'flex' || $('viewer-player').style.display === 'none') return;
  const tag = e.target.tagName.toLowerCase();
  if (tag === 'input' || tag === 'textarea' || tag === 'select') return;

  if (e.code === 'Space') {
    e.preventDefault();
    if (videoEl.paused) videoEl.play(); else videoEl.pause();
  } else if (e.code === 'ArrowRight' && !e.repeat) {
    e.preventDefault();
    videoEl.currentTime = Math.min(videoEl.duration, videoEl.currentTime + 5);
    seekTimeout = setTimeout(() => {
      seekInterval = setInterval(() => {
        videoEl.currentTime = Math.min(videoEl.duration, videoEl.currentTime + 2);
      }, 100);
    }, 300);
  } else if (e.code === 'ArrowLeft' && !e.repeat) {
    e.preventDefault();
    videoEl.currentTime = Math.max(0, videoEl.currentTime - 5);
    seekTimeout = setTimeout(() => {
      seekInterval = setInterval(() => {
        videoEl.currentTime = Math.max(0, videoEl.currentTime - 2);
      }, 100);
    }, 300);
  }
});

document.addEventListener('keyup', (e) => {
  if (e.code === 'ArrowRight' || e.code === 'ArrowLeft') {
    clearTimeout(seekTimeout);
    if (seekInterval) {
      clearInterval(seekInterval);
      seekInterval = null;
    }
  }
});

async function fetchComments(title) {
  try {
    const res = await fetch(`${API_BASE}/api/comments?title=${encodeURIComponent(title)}`);
    return res.ok ? await res.json() : [];
  } catch(e) { return []; }
}

function renderComments() {
  const list = $('review-list');
  if (!list) return;
  const filter = $('review-filter').value;
  
  let displayed = currentViewerComments;
  if (filter !== 'All') displayed = displayed.filter(c => c.type === filter);
  displayed.sort((a,b) => new Date(b.time) - new Date(a.time));
  
  if (displayed.length === 0) {
    list.innerHTML = '<div class="review-empty">No reviews yet.</div>';
    return;
  }
  
  list.innerHTML = displayed.map(c => `
    <div class="review-item review-type-${c.type.toLowerCase()}">
      <div class="review-meta">
        <span>${c.type}</span>
        <span>${c.time}</span>
      </div>
      <div class="review-text">${escHtml(c.text)}</div>
    </div>
  `).join('');
}

async function submitComment() {
  const text = $('review-text').value.trim();
  const type = $('review-type').value;
  if (!text || !currentViewerMovie) return;
  
  $('review-text').value = '';
  try {
    const res = await fetch(`${API_BASE}/api/comment`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ title: currentViewerMovie.title, text, type, did: getDeviceId() })
    });
    
    if (res.ok) {
      showToast('✓ Review submitted');
      currentViewerComments = await fetchComments(currentViewerMovie.title);
      renderComments();
    }
  } catch(e) {
    showToast('⚠ Failed to submit review');
  }
}

 $('submit-review-btn').addEventListener('click', submitComment);
 $('review-filter').addEventListener('change', renderComments);

// ─── EVENTS ───────────────────────────────────────────────────
let searchTimer;
if (searchInput) {
  searchInput.addEventListener('input', () => {
    if (clearSearch) clearSearch.classList.toggle('visible', searchInput.value.length > 0);
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => { render(); saveSettings(); }, 200);
  });
}
if (clearSearch) clearSearch.addEventListener('click', () => { searchInput.value = ''; clearSearch.classList.remove('visible'); render(); searchInput.focus(); });
if (sortBy) sortBy.addEventListener('change', () => { currentSort = sortBy.value; if (hasActiveFilters()) applySort(); saveSettings(); });
if (sortDirBtn) sortDirBtn.addEventListener('click', () => { currentDir = currentDir === 'desc' ? 'asc' : 'desc'; sortDirBtn.textContent = currentDir === 'desc' ? '↓' : '↑'; if (hasActiveFilters()) applySort(); saveSettings(); });
if (sidebarClearBtn) sidebarClearBtn.addEventListener('click', clearAllFilters);

document.body.addEventListener('click', e => {
    const btn = e.target.closest('.rating-btn'); if (!btn) return;
    e.stopPropagation();
    
    const title = btn.dataset.ratingTitle, type = btn.dataset.ratingType; if (!title || !type) return;
    const key = normalize(title); if (ratingInflight.has(key)) return; ratingInflight.add(key);

    const prevVote = getUserRating(title), prevUp = getRatingCount(title, 'up'), prevDown = getRatingCount(title, 'down');
    const nextVote = prevVote === type ? null : type;
    const delta = { up: prevUp, down: prevDown };
    if (prevVote === 'up') delta.up = Math.max(0, delta.up - 1);
    if (prevVote === 'down') delta.down = Math.max(0, delta.down - 1);
    if (nextVote === 'up') delta.up += 1;
    if (nextVote === 'down') delta.down += 1;
    ratingCounts[key] = { up: delta.up, down: delta.down };
    if (nextVote) userRatings[key] = nextVote; else delete userRatings[key];
    saveUserRatings(); applyRatingDOM(title, nextVote, delta.up, delta.down, btn);
    if (nextVote === 'up') showToast('👍 Liked ' + title); else if (nextVote === 'down') showToast('👎 Disliked ' + title); else showToast('Rating removed');
    
    fetch(`${API_BASE}/api/rate`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ title, type: nextVote, prev: prevVote, did: getDeviceId() }) })
    .then(res => res.json()).then(data => { if (data.ratings) { ratingCounts[key] = data.ratings; applyRatingDOM(title, getUserRating(title), data.ratings.up, data.ratings.down, null); } })
    .catch(e => console.error('Rating failed:', e)).finally(() => ratingInflight.delete(key));
});

// ─── FOOTER FORM ──────────────────────────────────────────────
(function() {
  const submitBtn = $('footer-submit'); if (!submitBtn) return;
  submitBtn.addEventListener('click', async () => {
    const msgInput = $('footer-message'), statusEl = $('footer-form-status');
    if (!msgInput.value.trim()) { statusEl.textContent = 'Please enter a message.'; statusEl.className = 'footer-form-status error'; statusEl.hidden = false; return; }
    statusEl.textContent = '⚠ Messaging disabled.'; statusEl.className = 'footer-form-status error'; statusEl.hidden = false;
  });
})();

// ─── LOAD DATA ────────────────────────────────────────────────
async function loadData() {
  setProgress(5);
  try {
    const r = await fetch(API_BASE + '/videos?_cb=' + Date.now());
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const videos = await r.json();
    allMovies = videos.map(v => ({
      title: v.title, runtime: v.runtime || '', resolution: v.resolution || '', maturityRating: v.maturityRating || '',
      year: v.year || '—', imdbRating: v.imdbRating || '', plot: v.plot || '', cast: v.cast || [], genres: v.genres || [],
      driveLink: API_BASE + v.video, poster: v.poster ? (API_BASE + v.poster) : null,
      added_date: v.added_date || '2024-01-01'
    }));
    try { const rR = await fetch(`${API_BASE}/api/ratings`); if (rR.ok) ratingCounts = await rR.json(); } catch(e) {}
    setProgress(100); render(); populateFilterCheckboxes(); updateCounts(); updateLastUpdated();
    setTimeout(() => scanBar.classList.add('hidden'), 300);
  } catch(e) {
    console.error('Failed:', e); showToast('⚠ Could not reach backend.');
    setProgress(100); setTimeout(() => scanBar.classList.add('hidden'), 300);
  }
}

// ─── HEARTBEAT ────────────────────────────────────────────────
async function pushPresencePing() {
  try {
    const res = await fetch(`${API_BASE}/api/heartbeat?did=${getDeviceId()}`);
    const data = await res.json(); const el = $('online-count');
    if (el && data && typeof data.online === 'number') el.textContent = data.online;
  } catch(e) {}
}

// ─── SETTINGS ────────────────────────────────────────────────
async function initSettingsTab() {
  const did = getDeviceId();
  const elId   = $('settings-device-id');
  const elName = $('settings-device-name');
  const elKey  = $('settings-access-key');
  const elSeen = $('settings-first-seen');
  if (elId) elId.textContent = did;
  try {
    const res = await fetch(`${API_BASE}/api/device/data?did=${encodeURIComponent(did)}`);
    if (res.ok) {
      const d = await res.json();
      if (elName) elName.textContent = d.device_name || 'Unnamed Device';
      if (elKey)  elKey.textContent  = d.access_key || '—';
      if (elSeen) elSeen.textContent = d.first_seen || '—';
    }
  } catch(e) {
    if (elKey) elKey.textContent = '—';
  }
  const btn = $('rename-btn');
  if (btn) {
    btn.onclick = startRename;
  }
}

function startRename() {
  const nameEl = $('settings-device-name');
  const btn    = $('rename-btn');
  if (!nameEl || !btn) return;
  const wrap = nameEl.parentElement;
  const current = nameEl.textContent === '—' ? '' : nameEl.textContent.replace(/^Unnamed Device$/, '');
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'rename-input';
  input.value = current;
  input.maxLength = 64;
  input.placeholder = 'Device name';
  nameEl.replaceWith(input);
  btn.textContent = 'SAVE';
  input.focus();
  input.select();
  
  const submit = async () => {
    const name = input.value.trim();
    if (!name) { 
      input.replaceWith(nameEl); 
      btn.textContent = 'RENAME'; 
      btn.onclick = startRename; 
      return; 
    }
    try {
      const res = await fetch(`${API_BASE}/api/device/rename`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ did: getDeviceId(), name })
      });
      const d = await res.json();
      nameEl.textContent = d.name || name;
    } catch(e) { nameEl.textContent = name; }
    input.replaceWith(nameEl);
    btn.textContent = 'RENAME';
    btn.onclick = startRename; 
  };
  
  btn.onclick = submit;
  input.addEventListener('keydown', e => { 
    if (e.key === 'Enter') submit(); 
    if (e.key === 'Escape') { 
      input.replaceWith(nameEl); 
      btn.textContent = 'RENAME'; 
      btn.onclick = startRename; 
    } 
  });
}

// ─── MAIN INIT ────────────────────────────────────────────────
(async function init() {
  currentSort = 'title'; currentDir = 'asc';
  if (sortBy) sortBy.value = 'title'; if (sortDirBtn) sortDirBtn.textContent = '↓';
  try { const cR = await fetch('config.json?t=' + Date.now()); if (cR.ok) { const c = await cR.json(); API_BASE = c.API_BASE || ''; } } catch(e) { API_BASE = ''; }
  await initWithGate();
  await Promise.all([loadData(), hydrateProgress(), loadLibrary()]);
  loadShowsData();

  pushPresencePing(); setInterval(pushPresencePing, 5000);

  renderRows();
  
  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.dataset.view;
      activeTab = view;
      
      document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      const panel = $('tab-' + view);
      if (panel) panel.classList.add('active');
      
      updateCounts();
      
      if (view === 'settings') initSettingsTab(); 
      else if (view === 'home') renderRows();
      else if (view === 'movies') { if(hasActiveFilters()) applyFilters(); else { filtered = [...allMovies]; applySort(); } }
      else if (view === 'shows') renderShows();
      else if (view === 'library') loadLibrary().then(() => renderLibrary());
      else if (view === 'updates') loadUpdates(); 
    });
  });
})();