/* =============================================================
   THE DRIVE — script.js
   ============================================================= */

// ─── CONFIG ───────────────────────────────────────────────────
let API_BASE = ''; 
const SHEET_CSV_URL    = '';
const SHOWS_CSV_URL    = '';
const DRIVE_SCRIPT_URL = '';

// Auto page reload disabled — it was resetting users' scroll positions and view state.
// const AUTO_RELOAD_MS = 30 * 60 * 1000;
// setTimeout(() => location.reload(), AUTO_RELOAD_MS);

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
  let s = String(str).toUpperCase().trim();
  if (/\bNC[-\s]?17\b/.test(s)) return 'NC-17';
  if (/\bPG[-\s]?13\b/.test(s)) return 'PG-13';
  if (/\bPG\b/.test(s)) return 'PG';
  if (/\bR\b/.test(s)) return 'R'; 
  if (/\bG\b/.test(s)) return 'G';
  return 'NR';
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

// Request images at 2x their display dimensions for sharpness on retina screens.
// The backend resizes to fit within w×h (maintaining aspect ratio).
function sizedImg(url, cssW, cssH) {
  if (!url) return url;
  const w = Math.round(cssW * 2);
  const h = Math.round(cssH * 2);
  const sep = url.includes('?') ? '&' : '?';
  return `${url}${sep}w=${w}&h=${h}`;
}

// ─── LAZY IMAGE LOADING ───────────────────────────────────────
// Only download poster images when the card is near the viewport. This avoids
// downloading hundreds of high-res posters the user will never see.
let _lazyObserver = null;
function initLazyImages() {
  if (_lazyObserver) return;
  if (!('IntersectionObserver' in window)) {
    // Fallback: just load everything immediately
    document.querySelectorAll('img[data-src]').forEach(img => { img.src = img.dataset.src; img.classList.add('loaded'); });
    return;
  }
  _lazyObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const img = entry.target;
        if (img.dataset.src) {
          img.src = img.dataset.src;
          img.addEventListener('load', () => img.classList.add('loaded'), { once: true });
          img.addEventListener('error', () => img.classList.add('loaded'), { once: true });
          delete img.dataset.src;
        }
        _lazyObserver.unobserve(img);
      }
    });
  }, { rootMargin: '300px 0px', threshold: 0.01 });
}
function observeLazyImages() {
  if (!_lazyObserver) initLazyImages();
  if (!_lazyObserver) return;
  document.querySelectorAll('img[data-src]:not([data-watched])').forEach(img => {
    img.setAttribute('data-watched', '1');
    _lazyObserver.observe(img);
  });
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
      driveLink: API_BASE + v.video, poster: v.poster ? (API_BASE + v.poster) : null,
      fanart: v.fanart ? (API_BASE + v.fanart) : null,
      clearlogo: v.clearlogo ? (API_BASE + v.clearlogo) : null,
      trailer: v.trailer ? (API_BASE + v.trailer) : null
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
  observeLazyImages();
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
        <div class="movie-row-scroll snap">
          <div id="${id}" class="movie-row"></div>
        </div>
      </div>
    </section>
  `;
}

// ─── HERO (Apple TV superhero lockup) ─────────────────────────
let heroFilms = [];
let heroIndex = 0;
let heroTimer = null;

function pickHeroFilms() {
  // Top-rated by USER ratings (thumbs up minus thumbs down), then recently added
  const byUserRating = [...allMovies]
    .sort((a, b) => getRatingScore(b.title) - getRatingScore(a.title));
  const byRecent = [...allMovies].sort((a, b) => new Date(b.added_date || 0) - new Date(a.added_date || 0));
  const seen = new Set();
  const picks = [];
  // First pass: top-user-rated films that have fanart
  for (const m of byUserRating) { const k = normalize(m.title); if (!seen.has(k) && m.fanart && picks.length < 4) { seen.add(k); picks.push(m); } }
  // Second pass: recently-added films that have fanart (for variety)
  for (const m of byRecent) { const k = normalize(m.title); if (!seen.has(k) && m.fanart && picks.length < 6) { seen.add(k); picks.push(m); } }
  // Fallback: films with at least a poster if we don't have enough fanart films
  for (const m of byUserRating) { const k = normalize(m.title); if (!seen.has(k) && m.poster && picks.length < 6) { seen.add(k); picks.push(m); } }
  return picks.filter(m => m.fanart || m.poster).slice(0, 6);
}

function renderHero() {
  const hero = $('hero');
  if (!hero) return;
  heroFilms = pickHeroFilms();
  if (heroFilms.length === 0) { hero.style.display = 'none'; return; }
  hero.style.display = '';
  hero.setAttribute('aria-hidden', 'false');
  heroIndex = 0;
  // dots
  const dotsEl = $('hero-dots');
  if (dotsEl) dotsEl.innerHTML = heroFilms.map((_, i) => `<button class="hero__dot${i === 0 ? ' active' : ''}" data-i="${i}" aria-label="Slide ${i+1}"></button>`).join('');
  showHeroSlide(0);
  startHeroRotation();
}

// Track which backdrop layer is currently visible for the crossfade
let _heroActiveLayer = 'a';

function showHeroSlide(i) {
  if (!heroFilms.length) return;
  heroIndex = (i + heroFilms.length) % heroFilms.length;
  const m = heroFilms[heroIndex];
  if (!m) return;

  // Two-layer crossfade: load new image into the inactive layer, then swap
  const bgImg = m.fanart ? sizedImg(m.fanart, 1400, 700) : (m.poster ? sizedImg(m.poster, 1400, 700) : null);
  const activeCls = _heroActiveLayer === 'a' ? 'hero__backdrop--a' : 'hero__backdrop--b';
  const inactiveCls = _heroActiveLayer === 'a' ? 'hero__backdrop--b' : 'hero__backdrop--a';
  const activeLayer = hero.querySelector('.' + activeCls);
  const inactiveLayer = hero.querySelector('.' + inactiveCls);

  if (bgImg && inactiveLayer) {
    // Preload the image, then crossfade once it's ready
    const img = new Image();
    img.onload = () => {
      inactiveLayer.style.backgroundImage = `url("${bgImg}")`;
      // Start the glide on the new layer
      inactiveLayer.classList.add('active');
      inactiveLayer.classList.remove('leaving');
      // Fade out the old layer
      if (activeLayer) {
        activeLayer.classList.remove('active');
        activeLayer.classList.add('leaving');
      }
      // Swap active layer for next time
      _heroActiveLayer = _heroActiveLayer === 'a' ? 'b' : 'a';
      // Clean up the leaving class after the fade completes
      setTimeout(() => { if (activeLayer) activeLayer.classList.remove('leaving'); }, 1300);
    };
    img.onerror = () => {
      // No image — just clear and swap
      if (activeLayer) { activeLayer.classList.remove('active'); activeLayer.classList.add('leaving'); }
      inactiveLayer.style.backgroundImage = '';
      inactiveLayer.classList.add('active');
      _heroActiveLayer = _heroActiveLayer === 'a' ? 'b' : 'a';
    };
    img.src = bgImg;
  }

  // Title: prefer the clearlogo image, fall back to text
  const titleEl = $('hero-title');
  if (titleEl) {
    titleEl.innerHTML = '';
    if (m.clearlogo) {
      const logoImg = document.createElement('img');
      logoImg.className = 'hero__title-logo';
      logoImg.alt = m.title;
      logoImg.src = sizedImg(m.clearlogo, 420, 150);
      logoImg.onload = () => logoImg.classList.add('loaded');
      logoImg.onerror = () => { titleEl.textContent = m.title; };
      titleEl.appendChild(logoImg);
    } else {
      titleEl.textContent = m.title;
    }
  }

  const maturity = normalizeMaturity(m.maturityRating);
  const metaEl = $('hero-meta');
  if (metaEl) metaEl.innerHTML = `
    <span>${m.year || '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-rating ${ratingClass(maturity)}">${maturity}</span>
    <span class="card-sep">·</span>
    <span>${m.runtime || '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-imdb">${m.imdbRating ? '★ ' + m.imdbRating : ''}</span>
  `;

  const plotEl = $('hero-plot'); if (plotEl) plotEl.textContent = m.plot || '';

  const badgesEl = $('hero-badges');
  if (badgesEl) {
    const badges = [];
    if (m.resolution && /4k|2160/i.test(m.resolution)) badges.push('<span class="badge badge-gold">4K</span>');
    else if (m.resolution) badges.push(`<span class="badge">${escHtml(m.resolution)}</span>`);
    if (m.imdbRating && parseFloat(m.imdbRating) >= 8) badges.push('<span class="badge badge-accent">Top Rated</span>');
    if (m.genres && m.genres[0]) badges.push(`<span class="badge">${escHtml(m.genres[0])}</span>`);
    badgesEl.innerHTML = badges.join('');
  }

  const playText = $('hero-play-text');
  const progObj = getVideoProgress(m.title);
  if (playText) playText.textContent = (progObj.time > 10 && progObj.duration > 0) ? 'Resume' : 'Play';

  document.querySelectorAll('.hero__dot').forEach((d, idx) => d.classList.toggle('active', idx === heroIndex));
}

function startHeroRotation() {
  if (heroTimer) clearInterval(heroTimer);
  if (heroFilms.length <= 1) return;
  heroTimer = setInterval(() => showHeroSlide(heroIndex + 1), 8000);
}

// Hero interactions (bound once)
(function initHeroEvents() {
  document.addEventListener('click', (e) => {
    const dot = e.target.closest('.hero__dot');
    if (dot) { showHeroSlide(parseInt(dot.dataset.i, 10)); startHeroRotation(); return; }
    if (e.target.closest('#hero-play-btn')) {
      if (heroFilms[heroIndex]) openMovieViewer(heroFilms[heroIndex]);
      return;
    }
    if (e.target.closest('#hero-info-btn')) {
      if (heroFilms[heroIndex]) openMovieViewer(heroFilms[heroIndex]);
      return;
    }
  });
})();

function renderRows() {
  if (!rowView) return;
  renderHero();
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
  const scroller = container.closest('.movie-row-scroll');
  if (scroller) {
    // Update button visibility now and after layout settles (images loading
    // changes the row width, which affects whether the arrows should show).
    updateRowScrollBtns(scroller);
    requestAnimationFrame(() => updateRowScrollBtns(scroller));
    setTimeout(() => updateRowScrollBtns(scroller), 200);
    setTimeout(() => updateRowScrollBtns(scroller), 600);
    // When each poster image finishes loading, the row width changes — re-check.
    scroller.querySelectorAll('img').forEach(img => {
      if (img.complete) return;
      img.addEventListener('load', () => updateRowScrollBtns(scroller), { once: true });
      img.addEventListener('error', () => updateRowScrollBtns(scroller), { once: true });
    });
  }
  observeLazyImages();
}

(function initRowScrollBtns() {
  // Click → scroll by ~3 card widths (based on the actual first card size)
  document.addEventListener('click', function(e) {
    const btn = e.target.closest('.row-scroll-btn'); if (!btn) return;
    const track = document.getElementById(btn.dataset.target); if (!track) return;
    const scroller = track.closest('.movie-row-scroll'); if (!scroller) return;
    const firstCard = scroller.querySelector('.row-card');
    const cardW = firstCard ? (firstCard.offsetWidth + 14) : 200; // card + gap
    scroller.scrollBy({ left: parseInt(btn.dataset.dir, 10) * cardW * 3, behavior: 'smooth' });
  });
  // Update buttons on scroll (capture so we catch the scroll event on the scroller)
  document.addEventListener('scroll', (e) => {
    if (e.target.classList && e.target.classList.contains('movie-row-scroll')) updateRowScrollBtns(e.target);
  }, true);
  // Update all rows on resize — a row can become scrollable/unscrollable
  let resizeTimer;
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      document.querySelectorAll('.movie-row-scroll').forEach(updateRowScrollBtns);
    }, 100);
  });
})();

function updateRowScrollBtns(scroller) {
  const wrapper = scroller.closest('.row-scroll-wrapper'); if (!wrapper) return;
  const lBtn = wrapper.querySelector('.row-scroll-btn--left');
  const rBtn = wrapper.querySelector('.row-scroll-btn--right');
  const maxScroll = scroller.scrollWidth - scroller.clientWidth;
  // Only show arrows when the row actually overflows
  const canScroll = maxScroll > 4;
  // Tolerance accounts for the scroller's left/right padding (e.g. 48px) so
  // the button doesn't show when the row is resting at its padded start, and
  // the right button hides when we're within padding of the end.
  const pad = parseInt(getComputedStyle(scroller).paddingLeft || '0', 10) || 0;
  const tol = Math.max(pad, 4) + 4;
  if (lBtn) lBtn.dataset.hidden = (!canScroll || scroller.scrollLeft <= tol) ? '1' : '0';
  if (rBtn) rBtn.dataset.hidden = (!canScroll || scroller.scrollLeft >= maxScroll - tol) ? '1' : '0';
}
function renderGrid() {
  if (!movieGrid) return; movieGrid.innerHTML = '';
  if (filtered.length === 0) { if (gridEmpty) gridEmpty.hidden = false; return; }
  if (gridEmpty) gridEmpty.hidden = true;
  const frag = document.createDocumentFragment();
  filtered.forEach((m, i) => frag.appendChild(buildCard(m, i, false)));
  movieGrid.appendChild(frag);
  observeLazyImages();
}

function renderLibrary() {
  const watchedGrid = $('library-watched-grid');
  const unwatchedGrid = $('library-unwatched-grid');
  const watchedEmpty = $('library-watched-empty');
  const unwatchedEmpty = $('library-unwatched-empty');

  if (!watchedGrid || !unwatchedGrid) return;

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
  observeLazyImages();
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
      ${m.poster ? `<img data-src="${sizedImg(m.poster, 168, 252)}" alt="${escHtml(m.title)}" />` : ''}
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
let _applying_party_action = false; // Flag to prevent WS echo loops
let _playing_trailer = false; // True while the trailer (not the movie) is loaded

// ─── BUFFERING SPINNER ──────────────────────────────────────
// Shows a spinner any time the video is not immediately playable: initial
// load (waiting for canplay), and any mid-playback stall/rebuffer (waiting).
// Hides on playing/canplay/canplaythrough, and always on pause/error/ended
// so it never gets stuck visible.
const bufferingSpinnerEl = $('viewer-buffering-spinner');
function showBufferingSpinner() { if (bufferingSpinnerEl) bufferingSpinnerEl.hidden = false; }
function hideBufferingSpinner() { if (bufferingSpinnerEl) bufferingSpinnerEl.hidden = true; }
videoEl.addEventListener('waiting', showBufferingSpinner);      // ran out of buffer mid-playback
videoEl.addEventListener('loadstart', showBufferingSpinner);    // new src just set
videoEl.addEventListener('playing', hideBufferingSpinner);      // actually resumed/started producing frames
videoEl.addEventListener('canplaythrough', hideBufferingSpinner);
videoEl.addEventListener('pause', hideBufferingSpinner);
videoEl.addEventListener('ended', hideBufferingSpinner);
videoEl.addEventListener('error', hideBufferingSpinner);

// ─── DISABLE EMBEDDED SUBTITLES ─────────────────────────────
// Some video files have a subtitle/caption track muxed into the container
// itself. We never add a <track> element or turn subtitles on ourselves,
// but certain browsers (Safari in particular) auto-enable an embedded text
// track by default. Force every text track off whenever the track list
// changes so subtitles never show up uninvited.
function disableAllTextTracks() {
  const tracks = videoEl.textTracks;
  if (!tracks) return;
  for (let i = 0; i < tracks.length; i++) tracks[i].mode = 'disabled';
}
videoEl.textTracks.addEventListener('addtrack', disableAllTextTracks);
videoEl.addEventListener('loadedmetadata', disableAllTextTracks);

// ─── PREFETCH (warm the browser cache before the user clicks Play) ──
// When the detail view opens, we start fetching the movie (at the resume
// position) and the trailer into hidden video elements. The browser caches
// those bytes, so when the user clicks Play/Trailer the main video element
// gets the data instantly from cache → near-zero buffering time.
let _moviePrefetchEl = null;
let _trailerPrefetchEl = null;

function prefetchVideo(url) {
  if (!url) return null;
  const v = document.createElement('video');
  // IMPORTANT: 'auto' tells the browser to eagerly download as much of the
  // file as it can in the background. On a fast connection that's a nice
  // head start; on a slow/constrained connection (1-2MB/s) it competes for
  // real bandwidth with actual playback and gets counted server-side as a
  // second "active stream", cutting the real stream's bandwidth share.
  // 'metadata' only fetches enough to get duration/dimensions — near-zero
  // bandwidth cost, no competing stream.
  v.preload = 'metadata';
  v.muted = true;
  v.src = url;
  // Some browsers won't fetch without the element being in the DOM
  v.style.cssText = 'position:absolute;width:1px;height:1px;opacity:0;pointer-events:none;';
  document.body.appendChild(v);
  return v;
}

function startPrefetch(m) {
  // Prefetch the movie + trailer using the EXACT same URLs that playVideo/
  // playTrailer will use, so the browser's HTTP cache applies when the user
  // clicks Play. We don't use #t= fragments because playVideo doesn't — using
  // a different URL would bypass the cache.
  stopPrefetch();
  _moviePrefetchEl = prefetchVideo(m.driveLink);
  if (m.trailer) {
    _trailerPrefetchEl = prefetchVideo(m.trailer);
  }
}

function stopPrefetch() {
  if (_moviePrefetchEl) { _moviePrefetchEl.removeAttribute('src'); _moviePrefetchEl.load(); _moviePrefetchEl.remove(); _moviePrefetchEl = null; }
  if (_trailerPrefetchEl) { _trailerPrefetchEl.removeAttribute('src'); _trailerPrefetchEl.load(); _trailerPrefetchEl.remove(); _trailerPrefetchEl = null; }
}

async function openMovieViewer(m, fromParty = false) {
  currentViewerMovie = m;
  viewerContent.classList.remove('player-active');
  $('viewer-details').style.display = 'flex';
  $('viewer-player').style.display = 'none';
  viewer.style.display = 'flex';
  document.body.style.overflow = 'hidden';

  const vPoster = $('viewer-poster'); if (vPoster) vPoster.src = m.poster ? sizedImg(m.poster, 200, 300) : '';

  // Apple-TV detail hero: prefer fanart for the backdrop, fall back to poster
  const heroBackdrop = $('viewer-hero-backdrop');
  const detailBg = m.fanart ? sizedImg(m.fanart, 1400, 600) : (m.poster ? sizedImg(m.poster, 1400, 600) : null);
  if (heroBackdrop) {
    heroBackdrop.classList.remove('loaded');
    if (detailBg) {
      heroBackdrop.style.backgroundImage = `url("${detailBg}")`;
      const img = new Image();
      img.onload = () => heroBackdrop.classList.add('loaded');
      img.src = detailBg;
    } else {
      heroBackdrop.style.backgroundImage = '';
    }
  }
  const vBadges = $('viewer-badges');
  if (vBadges) {
    const badges = [];
    if (m.resolution && /4k|2160/i.test(m.resolution)) badges.push('<span class="badge badge-gold">4K</span>');
    else if (m.resolution) badges.push(`<span class="badge">${escHtml(m.resolution)}</span>`);
    if (m.imdbRating && parseFloat(m.imdbRating) >= 8) badges.push('<span class="badge badge-accent">Top Rated</span>');
    if (m.genres && m.genres[0]) badges.push(`<span class="badge">${escHtml(m.genres[0])}</span>`);
    vBadges.innerHTML = badges.join('');
  }

  // Title: prefer the clearlogo image, fall back to text
  const vTitle = $('viewer-title');
  if (vTitle) {
    vTitle.innerHTML = '';
    if (m.clearlogo) {
      const logoImg = document.createElement('img');
      logoImg.className = 'viewer-hero__title-logo';
      logoImg.alt = m.title;
      logoImg.src = sizedImg(m.clearlogo, 380, 130);
      // Fade in once decoded — avoids the jarring top-to-bottom progressive load
      logoImg.onload = () => logoImg.classList.add('loaded');
      logoImg.onerror = () => { vTitle.textContent = m.title; };
      vTitle.appendChild(logoImg);
    } else {
      vTitle.textContent = m.title;
    }
  }

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
  // Trailer button: "Trailer" (playable) if a trailer exists, else "Request Trailer"
  const trailerBtn = $('viewer-trailer-btn');
  const trailerBtnText = $('viewer-trailer-btn-text');
  const trailerProgress = $('viewer-trailer-progress');
  if (trailerBtn) {
    trailerBtn.style.display = ''; // always show now
    trailerBtn.dataset.mode = m.trailer ? 'play' : 'request';
    if (trailerBtnText) trailerBtnText.textContent = m.trailer ? 'Trailer' : 'Request Trailer';
    if (trailerProgress) trailerProgress.classList.remove('active');
    // Check if there's an in-progress download for this movie
    if (!m.trailer) checkTrailerDownloadStatus(m.title);
  }
  renderViewerInfo(m);
  renderViewerRelated(m);

  // Prefetch the movie (at resume position) and trailer into the browser cache
  // so playback starts near-instantly when the user clicks Play/Trailer.
  startPrefetch(m);

  if (partyWS && !fromParty) {
    sendParty({ type: 'load_video', movie: m });
  }
}

// Metadata info strip below the hero — fills the page out with real details
function renderViewerInfo(m) {
  const info = $('viewer-info');
  if (!info) return;
  const fields = [];
  if (m.director) fields.push({ label: 'Director', value: m.director });
  if (m.cast && m.cast.length) fields.push({ label: 'Starring', value: m.cast.slice(0, 6).join(', ') });
  if (m.genres && m.genres.length) fields.push({ label: 'Genres', chips: m.genres });
  fields.push({ label: 'Released', value: m.year || '—' });
  fields.push({ label: 'Runtime', value: m.runtime || '—' });
  if (m.resolution) fields.push({ label: 'Quality', value: m.resolution });
  if (m.imdbRating) fields.push({ label: 'IMDb Rating', value: '★ ' + m.imdbRating });
  info.innerHTML = fields.map(f => {
    if (f.chips) {
      return `<div class="viewer-info__field">
        <span class="viewer-info__label">${escHtml(f.label)}</span>
        <div class="viewer-info__chips">${f.chips.map(c => `<span class="viewer-info__chip">${escHtml(c)}</span>`).join('')}</div>
      </div>`;
    }
    return `<div class="viewer-info__field">
      <span class="viewer-info__label">${escHtml(f.label)}</span>
      <span class="viewer-info__value">${escHtml(f.value)}</span>
    </div>`;
  }).join('');
}

// "More Like This" — same-genre films, excluding the current one
function renderViewerRelated(m) {
  const section = $('viewer-related');
  const scroll = $('viewer-related-scroll');
  if (!section || !scroll) return;
  const norm = normalize(m.title);
  const related = allMovies.filter(x => normalize(x.title) !== norm && x.genres && x.genres.some(g => m.genres && m.genres.includes(g)))
    .sort((a, b) => getRatingScore(b.title) - getRatingScore(a.title))
    .slice(0, 14);
  if (related.length === 0) { section.style.display = 'none'; scroll.innerHTML = ''; return; }
  section.style.display = '';
  scroll.innerHTML = '';
  const frag = document.createDocumentFragment();
  related.forEach((movie, i) => frag.appendChild(buildCard(movie, i, false)));
  scroll.appendChild(frag);
  observeLazyImages();
}

function closeMovieViewer() {
  viewer.style.display = 'none';
  viewerContent.classList.remove('player-active');
  _playing_trailer = false;
  if (!videoEl.paused) videoEl.pause();
  videoEl.removeAttribute('src'); videoEl.load();
  stopPrefetch(); // clean up prefetch elements + free their cache pressure
  document.body.style.overflow = '';
  // Don't re-render — preserve the user's scroll position and view state.
  // Targeted updates (ratings, library buttons) already happened in real-time.
  updateLibraryButtons();
}

// Close just the trailer/player and return to the info card (with a fade-out)
function closeTrailer() {
  if (!_playing_trailer) return;
  _playing_trailer = false;
  videoEl.pause();
  const player = $('viewer-player');
  if (player) player.style.opacity = '0';
  setTimeout(() => {
    videoEl.removeAttribute('src'); videoEl.load();
    viewerContent.classList.remove('player-active');
    $('viewer-details').style.display = 'flex';
    $('viewer-player').style.display = 'none';
    if (player) player.style.opacity = '';
  }, 400);
}

// Reset the video element + progress UI when switching videos, so stale data
// from the previous video (duration, currentTime, buffered ranges) doesn't
// leak into the new one.
function resetVideoForNewSrc() {
  // Fully clear the old video so the browser doesn't reuse its state
  videoEl.pause();
  videoEl.removeAttribute('src');
  videoEl.load();
  // Reset all progress UI to zeros
  const played = $('ctrl-progress-played'); if (played) played.style.width = '0%';
  const cb = $('ctrl-progress-buffered'); if (cb) cb.style.width = '0%';
  const vb = $('viewer-progress-buffered'); if (vb) vb.style.width = '0%';
  const time = $('ctrl-time'); if (time) time.textContent = '0:00 / 0:00';
}

function playVideo() {
  _playing_trailer = false;
  $('viewer-details').style.display = 'none';
  $('viewer-player').style.display = 'flex';
  viewerContent.classList.add('player-active');
  // Stop the movie prefetch — the main video element takes over. Keep the
  // trailer prefetch running in case the user watches the trailer next.
  if (_moviePrefetchEl) { _moviePrefetchEl.removeAttribute('src'); _moviePrefetchEl.load(); _moviePrefetchEl.remove(); _moviePrefetchEl = null; }
  resetVideoForNewSrc();
  videoEl.preload = 'auto';
  videoEl.src = currentViewerMovie.driveLink;

  const startTime = getVideoProgress(currentViewerMovie.title).time || 0;
  // Start as soon as the very first frame is available (loadeddata) rather
  // than waiting for canplay, which holds off until the browser estimates
  // it has "enough" buffered to play through smoothly. We want playback to
  // begin immediately on whatever data has arrived, buffering handles the rest.
  const startPlayback = () => {
    if (startTime > 10 && startTime < videoEl.duration - 10) {
      try { videoEl.currentTime = startTime; } catch(e) {}
    }
    videoEl.play().catch(() => {});
  };
  videoEl.addEventListener('loadeddata', startPlayback, { once: true });
}

function playTrailer() {
  if (!currentViewerMovie || !currentViewerMovie.trailer) return;
  _playing_trailer = true;
  $('viewer-details').style.display = 'none';
  $('viewer-player').style.display = 'flex';
  viewerContent.classList.add('player-active');
  // Stop the trailer prefetch — the main video element takes over.
  if (_trailerPrefetchEl) { _trailerPrefetchEl.removeAttribute('src'); _trailerPrefetchEl.load(); _trailerPrefetchEl.remove(); _trailerPrefetchEl = null; }
  resetVideoForNewSrc();
  videoEl.preload = 'auto';
  videoEl.src = currentViewerMovie.trailer;
  // Trailers start from the beginning — no progress restore
  const startPlayback = () => { videoEl.play().catch(() => {}); };
  videoEl.addEventListener('loadeddata', startPlayback, { once: true });
}

// ─── TRAILER DOWNLOAD (Request Trailer) ───────────────────────
let _trailerPollTimer = null;

async function requestTrailer() {
  if (!currentViewerMovie) return;
  const m = currentViewerMovie;
  const btn = $('viewer-trailer-btn');
  const btnText = $('viewer-trailer-btn-text');
  const progressEl = $('viewer-trailer-progress');
  if (btnText) btnText.textContent = 'Searching...';
  btn.disabled = true;
  if (progressEl) { progressEl.classList.add('active'); progressEl.style.setProperty('--trailer-progress', '0%'); }
  try {
    const res = await fetch(`${API_BASE}/api/trailer/request`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ did: getDeviceId(), title: m.title, year: m.year })
    });
    const data = await res.json();
    if (!data.started) {
      if (btnText) btnText.textContent = data.error || 'Failed';
      btn.disabled = false;
      if (progressEl) progressEl.classList.remove('active');
      setTimeout(() => { if (btnText) btnText.textContent = 'Request Trailer'; }, 3000);
      return;
    }
    // Start polling for progress
    pollTrailerDownload(m.title);
  } catch(e) {
    if (btnText) btnText.textContent = 'Failed';
    btn.disabled = false;
    if (progressEl) progressEl.classList.remove('active');
    setTimeout(() => { if (btnText) btnText.textContent = 'Request Trailer'; }, 3000);
  }
}

async function checkTrailerDownloadStatus(title) {
  // If a download is already in progress for this movie, resume polling
  try {
    const res = await fetch(`${API_BASE}/api/trailer/status?title=${encodeURIComponent(title)}`);
    const data = await res.json();
    if (data.status === 'downloading' || data.status === 'searching') {
      pollTrailerDownload(title);
    } else if (data.status === 'done') {
      // Trailer was downloaded since the page loaded — update the movie + button
      finishTrailerDownload(title, data.trailer_url);
    }
  } catch(e) {}
}

function pollTrailerDownload(title) {
  if (_trailerPollTimer) clearInterval(_trailerPollTimer);
  const btn = $('viewer-trailer-btn');
  const btnText = $('viewer-trailer-btn-text');
  const progressEl = $('viewer-trailer-progress');
  if (progressEl) progressEl.classList.add('active');
  _trailerPollTimer = setInterval(async () => {
    try {
      const res = await fetch(`${API_BASE}/api/trailer/status?title=${encodeURIComponent(title)}`);
      const data = await res.json();
      if (data.status === 'searching') {
        if (btnText) btnText.textContent = 'Searching...';
        if (progressEl) progressEl.style.setProperty('--trailer-progress', '5%');
      } else if (data.status === 'downloading') {
        if (btnText) btnText.textContent = 'Downloading...';
        if (progressEl) progressEl.style.setProperty('--trailer-progress', (data.progress || 0) + '%');
      } else if (data.status === 'done') {
        clearInterval(_trailerPollTimer); _trailerPollTimer = null;
        finishTrailerDownload(title, data.trailer_url);
      } else if (data.status === 'error') {
        clearInterval(_trailerPollTimer); _trailerPollTimer = null;
        if (btnText) btnText.textContent = 'Failed';
        if (progressEl) progressEl.classList.remove('active');
        if (btn) btn.disabled = false;
        setTimeout(() => { if (btnText) btnText.textContent = 'Request Trailer'; }, 3000);
      }
    } catch(e) {}
  }, 1500);
}

function finishTrailerDownload(title, trailerUrl) {
  const btn = $('viewer-trailer-btn');
  const btnText = $('viewer-trailer-btn-text');
  const progressEl = $('viewer-trailer-progress');
  if (btnText) btnText.textContent = 'Trailer';
  if (btn) { btn.disabled = false; btn.dataset.mode = 'play'; }
  if (progressEl) { progressEl.classList.remove('active'); progressEl.style.setProperty('--trailer-progress', '100%'); }
  // Update the movie object so playTrailer() works
  if (currentViewerMovie && currentViewerMovie.title === title && trailerUrl) {
    currentViewerMovie.trailer = trailerUrl;
  }
  // Also update allMovies so re-opening the viewer has the trailer
  const movie = allMovies.find(m => m.title === title);
  if (movie && trailerUrl) movie.trailer = trailerUrl;
  showToast('✓ Trailer downloaded');
  setTimeout(() => { if (progressEl) progressEl.style.setProperty('--trailer-progress', '0%'); }, 1000);
}

function updateProgressBar() {
  // Only update when we have real metadata for the CURRENT video (readyState > 0
  // means metadata is loaded). This prevents stale duration/currentTime from the
  // previous video leaking into the progress bar while the new one loads.
  if (videoEl.readyState > 0 && videoEl.duration && isFinite(videoEl.duration)) {
    const pct = (videoEl.currentTime / videoEl.duration) * 100;
    $('ctrl-progress-played').style.width = pct + '%';
    $('ctrl-time').textContent = `${formatTime(videoEl.currentTime)} / ${formatTime(videoEl.duration)}`;
    updateBufferedBar();
  }
  progressRaf = requestAnimationFrame(updateProgressBar);
}

// YouTube-style buffered bar: show how much of the video is downloaded ahead
// of the playhead. Uses the video element's buffered TimeRanges.
function updateBufferedBar() {
  if (!videoEl.duration) return;
  const bufferedEl = $('ctrl-progress-buffered');
  const vBufferedEl = $('viewer-progress-buffered');
  // Find the buffered range that contains the playhead, and report its end.
  // If the playhead is in a gap (unlikely but possible after a seek), report 0
  // so the bar accurately reflects "nothing buffered ahead right now."
  let bufEnd = -1;
  const ct = videoEl.currentTime;
  for (let i = 0; i < videoEl.buffered.length; i++) {
    const start = videoEl.buffered.start(i);
    const end = videoEl.buffered.end(i);
    if (ct >= start - 0.5 && ct <= end + 0.5) {
      bufEnd = end;
      break;
    }
  }
  // If no range contains the playhead, show the bar at the playhead position
  // (flat) rather than jumping to some unrelated range.
  const effectiveEnd = bufEnd >= 0 ? bufEnd : ct;
  const bufPct = Math.min(100, Math.max(0, (effectiveEnd / videoEl.duration) * 100));
  if (bufferedEl) bufferedEl.style.width = bufPct + '%';
  if (vBufferedEl) vBufferedEl.style.width = bufPct + '%';
}

videoEl.addEventListener('timeupdate', () => {
  if (_playing_trailer) return; // don't save progress while playing a trailer
  if (!videoEl.paused) saveVideoProgress(currentViewerMovie.title, videoEl.currentTime, videoEl.duration);
});

// Browser downloaded more media — refresh the buffered bar immediately
videoEl.addEventListener('progress', updateBufferedBar);

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
  if (partyWS && !_applying_party_action) sendParty({ type: 'play' });
});
videoEl.addEventListener('pause', () => {
  $('ctrl-play-pause').innerHTML = '&#9654;';
  if (progressRaf) { cancelAnimationFrame(progressRaf); progressRaf = null; }
  updateProgressBar();
  if (partyWS && !_applying_party_action) sendParty({ type: 'pause' });
});
videoEl.addEventListener('seeked', () => {
  if (partyWS && !_applying_party_action) sendParty({ type: 'seek', time: videoEl.currentTime });
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
 $('viewer-trailer-btn').addEventListener('click', () => {
   const btn = $('viewer-trailer-btn');
   if (btn && btn.dataset.mode === 'play') {
     playTrailer();
   } else {
     requestTrailer();
   }
 });
 // Close button: if a trailer is playing, go back to info card; else close everything
 $('viewer-close').addEventListener('click', () => {
  if (_playing_trailer) closeTrailer();
  else closeMovieViewer();
 });
 // Backdrop click: same logic
 $('viewer-backdrop').addEventListener('click', () => {
  if (_playing_trailer) closeTrailer();
  else closeMovieViewer();
 });
 // Trailer ended → fade out and return to info card
 videoEl.addEventListener('ended', () => {
  if (_playing_trailer) closeTrailer();
 });

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
      director: v.director || '',
      driveLink: API_BASE + v.video, poster: v.poster ? (API_BASE + v.poster) : null,
      fanart: v.fanart ? (API_BASE + v.fanart) : null,
      clearlogo: v.clearlogo ? (API_BASE + v.clearlogo) : null,
      trailer: v.trailer ? (API_BASE + v.trailer) : null,
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
      cachedDeviceData = d; // share with watch-party logic
      const unnamed = isUnnamedDeviceName(d.device_name);
      if (elName) elName.textContent = unnamed ? 'Unnamed Device' : d.device_name;
      if (elKey)  elKey.textContent  = d.access_key || '—';
      if (elSeen) elSeen.textContent = d.first_seen || '—';

      // Automatically prompt for a name (once per session) if the device is
      // still "Unnamed Device" — the name shows beside the cursor in Watch
      // Parties, so we want every device named before joining.
      if (unnamed && !_proactiveNamePromptShown) {
        _proactiveNamePromptShown = true;
        const newName = await ensureDeviceName();
        if (newName && elName) elName.textContent = newName;
      }
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
  const current = (nameEl.textContent === '—' || isUnnamedDeviceName(nameEl.textContent)) ? '' : nameEl.textContent;
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
    // saveDeviceName updates the DOM, cachedDeviceData, and notifies the party
    const saved = await saveDeviceName(name);
    if (!saved) nameEl.textContent = name; // fallback on failure
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

// ─── WATCH PARTY LOGIC ───────────────────────────────────────
let partyWS = null;
let userColor = null;
let cursorTimers = {};
let partyPingInterval = null;

// Tracks whether we've already proactively prompted for a name this session,
// so opening the Settings tab repeatedly doesn't keep nagging the user.
let _proactiveNamePromptShown = false;

// Returns true if the given device name should be treated as "unnamed".
// Case-insensitive to catch "Unnamed device", "Unnamed Device", "unnamed", etc.
function isUnnamedDeviceName(name) {
  if (name === null || name === undefined) return true;
  const n = String(name).trim().toLowerCase();
  return n === '' || n === 'unnamed device' || n === 'unnamed' || n === 'unknown' || n === 'null' || n === 'none';
}

// Persist a freshly chosen name to the server, update local state + UI,
// and notify any active watch party so remote cursor labels refresh live.
async function saveDeviceName(name) {
  name = name.trim().substring(0, 64);
  if (!name) return null;
  try {
    const res = await fetch(`${API_BASE}/api/device/rename`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ did: getDeviceId(), name })
    });
    if (res.ok) {
      const d = await res.json();
      if (!cachedDeviceData) cachedDeviceData = {};
      cachedDeviceData.device_name = d.name;
      if ($('settings-device-name')) $('settings-device-name').textContent = d.name;
      // Let other party members see the new name beside our cursor
      sendParty({ type: 'rename', name: d.name });
      showToast(`Name set to ${d.name}`);
      return d.name;
    }
  } catch(e) {}
  return null;
}

// Prompt for a name if the device is still unnamed, then rename it.
// Returns the chosen name, or null if the user cancelled / entered nothing.
async function ensureDeviceName() {
  if (!cachedDeviceData) {
    try {
      const res = await fetch(`${API_BASE}/api/device/data?did=${encodeURIComponent(getDeviceId())}`);
      if (res.ok) cachedDeviceData = await res.json();
    } catch(e) {}
  }

  let name = cachedDeviceData?.device_name;
  if (isUnnamedDeviceName(name)) {
    name = prompt('Please enter a name for this device (shown beside your cursor in Watch Parties):', '');
    if (name && name.trim()) {
      return await saveDeviceName(name);
    }
    return null; // User cancelled or entered empty string
  }
  return name;
}

async function joinParty(code) {
  if (!API_BASE) return;
  
  // 1. Ensure the user has a name before connecting
  const name = await ensureDeviceName();
  if (!name) {
    showToast('You must enter a name to join a party.');
    return;
  }

  // 2. Validate the code
  code = code.toUpperCase().substring(0, 6);
  if (code.length < 6) {
    showToast('Code must be 6 characters.');
    return;
  }

  // 3. Connect to WebSocket
  const wsUrl = `${API_BASE.replace('http', 'ws')}/ws/party/${code}?did=${getDeviceId()}`;
  partyWS = new WebSocket(wsUrl);
  
  partyWS.onopen = () => {
    $('party-status').textContent = `Connected to Party: ${code}`;
    $('party-create-btn').style.display = 'none';
    $('party-join-btn').style.display = 'none';
    $('party-leave-btn').style.display = 'block';
    document.querySelectorAll('.party-code-digit').forEach(inp => inp.disabled = true);
    showToast('Joined Watch Party!');
    
    // Send heartbeat every 30s to keep WS alive through Cloudflare
    partyPingInterval = setInterval(() => {
      if (partyWS && partyWS.readyState === WebSocket.OPEN) {
        partyWS.send(JSON.stringify({ type: 'ping' }));
      }
    }, 30000);
  };
  
  partyWS.onmessage = (event) => {
    const msg = JSON.parse(event.data);
    
    if (msg.type === 'assigned_color') {
      userColor = msg.color;
      document.documentElement.style.setProperty('--user-color', userColor);
      document.querySelector('.app-layout').classList.add('party-active');
    } else if (msg.type === 'user_joined' || msg.type === 'user_left') {
      // Could update a user list here
    } else if (msg.type === 'cursor') {
      moveRemoteCursor(msg.color, msg.x, msg.y, msg.name);
    } else if (msg.type === 'rename') {
      // A party member renamed their device — refresh their cursor label.
      updateRemoteCursorName(msg.color, msg.name);
    } else if (msg.type === 'load_video') {
      if (!currentViewerMovie || currentViewerMovie.title !== msg.movie.title) {
        openMovieViewer(msg.movie, true);
      }
    } else if (msg.type === 'play' || msg.type === 'pause' || msg.type === 'seek') {
      _applying_party_action = true;
      applyVideoAction(msg);
      setTimeout(() => { _applying_party_action = false; }, 1000); // Give it a second to load
    }
  };
  
  partyWS.onclose = () => {
    leaveParty(true);
    showToast('Left Watch Party.');
  };
  
  partyWS.onerror = () => {
    showToast('⚠ Party connection failed.');
  };
}

function applyVideoAction(msg) {
  if (msg.type === 'play') {
    // If the video player isn't active yet, we need to trigger playVideo() which sets the src
    if (viewer.style.display !== 'flex' || $('viewer-player').style.display === 'none') {
      playVideo();
      // If they are seeking to a specific time, apply it once metadata loads
      if (msg.time) {
        videoEl.addEventListener('loadedmetadata', () => { 
          videoEl.currentTime = msg.time; 
        }, { once: true });
      }
    } else {
      // Already playing, just ensure it's playing and seek if needed
      if (videoEl.paused) videoEl.play();
      if (msg.time) videoEl.currentTime = msg.time;
    }
  } else if (msg.type === 'pause') {
    if (!videoEl.paused) videoEl.pause();
  } else if (msg.type === 'seek') {
    if (videoEl.duration) videoEl.currentTime = msg.time;
  }
}

function leaveParty(silent = false) {
  if (partyPingInterval) clearInterval(partyPingInterval);
  if (partyWS) {
    partyWS.close();
    partyWS = null;
  }
  userColor = null;
  document.querySelector('.app-layout').classList.remove('party-active');
  $('party-status').textContent = 'Not connected.';
  $('party-create-btn').style.display = 'block';
  $('party-join-btn').style.display = 'block';
  $('party-leave-btn').style.display = 'none';
  document.querySelectorAll('.party-code-digit').forEach(inp => {
    inp.disabled = false;
    inp.value = '';
  });
  // Clear cursors
  $('remote-cursors-container').innerHTML = '';
}

function sendParty(data) {
  if (partyWS && partyWS.readyState === WebSocket.OPEN) {
    partyWS.send(JSON.stringify(data));
  }
}

// Throttle cursor sending
let lastMouseSend = 0;
document.addEventListener('mousemove', (e) => {
  if (!partyWS || partyWS.readyState !== WebSocket.OPEN) return;
  const now = Date.now();
  if (now - lastMouseSend < 40) return; // ~25fps
  lastMouseSend = now;
  sendParty({ type: 'cursor', x: e.clientX, y: e.clientY });
});

function moveRemoteCursor(color, x, y, name) {
  const container = $('remote-cursors-container');
  let cursor = container.querySelector(`.remote-cursor[data-color="${color}"]`);
  const displayName = (name && String(name).trim()) ? String(name).trim() : 'Unknown';
  
  if (!cursor) {
    cursor = document.createElement('div');
    cursor.className = 'remote-cursor visible';
    cursor.dataset.color = color;
    cursor.innerHTML = `
      <svg class="remote-cursor-arrow" viewBox="0 0 24 24" fill="${color}" stroke="#000" stroke-width="1">
        <path d="M5.5 3.21V20.79c0 .45.54.67.85.35l4.86-4.86a.5.5 0 01.35-.15h6.87a.5.5 0 00.35-.85L6.35 2.85a.5.5 0 00-.85.36z"/>
      </svg>
      <div class="remote-cursor-label" style="background: ${color}"></div>
    `;
    container.appendChild(cursor);
  }

  // Always sync the label with the latest device name (handles live renames)
  const label = cursor.querySelector('.remote-cursor-label');
  if (label && label.textContent !== displayName) label.textContent = displayName;
  
  cursor.style.transform = `translate(${x}px, ${y}px)`;
  cursor.classList.add('visible');
  
  // Fade out after 2 seconds of no movement
  clearTimeout(cursorTimers[color]);
  cursorTimers[color] = setTimeout(() => {
    cursor.classList.remove('visible');
  }, 2000);
}

// Update just the label of an existing remote cursor (used when a party
// member renames their device mid-session, without moving the cursor).
function updateRemoteCursorName(color, name) {
  const cursor = $('remote-cursors-container').querySelector(`.remote-cursor[data-color="${color}"]`);
  if (!cursor) return;
  const displayName = (name && String(name).trim()) ? String(name).trim() : 'Unknown';
  const label = cursor.querySelector('.remote-cursor-label');
  if (label && label.textContent !== displayName) label.textContent = displayName;
}

// Party UI Logic
function generatePartyCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // No confusing chars like 0/O, 1/I
  let code = '';
  for (let i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

function setPartyCodeInputs(code) {
  const inputs = document.querySelectorAll('.party-code-digit');
  inputs.forEach((inp, i) => {
    inp.value = code[i] || '';
  });
}

function getPartyCodeFromInputs() {
  let code = '';
  document.querySelectorAll('.party-code-digit').forEach(inp => {
    code += inp.value.toUpperCase();
  });
  return code;
}

 $('party-create-btn').addEventListener('click', () => {
  const code = generatePartyCode();
  setPartyCodeInputs(code);
  joinParty(code);
});

 $('party-join-btn').addEventListener('click', () => {
  const code = getPartyCodeFromInputs();
  joinParty(code);
});

 $('party-leave-btn').addEventListener('click', () => leaveParty());

// Setup 6-box input auto-advance
const partyInputs = document.querySelectorAll('.party-code-digit');
partyInputs.forEach((inp, index) => {
  inp.addEventListener('input', (e) => {
    // Ensure uppercase and letters/numbers only
    inp.value = inp.value.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (inp.value && index < 5) {
      partyInputs[index + 1].focus();
    }
  });
  
  inp.addEventListener('keydown', (e) => {
    if (e.key === 'Backspace' && !inp.value && index > 0) {
      partyInputs[index - 1].focus();
    }
  });
  
  inp.addEventListener('paste', (e) => {
    e.preventDefault();
    const paste = (e.clipboardData || window.clipboardData).getData('text').toUpperCase().substring(0, 6);
    if (paste) {
      setPartyCodeInputs(paste);
      // Focus the last filled input or the 6th one
      const focusIndex = Math.min(paste.length, 5);
      if (partyInputs[focusIndex]) partyInputs[focusIndex].focus();
    }
  });
});

// ─── APPLE TV UI ENHANCEMENTS (user menu, nav scroll, parallax) ─
(function initAppleTVUI() {
  // User / settings dropdown
  const menuBtn = $('user-menu-btn');
  const menu = $('user-menu');
  if (menuBtn && menu) {
    menuBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const open = menu.hasAttribute('hidden') ? false : menu.hidden;
      menu.hidden = !open;
      menuBtn.setAttribute('aria-expanded', String(open));
    });
    document.addEventListener('click', (e) => {
      if (menu.hidden) return;
      if (!menu.contains(e.target) && e.target !== menuBtn) {
        menu.hidden = true;
        menuBtn.setAttribute('aria-expanded', 'false');
      }
    });
    // Close menu after a settings click
    menu.addEventListener('click', (e) => {
      if (e.target.closest('.nav-btn')) { menu.hidden = true; menuBtn.setAttribute('aria-expanded', 'false'); }
    });
  }

  // Nav darkens on scroll
  const topNav = $('top-nav');
  if (topNav) {
    const onScroll = () => topNav.classList.toggle('scrolled', window.scrollY > 24);
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
  }

  // Subtle parallax tilt on poster cards (pointer-driven). The tilt is applied
  // to .card-poster while the hover scale lives on .movie-card, so they compose.
  if (!window.matchMedia('(prefers-reduced-motion: reduce)').matches && window.matchMedia('(hover: hover)').matches) {
    let activePoster = null;
    document.addEventListener('mousemove', (e) => {
      const card = e.target.closest && e.target.closest('.movie-card .card-poster');
      if (card !== activePoster) {
        if (activePoster) activePoster.style.transform = '';
        activePoster = card;
      }
      if (!card) return;
      const r = card.getBoundingClientRect();
      const px = (e.clientX - r.left) / r.width - 0.5;
      const py = (e.clientY - r.top) / r.height - 0.5;
      card.style.transform = `perspective(600px) rotateY(${px * 6}deg) rotateX(${-py * 6}deg)`;
    }, { passive: true });
    document.addEventListener('mouseleave', () => {
      if (activePoster) { activePoster.style.transform = ''; activePoster = null; }
    });
  }
})();

// ─── DEBUG OVERLAY (Konami code: ↑↑↓↓←→←→) ───────────────────
(function initDebugOverlay() {
  const SEQUENCE = ['ArrowUp','ArrowUp','ArrowDown','ArrowDown','ArrowLeft','ArrowRight','ArrowLeft','ArrowRight'];
  let _konmiIdx = 0;
  let _debugActive = false;
  let _debugTimer = null;
  const overlay = document.getElementById('debug-overlay');

  function fmtBytes(n) {
    if (!n || n < 1) return '0 B';
    const units = ['B','KB','MB','GB'];
    let i = 0; while (n >= 1024 && i < units.length-1) { n /= 1024; i++; }
    return n.toFixed(1) + ' ' + units[i];
  }
  function fmtRate(n) { return fmtBytes(n) + '/s'; }

  async function updateDebugOverlay() {
    if (!_debugActive) return;
    // Fetch server-side stats
    let server = {};
    try {
      const res = await fetch(`${API_BASE}/api/debug?t=${Date.now()}`);
      if (res.ok) server = await res.json();
    } catch(e) {}

    // Client-side video stats
    const v = videoEl;
    const watchPos = v && v.duration ? formatTime(v.currentTime) : '—';
    const duration = v && v.duration ? formatTime(v.duration) : '—';
    const readyState = v ? v.readyState : '—';
    let bufInfo = '—';
    if (v && v.buffered.length > 0 && v.duration) {
      // Find buffered range containing playhead
      let end = 0;
      for (let i = 0; i < v.buffered.length; i++) {
        if (v.currentTime >= v.buffered.start(i) - 0.5 && v.currentTime <= v.buffered.end(i) + 0.5) {
          end = v.buffered.end(i);
          break;
        }
      }
      const bufPct = v.duration ? ((end / v.duration) * 100).toFixed(1) : 0;
      const aheadSec = Math.max(0, end - v.currentTime);
      bufInfo = `${bufPct}% (+${aheadSec.toFixed(0)}s)`;
    }
    const networkState = v ? ['EMPTY','IDLE','LOADING','NO_SOURCE'][v.networkState] || v.networkState : '—';

    // Connection stats (if available via Navigation Timing API)
    let connInfo = '—';
    if (navigator.connection) {
      const c = navigator.connection;
      connInfo = `${c.downlink || '?'} Mbps (${c.effectiveType || '?'})`;
    }

    const rows = [
      { label: 'Server Upload', value: fmtRate(server.upload_rate || 0), cls: 'good' },
      { label: 'Server Download', value: fmtRate(server.download_rate || 0) },
      { label: 'Upload Cap', value: fmtRate(server.upload_cap || 0) },
      { label: 'Stream Share', value: fmtRate(server.stream_share || 0) + ` (${server.active_streams||0} stream${(server.active_streams||0)!==1?'s':''})` },
      { label: 'Peak Upload', value: fmtRate(server.peak_up || 0) },
      { label: 'Total Up/Down', value: `${fmtBytes(server.total_up||0)} / ${fmtBytes(server.total_down||0)}` },
      { label: 'Online', value: server.online || 0 },
      { label: '─ Player ─', value: '' },
      { label: 'Watch Position', value: `${watchPos} / ${duration}` },
      { label: 'Buffered', value: bufInfo, cls: bufInfo !== '—' && bufInfo.startsWith('0%') ? 'bad' : 'good' },
      { label: 'Ready State', value: readyState + (v && v.paused ? ' (paused)' : ' (playing)') },
      { label: 'Network State', value: networkState, cls: networkState === 'LOADING' ? 'warn' : '' },
      { label: 'Your Connection', value: connInfo },
    ];

    if (overlay) {
      overlay.innerHTML = `<div class="debug-overlay__title">Debug (↑↑↓↓←→←→ to close)</div>` +
        rows.map(r => `<div class="debug-overlay__row"><span class="debug-overlay__label">${r.label}</span><span class="debug-overlay__value ${r.cls||''}">${r.value}</span></div>`).join('');
    }
  }

  document.addEventListener('keydown', (e) => {
    // Only watch arrow keys for the Konami sequence
    if (e.key === SEQUENCE[_konmiIdx]) {
      _konmiIdx++;
      if (_konmiIdx === SEQUENCE.length) {
        _konmiIdx = 0;
        _debugActive = !_debugActive;
        if (overlay) overlay.style.display = _debugActive ? '' : 'none';
        if (_debugActive) {
          updateDebugOverlay();
          _debugTimer = setInterval(updateDebugOverlay, 1000);
        } else {
          if (_debugTimer) { clearInterval(_debugTimer); _debugTimer = null; }
        }
      }
    } else {
      // Reset on wrong key (but allow arrow keys to restart matching)
      _konmiIdx = (e.key === SEQUENCE[0]) ? 1 : 0;
    }
  });
})();

// ─── MAIN INIT ────────────────────────────────────────────────
(async function init() {
  currentSort = 'title'; currentDir = 'asc';
  if (sortBy) sortBy.value = 'title'; if (sortDirBtn) sortDirBtn.textContent = '↓';
  try { const cR = await fetch('config.json?t=' + Date.now()); if (cR.ok) { const c = await cR.json(); API_BASE = c.API_BASE || ''; } } catch(e) { API_BASE = ''; }

  // Show the scan bar immediately so the user sees loading activity
  if (scanBar) scanBar.classList.remove('hidden');

  await initWithGate();

  // Load movie data first and render immediately — don't make the user stare
  // at a black screen while library/progress data loads in parallel.
  await loadData();
  // These can finish in the background; the home page is already usable
  hydrateProgress().then(() => { if (activeTab === 'home') renderRows(); });
  loadLibrary().then(() => { if (activeTab === 'library') renderLibrary(); });
  loadShowsData();

  pushPresencePing(); setInterval(pushPresencePing, 5000);
  
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