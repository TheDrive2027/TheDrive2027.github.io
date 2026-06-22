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
const LOCAL_KEY_STORE = 'thedrive_access_key_v1';
const LOCAL_DEVICE_ID = 'thedrive_device_id_v1';

function getSavedKey() { try { return localStorage.getItem(LOCAL_KEY_STORE) || null; } catch(e) { return null; } }
function saveKey(key) { try { localStorage.setItem(LOCAL_KEY_STORE, key); } catch(e) {} }
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
  const savedKey = getSavedKey();

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
      } else { localStorage.removeItem(LOCAL_KEY_STORE); }
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
    submitBtn.addEventListener('click', attempt, { once: true });
    input.addEventListener('keydown', e => { if (e.key === 'Enter') attempt(); }, { once: true });
    setTimeout(() => input.focus(), 100);
  });
}

// ─── STATE ────────────────────────────────────────────────────
let allMovies   = [], allShows = [], filtered = [];
let currentSort = 'title', currentDir = 'asc', activeTab = 'movies'; 
let activeFilters = { maturity: new Set(), status: new Set(), resolution: new Set() };

function hasActiveFilters() {
  const search = searchInput ? searchInput.value.trim() : '';
  return search.length > 0 || activeFilters.maturity.size > 0 || activeFilters.status.size > 0 || activeFilters.resolution.size > 0;
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
const scanBar = $('scan-bar'), lastUpdatedEl = $('last-updated'), refreshBtn = $('refresh-btn'), scanFill = $('scan-fill');
const toast = $('toast'), rowView = $('row-view'), gridView = $('grid-view'), movieGrid = $('movie-grid');
const gridEmpty = $('grid-empty'), sidebarClearBtn = $('sidebar-clear-btn');

// ─── UTILITIES ────────────────────────────────────────────────
function normalize(str) { return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function extractYear(dateStr) { if (!dateStr) return '—'; const m = String(dateStr).match(/\d{4}/); return m ? m[0] : '—'; }
function parseSizeGB(sizeStr) { if (!sizeStr) return 0; const n = parseFloat(sizeStr), s = sizeStr.toUpperCase(); if (s.includes('TB')) return n * 1024; if (s.includes('GB')) return n; if (s.includes('MB')) return n / 1024; return n; }
function parseRuntimeMinutes(str) { if (!str) return 0; const hm = str.match(/(\d+)\s*h(?:r|ours?)?\s*(\d+)?\s*m?/i); if (hm) return parseInt(hm[1]) * 60 + (parseInt(hm[2]) || 0); const m = str.match(/(\d+)/); return m ? parseInt(m[1]) : 0; }
const MATURITY_ORDER = { 'G': 1, 'PG': 2, 'PG-13': 3, 'PG13': 3, 'R': 4, 'NC-17': 5, 'NR': 6 };
function parseResolutionScore(res) { if (!res) return 0; const s = String(res).toUpperCase().trim(); if (s === '4K' || s === 'UHD' || s.includes('2160')) return 2160; const m = s.match(/(\d+)/); return m ? parseInt(m[1], 10) : 0; }
function imdbClass(rating) { const r = parseFloat(rating); if (r >= 8) return 'imdb-high'; if (r >= 6.5) return 'imdb-mid'; return 'imdb-low'; }
function resClass(res) { const r = String(res).toUpperCase(); if (r.includes('4K') || r.includes('2160')) return 'res-4k'; if (r.includes('1080')) return 'res-1080'; if (r.includes('720') || r.includes('576')) return 'res-720'; return 'res-other'; }
function ratingClass(rating) { const r = String(rating || '').toUpperCase().replace(/[\s-]/g, ''); if (r === 'G') return 'rating-g'; if (r === 'PG') return 'rating-pg'; if (r === 'PG13') return 'rating-pg13'; if (r === 'R') return 'rating-r'; return ''; }
let toastTimer;
function showToast(msg, duration = 3000) { toast.textContent = msg; toast.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove('show'), duration); }
function updateLastUpdated(date) { const d = (date instanceof Date && !isNaN(date)) ? date : new Date(); let h = d.getHours(), m = d.getMinutes(); const ampm = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12; const mm = String(m).padStart(2, '0'); if (lastUpdatedEl) lastUpdatedEl.textContent = h + ':' + mm + ' ' + ampm; }
function setProgress(pct) { scanFill.style.width = pct + '%'; }
function formatTime(sec) {
  if (isNaN(sec)) return '0:00';
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = Math.floor(sec % 60);
  return h > 0 ? `${h}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}` : `${m}:${String(s).padStart(2,'0')}`;
}

// ─── SHOWS (Local Stubs) ──────────────────────────────────────
async function loadShowsData() { allShows = []; renderShows(); }
function renderShows() { const c = document.getElementById('shows-grid'); if (c) c.innerHTML = '<div class="empty-state"><span class="empty-icon">◻</span><p>No shows found.</p></div>'; updateCounts(); }
function filterAndRenderShows() { renderShows(); }

// ─── SIDEBAR FILTERS ──────────────────────────────────────────
function populateFilterCheckboxes() {
  const maturityEl = $('filter-maturity-checks');
  if (maturityEl) {
    const ratings = [...new Set(allMovies.map(m => m.maturityRating).filter(Boolean))].sort((a, b) => (MATURITY_ORDER[a.toUpperCase().replace(/[\s-]/g,'')] || 99) - (MATURITY_ORDER[b.toUpperCase().replace(/[\s-]/g,'')] || 99));
    maturityEl.innerHTML = ratings.map(r => `<label class="check-row"><input type="checkbox" value="${escHtml(r)}" data-filter="maturity" ${activeFilters.maturity.has(r) ? 'checked' : ''} /><span class="check-label ${ratingClass(r)}">${escHtml(r)}</span></label>`).join('');
  }
  const resEl = $('filter-resolution-checks');
  if (resEl) {
    const resolutions = [...new Set(allMovies.map(m => m.resolution).filter(Boolean))].sort((a, b) => parseResolutionScore(b) - parseResolutionScore(a));
    resEl.innerHTML = resolutions.map(r => `<label class="check-row"><input type="checkbox" value="${escHtml(r)}" data-filter="resolution" ${activeFilters.resolution.has(r) ? 'checked' : ''} /><span class="check-label ${resClass(r)}">${escHtml(r)}</span></label>`).join('');
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
function updateClearBtn() { if (sidebarClearBtn) sidebarClearBtn.hidden = !(activeFilters.maturity.size > 0 || activeFilters.status.size > 0 || activeFilters.resolution.size > 0); }
function clearAllFilters() { activeFilters.maturity.clear(); activeFilters.status.clear(); activeFilters.resolution.clear(); if (searchInput) { searchInput.value = ''; clearSearch && clearSearch.classList.remove('visible'); } document.querySelectorAll('.sidebar-checks input[type="checkbox"]').forEach(cb => cb.checked = false); updateClearBtn(); render(); saveSettings(); }
function updateCounts() {
  const totalMovies = allMovies.length, totalEps = 0;
  let totalText = activeTab === 'shows' ? totalEps + ' Episodes' : activeTab === 'stats' ? (totalMovies + totalEps) + ' Files' : totalMovies + ' Movies';
  if (movieCount) movieCount.textContent = totalText;
  if (availCount) availCount.style.display = 'none';
}

// ─── SORT & FILTER ────────────────────────────────────────────
function applyFilters() {
  const q = normalize(searchInput ? searchInput.value : '');
  filtered = allMovies.filter(m => {
    if (q && !normalize(m.title).includes(q)) return false;
    if (activeFilters.maturity.size > 0 && !activeFilters.maturity.has(m.maturityRating)) return false;
    if (activeFilters.resolution.size > 0 && !activeFilters.resolution.has(m.resolution)) return false;
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
    else if (key === 'maturity') { va = MATURITY_ORDER[a.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99; vb = MATURITY_ORDER[b.maturityRating?.toUpperCase().replace(/[\s-]/g,'')] || 99; }
    else if (key === 'res') { va = parseResolutionScore(a.resolution); vb = parseResolutionScore(b.resolution); }
    if (va < vb) return dir === 'asc' ? -1 : 1; if (va > vb) return dir === 'asc' ? 1 : -1; return 0;
  });
  renderCurrentView();
}

// ─── RENDER ───────────────────────────────────────────────────
function render() { applyFilters(); }
function renderCurrentView() {
  if (hasActiveFilters()) {
    rowView.classList.remove('active'); gridView.classList.add('active'); renderGrid();
    if (resultsSummary) resultsSummary.textContent = `Showing ${filtered.length} of ${allMovies.length} movies`;
  } else {
    gridView.classList.remove('active'); rowView.classList.add('active'); renderRows();
    if (resultsSummary) resultsSummary.textContent = `${allMovies.length} movies in the library`;
  }
}
function renderRows() {
  const availableMovies = [...allMovies].sort((a, b) => getRatingScore(b.title) - getRatingScore(a.title));
  const imdbMovies = [...allMovies].filter(m => parseFloat(m.imdbRating) > 0).sort((a, b) => (parseFloat(b.imdbRating) || 0) - (parseFloat(a.imdbRating) || 0));
  if ($('row-available-cards')) renderRowCards($('row-available-cards'), availableMovies.slice(0, 30));
  if ($('row-requested')) $('row-requested').style.display = 'none';
  if ($('row-imdb-cards')) renderRowCards($('row-imdb-cards'), imdbMovies.slice(0, 30));
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

function buildCard(m, i, isRowCard) {
  const card = document.createElement('div');
  card.className = isRowCard ? 'movie-card row-card' : 'movie-card';
  card.dataset.key = normalize(m.title);
  card.style.animationDelay = Math.min(i * 30, 400) + 'ms';
  card.style.cursor = 'pointer';

  const progress = getVideoProgress(m.title);
  const progressHtml = progress > 5 ? `<div class="card-progress-bar"><div class="card-progress-fill" style="width:${Math.min(100, (progress / m.duration) * 100)}%"></div></div>` : '';

  card.innerHTML = `
    <div class="card-poster card-poster--playable">
      ${m.poster ? `<img src="${m.poster}" alt="${escHtml(m.title)}" loading="lazy" onload="this.classList.add('loaded')" />` : ''}
      <div class="card-play-overlay"><div class="card-play-btn"><span class="card-play-icon">&#9654;</span></div></div>
      ${progressHtml}
    </div>
    <div class="card-title">${escHtml(m.title)}</div>
    <div class="card-meta">
      <span class="card-year">${escHtml(m.year)}</span><span class="card-sep">·</span>
      <span class="card-rating ${ratingClass(m.maturityRating)}">${escHtml(m.maturityRating) || '—'}</span>
      ${m.runtime ? `<span class="card-sep">·</span><span class="card-runtime">${escHtml(m.runtime)}</span>` : ''}
    </div>
    <div class="card-row">
      <span class="card-imdb ${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span>
      <span class="card-res ${resClass(m.resolution)}">${escHtml(m.resolution) || '—'}</span>
    </div>
    <div class="card-footer">${ratingHTML(m.title)}</div>
  `;
  card.addEventListener('click', () => openMovieViewer(m));
  return card;
}

function ratingHTML(title) {
  const userVote = getUserRating(title), ups = getRatingCount(title, 'up'), downs = getRatingCount(title, 'down');
  return `<div class="rating-wrap">
    <button class="rating-btn rating-btn--up ${userVote === 'up' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="up" title="Liked it"><span class="rating-icon">👍</span><span class="rating-count">${ups || 0}</span></button>
    <button class="rating-btn rating-btn--down ${userVote === 'down' ? 'active' : ''}" data-rating-title="${escHtml(title)}" data-rating-type="down" title="Didn't like it"><span class="rating-icon">👎</span><span class="rating-count">${downs || 0}</span></button>
  </div>`;
}
function escHtml(str) { return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ─── MOVIE VIEWER & PLAYER LOGIC ──────────────────────────────
function getVideoProgress(title) { return parseFloat(localStorage.getItem('thedrive_progress_' + normalize(title)) || '0'); }
function saveVideoProgress(title, time) { localStorage.setItem('thedrive_progress_' + normalize(title), String(time)); }

let currentViewerMovie = null;
const viewer = $('movie-viewer');
const videoEl = $('video-el');

function openMovieViewer(m) {
  currentViewerMovie = m;
  $('viewer-details').hidden = false;
  $('viewer-player').hidden = true;
  viewer.hidden = false;
  document.body.style.overflow = 'hidden';

  $('viewer-poster').src = m.poster || '';
  $('viewer-title').textContent = m.title;
  $('viewer-meta').innerHTML = `
    <span>${m.year || '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-rating ${ratingClass(m.maturityRating)}">${m.maturityRating || 'NR'}</span>
    <span class="card-sep">·</span>
    <span>${m.runtime || '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-imdb ${imdbClass(m.imdbRating)}">${m.imdbRating ? '★ ' + m.imdbRating : '—'}</span>
    <span class="card-sep">·</span>
    <span class="card-res ${resClass(m.resolution)}">${m.resolution || '—'}</span>
  `;
  $('viewer-plot').textContent = m.plot || 'No plot available.';
  $('viewer-cast').innerHTML = m.cast.length ? `<b>CAST:</b> ${m.cast.join(', ')}` : '';
  
  const progress = getVideoProgress(m.title);
  $('play-btn-text').textContent = progress > 10 ? 'Resume' : 'Play';
}

function closeMovieViewer() {
  viewer.hidden = true;
  if (!videoEl.paused) videoEl.pause();
  videoEl.removeAttribute('src'); videoEl.load();
  document.body.style.overflow = '';
  render(); // Refresh cards to show new progress bar
}

function playVideo() {
  $('viewer-details').hidden = true;
  $('viewer-player').hidden = false;
  videoEl.src = currentViewerMovie.driveLink;
  
  const startTime = getVideoProgress(currentViewerMovie.title);
  videoEl.addEventListener('loadedmetadata', () => {
    if (startTime > 10 && startTime < videoEl.duration - 10) {
      videoEl.currentTime = startTime;
    }
    videoEl.play();
  }, { once: true });
}

// Video controls
videoEl.addEventListener('timeupdate', () => {
  if (!videoEl.paused) saveVideoProgress(currentViewerMovie.title, videoEl.currentTime);
  $('ctrl-progress-played').style.width = ((videoEl.currentTime / videoEl.duration) * 100) + '%';
  $('ctrl-time').textContent = `${formatTime(videoEl.currentTime)} / ${formatTime(videoEl.duration)}`;
});
 $('ctrl-play-pause').addEventListener('click', () => {
  if (videoEl.paused) { videoEl.play(); $('ctrl-play-pause').innerHTML = '&#10074;&#10074;'; } 
  else { videoEl.pause(); $('ctrl-play-pause').innerHTML = '&#9654;'; }
});
videoEl.addEventListener('play', () => $('ctrl-play-pause').innerHTML = '&#10074;&#10074;');
videoEl.addEventListener('pause', () => $('ctrl-play-pause').innerHTML = '&#9654;');
 $('ctrl-progress-track').addEventListener('click', (e) => {
  const rect = e.currentTarget.getBoundingClientRect();
  const pct = (e.clientX - rect.left) / rect.width;
  if (videoEl.duration) videoEl.currentTime = videoEl.duration * pct;
});
 $('ctrl-fullscreen').addEventListener('click', () => {
  if (!document.fullscreenElement) $('viewer-player').requestFullscreen();
  else document.exitFullscreen();
});
 $('viewer-play-btn').addEventListener('click', playVideo);
 $('viewer-close').addEventListener('click', closeMovieViewer);
 $('viewer-backdrop').addEventListener('click', closeMovieViewer);
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !viewer.hidden) closeMovieViewer(); });

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

const mainContent = document.getElementById('main-content');
if (mainContent) {
  mainContent.addEventListener('click', e => {
    const btn = e.target.closest('.rating-btn'); if (!btn) return;
    // Prevent rating click from triggering card open
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
}

// ─── NOTIFICATIONS (Local Stubs) ─────────────────────────────
function toggleNotificationPanel() { const p = $('notif-panel'); if (p) p.style.display = p.style.display === 'block' ? 'none' : 'block'; }
(function initNotificationBell() {
  if (!refreshBtn) return;
  refreshBtn.id = 'notif-btn'; refreshBtn.classList.remove('spinning');
  refreshBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg><span class="notif-dot" style="display:none; position:absolute; top:4px; right:4px; width:8px; height:8px; background:var(--red); border-radius:50%; border:1px solid var(--bg);"></span>`;
  const panel = document.createElement('div'); panel.id = 'notif-panel';
  panel.style.cssText = 'display:none;position:absolute;top:calc(100% + 8px);right:0;width:320px;max-height:400px;overflow-y:auto;background:var(--surface);border:1px solid var(--border);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,0.4);z-index:1000;padding:12px;font-size:13px;color:var(--text)';
  panel.innerHTML = '<div class="notif-list"><div class="notif-empty">No notifications</div></div>';
  refreshBtn.parentNode.style.position = 'relative'; refreshBtn.parentNode.appendChild(panel);
  refreshBtn.addEventListener('click', (e) => { e.stopPropagation(); toggleNotificationPanel(); });
  document.addEventListener('click', (e) => { if (!refreshBtn.contains(e.target) && !panel.contains(e.target)) panel.style.display = 'none'; });
})();

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
      year: v.year || '—', imdbRating: v.imdbRating || '', plot: v.plot || '', cast: v.cast || [],
      driveLink: API_BASE + v.video, poster: v.poster ? (API_BASE + v.poster) : null
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

// ─── STATS ────────────────────────────────────────────────────
let statsLoaded = false, statsLoadedAt = 0, chartLibrary = null, chartUsers = null, chartPresence = null;
function initStatsTab() { renderLocalStats(); if (!statsLoaded || Date.now() - statsLoadedAt > 60000) { fetchStatsData(); statsLoaded = true; statsLoadedAt = Date.now(); } }
function renderLocalStats() {
  if (!allMovies.length) return;
  const t = allMovies.length;
  if ($('upload-fraction')) $('upload-fraction').textContent = t + ' movies uploaded';
  if ($('upload-pct')) $('upload-pct').textContent = '100%';
  if ($('upload-fill')) $('upload-fill').style.width = '100%';
  setText('stat-total-films', t);
  if ($('stat-available')?.parentElement) $('stat-available').parentElement.style.display = 'none';
  let m = 0; allMovies.forEach(v => { m += parseRuntimeMinutes(v.runtime); });
  if (m > 0) setText('stat-total-runtime', Math.floor(m / 60) + 'h ' + (m % 60) + 'm');
