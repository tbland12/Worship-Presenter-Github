import { ensureApiBridge } from './shared/bridge.js';

const api = ensureApiBridge();

const elements = {
  grid: document.getElementById('library-grid'),
  empty: document.getElementById('library-empty'),
  path: document.getElementById('library-path'),
  refresh: document.getElementById('refresh-library'),
  importPack: document.getElementById('import-pack'),
  close: document.getElementById('close-library'),
  filterAll: document.getElementById('filter-all'),
  filterImages: document.getElementById('filter-images'),
  filterVideos: document.getElementById('filter-videos'),
  filters: document.querySelector('.library-filters')
};

let libraryItems = [];
let filterMode = 'all';
let libraryScope = new URLSearchParams(window.location.search).get('scope') || 'background';

function clearGrid() {
  elements.grid.innerHTML = '';
}

function renderItem(item) {
  const card = document.createElement('button');
  card.type = 'button';
  card.className = 'library-card';

  const thumb = document.createElement('div');
  thumb.className = 'library-thumb';

  if (item.type === 'image') {
    const img = document.createElement('img');
    img.src = item.fileUrl;
    img.alt = item.name;
    thumb.appendChild(img);
  } else {
    const video = document.createElement('video');
    video.src = item.fileUrl;
    video.muted = true;
    video.playsInline = true;
    video.preload = 'metadata';
    video.loop = true;
    thumb.appendChild(video);

    card.addEventListener('mouseenter', () => {
      video.play().catch(() => {});
    });
    card.addEventListener('mouseleave', () => {
      video.pause();
      video.currentTime = 0;
    });
  }

  const tag = document.createElement('div');
  tag.className = 'library-tag';
  tag.textContent = item.type === 'video' ? 'Video' : 'Image';
  thumb.appendChild(tag);

  const name = document.createElement('div');
  name.className = 'library-name';
  name.textContent = item.name;

  card.appendChild(thumb);
  card.appendChild(name);

  card.addEventListener('click', () => {
    if (window.api && window.api.selectLibraryItem) {
      const normalized = item.relativePath.replace(/\\/g, '/');
      let scoped = normalized;
      if (libraryScope === 'announcements') {
        scoped = `Announcements/${normalized.replace(/^Announcements\//i, '')}`;
      } else if (libraryScope === 'timer') {
        scoped = `Timers/${normalized.replace(/^Timers\//i, '')}`;
      }
      const relativePath = `library/${scoped}`;
      window.api.selectLibraryItem({ scope: libraryScope, sourcePath: relativePath });
    }
  });

  elements.grid.appendChild(card);
}

async function loadLibrary() {
  if (!api || !window.api || !window.api.listLibraryItems) {
    elements.path.textContent = 'Library API not available.';
    return;
  }
  const payload = await window.api.listLibraryItems({ scope: libraryScope });
  elements.path.textContent = payload.folder || 'Library folder unavailable';
  clearGrid();
  libraryItems = payload.items || [];
  renderLibrary();
}

function renderLibrary() {
  clearGrid();
  const items = libraryItems.filter((item) => {
    if (filterMode === 'images') {
      return item.type === 'image';
    }
    if (filterMode === 'videos') {
      return item.type === 'video';
    }
    return true;
  });
  if (items.length === 0) {
    elements.empty.style.display = 'block';
    return;
  }
  elements.empty.style.display = 'none';
  items.forEach(renderItem);
}

function setFilter(mode) {
  filterMode = mode;
  elements.filterAll.classList.toggle('active', mode === 'all');
  elements.filterImages.classList.toggle('active', mode === 'images');
  elements.filterVideos.classList.toggle('active', mode === 'videos');
  renderLibrary();
}

function updateImportPackVisibility() {
  if (!elements.importPack) {
    return;
  }
  elements.importPack.hidden = libraryScope !== 'background';
}

async function importContentPack() {
  if (!elements.importPack) {
    return;
  }
  if (!window.api || !window.api.importContentPack) {
    window.alert('Content pack import is unavailable.');
    return;
  }
  const originalLabel = elements.importPack.textContent;
  elements.importPack.disabled = true;
  elements.importPack.textContent = 'Importing...';
  try {
    const result = await window.api.importContentPack();
    if (!result || result.canceled) {
      return;
    }
    if (result.error) {
      window.alert(result.error);
      return;
    }
    const imported = Number(result.imported) || 0;
    const skipped = Number(result.skipped) || 0;
    const failed = Number(result.failed) || 0;
    const renamed = Number(result.renamed) || 0;
    const name = result.packName ? ` "${result.packName}"` : '';
    window.alert(
      `Content pack${name} imported.\n\nImported: ${imported}\nSkipped: ${skipped}\nRenamed: ${renamed}\nFailed: ${failed}`
    );
    if (imported > 0 || renamed > 0) {
      await loadLibrary();
    }
  } finally {
    elements.importPack.disabled = false;
    elements.importPack.textContent = originalLabel;
  }
}

elements.refresh.addEventListener('click', loadLibrary);
if (elements.importPack) {
  elements.importPack.addEventListener('click', importContentPack);
}
elements.close.addEventListener('click', () => {
  if (window.api && window.api.closeLibrary) {
    window.api.closeLibrary();
  } else {
    window.close();
  }
});

elements.filterAll.addEventListener('click', () => setFilter('all'));
elements.filterImages.addEventListener('click', () => setFilter('images'));
elements.filterVideos.addEventListener('click', () => setFilter('videos'));

if (window.api && window.api.onLibraryScope) {
  window.api.onLibraryScope((scope) => {
    libraryScope = scope || 'background';
    if (libraryScope === 'announcements') {
      setFilter('images');
      if (elements.filters) {
        elements.filters.style.display = 'none';
      }
    } else if (elements.filters) {
      elements.filters.style.display = '';
    }
    updateImportPackVisibility();
    loadLibrary();
  });
}

if (libraryScope === 'announcements') {
  setFilter('images');
  if (elements.filters) {
    elements.filters.style.display = 'none';
  }
}
updateImportPackVisibility();
loadLibrary();
