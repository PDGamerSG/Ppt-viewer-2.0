(function () {
  'use strict';

  const { loadDocument, renderSlide, renderThumbnail, cleanup } = window.PdfRenderer;

  // State
  let currentSlide = 1;
  let totalSlides = 0;
  let zoomLevel = 1;
  let currentFilePath = null;
  let isFullscreen = false;
  let fsCounterTimeout = null;
  let resizeTimeout = null;
  let libreOfficeReady = false;

  // DOM Elements
  const loadingOverlay = document.getElementById('loading-overlay');
  const errorOverlay = document.getElementById('error-overlay');
  const errorMessage = document.getElementById('error-message');
  const errorRetryBtn = document.getElementById('error-retry-btn');
  const errorCloseBtn = document.getElementById('error-close-btn');
  const loScreen = document.getElementById('libreoffice-screen');
  const welcomeScreen = document.getElementById('welcome-screen');
  const viewer = document.getElementById('viewer');
  const slideCanvas = document.getElementById('slide-canvas');
  const thumbnailList = document.getElementById('thumbnail-list');
  const recentList = document.getElementById('recent-list');
  const fsCounter = document.getElementById('fs-counter');

  // Toolbar
  const toolbarOpen = document.getElementById('toolbar-open');
  const toolbarPrev = document.getElementById('toolbar-prev');
  const toolbarNext = document.getElementById('toolbar-next');
  const toolbarSlideCounter = document.getElementById('toolbar-slide-counter');
  const toolbarZoomIn = document.getElementById('toolbar-zoom-in');
  const toolbarZoomOut = document.getElementById('toolbar-zoom-out');
  const toolbarZoomReset = document.getElementById('toolbar-zoom-reset');
  const toolbarZoom = document.getElementById('toolbar-zoom');
  const toolbarPresent = document.getElementById('toolbar-present');

  // Status bar
  const statusFilename = document.getElementById('status-filename');
  const statusSlide = document.getElementById('status-slide');
  const statusZoom = document.getElementById('status-zoom');
  const statusLoDot = document.querySelector('.lo-dot');
  const statusLoText = document.getElementById('status-lo-text');

  // Welcome screen
  const dropZone = document.getElementById('drop-zone');
  const welcomeOpenBtn = document.getElementById('welcome-open-btn');

  // LibreOffice screen
  const downloadLoBtn = document.getElementById('download-lo-btn');
  const refreshLoBtn = document.getElementById('refresh-lo-btn');
  const loPathInput = document.getElementById('lo-path-input');
  const loPathSaveBtn = document.getElementById('lo-path-save-btn');

  // Initialize
  async function init() {
    const loPath = await window.api.detectLibreOffice();
    if (loPath) {
      libreOfficeReady = true;
      statusLoDot.classList.add('ready');
      statusLoText.textContent = 'LibreOffice Ready';
      showWelcome();
    } else {
      libreOfficeReady = false;
      statusLoDot.classList.add('missing');
      statusLoText.textContent = 'Not Found';
      showLibreOfficeScreen();
    }

    loadRecentFiles();
    setupEventListeners();
  }

  function showScreen(screen) {
    loScreen.classList.add('hidden');
    welcomeScreen.classList.add('hidden');
    viewer.classList.add('hidden');
    screen.classList.remove('hidden');
  }

  function showLibreOfficeScreen() {
    showScreen(loScreen);
  }

  function showWelcome() {
    showScreen(welcomeScreen);
  }

  function showViewer() {
    showScreen(viewer);
  }

  // Recent files
  async function loadRecentFiles() {
    const files = await window.api.getRecentFiles();
    recentList.innerHTML = '';

    if (files.length === 0) {
      document.getElementById('recent-files').classList.add('hidden');
      return;
    }

    document.getElementById('recent-files').classList.remove('hidden');

    for (const file of files) {
      const li = document.createElement('li');
      const date = new Date(file.openedAt);
      const dateStr = date.toLocaleDateString();

      li.innerHTML =
        '<div>' +
          '<div class="recent-file-name">' + escapeHtml(file.name) + '</div>' +
          '<div class="recent-file-path">' + escapeHtml(file.path) + '</div>' +
        '</div>' +
        '<div class="recent-file-date">' + dateStr + '</div>';

      li.addEventListener('click', function () { openFile(file.path); });
      recentList.appendChild(li);
    }
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // File opening
  async function openFile(filePath) {
    if (!libreOfficeReady) return;

    currentFilePath = filePath;
    const parts = filePath.replace(/\\/g, '/').split('/');
    const fileName = parts[parts.length - 1];

    loadingOverlay.classList.remove('hidden');

    try {
      const pdfPath = await window.api.convertFile(filePath);
      const data = await window.api.readFile(pdfPath);

      cleanup();
      totalSlides = await loadDocument(data);
      currentSlide = 1;
      zoomLevel = 1;

      await window.api.saveRecentFile(filePath, fileName);

      showViewer();
      updateSlideCounter();
      updateZoomDisplay();
      statusFilename.textContent = fileName;
      await window.api.setTitle(fileName + ' \u2014 PPT Viewer');

      await renderCurrentSlide();
      renderAllThumbnails();
    } catch (err) {
      showError(err.message || String(err));
    } finally {
      loadingOverlay.classList.add('hidden');
    }
  }

  function showError(msg) {
    errorMessage.textContent = msg;
    errorOverlay.classList.remove('hidden');
  }

  // Slide rendering
  async function renderCurrentSlide() {
    if (totalSlides === 0) return;

    const mainView = document.getElementById('main-view');
    const maxWidth = mainView.clientWidth - 40;
    const maxHeight = mainView.clientHeight - 40;

    // Get base page dimensions
    const tempCanvas = document.createElement('canvas');
    const tempResult = await renderSlide(tempCanvas, currentSlide, 1);
    if (!tempResult) return;

    const baseWidth = tempResult.width;
    const baseHeight = tempResult.height;

    // Calculate fit scale
    const fitScale = Math.min(maxWidth / baseWidth, maxHeight / baseHeight);
    const effectiveZoom = fitScale * zoomLevel;

    await renderSlide(slideCanvas, currentSlide, effectiveZoom);
    updateSlideCounter();
    updateThumbnailHighlight();
  }

  async function renderAllThumbnails() {
    thumbnailList.innerHTML = '';

    for (let i = 1; i <= totalSlides; i++) {
      const item = document.createElement('div');
      item.className = 'thumbnail-item' + (i === currentSlide ? ' active' : '');
      item.dataset.page = i;

      const canvas = await renderThumbnail(i, 160);
      if (canvas) {
        item.appendChild(canvas);
      }

      const num = document.createElement('span');
      num.className = 'thumbnail-number';
      num.textContent = i;
      item.appendChild(num);

      (function (pageNum) {
        item.addEventListener('click', function () { goToSlide(pageNum); });
      })(i);

      thumbnailList.appendChild(item);
    }
  }

  function updateThumbnailHighlight() {
    const items = thumbnailList.querySelectorAll('.thumbnail-item');
    items.forEach(function (item) {
      var page = parseInt(item.dataset.page);
      if (page === currentSlide) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });

    // Scroll active into view
    var active = thumbnailList.querySelector('.thumbnail-item.active');
    if (active) {
      active.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }

  // Navigation
  function goToSlide(num) {
    if (num < 1 || num > totalSlides) return;
    currentSlide = num;
    renderCurrentSlide();
    showFsCounter();
  }

  function prevSlide() {
    goToSlide(currentSlide - 1);
  }

  function nextSlide() {
    goToSlide(currentSlide + 1);
  }

  // Zoom
  function setZoom(level) {
    zoomLevel = Math.max(0.25, Math.min(4, level));
    updateZoomDisplay();
    renderCurrentSlide();
  }

  function zoomIn() {
    setZoom(zoomLevel + 0.1);
  }

  function zoomOut() {
    setZoom(zoomLevel - 0.1);
  }

  function zoomReset() {
    setZoom(1);
  }

  function updateZoomDisplay() {
    var pct = Math.round(zoomLevel * 100) + '%';
    toolbarZoom.textContent = pct;
    statusZoom.textContent = pct;
  }

  // Slide counter
  function updateSlideCounter() {
    var text = 'Slide ' + currentSlide + ' of ' + totalSlides;
    toolbarSlideCounter.textContent = text;
    statusSlide.textContent = text;
    fsCounter.textContent = text;
  }

  // Fullscreen
  async function toggleFullscreen() {
    if (totalSlides === 0) return;

    isFullscreen = !isFullscreen;
    if (isFullscreen) {
      await window.api.enterFullscreen();
      document.body.classList.add('fullscreen');
      showFsCounter();
    } else {
      await window.api.exitFullscreen();
      document.body.classList.remove('fullscreen');
      fsCounter.classList.add('hidden');
    }

    setTimeout(function () { renderCurrentSlide(); }, 100);
  }

  function showFsCounter() {
    if (!isFullscreen) return;
    fsCounter.classList.remove('hidden');
    fsCounter.classList.remove('fade');
    clearTimeout(fsCounterTimeout);
    fsCounterTimeout = setTimeout(function () {
      fsCounter.classList.add('fade');
      setTimeout(function () { fsCounter.classList.add('hidden'); }, 500);
    }, 2000);
  }

  // Event Listeners
  function setupEventListeners() {
    // Toolbar
    toolbarOpen.addEventListener('click', function () { window.api.openFileDialog(); });
    toolbarPrev.addEventListener('click', prevSlide);
    toolbarNext.addEventListener('click', nextSlide);
    toolbarZoomIn.addEventListener('click', zoomIn);
    toolbarZoomOut.addEventListener('click', zoomOut);
    toolbarZoomReset.addEventListener('click', zoomReset);
    toolbarPresent.addEventListener('click', toggleFullscreen);

    // Welcome screen
    welcomeOpenBtn.addEventListener('click', function () { window.api.openFileDialog(); });

    // Error
    errorRetryBtn.addEventListener('click', function () {
      errorOverlay.classList.add('hidden');
      if (currentFilePath) openFile(currentFilePath);
    });
    errorCloseBtn.addEventListener('click', function () {
      errorOverlay.classList.add('hidden');
    });

    // LibreOffice screen
    downloadLoBtn.addEventListener('click', function () {
      window.api.openExternal('https://www.libreoffice.org/download/download-libreoffice/');
    });

    refreshLoBtn.addEventListener('click', async function () {
      var loPath = await window.api.detectLibreOffice();
      if (loPath) {
        libreOfficeReady = true;
        statusLoDot.classList.remove('missing');
        statusLoDot.classList.add('ready');
        statusLoText.textContent = 'LibreOffice Ready';
        showWelcome();
        loadRecentFiles();
      }
    });

    loPathSaveBtn.addEventListener('click', async function () {
      var p = loPathInput.value.trim();
      if (p) {
        await window.api.setLibreOfficePath(p);
        var loPath = await window.api.detectLibreOffice();
        if (loPath) {
          libreOfficeReady = true;
          statusLoDot.classList.remove('missing');
          statusLoDot.classList.add('ready');
          statusLoText.textContent = 'LibreOffice Ready';
          showWelcome();
          loadRecentFiles();
        }
      }
    });

    // Keyboard navigation
    document.addEventListener('keydown', function (e) {
      if (totalSlides === 0) return;

      switch (e.key) {
        case 'ArrowLeft':
        case 'PageUp':
          e.preventDefault();
          prevSlide();
          break;
        case 'ArrowRight':
        case 'PageDown':
          e.preventDefault();
          nextSlide();
          break;
        case 'Escape':
          if (isFullscreen) {
            e.preventDefault();
            toggleFullscreen();
          }
          break;
        case 'F5':
          e.preventDefault();
          toggleFullscreen();
          break;
      }
    });

    // Click on main slide to advance
    slideCanvas.addEventListener('click', function () {
      if (totalSlides > 0) nextSlide();
    });

    // Ctrl+Scroll zoom
    document.getElementById('main-view').addEventListener('wheel', function (e) {
      if (e.ctrlKey && totalSlides > 0) {
        e.preventDefault();
        if (e.deltaY < 0) zoomIn();
        else zoomOut();
      }
    }, { passive: false });

    // Window resize
    window.addEventListener('resize', function () {
      clearTimeout(resizeTimeout);
      resizeTimeout = setTimeout(function () {
        if (totalSlides > 0) renderCurrentSlide();
      }, 150);
    });

    // Drag & drop
    document.body.addEventListener('dragover', function (e) {
      e.preventDefault();
      e.stopPropagation();
    });

    document.body.addEventListener('drop', function (e) {
      e.preventDefault();
      e.stopPropagation();
      handleDrop(e);
    });

    dropZone.addEventListener('dragover', function (e) {
      e.preventDefault();
      dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', function () {
      dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', function (e) {
      e.preventDefault();
      dropZone.classList.remove('drag-over');
      handleDrop(e);
    });

    // IPC events from main process
    window.api.onOpenFile(function (filePath) { openFile(filePath); });
    window.api.onToggleFullscreen(function () { toggleFullscreen(); });
    window.api.onZoomIn(function () { zoomIn(); });
    window.api.onZoomOut(function () { zoomOut(); });
    window.api.onZoomReset(function () { zoomReset(); });
  }

  function handleDrop(e) {
    var files = e.dataTransfer.files;
    if (files.length > 0) {
      var file = files[0];
      var ext = file.name.split('.').pop().toLowerCase();
      if (ext === 'pptx' || ext === 'ppt') {
        openFile(file.path);
      }
    }
  }

  // Start
  init();
})();
