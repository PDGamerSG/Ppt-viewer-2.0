(function () {
  'use strict';

  var PdfRenderer = window.PdfRenderer;

  // ---- Multi-tab state ----
  // Each tab: { id, filePath, fileName, pdfData, currentSlide, totalSlides, zoomLevel, thumbnailsHtml }
  var tabs = [];
  var activeTabId = null;
  var tabIdCounter = 0;
  var sidebarVisible = true;

  // Per-tab PDF documents (keyed by tab id)
  var tabDocs = {};

  // Global
  var isFullscreen = false;
  var fsCounterTimeout = null;
  var resizeTimeout = null;
  var libreOfficeReady = false;

  // DOM Elements
  var loadingOverlay = document.getElementById('loading-overlay');
  var errorOverlay = document.getElementById('error-overlay');
  var errorMessage = document.getElementById('error-message');
  var errorRetryBtn = document.getElementById('error-retry-btn');
  var errorCloseBtn = document.getElementById('error-close-btn');
  var loScreen = document.getElementById('libreoffice-screen');
  var welcomeScreen = document.getElementById('welcome-screen');
  var viewer = document.getElementById('viewer');
  var slideCanvas = document.getElementById('slide-canvas');
  var thumbnailList = document.getElementById('thumbnail-list');
  var recentList = document.getElementById('recent-list');
  var fsCounter = document.getElementById('fs-counter');

  // Tab bar
  var tabList = document.getElementById('tab-list');
  var tabOpenBtn = document.getElementById('tab-open-btn');

  // Toolbar
  var toolbarOpen = document.getElementById('toolbar-open');
  var toolbarToggleSidebar = document.getElementById('toolbar-toggle-sidebar');
  var toolbarPrev = document.getElementById('toolbar-prev');
  var toolbarNext = document.getElementById('toolbar-next');
  var toolbarSlideCounter = document.getElementById('toolbar-slide-counter');
  var toolbarZoomIn = document.getElementById('toolbar-zoom-in');
  var toolbarZoomOut = document.getElementById('toolbar-zoom-out');
  var toolbarZoomReset = document.getElementById('toolbar-zoom-reset');
  var toolbarZoom = document.getElementById('toolbar-zoom');
  var toolbarPresent = document.getElementById('toolbar-present');

  // Welcome screen
  var dropZone = document.getElementById('drop-zone');
  var welcomeOpenBtn = document.getElementById('welcome-open-btn');

  // LibreOffice screen
  var downloadLoBtn = document.getElementById('download-lo-btn');
  var refreshLoBtn = document.getElementById('refresh-lo-btn');
  var loPathInput = document.getElementById('lo-path-input');
  var loPathSaveBtn = document.getElementById('lo-path-save-btn');

  // ---- Helpers ----
  function getActiveTab() {
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].id === activeTabId) return tabs[i];
    }
    return null;
  }

  function escapeHtml(text) {
    var div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  function extractFileName(filePath) {
    return filePath.replace(/\\/g, '/').split('/').pop();
  }

  // ---- Init ----
  async function init() {
    var loPath = await window.api.detectLibreOffice();
    if (loPath) {
      libreOfficeReady = true;
      showWelcome();
    } else {
      libreOfficeReady = false;
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

  function showLibreOfficeScreen() { showScreen(loScreen); }
  function showWelcome() { showScreen(welcomeScreen); }
  function showViewer() { showScreen(viewer); }

  // ---- Recent files ----
  async function loadRecentFiles() {
    var files = await window.api.getRecentFiles();
    recentList.innerHTML = '';

    if (files.length === 0) {
      document.getElementById('recent-files').classList.add('hidden');
      return;
    }

    document.getElementById('recent-files').classList.remove('hidden');

    for (var i = 0; i < files.length; i++) {
      (function (file) {
        var li = document.createElement('li');
        var date = new Date(file.openedAt);
        var dateStr = date.toLocaleDateString();
        li.innerHTML =
          '<div>' +
            '<div class="recent-file-name">' + escapeHtml(file.name) + '</div>' +
            '<div class="recent-file-path">' + escapeHtml(file.path) + '</div>' +
          '</div>' +
          '<div class="recent-file-date">' + dateStr + '</div>';
        li.addEventListener('click', function () { openFile(file.path); });
        recentList.appendChild(li);
      })(files[i]);
    }
  }

  // ---- Tab bar rendering ----
  function renderTabs() {
    tabList.innerHTML = '';
    for (var i = 0; i < tabs.length; i++) {
      (function (tab) {
        var el = document.createElement('div');
        el.className = 'tab-item' + (tab.id === activeTabId ? ' active' : '');
        el.innerHTML =
          '<span class="tab-item-name">' + escapeHtml(tab.fileName) + '</span>' +
          '<span class="tab-close" title="Close tab">&times;</span>';

        // Click tab to switch
        el.addEventListener('click', function (e) {
          if (e.target.classList.contains('tab-close')) return;
          switchTab(tab.id);
        });

        // Close tab
        el.querySelector('.tab-close').addEventListener('click', function (e) {
          e.stopPropagation();
          closeTab(tab.id);
        });

        tabList.appendChild(el);
      })(tabs[i]);
    }
  }

  // ---- Tab management ----
  function switchTab(tabId) {
    if (activeTabId === tabId) return;

    activeTabId = tabId;
    var tab = getActiveTab();
    if (!tab) return;

    renderTabs();
    restoreTab(tab);
  }

  async function restoreTab(tab) {
    // Reload the PDF document for this tab
    PdfRenderer.cleanup();
    await PdfRenderer.loadDocument(tab.pdfData);

    // Always re-render thumbnails fresh (canvas pixels don't survive innerHTML)
    await renderAllThumbnails(tab);

    updateSlideCounter();
    updateZoomDisplay();
    window.api.setTitle(tab.fileName + ' \u2014 PPT Viewer');
    await renderCurrentSlide();
  }

  function closeTab(tabId) {
    var idx = -1;
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].id === tabId) { idx = i; break; }
    }
    if (idx === -1) return;

    tabs.splice(idx, 1);

    if (tabs.length === 0) {
      activeTabId = null;
      PdfRenderer.cleanup();
      showWelcome();
      loadRecentFiles();
      return;
    }

    // If we closed the active tab, switch to nearest
    if (activeTabId === tabId) {
      var newIdx = Math.min(idx, tabs.length - 1);
      activeTabId = tabs[newIdx].id;
      renderTabs();
      restoreTab(tabs[newIdx]);
    } else {
      renderTabs();
    }
  }

  // ---- File opening ----
  async function openFile(filePath) {
    if (!libreOfficeReady) return;

    // Check if already open in a tab
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].filePath === filePath) {
        switchTab(tabs[i].id);
        return;
      }
    }

    var fileName = extractFileName(filePath);
    loadingOverlay.classList.remove('hidden');

    try {
      var pdfPath = await window.api.convertFile(filePath);
      var data = await window.api.readFile(pdfPath);

      PdfRenderer.cleanup();
      var numPages = await PdfRenderer.loadDocument(data);

      var tab = {
        id: ++tabIdCounter,
        filePath: filePath,
        fileName: fileName,
        pdfData: data,
        currentSlide: 1,
        totalSlides: numPages,
        zoomLevel: 1,
      };

      tabs.push(tab);
      activeTabId = tab.id;

      await window.api.saveRecentFile(filePath, fileName);

      showViewer();
      renderTabs();
      updateSlideCounter();
      updateZoomDisplay();
      window.api.setTitle(fileName + ' \u2014 PPT Viewer');

      await renderCurrentSlide();
      renderAllThumbnails(tab);
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

  // ---- Slide rendering ----
  async function renderCurrentSlide() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;

    var mainView = document.getElementById('main-view');
    var maxWidth = mainView.clientWidth - 40;
    var maxHeight = mainView.clientHeight - 40;

    var tempCanvas = document.createElement('canvas');
    var tempResult = await PdfRenderer.renderSlide(tempCanvas, tab.currentSlide, 1);
    if (!tempResult) return;

    var fitScale = Math.min(maxWidth / tempResult.width, maxHeight / tempResult.height);
    var effectiveZoom = fitScale * tab.zoomLevel;

    await PdfRenderer.renderSlide(slideCanvas, tab.currentSlide, effectiveZoom);
    updateSlideCounter();
    updateThumbnailHighlight();
  }

  async function renderAllThumbnails(tab) {
    thumbnailList.innerHTML = '';

    for (var i = 1; i <= tab.totalSlides; i++) {
      var item = document.createElement('div');
      item.className = 'thumbnail-item' + (i === tab.currentSlide ? ' active' : '');
      item.dataset.page = i;

      var canvas = await PdfRenderer.renderThumbnail(i, 160);
      if (canvas) item.appendChild(canvas);

      var num = document.createElement('span');
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
    var tab = getActiveTab();
    if (!tab) return;

    var items = thumbnailList.querySelectorAll('.thumbnail-item');
    items.forEach(function (item) {
      var page = parseInt(item.dataset.page);
      if (page === tab.currentSlide) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });

    var active = thumbnailList.querySelector('.thumbnail-item.active');
    if (active) {
      active.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }

  // ---- Navigation ----
  function goToSlide(num) {
    var tab = getActiveTab();
    if (!tab || num < 1 || num > tab.totalSlides) return;
    tab.currentSlide = num;
    renderCurrentSlide();
    showFsCounter();
  }

  function prevSlide() {
    var tab = getActiveTab();
    if (tab) goToSlide(tab.currentSlide - 1);
  }

  function nextSlide() {
    var tab = getActiveTab();
    if (tab) goToSlide(tab.currentSlide + 1);
  }

  // ---- Zoom ----
  function setZoom(level) {
    var tab = getActiveTab();
    if (!tab) return;
    tab.zoomLevel = Math.max(0.25, Math.min(4, level));
    updateZoomDisplay();
    renderCurrentSlide();
  }

  function zoomIn() {
    var tab = getActiveTab();
    if (tab) setZoom(tab.zoomLevel + 0.1);
  }

  function zoomOut() {
    var tab = getActiveTab();
    if (tab) setZoom(tab.zoomLevel - 0.1);
  }

  function zoomReset() { setZoom(1); }

  function updateZoomDisplay() {
    var tab = getActiveTab();
    var pct = tab ? Math.round(tab.zoomLevel * 100) + '%' : '100%';
    toolbarZoom.textContent = pct;
  }

  // ---- Slide counter ----
  function updateSlideCounter() {
    var tab = getActiveTab();
    if (!tab) {
      toolbarSlideCounter.textContent = 'Slide 0 of 0';
      fsCounter.textContent = 'Slide 0 of 0';
      return;
    }
    var text = 'Slide ' + tab.currentSlide + ' of ' + tab.totalSlides;
    toolbarSlideCounter.textContent = text;
    fsCounter.textContent = text;
  }

  // ---- Tab cycling ----
  function nextTab() {
    if (tabs.length <= 1) return;
    var idx = -1;
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].id === activeTabId) { idx = i; break; }
    }
    var next = (idx + 1) % tabs.length;
    switchTab(tabs[next].id);
  }

  function prevTab() {
    if (tabs.length <= 1) return;
    var idx = -1;
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].id === activeTabId) { idx = i; break; }
    }
    var prev = (idx - 1 + tabs.length) % tabs.length;
    switchTab(tabs[prev].id);
  }

  function closeActiveTab() {
    if (activeTabId !== null) {
      closeTab(activeTabId);
    }
  }

  // ---- Clear cache ----
  async function clearCache() {
    await window.api.clearCache();
    // Brief visual feedback
    var btn = document.getElementById('toolbar-clear-cache');
    var original = btn.innerHTML;
    btn.innerHTML = '<span>✓</span> Cleared';
    setTimeout(function () { btn.innerHTML = original; }, 1500);
  }

  // ---- Sidebar toggle ----
  function toggleSidebar() {
    sidebarVisible = !sidebarVisible;
    var sidebar = document.getElementById('sidebar');
    if (sidebarVisible) {
      sidebar.classList.remove('sidebar-hidden');
    } else {
      sidebar.classList.add('sidebar-hidden');
    }
    // Re-render slide to fill new space
    setTimeout(function () { renderCurrentSlide(); }, 50);
  }

  // ---- Fullscreen ----
  async function toggleFullscreen() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;

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

  // ---- Event Listeners ----
  function setupEventListeners() {
    // Toolbar
    toolbarOpen.addEventListener('click', function () { window.api.openFileDialog(); });
    tabOpenBtn.addEventListener('click', function () { window.api.openFileDialog(); });
    toolbarToggleSidebar.addEventListener('click', toggleSidebar);
    toolbarPrev.addEventListener('click', prevSlide);
    toolbarNext.addEventListener('click', nextSlide);
    toolbarZoomIn.addEventListener('click', zoomIn);
    toolbarZoomOut.addEventListener('click', zoomOut);
    toolbarZoomReset.addEventListener('click', zoomReset);
    toolbarPresent.addEventListener('click', toggleFullscreen);
    document.getElementById('toolbar-clear-cache').addEventListener('click', clearCache);

    // Welcome screen
    welcomeOpenBtn.addEventListener('click', function () { window.api.openFileDialog(); });

    // Error
    errorRetryBtn.addEventListener('click', function () {
      errorOverlay.classList.add('hidden');
      var tab = getActiveTab();
      if (tab) openFile(tab.filePath);
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
          showWelcome();
          loadRecentFiles();
        }
      }
    });

    // Keyboard shortcuts
    document.addEventListener('keydown', function (e) {
      // Ctrl shortcuts — work even without a file open
      if (e.ctrlKey && e.key === 'Tab') {
        e.preventDefault();
        if (e.shiftKey) prevTab();
        else nextTab();
        return;
      }
      if (e.ctrlKey && (e.key === 'w' || e.key === 'W')) {
        e.preventDefault();
        closeActiveTab();
        return;
      }
      if (e.ctrlKey && (e.key === 'b' || e.key === 'B')) {
        e.preventDefault();
        toggleSidebar();
        return;
      }

      // Slide navigation — only when a file is open
      var tab = getActiveTab();
      if (!tab || tab.totalSlides === 0) return;

      switch (e.key) {
        case 'ArrowLeft':
        case 'ArrowUp':
        case 'PageUp':
          e.preventDefault();
          prevSlide();
          break;
        case 'ArrowRight':
        case 'ArrowDown':
        case 'PageDown':
          e.preventDefault();
          nextSlide();
          break;
        case 'Home':
          e.preventDefault();
          goToSlide(1);
          break;
        case 'End':
          e.preventDefault();
          goToSlide(tab.totalSlides);
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
      var tab = getActiveTab();
      if (tab && tab.totalSlides > 0) nextSlide();
    });

    // Ctrl+Scroll zoom
    document.getElementById('main-view').addEventListener('wheel', function (e) {
      var tab = getActiveTab();
      if (e.ctrlKey && tab && tab.totalSlides > 0) {
        e.preventDefault();
        if (e.deltaY < 0) zoomIn();
        else zoomOut();
      }
    }, { passive: false });

    // Window resize
    window.addEventListener('resize', function () {
      clearTimeout(resizeTimeout);
      resizeTimeout = setTimeout(function () {
        var tab = getActiveTab();
        if (tab && tab.totalSlides > 0) renderCurrentSlide();
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
    for (var i = 0; i < files.length; i++) {
      var file = files[i];
      var ext = file.name.split('.').pop().toLowerCase();
      if (ext === 'pptx' || ext === 'ppt') {
        openFile(file.path);
      }
    }
  }

  // Start
  init();
})();
