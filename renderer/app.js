(function () {
  'use strict';

  var PdfRenderer = window.PdfRenderer;

  // ---- Multi-tab state ----
  var tabs = [];
  var activeTabId = null;
  var tabIdCounter = 0;
  var sidebarVisible = true;
  var barsCollapsed = false;

  // Drag reorder state
  var dragTabId = null;

  // Pan/drag state for zoomed slides
  var isPanning = false;
  var panStartX = 0;
  var panStartY = 0;
  var scrollStartX = 0;
  var scrollStartY = 0;
  var didPan = false;

  // Global
  var isFullscreen = false;
  var fsCounterTimeout = null;
  var resizeTimeout = null;
  var libreOfficeReady = false;

  // File open queue — serialize opens to prevent PdfRenderer race conditions
  var openFileQueue = Promise.resolve();

  // DOM Elements
  var textLayer = document.getElementById('text-layer');

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

  function getTabIndex(tabId) {
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].id === tabId) return i;
    }
    return -1;
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

  // ---- Tab bar rendering (with drag-to-reorder) ----
  function renderTabs() {
    tabList.innerHTML = '';
    for (var i = 0; i < tabs.length; i++) {
      (function (tab, idx) {
        var el = document.createElement('div');
        el.className = 'tab-item' + (tab.id === activeTabId ? ' active' : '');
        el.draggable = true;
        el.dataset.tabId = tab.id;
        var loadingDot = tab.loading ? '<span class="tab-loading-dot"></span> ' : '';
        el.innerHTML =
          loadingDot +
          '<span class="tab-item-name">' + escapeHtml(tab.fileName) + '</span>' +
          '<span class="tab-close" title="Close tab">&times;</span>';

        // Click to switch
        el.addEventListener('click', function (e) {
          if (e.target.classList.contains('tab-close')) return;
          switchTab(tab.id);
        });

        // Close
        el.querySelector('.tab-close').addEventListener('click', function (e) {
          e.stopPropagation();
          closeTab(tab.id);
        });

        // Drag start
        el.addEventListener('dragstart', function (e) {
          dragTabId = tab.id;
          el.classList.add('dragging');
          e.dataTransfer.effectAllowed = 'move';
          e.dataTransfer.setData('text/plain', ''); // required for Firefox
        });

        el.addEventListener('dragend', function () {
          el.classList.remove('dragging');
          dragTabId = null;
          clearDragIndicators();
        });

        // Drag over — show drop indicator
        el.addEventListener('dragover', function (e) {
          if (dragTabId === null || dragTabId === tab.id) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = 'move';

          clearDragIndicators();
          var rect = el.getBoundingClientRect();
          var midX = rect.left + rect.width / 2;
          if (e.clientX < midX) {
            el.classList.add('drag-over-left');
          } else {
            el.classList.add('drag-over-right');
          }
        });

        el.addEventListener('dragleave', function () {
          el.classList.remove('drag-over-left');
          el.classList.remove('drag-over-right');
        });

        // Drop — reorder
        el.addEventListener('drop', function (e) {
          e.preventDefault();
          if (dragTabId === null || dragTabId === tab.id) return;

          var fromIdx = getTabIndex(dragTabId);
          var toIdx = getTabIndex(tab.id);
          if (fromIdx === -1 || toIdx === -1) return;

          // Determine if dropping before or after this tab
          var rect = el.getBoundingClientRect();
          var midX = rect.left + rect.width / 2;
          var insertBefore = e.clientX < midX;

          // Remove dragged tab from array
          var movedTab = tabs.splice(fromIdx, 1)[0];

          // Recalculate toIdx after removal
          var newToIdx = getTabIndex(tab.id);
          if (insertBefore) {
            tabs.splice(newToIdx, 0, movedTab);
          } else {
            tabs.splice(newToIdx + 1, 0, movedTab);
          }

          clearDragIndicators();
          dragTabId = null;
          renderTabs();
        });

        tabList.appendChild(el);
      })(tabs[i], i);
    }
  }

  function clearDragIndicators() {
    var items = tabList.querySelectorAll('.tab-item');
    items.forEach(function (item) {
      item.classList.remove('drag-over-left');
      item.classList.remove('drag-over-right');
    });
  }

  // ---- Tab management ----
  function switchTab(tabId) {
    if (activeTabId === tabId) return;

    activeTabId = tabId;
    var tab = getActiveTab();
    if (!tab) return;

    renderTabs();

    // Don't try to restore a tab that's still loading
    if (tab.loading) {
      thumbnailList.innerHTML = '';
      slideCanvas.width = 0;
      slideCanvas.height = 0;
      toolbarSlideCounter.textContent = 'Loading...';
      window.api.setTitle(tab.fileName + ' (loading) \u2014 PPT Viewer');
      return;
    }

    restoreTab(tab);
  }

  async function restoreTab(tab) {
    PdfRenderer.cleanup();
    await PdfRenderer.loadDocument(tab.pdfData);

    if (tab.thumbnailDataUrls && tab.thumbnailDataUrls.length === tab.totalSlides) {
      restoreThumbnailsFromCache(tab);
    } else {
      await renderAllThumbnails(tab);
    }

    updateSlideCounter();
    updateZoomDisplay();
    window.api.setTitle(tab.fileName + ' \u2014 PPT Viewer');
    await renderCurrentSlide();
  }

  function restoreThumbnailsFromCache(tab) {
    thumbnailList.innerHTML = '';
    for (var i = 0; i < tab.thumbnailDataUrls.length; i++) {
      var pageNum = i + 1;
      var item = document.createElement('div');
      item.className = 'thumbnail-item' + (pageNum === tab.currentSlide ? ' active' : '');
      item.dataset.page = pageNum;

      var img = document.createElement('img');
      img.src = tab.thumbnailDataUrls[i];
      img.style.width = '100%';
      img.style.height = 'auto';
      img.style.display = 'block';
      item.appendChild(img);

      var num = document.createElement('span');
      num.className = 'thumbnail-number';
      num.textContent = pageNum;
      item.appendChild(num);

      (function (pn) {
        item.addEventListener('click', function () { goToSlide(pn); });
      })(pageNum);

      thumbnailList.appendChild(item);
    }
  }

  function closeTab(tabId) {
    var idx = getTabIndex(tabId);
    if (idx === -1) return;

    tabs.splice(idx, 1);

    if (tabs.length === 0) {
      activeTabId = null;
      PdfRenderer.cleanup();
      showWelcome();
      loadRecentFiles();
      return;
    }

    if (activeTabId === tabId) {
      var newIdx = Math.min(idx, tabs.length - 1);
      activeTabId = tabs[newIdx].id;
      renderTabs();
      if (!tabs[newIdx].loading) {
        restoreTab(tabs[newIdx]);
      }
    } else {
      renderTabs();
    }
  }

  // ---- File opening ----
  function openFile(filePath) {
    var task = openFileQueue.then(function () { return doOpenFile(filePath); });
    openFileQueue = task.catch(function () {});
    return task;
  }

  async function doOpenFile(filePath) {
    var ext = filePath.split('.').pop().toLowerCase();
    var isPdf = (ext === 'pdf');
    if (!isPdf && !libreOfficeReady) return;

    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].filePath === filePath) {
        switchTab(tabs[i].id);
        return;
      }
    }

    var fileName = extractFileName(filePath);

    // Create the tab immediately with a loading state so UI stays responsive
    var tab = {
      id: ++tabIdCounter,
      filePath: filePath,
      fileName: fileName,
      pdfData: null,
      currentSlide: 1,
      totalSlides: 0,
      zoomLevel: 1,
      thumbnailDataUrls: [],
      loading: true,
    };

    tabs.push(tab);
    showViewer();
    renderTabs();

    // Only show full-page loading overlay if this is the only/active tab
    var isFirstTab = (tabs.length === 1);
    if (isFirstTab) {
      activeTabId = tab.id;
      loadingOverlay.classList.remove('hidden');
    }

    try {
      var pdfPath = await window.api.convertFile(filePath);
      var data = await window.api.readFile(pdfPath);

      // Tab may have been closed while loading
      if (getTabIndex(tab.id) === -1) return;

      tab.pdfData = data;
      tab.loading = false;

      // Load the document to get page count
      PdfRenderer.cleanup();
      var numPages = await PdfRenderer.loadDocument(data);
      tab.totalSlides = numPages;

      await window.api.saveRecentFile(filePath, fileName);

      // If this tab is active (or was the first), render it
      if (activeTabId === tab.id || isFirstTab) {
        activeTabId = tab.id;
        renderTabs();
        updateSlideCounter();
        updateZoomDisplay();
        window.api.setTitle(fileName + ' \u2014 PPT Viewer');
        await renderCurrentSlide();
        renderAllThumbnails(tab);
      } else {
        // Background tab finished loading — just update the tab indicator
        renderTabs();
      }
    } catch (err) {
      // Remove failed tab
      var failIdx = getTabIndex(tab.id);
      if (failIdx !== -1) tabs.splice(failIdx, 1);
      renderTabs();

      if (tabs.length === 0) {
        showWelcome();
        loadRecentFiles();
      }

      showError(err.message || String(err));
    } finally {
      if (isFirstTab) {
        loadingOverlay.classList.add('hidden');
      }
    }
  }

  function showError(msg) {
    errorMessage.textContent = msg;
    errorOverlay.classList.remove('hidden');
  }

  // ---- Slide rendering ----
  var renderTimer = null;
  var slideSpinner = null;

  async function renderCurrentSlide() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;

    if (!slideSpinner) slideSpinner = document.getElementById('slide-spinner');

    // Show spinner after a short delay so quick renders don't flicker
    clearTimeout(renderTimer);
    renderTimer = setTimeout(function () {
      slideSpinner.classList.remove('hidden');
    }, 80);

    var mainView = document.getElementById('main-view');
    var slideContainer = document.getElementById('slide-container');
    var maxWidth = mainView.clientWidth - 40;
    var maxHeight = mainView.clientHeight - 40;

    var tempCanvas = document.createElement('canvas');
    var tempResult = await PdfRenderer.renderSlide(tempCanvas, tab.currentSlide, 1);
    if (!tempResult) {
      clearTimeout(renderTimer);
      slideSpinner.classList.add('hidden');
      return;
    }

    var fitScale = Math.min(maxWidth / tempResult.width, maxHeight / tempResult.height);
    var effectiveZoom = fitScale * tab.zoomLevel;

    // Render to an offscreen canvas first, then copy to the visible canvas
    // in one step. This prevents the visual glitch where the old content is
    // cleared/resized before the new content is ready.
    var offscreen = document.createElement('canvas');
    await PdfRenderer.renderSlide(offscreen, tab.currentSlide, effectiveZoom);

    // Now swap: resize visible canvas and blit the finished frame
    slideCanvas.width = offscreen.width;
    slideCanvas.height = offscreen.height;
    slideCanvas.style.width = offscreen.style.width;
    slideCanvas.style.height = offscreen.style.height;
    var ctx = slideCanvas.getContext('2d');
    ctx.drawImage(offscreen, 0, 0);

    await PdfRenderer.renderTextLayer(textLayer, tab.currentSlide, effectiveZoom);

    clearTimeout(renderTimer);
    slideSpinner.classList.add('hidden');

    // When zoomed in beyond fit, add padding and enable pan cursor
    var canvasW = slideCanvas.offsetWidth;
    var canvasH = slideCanvas.offsetHeight;
    var isZoomedBeyondFit = canvasW > mainView.clientWidth || canvasH > mainView.clientHeight;
    slideContainer.style.padding = isZoomedBeyondFit ? '20px' : '0';
    if (isZoomedBeyondFit) {
      mainView.classList.add('pannable');
    } else {
      mainView.classList.remove('pannable');
    }

    updateSlideCounter();
    updateThumbnailHighlight();
  }

  async function renderAllThumbnails(tab) {
    thumbnailList.innerHTML = '';
    tab.thumbnailDataUrls = [];

    for (var i = 1; i <= tab.totalSlides; i++) {
      var item = document.createElement('div');
      item.className = 'thumbnail-item' + (i === tab.currentSlide ? ' active' : '');
      item.dataset.page = i;

      var canvas = await PdfRenderer.renderThumbnail(i, 160);
      if (canvas) {
        tab.thumbnailDataUrls.push(canvas.toDataURL('image/png'));
        item.appendChild(canvas);
      } else {
        tab.thumbnailDataUrls.push('');
      }

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
    var idx = getTabIndex(activeTabId);
    var next = (idx + 1) % tabs.length;
    switchTab(tabs[next].id);
  }

  function prevTab() {
    if (tabs.length <= 1) return;
    var idx = getTabIndex(activeTabId);
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
    var btn = document.getElementById('toolbar-clear-cache');
    var original = btn.innerHTML;
    btn.innerHTML = '<span>\u2713</span> Cleared';
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
    setTimeout(function () { renderCurrentSlide(); }, 50);
  }

  // ---- Collapse tab bar + toolbar ----
  function toggleBars() {
    barsCollapsed = !barsCollapsed;
    if (barsCollapsed) {
      viewer.classList.add('bars-collapsed');
    } else {
      viewer.classList.remove('bars-collapsed');
    }
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
    toolbarOpen.addEventListener('click', function () { window.api.openFileDialog(); });
    tabOpenBtn.addEventListener('click', function () { window.api.openFileDialog(); });
    toolbarToggleSidebar.addEventListener('click', toggleSidebar);
    document.getElementById('collapse-bars-btn').addEventListener('click', toggleBars);
    document.getElementById('collapse-restore-btn').addEventListener('click', toggleBars);
    toolbarPrev.addEventListener('click', prevSlide);
    toolbarNext.addEventListener('click', nextSlide);
    toolbarZoomIn.addEventListener('click', zoomIn);
    toolbarZoomOut.addEventListener('click', zoomOut);
    toolbarZoomReset.addEventListener('click', zoomReset);
    toolbarPresent.addEventListener('click', toggleFullscreen);
    document.getElementById('toolbar-clear-cache').addEventListener('click', clearCache);

    welcomeOpenBtn.addEventListener('click', function () { window.api.openFileDialog(); });

    errorRetryBtn.addEventListener('click', function () {
      errorOverlay.classList.add('hidden');
      var tab = getActiveTab();
      if (tab) openFile(tab.filePath);
    });
    errorCloseBtn.addEventListener('click', function () {
      errorOverlay.classList.add('hidden');
    });

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
      if (e.ctrlKey && (e.key === 'h' || e.key === 'H')) {
        e.preventDefault();
        toggleBars();
        return;
      }
      // Let the browser handle Ctrl+<key> combos we don't explicitly handle
      // (e.g. Ctrl+C for copy, Ctrl+A for select all)
      if (e.ctrlKey || e.metaKey) return;

      var tab = getActiveTab();
      if (!tab || tab.totalSlides === 0) return;

      // Don't hijack arrow keys when the user has a text selection active
      var sel = window.getSelection();
      if (sel && sel.toString().length > 0) return;

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

    // Click on slide area no longer advances slides — navigation is via
    // toolbar buttons, keyboard arrows, or thumbnail clicks only.

    // Pan with mouse drag when zoomed in
    var mainViewEl = document.getElementById('main-view');

    mainViewEl.addEventListener('mousedown', function (e) {
      if (!mainViewEl.classList.contains('pannable')) return;
      if (e.button !== 0) return; // left click only
      // Don't pan when the user is clicking on text — let the browser handle selection
      if (e.target !== textLayer && textLayer.contains(e.target)) return;
      isPanning = true;
      didPan = false;
      panStartX = e.clientX;
      panStartY = e.clientY;
      scrollStartX = mainViewEl.scrollLeft;
      scrollStartY = mainViewEl.scrollTop;
      mainViewEl.classList.add('panning');
      e.preventDefault();
    });

    window.addEventListener('mousemove', function (e) {
      if (!isPanning) return;
      var dx = e.clientX - panStartX;
      var dy = e.clientY - panStartY;
      if (Math.abs(dx) > 3 || Math.abs(dy) > 3) didPan = true;
      mainViewEl.scrollLeft = scrollStartX - dx;
      mainViewEl.scrollTop = scrollStartY - dy;
    });

    window.addEventListener('mouseup', function () {
      if (!isPanning) return;
      isPanning = false;
      mainViewEl.classList.remove('panning');
    });

    mainViewEl.addEventListener('wheel', function (e) {
      var tab = getActiveTab();
      if (e.ctrlKey && tab && tab.totalSlides > 0) {
        e.preventDefault();
        if (e.deltaY < 0) zoomIn();
        else zoomOut();
      }
    }, { passive: false });

    window.addEventListener('resize', function () {
      clearTimeout(resizeTimeout);
      resizeTimeout = setTimeout(function () {
        var tab = getActiveTab();
        if (tab && tab.totalSlides > 0) renderCurrentSlide();
      }, 150);
    });

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
      if (ext === 'pptx' || ext === 'ppt' || ext === 'pdf') {
        openFile(file.path);
      }
    }
  }

  // Start
  init();
})();
