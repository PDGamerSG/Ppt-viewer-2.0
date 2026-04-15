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

  // Document dark mode state
  var docDarkMode = localStorage.getItem('pptviewer-doc-dark') === 'true';

  // Auto fit page state
  var autoFitPage = localStorage.getItem('pptviewer-auto-fit') === 'true';

  // Global
  var isFullscreen = false;
  var fsCounterTimeout = null;
  var resizeTimeout = null;
  var libreOfficeReady = false;

  // File open queue — serialize opens to prevent PdfRenderer race conditions
  var openFileQueue = Promise.resolve();

  // Continuous scroll state
  var renderedPages = new Set();
  var pageBaseDims = null;
  var scrollTimer = null;
  var renderGeneration = 0;

  // Navigation lock — when true, the scroll handler will NOT update
  // tab.currentSlide.  goToSlide sets the authoritative value and the
  // lock stays active until the user manually scrolls (wheel / trackpad).
  var navLock = false;

  // DOM Elements
  var loadingOverlay = document.getElementById('loading-overlay');
  var errorOverlay = document.getElementById('error-overlay');
  var errorMessage = document.getElementById('error-message');
  var errorRetryBtn = document.getElementById('error-retry-btn');
  var errorCloseBtn = document.getElementById('error-close-btn');
  var loScreen = document.getElementById('libreoffice-screen');
  var welcomeScreen = document.getElementById('welcome-screen');
  var viewer = document.getElementById('viewer');
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
  var toolbarPageInput = document.getElementById('toolbar-page-input');
  var toolbarPageTotal = document.getElementById('toolbar-page-total');
  var toolbarZoomIn = document.getElementById('toolbar-zoom-in');
  var toolbarZoomOut = document.getElementById('toolbar-zoom-out');
  var toolbarZoomReset = document.getElementById('toolbar-zoom-reset');
  var toolbarZoom = document.getElementById('toolbar-zoom');
  var toolbarPresent = document.getElementById('toolbar-present');
  var toolbarRotateLeft = document.getElementById('toolbar-rotate-left');
  var toolbarRotateRight = document.getElementById('toolbar-rotate-right');

  // Find bar
  var findInput = document.getElementById('find-input');
  var findPrevBtn = document.getElementById('find-prev');
  var findNextBtn = document.getElementById('find-next');
  var findCount = document.getElementById('find-count');

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

  // ---- Persist PDF scroll position ----
  var positionSaveTimer = null;

  function saveTabPosition(tab) {
    if (!tab || !tab.filePath || tab.loading || tab.totalSlides === 0) return;
    var mainView = document.getElementById('main-view');
    window.api.setFilePosition(tab.filePath, {
      currentSlide: tab.currentSlide,
      scrollTop: mainView ? mainView.scrollTop : 0,
      savedAt: Date.now(),
    });
  }

  function debouncedSavePosition() {
    if (positionSaveTimer) clearTimeout(positionSaveTimer);
    positionSaveTimer = setTimeout(function () {
      var tab = getActiveTab();
      if (tab) saveTabPosition(tab);
    }, 1000);
  }

  // ---- Document dark mode ----
  var themeToastTimeout = null;

  function applyDocDark() {
    // Apply to all page wrappers (main view)
    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');
    wrappers.forEach(function (w) {
      if (docDarkMode) w.classList.add('doc-dark');
      else w.classList.remove('doc-dark');
    });

    // Apply to all thumbnails
    var thumbs = document.querySelectorAll('.thumbnail-item');
    thumbs.forEach(function (t) {
      if (docDarkMode) t.classList.add('doc-dark');
      else t.classList.remove('doc-dark');
    });

    updateDocDarkButton();
  }

  var svgMoon = '<svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"><path d="M13.5 10.5A5.5 5.5 0 0 1 5.5 2.5a5.5 5.5 0 1 0 8 8z"/></svg>';
  var svgSun = '<svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"><circle cx="8" cy="8" r="3"/><line x1="8" y1="1.5" x2="8" y2="3"/><line x1="8" y1="13" x2="8" y2="14.5"/><line x1="1.5" y1="8" x2="3" y2="8"/><line x1="13" y1="8" x2="14.5" y2="8"/><line x1="3.5" y1="3.5" x2="4.5" y2="4.5"/><line x1="11.5" y1="11.5" x2="12.5" y2="12.5"/><line x1="12.5" y1="3.5" x2="11.5" y2="4.5"/><line x1="4.5" y1="11.5" x2="3.5" y2="12.5"/></svg>';

  function updateDocDarkButton() {
    var btn = document.getElementById('toolbar-doc-dark');
    var label = document.getElementById('doc-dark-label');
    var icon = document.getElementById('doc-dark-icon');
    if (!btn) return;

    if (docDarkMode) {
      btn.classList.add('doc-dark-active');
      if (icon) icon.innerHTML = svgSun;
      if (label) label.textContent = 'Normal';
    } else {
      btn.classList.remove('doc-dark-active');
      if (icon) icon.innerHTML = svgMoon;
      if (label) label.textContent = 'Dark Read';
    }
  }

  // ---- Auto fit page ----
  function getFitScale() {
    var mainView = document.getElementById('main-view');
    var availW = mainView.clientWidth - 40;
    if (!pageBaseDims) return 1;
    if (!autoFitPage) return availW / pageBaseDims.width;
    var availH = mainView.clientHeight - 20;
    return Math.min(availW / pageBaseDims.width, availH / pageBaseDims.height);
  }

  function updateAutoFitButton() {
    var btn = document.getElementById('toolbar-auto-fit');
    var menuItem = document.getElementById('hmenu-auto-fit');
    if (btn) {
      if (autoFitPage) btn.classList.add('auto-fit-active');
      else btn.classList.remove('auto-fit-active');
    }
    if (menuItem) {
      menuItem.querySelector('.hmenu-check').textContent = autoFitPage ? '✓' : '';
    }
  }

  function toggleAutoFit() {
    autoFitPage = !autoFitPage;
    localStorage.setItem('pptviewer-auto-fit', autoFitPage ? 'true' : 'false');
    updateAutoFitButton();
    refreshView();
  }

  function toggleDocDark() {
    docDarkMode = !docDarkMode;
    localStorage.setItem('pptviewer-doc-dark', docDarkMode ? 'true' : 'false');
    applyDocDark();
    showThemeToast(docDarkMode ? 'Document dark mode ON' : 'Document dark mode OFF');
  }

  function showThemeToast(message) {
    var toast = document.getElementById('theme-toast');
    toast.textContent = message;
    toast.classList.add('show');
    clearTimeout(themeToastTimeout);
    themeToastTimeout = setTimeout(function () {
      toast.classList.remove('show');
    }, 1200);
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
    updateAutoFitButton();
  }

  var universalTitlebar = document.getElementById('universal-titlebar');

  function showScreen(screen) {
    loScreen.classList.add('hidden');
    welcomeScreen.classList.add('hidden');
    viewer.classList.add('hidden');
    screen.classList.remove('hidden');

    // Show universal title bar on welcome/LO screens, hide on viewer (it has its own tab bar)
    if (screen === viewer) {
      universalTitlebar.classList.add('hidden');
    } else {
      universalTitlebar.classList.remove('hidden');
    }
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

    // Save outgoing tab's position and DOM snapshot
    var outgoing = getActiveTab();
    if (outgoing && !outgoing.loading) {
      var slideContainer = document.getElementById('slide-container');
      var mainView = document.getElementById('main-view');
      // Capture scroll position BEFORE moving DOM — once content is removed
      // the browser clamps scrollTop to 0
      var savedScrollTop = mainView.scrollTop;
      saveTabPosition(outgoing);
      // Move the live DOM nodes into a DocumentFragment (cheap, no cloning)
      var frag = document.createDocumentFragment();
      while (slideContainer.firstChild) {
        frag.appendChild(slideContainer.firstChild);
      }
      outgoing.domSnapshot = frag;
      outgoing.snapshotScrollTop = savedScrollTop;
    }

    activeTabId = tabId;
    var tab = getActiveTab();
    if (!tab) return;

    renderTabs();

    // Don't try to restore a tab that's still loading
    if (tab.loading) {
      thumbnailList.innerHTML = '';
      document.getElementById('slide-container').innerHTML = '';
      toolbarPageInput.value = '';
      toolbarPageTotal.textContent = '…';
      window.api.setTitle(tab.fileName + ' (loading) \u2014 PPT Viewer');
      return;
    }

    restoreTab(tab);
  }

  async function restoreTab(tab) {
    // Clear find state for the new tab
    findInput.value = '';
    findTerm = '';
    findMatches = [];
    findCurrentIdx = -1;
    findCount.textContent = '';
    findInput.classList.remove('find-has-results', 'find-no-results');

    // Activate the document (no re-parsing needed if already loaded)
    if (PdfRenderer.isLoaded(tab.id)) {
      PdfRenderer.setActive(tab.id);
    } else {
      await PdfRenderer.loadDocument(tab.pdfData, tab.id);
    }

    updateSlideCounter();
    updateZoomDisplay();
    updateRotateButtons();
    window.api.setTitle(tab.fileName + ' \u2014 PPT Viewer');

    // Restore main view FIRST so the user sees slides immediately
    var slideContainer = document.getElementById('slide-container');
    var mainView = document.getElementById('main-view');

    // Lock scroll handler during restore — prevents updateCurrentPageFromScroll
    // from resetting currentSlide while we're restoring the view
    navLock = true;

    if (tab.domSnapshot) {
      slideContainer.innerHTML = '';
      slideContainer.appendChild(tab.domSnapshot);
      tab.domSnapshot = null;

      pageBaseDims = await PdfRenderer.getPageDimensions(1, tab.pageRotation);
      renderedPages.clear();
      textLayersRendered.clear();
      var wrappers = slideContainer.querySelectorAll('.page-wrapper');
      for (var w = 0; w < wrappers.length; w++) {
        var c = wrappers[w].querySelector('.page-canvas');
        if (c && c.width > 0) renderedPages.add(parseInt(wrappers[w].dataset.page));
      }
      renderGeneration++;
      updatePannable();
      mainView.scrollTop = tab.snapshotScrollTop;
      applyDocDark();
      renderVisiblePages();
    } else {
      // setupContinuousView scrolls to tab.currentSlide before rendering
      await setupContinuousView(tab);
    }

    // Release after a short delay so the programmatic scroll settles
    setTimeout(function () { navLock = false; }, 300);

    // Thumbnails AFTER main view — don't block the slide display
    if (tab.thumbnailDataUrls && tab.thumbnailDataUrls.length === tab.totalSlides) {
      restoreThumbnailsFromCache(tab);
    } else {
      renderAllThumbnails(tab);
    }
  }

  function restoreThumbnailsFromCache(tab) {
    // Verify we actually have valid cached thumbnails, not just empty placeholders
    var hasValid = tab.thumbnailDataUrls.some(function (url) { return url && url.length > 0; });
    if (!hasValid) {
      renderAllThumbnails(tab);
      return;
    }

    thumbnailList.innerHTML = '';
    for (var i = 0; i < tab.thumbnailDataUrls.length; i++) {
      var pageNum = i + 1;
      var item = document.createElement('div');
      item.className = 'thumbnail-item' + (pageNum === tab.currentSlide ? ' active' : '');
      item.dataset.page = pageNum;

      // Only create img if we have a valid data URL
      if (tab.thumbnailDataUrls[i]) {
        var img = document.createElement('img');
        img.src = tab.thumbnailDataUrls[i];
        img.style.width = '100%';
        img.style.height = 'auto';
        img.style.display = 'block';
        item.appendChild(img);
      }

      var num = document.createElement('span');
      num.className = 'thumbnail-number';
      num.textContent = pageNum;
      item.appendChild(num);

      (function (pn) {
        item.addEventListener('click', function () { goToSlide(pn); });
      })(pageNum);

      thumbnailList.appendChild(item);
    }
    applyDocDark();
  }

  function closeTab(tabId) {
    var idx = getTabIndex(tabId);
    if (idx === -1) return;

    // Save position before closing
    saveTabPosition(tabs[idx]);

    // Clean up only this tab's document
    PdfRenderer.cleanupDoc(tabId);
    tabs[idx].domSnapshot = null;
    tabs[idx].canvasCache = {};
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
      pageRotation: 0,
      thumbnailDataUrls: [],
      loading: true,
      // Canvas cache: pageNum -> { canvas, zoom, rotation }
      canvasCache: {},
      // Snapshot of the slide container DOM for instant restore
      domSnapshot: null,
      snapshotScrollTop: 0,
    };

    tabs.push(tab);
    showViewer();
    renderTabs();

    // Only show full-page loading overlay if this is the only/active tab
    var isFirstTab = (tabs.length === 1);
    if (isFirstTab) {
      activeTabId = tab.id;
      if (loadingMessage) loadingMessage.textContent = 'Loading…';
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
      if (loadingMessage) loadingMessage.textContent = progressMessages.rendering;
      var numPages = await PdfRenderer.loadDocument(data, tab.id);
      tab.totalSlides = numPages;

      // Save recent file in background — don't block rendering
      window.api.saveRecentFile(filePath, fileName);

      // If this tab is active (or was the first), render it
      if (activeTabId === tab.id || isFirstTab) {
        activeTabId = tab.id;

        // Restore saved page position BEFORE rendering so the correct
        // region is rendered first instead of always starting at page 1
        var savedPos = await window.api.getFilePosition(filePath);
        if (savedPos && savedPos.currentSlide > 1 && savedPos.currentSlide <= tab.totalSlides) {
          tab.currentSlide = savedPos.currentSlide;
        }

        renderTabs();
        updateSlideCounter();
        updateZoomDisplay();
        window.api.setTitle(fileName + ' \u2014 PPT Viewer');
        await setupContinuousView(tab);

        // Thumbnails render in background — defer to not block initial slide display
        setTimeout(function () { renderAllThumbnails(tab); }, 100);
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

  // ---- Continuous scroll rendering ----

  async function setupContinuousView(tab) {
    var slideContainer = document.getElementById('slide-container');
    renderedPages.clear();
    textLayersRendered.clear();
    renderGeneration++;
    pageBaseDims = null;

    if (tab.totalSlides === 0) {
      slideContainer.innerHTML = '';
      return;
    }

    // Get base dimensions from first page (all PPT slides share dimensions)
    pageBaseDims = await PdfRenderer.getPageDimensions(1, tab.pageRotation);
    if (!pageBaseDims) return;

    var mainView = document.getElementById('main-view');
    var fitScale = getFitScale();
    var effectiveZoom = fitScale * tab.zoomLevel;
    var cssWidth = pageBaseDims.width * effectiveZoom;
    var cssHeight = pageBaseDims.height * effectiveZoom;

    // Build all wrappers off-screen in a fragment, then swap in one shot
    var frag = document.createDocumentFragment();
    for (var i = 1; i <= tab.totalSlides; i++) {
      var wrapper = document.createElement('div');
      wrapper.className = 'page-wrapper';
      wrapper.dataset.page = i;
      wrapper.style.width = cssWidth + 'px';
      wrapper.style.height = cssHeight + 'px';

      var canvas = document.createElement('canvas');
      canvas.className = 'page-canvas';
      wrapper.appendChild(canvas);

      var textLayerDiv = document.createElement('div');
      textLayerDiv.className = 'page-text-layer';
      wrapper.appendChild(textLayerDiv);

      frag.appendChild(wrapper);
    }
    slideContainer.innerHTML = '';
    slideContainer.appendChild(frag);

    // If tab has a remembered page, scroll there BEFORE rendering
    // so renderVisiblePages targets the correct viewport region
    if (tab.currentSlide > 1) {
      var targetWrapper = slideContainer.querySelector('.page-wrapper[data-page="' + tab.currentSlide + '"]');
      if (targetWrapper) {
        targetWrapper.scrollIntoView({ block: 'start' });
      }
    }

    updatePannable();
    await renderVisiblePages();
    updateSlideCounter();
    updateThumbnailHighlight();
    applyDocDark();
  }

  function updatePannable() {
    var mainView = document.getElementById('main-view');
    var slideContainer = document.getElementById('slide-container');
    var isWider = slideContainer.scrollWidth > mainView.clientWidth;
    if (isWider) {
      mainView.classList.add('pannable');
    } else {
      mainView.classList.remove('pannable');
    }
  }

  async function renderVisiblePages() {
    var tab = getActiveTab();
    if (!tab || !pageBaseDims) return;

    var gen = renderGeneration;
    var mainView = document.getElementById('main-view');
    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');

    var viewTop = mainView.scrollTop;
    var viewBottom = viewTop + mainView.clientHeight;
    var buffer = mainView.clientHeight; // 1 screen buffer above and below

    var fitScale = getFitScale();
    var effectiveZoom = fitScale * tab.zoomLevel;

    // Collect pages that need rendering
    var toRender = [];
    for (var i = 0; i < wrappers.length; i++) {
      var wrapper = wrappers[i];
      var wTop = wrapper.offsetTop;
      var wBottom = wTop + wrapper.offsetHeight;
      var pageNum = parseInt(wrapper.dataset.page);

      if (wBottom >= viewTop - buffer && wTop <= viewBottom + buffer) {
        if (!renderedPages.has(pageNum)) {
          renderedPages.add(pageNum);
          toRender.push({ wrapper: wrapper, pageNum: pageNum });
        }
      }
    }

    // Render pages in parallel batches of 5
    var BATCH = 5;
    for (var s = 0; s < toRender.length; s += BATCH) {
      if (gen !== renderGeneration) return;
      var batch = toRender.slice(s, s + BATCH);
      await Promise.all(batch.map(function (item) {
        return (async function () {
          try {
            var canvas = item.wrapper.querySelector('.page-canvas');

            await PdfRenderer.renderSlide(canvas, item.pageNum, effectiveZoom, tab.pageRotation);

            // Generation changed — discard this render and allow retry
            if (gen !== renderGeneration) {
              renderedPages.delete(item.pageNum);
              return;
            }

            // Guard: if canvas has no content, allow retry
            if (canvas.width === 0 || canvas.height === 0) {
              renderedPages.delete(item.pageNum);
              return;
            }

            if (docDarkMode) applyDocDark();
          } catch (err) {
            renderedPages.delete(item.pageNum);
          }
        })();
      }));
    }

    // Schedule text layer rendering in background after canvases are painted
    scheduleTextLayerRender();
  }

  function updateCurrentPageFromScroll() {
    // Skip if an explicit goToSlide is still animating — its target is authoritative
    if (navLock) return;

    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;

    var mainView = document.getElementById('main-view');
    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');
    if (wrappers.length === 0) return;

    var viewMid = mainView.scrollTop + mainView.clientHeight / 2;

    var currentPage = 1;
    for (var i = 0; i < wrappers.length; i++) {
      var wrapper = wrappers[i];
      if (wrapper.offsetTop + wrapper.offsetHeight / 2 >= viewMid) {
        currentPage = i + 1;
        break;
      }
      currentPage = i + 1;
    }

    if (tab.currentSlide !== currentPage) {
      tab.currentSlide = currentPage;
      updateSlideCounter();
      updateThumbnailHighlight();
    }
  }

  // Recalculate page sizes and re-render visible pages (used after zoom, resize, layout changes)
  async function refreshView() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;

    // Invalidate DOM snapshot since layout changed
    tab.domSnapshot = null;

    // Text layers will be rebuilt — old find match refs are now stale
    findMatches = [];
    findCurrentIdx = -1;
    updateFindCount();

    // Re-fetch dimensions in case rotation changed
    pageBaseDims = await PdfRenderer.getPageDimensions(1, tab.pageRotation);
    if (!pageBaseDims) return;

    renderedPages.clear();
    textLayersRendered.clear();
    renderGeneration++;

    var mainView = document.getElementById('main-view');
    var fitScale = getFitScale();
    var effectiveZoom = fitScale * tab.zoomLevel;
    var cssWidth = pageBaseDims.width * effectiveZoom;
    var cssHeight = pageBaseDims.height * effectiveZoom;

    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');
    for (var i = 0; i < wrappers.length; i++) {
      wrappers[i].style.width = cssWidth + 'px';
      wrappers[i].style.height = cssHeight + 'px';
      // Clear rendered content so it re-renders at new zoom
      var canvas = wrappers[i].querySelector('.page-canvas');
      canvas.width = 0;
      canvas.height = 0;
      wrappers[i].querySelector('.page-text-layer').innerHTML = '';
    }

    updatePannable();
    renderVisiblePages();
  }

  async function renderAllThumbnails(tab) {
    thumbnailList.innerHTML = '';
    tab.thumbnailDataUrls = new Array(tab.totalSlides).fill('');

    // Create all placeholder items first for instant layout
    var items = [];
    var frag = document.createDocumentFragment();
    for (var i = 1; i <= tab.totalSlides; i++) {
      var item = document.createElement('div');
      item.className = 'thumbnail-item' + (i === tab.currentSlide ? ' active' : '');
      item.dataset.page = i;

      var num = document.createElement('span');
      num.className = 'thumbnail-number';
      num.textContent = i;
      item.appendChild(num);

      (function (pageNum) {
        item.addEventListener('click', function () { goToSlide(pageNum); });
      })(i);

      items.push(item);
      frag.appendChild(item);
    }
    thumbnailList.appendChild(frag);

    // Render thumbnails in parallel batches of 4 for speed
    var BATCH_SIZE = 4;
    for (var start = 0; start < tab.totalSlides; start += BATCH_SIZE) {
      var end = Math.min(start + BATCH_SIZE, tab.totalSlides);
      var promises = [];
      for (var j = start; j < end; j++) {
        (function (idx) {
          promises.push(
            PdfRenderer.renderThumbnail(idx + 1, 140, tab.pageRotation).then(function (canvas) {
              if (canvas) {
                tab.thumbnailDataUrls[idx] = canvas.toDataURL('image/jpeg', 0.7);
                items[idx].insertBefore(canvas, items[idx].firstChild);
              }
            }).catch(function () {})
          );
        })(j);
      }
      await Promise.all(promises);
      // Bail out if tab was closed or switched during rendering
      if (getTabIndex(tab.id) === -1 || activeTabId !== tab.id) return;
    }
    applyDocDark();
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
      active.scrollIntoView({ block: 'nearest' });
    }
  }

  // ---- Navigation ----
  function goToSlide(num) {
    var tab = getActiveTab();
    if (!tab || num < 1 || num > tab.totalSlides) return;
    tab.currentSlide = num;

    // Lock: the scroll handler must NOT touch currentSlide until the user
    // manually scrolls (wheel/trackpad).  This prevents scrollIntoView's
    // intermediate scroll positions from corrupting the page counter.
    navLock = true;

    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');
    if (wrappers[num - 1]) {
      wrappers[num - 1].scrollIntoView({ block: 'start' });
    }

    updateSlideCounter();
    updateThumbnailHighlight();
    showFsCounter();

    // Render the target page and its neighbours if not yet rendered
    renderVisiblePages();
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
    if (autoFitPage) {
      autoFitPage = false;
      localStorage.setItem('pptviewer-auto-fit', 'false');
      updateAutoFitButton();
    }
    tab.zoomLevel = Math.max(0.25, Math.min(4, level));
    updateZoomDisplay();
    refreshView();
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

  // ---- Rotation ----
  function rotateLeft() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;
    tab.pageRotation = ((tab.pageRotation - 90) + 360) % 360;
    tab.thumbnailDataUrls = []; // invalidate thumbnail cache
    updateRotateButtons();
    refreshView();
    renderAllThumbnails(tab);
  }

  function rotateRight() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;
    tab.pageRotation = (tab.pageRotation + 90) % 360;
    tab.thumbnailDataUrls = [];
    updateRotateButtons();
    refreshView();
    renderAllThumbnails(tab);
  }

  function updateRotateButtons() {
    var tab = getActiveTab();
    var rotated = tab && tab.pageRotation !== 0;
    toolbarRotateLeft.classList.toggle('rotated', rotated);
    toolbarRotateRight.classList.toggle('rotated', rotated);
  }

  // ---- Find text ----
  // ---- Lazy text layer rendering ----
  // Text layers are expensive — defer until the user needs them (find, text select).
  // After initial page render, we render text layers in the background with idle callbacks.
  var textLayersRendered = new Set();
  var textLayerIdleTimer = null;

  async function ensureTextLayers() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0) return;

    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');
    var fitScale = getFitScale();
    var effectiveZoom = fitScale * tab.zoomLevel;
    var pending = [];

    for (var i = 0; i < wrappers.length; i++) {
      var pageNum = parseInt(wrappers[i].dataset.page);
      if (!textLayersRendered.has(pageNum)) {
        pending.push({ wrapper: wrappers[i], pageNum: pageNum });
      }
    }

    for (var j = 0; j < pending.length; j++) {
      var textLayerEl = pending[j].wrapper.querySelector('.page-text-layer');
      if (textLayerEl) {
        try {
          await PdfRenderer.renderTextLayer(textLayerEl, pending[j].pageNum, effectiveZoom, tab.pageRotation);
          textLayersRendered.add(pending[j].pageNum);
        } catch (_) {}
      }
    }
  }

  // Render text layers for visible pages in background after initial render
  function scheduleTextLayerRender() {
    if (textLayerIdleTimer) clearTimeout(textLayerIdleTimer);
    textLayerIdleTimer = setTimeout(function () {
      renderVisibleTextLayers();
    }, 500);
  }

  async function renderVisibleTextLayers() {
    var tab = getActiveTab();
    if (!tab || tab.totalSlides === 0 || !pageBaseDims) return;

    var mainView = document.getElementById('main-view');
    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');
    var viewTop = mainView.scrollTop;
    var viewBottom = viewTop + mainView.clientHeight;
    var buffer = mainView.clientHeight;
    var fitScale = getFitScale();
    var effectiveZoom = fitScale * tab.zoomLevel;

    for (var i = 0; i < wrappers.length; i++) {
      var wrapper = wrappers[i];
      var wTop = wrapper.offsetTop;
      var wBottom = wTop + wrapper.offsetHeight;
      var pageNum = parseInt(wrapper.dataset.page);

      if (wBottom >= viewTop - buffer && wTop <= viewBottom + buffer) {
        if (!textLayersRendered.has(pageNum)) {
          var textLayerEl = wrapper.querySelector('.page-text-layer');
          if (textLayerEl) {
            try {
              await PdfRenderer.renderTextLayer(textLayerEl, pageNum, effectiveZoom, tab.pageRotation);
              textLayersRendered.add(pageNum);
            } catch (_) {}
          }
        }
      }
    }
    if (findTerm) applyFindHighlights();
  }

  var findTerm = '';
  var findMatches = [];
  var findCurrentIdx = -1;
  var findDebounceTimer = null;

  function openFind() {
    findInput.focus();
    findInput.select();
  }

  async function runFind() {
    clearFindHighlights();
    findMatches = [];
    findCurrentIdx = -1;

    findTerm = findInput.value;
    if (!findTerm) {
      updateFindCount();
      findInput.classList.remove('find-has-results', 'find-no-results');
      return;
    }

    // Ensure all text layers are rendered before searching
    await ensureTextLayers();

    var term = findTerm.toLowerCase();
    var spans = document.querySelectorAll('.page-text-layer span');

    spans.forEach(function (span) {
      if (!span.textContent) return;
      if (span.textContent.toLowerCase().indexOf(term) !== -1) {
        span.classList.add('find-match');
        findMatches.push(span);
      }
    });

    if (findMatches.length > 0) {
      findInput.classList.add('find-has-results');
      findInput.classList.remove('find-no-results');
      findCurrentIdx = 0;
      activateFindMatch(0);
    } else {
      findInput.classList.remove('find-has-results');
      findInput.classList.add('find-no-results');
    }

    updateFindCount();
  }

  function clearFindHighlights() {
    document.querySelectorAll('.page-text-layer span.find-match').forEach(function (el) {
      el.classList.remove('find-match', 'find-match-active');
    });
  }

  // Called after new pages render — re-apply highlights to newly rendered spans
  function applyFindHighlights() {
    if (!findTerm) return;
    var term = findTerm.toLowerCase();
    var spans = document.querySelectorAll('.page-text-layer span');
    spans.forEach(function (span) {
      if (!span.textContent) return;
      if (span.textContent.toLowerCase().indexOf(term) !== -1) {
        if (!span.classList.contains('find-match')) {
          span.classList.add('find-match');
          findMatches.push(span);
        }
      }
    });
    // Re-highlight active
    if (findCurrentIdx >= 0 && findMatches[findCurrentIdx]) {
      findMatches[findCurrentIdx].classList.add('find-match-active');
    }
    updateFindCount();
  }

  function activateFindMatch(idx) {
    if (findCurrentIdx >= 0 && findMatches[findCurrentIdx]) {
      findMatches[findCurrentIdx].classList.remove('find-match-active');
    }
    findCurrentIdx = idx;
    if (findMatches[findCurrentIdx]) {
      findMatches[findCurrentIdx].classList.add('find-match-active');
      findMatches[findCurrentIdx].scrollIntoView({ block: 'center' });
    }
    updateFindCount();
  }

  function findNext() {
    if (!findMatches.length) { runFind(); return; }
    activateFindMatch((findCurrentIdx + 1) % findMatches.length);
  }

  function findPrev() {
    if (!findMatches.length) { runFind(); return; }
    activateFindMatch((findCurrentIdx - 1 + findMatches.length) % findMatches.length);
  }

  function updateFindCount() {
    if (!findTerm) {
      findCount.textContent = '';
      return;
    }
    if (findMatches.length === 0) {
      findCount.textContent = 'No results';
      return;
    }
    findCount.textContent = (findCurrentIdx + 1) + ' / ' + findMatches.length;
  }

  // ---- Slide counter ----
  function updateSlideCounter() {
    var tab = getActiveTab();
    if (!tab) {
      toolbarPageInput.value = 0;
      toolbarPageTotal.textContent = 0;
      fsCounter.textContent = 'Slide 0 of 0';
      return;
    }
    toolbarPageInput.value = tab.currentSlide;
    toolbarPageInput.max = tab.totalSlides;
    toolbarPageTotal.textContent = tab.totalSlides;
    fsCounter.textContent = 'Slide ' + tab.currentSlide + ' of ' + tab.totalSlides;
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
    showThemeToast('Cache cleared');
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
    requestAnimationFrame(function () { refreshView(); });
  }

  // ---- Collapse tab bar + toolbar ----
  function toggleBars() {
    barsCollapsed = !barsCollapsed;
    if (barsCollapsed) {
      viewer.classList.add('bars-collapsed');
    } else {
      viewer.classList.remove('bars-collapsed');
    }
    requestAnimationFrame(function () { refreshView(); });
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

    requestAnimationFrame(function () { refreshView(); });
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
  var loadingMessage = document.getElementById('loading-message');
  var progressMessages = {
    converting: 'Converting presentation…',
    loading: 'Loading PDF…',
    rendering: 'Rendering slides…',
  };

  window.api.onConvertProgress(function (stage) {
    if (loadingMessage && progressMessages[stage]) {
      loadingMessage.textContent = progressMessages[stage];
    }
  });

  function setupEventListeners() {
    // Window controls (viewer tab bar)
    document.getElementById('win-minimize').addEventListener('click', function () { window.api.windowMinimize(); });
    document.getElementById('win-maximize').addEventListener('click', function () { window.api.windowMaximize(); });
    document.getElementById('win-close').addEventListener('click', function () { window.api.windowClose(); });

    // Universal title bar window controls (welcome/LO screens)
    document.getElementById('utb-minimize').addEventListener('click', function () { window.api.windowMinimize(); });
    document.getElementById('utb-maximize').addEventListener('click', function () { window.api.windowMaximize(); });
    document.getElementById('utb-close').addEventListener('click', function () { window.api.windowClose(); });

    // Hamburger menu toggle
    var hamburgerMenu = document.getElementById('hamburger-menu');
    document.getElementById('toolbar-hamburger').addEventListener('click', function (e) {
      e.stopPropagation();
      hamburgerMenu.classList.toggle('hidden');
    });
    // Close dropdown when clicking anywhere else
    document.addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
    });
    hamburgerMenu.addEventListener('click', function (e) { e.stopPropagation(); });

    // Hamburger menu items
    document.getElementById('hmenu-open').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      window.api.openFileDialog();
    });
    document.getElementById('hmenu-exit').addEventListener('click', function () {
      window.api.windowClose();
    });
    document.getElementById('hmenu-sidebar').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      toggleSidebar();
    });
    document.getElementById('hmenu-bars').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      toggleBars();
    });
    document.getElementById('hmenu-present').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      toggleFullscreen();
    });
    document.getElementById('hmenu-doc-dark').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      toggleDocDark();
    });
    document.getElementById('hmenu-devtools').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      window.api.toggleDevTools();
    });
    document.getElementById('hmenu-about').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      window.api.showAbout();
    });
    document.getElementById('hmenu-cache').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      clearCache();
    });

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
    document.getElementById('toolbar-doc-dark').addEventListener('click', toggleDocDark);
    document.getElementById('toolbar-auto-fit').addEventListener('click', toggleAutoFit);
    document.getElementById('hmenu-auto-fit').addEventListener('click', function () {
      hamburgerMenu.classList.add('hidden');
      toggleAutoFit();
    });
    toolbarRotateLeft.addEventListener('click', rotateLeft);
    toolbarRotateRight.addEventListener('click', rotateRight);

    // Page input — jump to page on Enter
    toolbarPageInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        var n = parseInt(toolbarPageInput.value);
        if (!isNaN(n)) goToSlide(n);
        toolbarPageInput.blur();
      }
    });
    toolbarPageInput.addEventListener('blur', function () {
      var tab = getActiveTab();
      if (tab) toolbarPageInput.value = tab.currentSlide;
    });

    // Find bar
    findInput.addEventListener('input', function () {
      clearTimeout(findDebounceTimer);
      findDebounceTimer = setTimeout(runFind, 100);
    });
    findInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.shiftKey) findPrev(); else findNext();
      }
      if (e.key === 'Escape') {
        findInput.value = '';
        clearFindHighlights();
        findMatches = [];
        findCurrentIdx = -1;
        findTerm = '';
        findCount.textContent = '';
        findInput.classList.remove('find-has-results', 'find-no-results');
        findInput.blur();
      }
    });
    findPrevBtn.addEventListener('click', findPrev);
    findNextBtn.addEventListener('click', findNext);

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
      if (e.ctrlKey && (e.key === 'o' || e.key === 'O')) {
        e.preventDefault();
        window.api.openFileDialog();
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
      if (e.ctrlKey && (e.key === 'f' || e.key === 'F')) {
        e.preventDefault();
        openFind();
        return;
      }
      if (e.ctrlKey && (e.key === '=' || e.key === '+')) {
        e.preventDefault();
        zoomIn();
        return;
      }
      if (e.ctrlKey && e.key === '-') {
        e.preventDefault();
        zoomOut();
        return;
      }
      if (e.ctrlKey && e.key === '0') {
        e.preventDefault();
        zoomReset();
        return;
      }
      // Let the browser handle Ctrl+<key> combos we don't explicitly handle
      if (e.ctrlKey || e.metaKey) return;

      // Don't handle shortcuts when find input or page input is focused
      var focused = document.activeElement;
      if (focused === findInput || focused === toolbarPageInput) return;

      // Document dark mode toggle — "i" key
      if (e.key === 'i' || e.key === 'I') {
        toggleDocDark();
        return;
      }

      // Auto fit page toggle — "f" key
      if (e.key === 'f' || e.key === 'F') {
        toggleAutoFit();
        return;
      }

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

    // Pan with mouse drag when zoomed wider than viewport
    var mainViewEl = document.getElementById('main-view');

    mainViewEl.addEventListener('mousedown', function (e) {
      navLock = false; // user is interacting directly — let scroll detection work
      if (!mainViewEl.classList.contains('pannable')) return;
      if (e.button !== 0) return; // left click only
      // Don't pan when the user is clicking on text — let the browser handle selection
      var closestTextLayer = e.target.closest('.page-text-layer');
      if (closestTextLayer && e.target !== closestTextLayer) return;
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

    // Scroll handler — update current page indicator and lazily render pages
    mainViewEl.addEventListener('scroll', function () {
      clearTimeout(scrollTimer);
      scrollTimer = setTimeout(function () {
        updateCurrentPageFromScroll();
        renderVisiblePages();
      }, 8);
      debouncedSavePosition();
    });

    // Wheel = genuine user scroll → release navLock so scroll detection resumes
    mainViewEl.addEventListener('wheel', function (e) {
      navLock = false;
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
        if (tab && tab.totalSlides > 0) refreshView();
      }, 50);
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
    window.api.onToggleTheme(function () { toggleDocDark(); });
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

  // Save all tab positions before window closes — must be synchronous
  // so the data persists before the process exits
  window.addEventListener('beforeunload', function () {
    for (var i = 0; i < tabs.length; i++) {
      var tab = tabs[i];
      if (!tab || !tab.filePath || tab.loading || tab.totalSlides === 0) continue;
      window.api.setFilePositionSync(tab.filePath, {
        currentSlide: tab.currentSlide,
        savedAt: Date.now(),
      });
    }
  });

  // Start
  init();
})();
