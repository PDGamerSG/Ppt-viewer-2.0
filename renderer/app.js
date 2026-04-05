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

    // Save outgoing tab's DOM snapshot
    var outgoing = getActiveTab();
    if (outgoing && !outgoing.loading) {
      var slideContainer = document.getElementById('slide-container');
      var mainView = document.getElementById('main-view');
      // Move the live DOM nodes into a DocumentFragment (cheap, no cloning)
      var frag = document.createDocumentFragment();
      while (slideContainer.firstChild) {
        frag.appendChild(slideContainer.firstChild);
      }
      outgoing.domSnapshot = frag;
      outgoing.snapshotScrollTop = mainView.scrollTop;
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

    if (tab.thumbnailDataUrls && tab.thumbnailDataUrls.length === tab.totalSlides) {
      restoreThumbnailsFromCache(tab);
    } else {
      await renderAllThumbnails(tab);
    }

    updateSlideCounter();
    updateZoomDisplay();
    updateRotateButtons();
    window.api.setTitle(tab.fileName + ' \u2014 PPT Viewer');

    // If we have a DOM snapshot, restore it instantly (no re-render)
    var slideContainer = document.getElementById('slide-container');
    var mainView = document.getElementById('main-view');
    if (tab.domSnapshot) {
      slideContainer.innerHTML = '';
      slideContainer.appendChild(tab.domSnapshot);
      tab.domSnapshot = null;

      // Restore page base dims from the first wrapper
      pageBaseDims = await PdfRenderer.getPageDimensions(1, tab.pageRotation);
      renderedPages.clear();
      // Mark all pages that have canvas content as rendered
      var wrappers = slideContainer.querySelectorAll('.page-wrapper');
      for (var w = 0; w < wrappers.length; w++) {
        var c = wrappers[w].querySelector('.page-canvas');
        if (c && c.width > 0) renderedPages.add(parseInt(wrappers[w].dataset.page));
      }
      renderGeneration++;
      updatePannable();
      mainView.scrollTop = tab.snapshotScrollTop;
      applyDocDark();
      return;
    }

    await setupContinuousView(tab);

    // Scroll to the page the user was on
    var wrappers2 = document.querySelectorAll('#slide-container .page-wrapper');
    if (wrappers2[tab.currentSlide - 1]) {
      wrappers2[tab.currentSlide - 1].scrollIntoView({ block: 'start' });
    }
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
    applyDocDark();
  }

  function closeTab(tabId) {
    var idx = getTabIndex(tabId);
    if (idx === -1) return;

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
      var numPages = await PdfRenderer.loadDocument(data, tab.id);
      tab.totalSlides = numPages;

      await window.api.saveRecentFile(filePath, fileName);

      // If this tab is active (or was the first), render it
      if (activeTabId === tab.id || isFirstTab) {
        activeTabId = tab.id;
        renderTabs();
        updateSlideCounter();
        updateZoomDisplay();
        window.api.setTitle(fileName + ' \u2014 PPT Viewer');
        await setupContinuousView(tab);
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

  // ---- Continuous scroll rendering ----

  async function setupContinuousView(tab) {
    var slideContainer = document.getElementById('slide-container');
    renderedPages.clear();
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
    var maxWidth = mainView.clientWidth - 40;
    var fitScale = maxWidth / pageBaseDims.width;
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

    var maxWidth = mainView.clientWidth - 40;
    var fitScale = maxWidth / pageBaseDims.width;
    var effectiveZoom = fitScale * tab.zoomLevel;

    for (var i = 0; i < wrappers.length; i++) {
      if (gen !== renderGeneration) return; // zoom/tab changed, abort stale render

      var wrapper = wrappers[i];
      var wTop = wrapper.offsetTop;
      var wBottom = wTop + wrapper.offsetHeight;
      var pageNum = parseInt(wrapper.dataset.page);

      // Check if page is in visible range (with buffer)
      if (wBottom >= viewTop - buffer && wTop <= viewBottom + buffer) {
        if (!renderedPages.has(pageNum)) {
          renderedPages.add(pageNum);

          try {
            var canvas = wrapper.querySelector('.page-canvas');
            var textLayerEl = wrapper.querySelector('.page-text-layer');

            // Render to offscreen canvas, then blit — prevents flash
            var offscreen = document.createElement('canvas');
            await PdfRenderer.renderSlide(offscreen, pageNum, effectiveZoom, tab.pageRotation);

            if (gen !== renderGeneration) return; // stale, discard

            canvas.width = offscreen.width;
            canvas.height = offscreen.height;
            canvas.style.width = offscreen.style.width;
            canvas.style.height = offscreen.style.height;
            var ctx = canvas.getContext('2d');
            ctx.drawImage(offscreen, 0, 0);

            await PdfRenderer.renderTextLayer(textLayerEl, pageNum, effectiveZoom, tab.pageRotation);
            if (docDarkMode) applyDocDark();
            applyFindHighlights();
          } catch (err) {
            renderedPages.delete(pageNum); // allow retry on next scroll
          }
        }
      }
    }
  }

  function updateCurrentPageFromScroll() {
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

    // Re-fetch dimensions in case rotation changed
    pageBaseDims = await PdfRenderer.getPageDimensions(1, tab.pageRotation);
    if (!pageBaseDims) return;

    renderedPages.clear();
    renderGeneration++;

    var mainView = document.getElementById('main-view');
    var maxWidth = mainView.clientWidth - 40;
    var fitScale = maxWidth / pageBaseDims.width;
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
    tab.thumbnailDataUrls = [];

    for (var i = 1; i <= tab.totalSlides; i++) {
      var item = document.createElement('div');
      item.className = 'thumbnail-item' + (i === tab.currentSlide ? ' active' : '');
      item.dataset.page = i;

      var canvas = await PdfRenderer.renderThumbnail(i, 160, tab.pageRotation);
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

    var wrappers = document.querySelectorAll('#slide-container .page-wrapper');
    if (wrappers[num - 1]) {
      wrappers[num - 1].scrollIntoView({ block: 'center' });
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
  var findTerm = '';
  var findMatches = [];
  var findCurrentIdx = -1;
  var findDebounceTimer = null;

  function openFind() {
    findInput.focus();
    findInput.select();
  }

  function runFind() {
    clearFindHighlights();
    findMatches = [];
    findCurrentIdx = -1;

    findTerm = findInput.value;
    if (!findTerm) {
      updateFindCount();
      findInput.classList.remove('find-has-results', 'find-no-results');
      return;
    }

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
  function setupEventListeners() {
    // Window controls
    document.getElementById('win-minimize').addEventListener('click', function () { window.api.windowMinimize(); });
    document.getElementById('win-maximize').addEventListener('click', function () { window.api.windowMaximize(); });
    document.getElementById('win-close').addEventListener('click', function () { window.api.windowClose(); });

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
      }, 16);
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

  // Start
  init();
})();
