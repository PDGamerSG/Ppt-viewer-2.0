// PDF.js rendering utilities — loaded as UMD script in renderer
// pdfjsLib is available globally from pdf.min.js

(function () {
  'use strict';

  const workerSrc = '../node_modules/pdfjs-dist/build/pdf.worker.min.js';
  pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;

  let currentDoc = null;
  const pageCache = new Map();

  async function loadDocument(data) {
    pageCache.clear();
    const uint8 = new Uint8Array(data);
    currentDoc = await pdfjsLib.getDocument({ data: uint8 }).promise;
    return currentDoc.numPages;
  }

  async function getPage(pageNum) {
    if (!currentDoc) return null;
    if (pageCache.has(pageNum)) return pageCache.get(pageNum);
    const page = await currentDoc.getPage(pageNum);
    pageCache.set(pageNum, page);
    return page;
  }

  async function renderSlide(canvas, pageNum, zoom, rotation) {
    const page = await getPage(pageNum);
    if (!page) return null;

    const rot = rotation || 0;
    const baseViewport = page.getViewport({ scale: 1, rotation: rot });
    const dpr = window.devicePixelRatio || 1;
    const scale = zoom * dpr;
    const viewport = page.getViewport({ scale, rotation: rot });

    canvas.width = viewport.width;
    canvas.height = viewport.height;
    canvas.style.width = (viewport.width / dpr) + 'px';
    canvas.style.height = (viewport.height / dpr) + 'px';

    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    await page.render({
      canvasContext: ctx,
      viewport: viewport,
    }).promise;

    return {
      width: baseViewport.width,
      height: baseViewport.height,
    };
  }

  async function renderThumbnail(pageNum, width, rotation) {
    const page = await getPage(pageNum);
    if (!page) return null;

    const rot = rotation || 0;
    const baseViewport = page.getViewport({ scale: 1, rotation: rot });
    const scale = width / baseViewport.width;
    const dpr = window.devicePixelRatio || 1;
    const viewport = page.getViewport({ scale: scale * dpr, rotation: rot });

    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    canvas.style.width = (viewport.width / dpr) + 'px';
    canvas.style.height = (viewport.height / dpr) + 'px';

    const ctx = canvas.getContext('2d');
    await page.render({
      canvasContext: ctx,
      viewport: viewport,
    }).promise;

    return canvas;
  }

  function getPageCount() {
    return currentDoc ? currentDoc.numPages : 0;
  }

  async function getPageDimensions(pageNum, rotation) {
    const page = await getPage(pageNum);
    if (!page) return null;
    const rot = rotation || 0;
    const vp = page.getViewport({ scale: 1, rotation: rot });
    return { width: vp.width, height: vp.height };
  }

  async function renderTextLayer(container, pageNum, zoom, rotation) {
    const page = await getPage(pageNum);
    if (!page) return;

    const rot = rotation || 0;
    const viewport = page.getViewport({ scale: zoom, rotation: rot });

    container.innerHTML = '';
    container.style.width = viewport.width + 'px';
    container.style.height = viewport.height + 'px';
    container.style.setProperty('--scale-factor', zoom);

    const textContent = await page.getTextContent();
    if (!textContent.items.length) return;

    pdfjsLib.renderTextLayer({
      textContentSource: textContent,
      container: container,
      viewport: viewport,
      textDivs: [],
    });
  }

  function cleanupPdf() {
    if (currentDoc) {
      currentDoc.destroy();
      currentDoc = null;
    }
    pageCache.clear();
  }

  // Expose globally
  window.PdfRenderer = {
    loadDocument,
    renderSlide,
    renderTextLayer,
    renderThumbnail,
    getPageCount,
    getPageDimensions,
    cleanup: cleanupPdf,
  };
})();
