// PDF.js rendering utilities — loaded as UMD script in renderer
// pdfjsLib is available globally from pdf.min.js
//
// Supports multiple concurrently-loaded documents so tab switching
// doesn't require re-parsing the PDF from raw bytes.

(function () {
  'use strict';

  const workerSrc = '../node_modules/pdfjs-dist/build/pdf.worker.min.js';
  pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;

  // Map<docId, { doc, pageCache }>
  const docs = new Map();
  let activeId = null;

  function active() {
    return activeId !== null ? docs.get(activeId) : null;
  }

  // Load a document under the given id.  If already loaded, just activate it.
  async function loadDocument(data, docId) {
    if (docId !== undefined) {
      if (docs.has(docId)) {
        activeId = docId;
        return docs.get(docId).doc.numPages;
      }
    }

    const uint8 = new Uint8Array(data);
    const doc = await pdfjsLib.getDocument({ data: uint8 }).promise;
    const id = docId !== undefined ? docId : Symbol();
    docs.set(id, { doc, pageCache: new Map() });
    activeId = id;
    return doc.numPages;
  }

  function setActive(docId) {
    activeId = docId;
  }

  async function getPage(pageNum) {
    const entry = active();
    if (!entry) return null;
    if (entry.pageCache.has(pageNum)) return entry.pageCache.get(pageNum);
    const page = await entry.doc.getPage(pageNum);
    entry.pageCache.set(pageNum, page);
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
    const entry = active();
    return entry ? entry.doc.numPages : 0;
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

  // Destroy a specific document by id
  function cleanupDoc(docId) {
    const entry = docs.get(docId);
    if (entry) {
      entry.doc.destroy();
      entry.pageCache.clear();
      docs.delete(docId);
    }
    if (activeId === docId) activeId = null;
  }

  // Destroy all documents
  function cleanupAll() {
    docs.forEach(function (entry) {
      entry.doc.destroy();
      entry.pageCache.clear();
    });
    docs.clear();
    activeId = null;
  }

  function isLoaded(docId) {
    return docs.has(docId);
  }

  // Expose globally
  window.PdfRenderer = {
    loadDocument,
    setActive,
    isLoaded,
    renderSlide,
    renderTextLayer,
    renderThumbnail,
    getPageCount,
    getPageDimensions,
    cleanupDoc,
    cleanup: cleanupAll,
  };
})();
