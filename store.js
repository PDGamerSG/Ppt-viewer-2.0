const Store = require('electron-store');

const store = new Store({
  defaults: {
    recentFiles: [],
    libreOfficePath: '',
    windowBounds: { width: 1280, height: 800 },
    filePositions: {},
  },
});

function getRecentFiles() {
  return store.get('recentFiles', []);
}

function addRecentFile(filePath, fileName) {
  let recent = getRecentFiles();
  // Remove duplicate
  recent = recent.filter((f) => f.path !== filePath);
  // Add to front
  recent.unshift({
    path: filePath,
    name: fileName,
    openedAt: new Date().toISOString(),
  });
  // Keep max 10
  recent = recent.slice(0, 10);
  store.set('recentFiles', recent);
}

function getLibreOfficePath() {
  return store.get('libreOfficePath', '');
}

function setLibreOfficePath(p) {
  store.set('libreOfficePath', p);
}

function getWindowBounds() {
  return store.get('windowBounds', { width: 1280, height: 800 });
}

function setWindowBounds(bounds) {
  store.set('windowBounds', bounds);
}

function getFilePosition(filePath) {
  const positions = store.get('filePositions', {});
  return positions[filePath] || null;
}

function setFilePosition(filePath, position) {
  const positions = store.get('filePositions', {});
  positions[filePath] = position;
  // Keep max 50 entries to avoid unbounded growth
  const keys = Object.keys(positions);
  if (keys.length > 50) {
    // Remove oldest entries by openedAt
    keys.sort((a, b) => (positions[a].savedAt || 0) - (positions[b].savedAt || 0));
    for (let i = 0; i < keys.length - 50; i++) {
      delete positions[keys[i]];
    }
  }
  store.set('filePositions', positions);
}

function getSession() {
  return store.get('session', null);
}

function setSession(session) {
  store.set('session', session);
}

module.exports = {
  getRecentFiles,
  addRecentFile,
  getLibreOfficePath,
  setLibreOfficePath,
  getWindowBounds,
  setWindowBounds,
  getFilePosition,
  setFilePosition,
  getSession,
  setSession,
};
