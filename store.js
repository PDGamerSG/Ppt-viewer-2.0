const Store = require('electron-store');

const store = new Store({
  defaults: {
    recentFiles: [],
    libreOfficePath: '',
    windowBounds: { width: 1280, height: 800 },
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

module.exports = {
  getRecentFiles,
  addRecentFile,
  getLibreOfficePath,
  setLibreOfficePath,
  getWindowBounds,
  setWindowBounds,
};
