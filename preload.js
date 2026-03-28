const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
  convertFile: (filePath) => ipcRenderer.invoke('convert-file', filePath),
  detectLibreOffice: () => ipcRenderer.invoke('detect-libreoffice'),
  getRecentFiles: () => ipcRenderer.invoke('get-recent-files'),
  saveRecentFile: (filePath, fileName) => ipcRenderer.invoke('save-recent-file', filePath, fileName),
  getLibreOfficePath: () => ipcRenderer.invoke('get-libreoffice-path'),
  setLibreOfficePath: (p) => ipcRenderer.invoke('set-libreoffice-path', p),
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  setTitle: (title) => ipcRenderer.invoke('set-title', title),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  enterFullscreen: () => ipcRenderer.invoke('enter-fullscreen'),
  exitFullscreen: () => ipcRenderer.invoke('exit-fullscreen'),
  isFullscreen: () => ipcRenderer.invoke('is-fullscreen'),

  // Events from main process
  onOpenFile: (callback) => ipcRenderer.on('open-file', (_e, filePath) => callback(filePath)),
  onToggleFullscreen: (callback) => ipcRenderer.on('toggle-fullscreen', () => callback()),
  onZoomIn: (callback) => ipcRenderer.on('zoom-in', () => callback()),
  onZoomOut: (callback) => ipcRenderer.on('zoom-out', () => callback()),
  onZoomReset: (callback) => ipcRenderer.on('zoom-reset', () => callback()),
});
