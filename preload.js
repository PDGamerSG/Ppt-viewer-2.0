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
  getFilePosition: (filePath) => ipcRenderer.invoke('get-file-position', filePath),
  setFilePosition: (filePath, position) => ipcRenderer.invoke('set-file-position', filePath, position),
  setTitle: (title) => ipcRenderer.invoke('set-title', title),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  enterFullscreen: () => ipcRenderer.invoke('enter-fullscreen'),
  exitFullscreen: () => ipcRenderer.invoke('exit-fullscreen'),
  isFullscreen: () => ipcRenderer.invoke('is-fullscreen'),
  clearCache: () => ipcRenderer.invoke('clear-cache'),

  // Events from main process
  onOpenFile: (callback) => ipcRenderer.on('open-file', (_e, filePath) => callback(filePath)),
  onConvertProgress: (callback) => ipcRenderer.on('convert-progress', (_e, stage) => callback(stage)),
  onToggleFullscreen: (callback) => ipcRenderer.on('toggle-fullscreen', () => callback()),
  onZoomIn: (callback) => ipcRenderer.on('zoom-in', () => callback()),
  onZoomOut: (callback) => ipcRenderer.on('zoom-out', () => callback()),
  onZoomReset: (callback) => ipcRenderer.on('zoom-reset', () => callback()),
  onToggleTheme: (callback) => ipcRenderer.on('toggle-theme', () => callback()),

  // Window controls
  windowMinimize: () => ipcRenderer.invoke('window-minimize'),
  windowMaximize: () => ipcRenderer.invoke('window-maximize'),
  windowClose: () => ipcRenderer.invoke('window-close'),
  windowIsMaximized: () => ipcRenderer.invoke('window-is-maximized'),
  showAbout: () => ipcRenderer.invoke('show-about'),
  toggleDevTools: () => ipcRenderer.invoke('toggle-devtools'),
});
