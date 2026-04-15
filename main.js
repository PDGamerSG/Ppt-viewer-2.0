const { app, BrowserWindow, ipcMain, dialog, Menu, shell, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');
const { detectLibreOffice, convertToPdf, prewarmLibreOffice, cleanupCache } = require('./lib/converter');
const {
  getRecentFiles,
  addRecentFile,
  getLibreOfficePath,
  setLibreOfficePath,
  getWindowBounds,
  setWindowBounds,
  getFilePosition,
  setFilePosition,
} = require('./store');

let mainWindow = null;

// Extract file argument from argv
function findFileArg(argv) {
  return argv.find((arg) => {
    const ext = path.extname(arg).toLowerCase();
    return ext === '.pptx' || ext === '.ppt' || ext === '.pdf';
  });
}

// Send file to renderer when ready
function openFileInWindow(filePath) {
  if (!mainWindow) return;
  if (mainWindow.webContents.isLoading()) {
    mainWindow.webContents.once('did-finish-load', () => {
      mainWindow.webContents.send('open-file', filePath);
    });
  } else {
    mainWindow.webContents.send('open-file', filePath);
  }
}

// Single instance lock — second launch sends file to existing window
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv) => {
    const fileArg = findFileArg(argv);
    if (fileArg && fs.existsSync(fileArg)) {
      openFileInWindow(fileArg);
    }
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}

function createWindow() {
  const bounds = getWindowBounds();

  mainWindow = new BrowserWindow({
    width: bounds.width,
    height: bounds.height,
    minWidth: 800,
    minHeight: 600,
    title: 'PPT Viewer',
    icon: path.join(__dirname, 'assets', 'icon.ico'),
    backgroundColor: '#1e1e1e',
    show: false,
    frame: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    // Handle CLI argument from first launch
    const fileArg = findFileArg(process.argv);
    if (fileArg && fs.existsSync(fileArg)) {
      mainWindow.webContents.send('open-file', fileArg);
    }
  });

  mainWindow.on('resize', () => {
    const [width, height] = mainWindow.getSize();
    setWindowBounds({ width, height });
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  Menu.setApplicationMenu(null);
}

// Window control IPC handlers
ipcMain.handle('window-minimize', () => { if (mainWindow) mainWindow.minimize(); });
ipcMain.handle('window-maximize', () => {
  if (!mainWindow) return;
  if (mainWindow.isMaximized()) mainWindow.unmaximize();
  else mainWindow.maximize();
});
ipcMain.handle('window-close', () => { if (mainWindow) mainWindow.close(); });
ipcMain.handle('window-is-maximized', () => mainWindow ? mainWindow.isMaximized() : false);

ipcMain.handle('toggle-devtools', () => {
  if (mainWindow) mainWindow.webContents.toggleDevTools();
});

ipcMain.handle('show-about', () => {
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'About PPT Viewer',
    message: 'PPT Viewer v1.0.0',
    detail: 'A fast, native PowerPoint viewer for Windows.\nPowered by LibreOffice and PDF.js.',
  });
});

async function openFileDialog() {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Open Presentation',
    filters: [
      { name: 'Presentations & PDFs', extensions: ['pptx', 'ppt', 'pdf'] },
      { name: 'PowerPoint Files', extensions: ['pptx', 'ppt'] },
      { name: 'PDF Files', extensions: ['pdf'] },
      { name: 'All Files', extensions: ['*'] },
    ],
    properties: ['openFile', 'multiSelections'],
  });

  if (!result.canceled && result.filePaths.length > 0) {
    for (const filePath of result.filePaths) {
      mainWindow.webContents.send('open-file', filePath);
    }
  }
}

// IPC Handlers
ipcMain.handle('open-file-dialog', async () => {
  await openFileDialog();
});

ipcMain.handle('detect-libreoffice', () => {
  const customPath = getLibreOfficePath();
  return detectLibreOffice(customPath);
});

ipcMain.handle('convert-file', async (_event, filePath) => {
  // PDF files need no conversion
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.pdf') {
    return filePath;
  }

  const customPath = getLibreOfficePath();
  const sofficePath = detectLibreOffice(customPath);
  if (!sofficePath) {
    throw new Error('LibreOffice not found');
  }
  const pdfPath = await convertToPdf(sofficePath, filePath);
  return pdfPath;
});

ipcMain.handle('get-recent-files', () => {
  return getRecentFiles();
});

ipcMain.handle('save-recent-file', (_event, filePath, fileName) => {
  addRecentFile(filePath, fileName);
});

ipcMain.handle('get-libreoffice-path', () => {
  return getLibreOfficePath();
});

ipcMain.handle('set-libreoffice-path', (_event, p) => {
  setLibreOfficePath(p);
});

ipcMain.handle('read-file', async (_event, filePath) => {
  return fs.promises.readFile(filePath);
});

ipcMain.handle('get-file-position', (_event, filePath) => {
  return getFilePosition(filePath);
});

ipcMain.handle('set-file-position', (_event, filePath, position) => {
  setFilePosition(filePath, position);
});

ipcMain.handle('set-title', (_event, title) => {
  if (mainWindow) mainWindow.setTitle(title);
});

ipcMain.handle('open-external', (_event, url) => {
  shell.openExternal(url);
});

ipcMain.handle('enter-fullscreen', () => {
  if (mainWindow) mainWindow.setFullScreen(true);
});

ipcMain.handle('exit-fullscreen', () => {
  if (mainWindow) mainWindow.setFullScreen(false);
});

ipcMain.handle('is-fullscreen', () => {
  return mainWindow ? mainWindow.isFullScreen() : false;
});

ipcMain.handle('clear-cache', () => {
  cleanupCache();
  return true;
});

// App lifecycle
app.whenReady().then(() => {
  createWindow();
  // Pre-warm LibreOffice in the background so the first conversion is faster
  const customPath = getLibreOfficePath();
  const sofficePath = detectLibreOffice(customPath);
  prewarmLibreOffice(sofficePath);
});

app.on('window-all-closed', () => {
  app.quit();
});

app.on('before-quit', () => {
  cleanupCache();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
