const { app, BrowserWindow, ipcMain, dialog, Menu, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { detectLibreOffice, convertToPdf, cleanupCache } = require('./lib/converter');
const {
  getRecentFiles,
  addRecentFile,
  getLibreOfficePath,
  setLibreOfficePath,
  getWindowBounds,
  setWindowBounds,
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

  buildMenu();
}

function buildMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        {
          label: 'Open...',
          accelerator: 'CmdOrCtrl+O',
          click: () => openFileDialog(),
        },
        { type: 'separator' },
        {
          label: 'Exit',
          accelerator: 'Alt+F4',
          click: () => app.quit(),
        },
      ],
    },
    {
      label: 'View',
      submenu: [
        {
          label: 'Present (Fullscreen)',
          accelerator: 'F5',
          click: () => {
            if (mainWindow) mainWindow.webContents.send('toggle-fullscreen');
          },
        },
        { type: 'separator' },
        {
          label: 'Zoom In',
          accelerator: 'CmdOrCtrl+=',
          click: () => {
            if (mainWindow) mainWindow.webContents.send('zoom-in');
          },
        },
        {
          label: 'Zoom Out',
          accelerator: 'CmdOrCtrl+-',
          click: () => {
            if (mainWindow) mainWindow.webContents.send('zoom-out');
          },
        },
        {
          label: 'Reset Zoom',
          accelerator: 'CmdOrCtrl+0',
          click: () => {
            if (mainWindow) mainWindow.webContents.send('zoom-reset');
          },
        },
        { type: 'separator' },
        { role: 'toggleDevTools' },
      ],
    },
    {
      label: 'Help',
      submenu: [
        {
          label: 'About PPT Viewer',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: 'About PPT Viewer',
              message: 'PPT Viewer v1.0.0',
              detail: 'A fast, native PowerPoint viewer for Windows.\nPowered by LibreOffice and PDF.js.',
            });
          },
        },
      ],
    },
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

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
  return fs.readFileSync(filePath);
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
app.whenReady().then(createWindow);

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
