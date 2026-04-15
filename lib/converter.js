const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

const SEARCH_PATHS = [
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files\\LibreOffice 7\\program\\soffice.exe',
  'C:\\Program Files\\LibreOffice 6\\program\\soffice.exe',
];

const CACHE_DIR = path.join(os.tmpdir(), 'pptviewer-cache');

// Ensure cache dir exists once at module load — skip repeated checks
try { fs.mkdirSync(CACHE_DIR, { recursive: true }); } catch (_) {}

// Cache the detected soffice path so we only stat the filesystem once.
// Custom paths always take priority over the cache so the user can
// change the path at any time via the LO screen.
let cachedSofficePath = null;

function detectLibreOffice(customPath) {
  // Custom path always checked first — user may have changed it
  if (customPath && fs.existsSync(customPath)) {
    cachedSofficePath = customPath;
    return customPath;
  }
  if (cachedSofficePath && fs.existsSync(cachedSofficePath)) return cachedSofficePath;

  const envPath = process.env.LIBREOFFICE_PATH;
  if (envPath && fs.existsSync(envPath)) { cachedSofficePath = envPath; return envPath; }
  for (const p of SEARCH_PATHS) {
    if (fs.existsSync(p)) { cachedSofficePath = p; return p; }
  }
  return null;
}

function getCacheKey(filePath) {
  const stat = fs.statSync(filePath);
  const baseName = path.basename(filePath, path.extname(filePath));
  const mtime = stat.mtimeMs.toString(36);
  return `${baseName}_${mtime}`;
}

function getCachedPdf(filePath) {
  const key = getCacheKey(filePath);
  const cachedPath = path.join(CACHE_DIR, `${key}.pdf`);
  if (fs.existsSync(cachedPath)) {
    return cachedPath;
  }
  return null;
}

// Use fs.watch for instant detection, with polling fallback and timeout
function waitForFile(filePath, timeout) {
  return new Promise((resolve, reject) => {
    if (fs.existsSync(filePath)) return resolve(filePath);

    let settled = false;
    const dir = path.dirname(filePath);
    const base = path.basename(filePath);

    // fs.watch fires immediately when the file appears — much faster than polling
    let watcher;
    try {
      watcher = fs.watch(dir, (eventType, filename) => {
        if (settled) return;
        if (filename === base && fs.existsSync(filePath)) {
          settled = true;
          watcher.close();
          clearInterval(poll);
          clearTimeout(deadline);
          resolve(filePath);
        }
      });
    } catch (_) {
      // fs.watch can fail on some systems — fall through to polling
    }

    // Polling fallback every 50ms (fast enough, covers fs.watch edge cases)
    const poll = setInterval(() => {
      if (settled) return;
      if (fs.existsSync(filePath)) {
        settled = true;
        if (watcher) watcher.close();
        clearInterval(poll);
        clearTimeout(deadline);
        resolve(filePath);
      }
    }, 50);

    const deadline = setTimeout(() => {
      if (settled) return;
      settled = true;
      if (watcher) watcher.close();
      clearInterval(poll);
      reject(new Error(`File not found after ${timeout}ms: ${filePath}`));
    }, timeout);
  });
}

// Conversion queue — LibreOffice only supports one headless instance at a time
let conversionQueue = Promise.resolve();

function convertToPdf(sofficePath, filePath) {
  const task = conversionQueue.then(() => doConvertToPdf(sofficePath, filePath));
  conversionQueue = task.catch(() => {});
  return task;
}

function doConvertToPdf(sofficePath, filePath) {
  return new Promise((resolve, reject) => {
    // Check cache first — instant return
    const cached = getCachedPdf(filePath);
    if (cached) {
      return resolve(cached);
    }

    const cacheKey = getCacheKey(filePath);
    const expectedOutput = path.join(CACHE_DIR, `${cacheKey}.pdf`);

    // Copy input file with cache key name so output PDF gets the right name
    const ext = path.extname(filePath);
    const tempInput = path.join(CACHE_DIR, `${cacheKey}${ext}`);

    // Async copy — don't block the main process
    fs.promises.copyFile(filePath, tempInput).then(() => {
      const args = [
        '--headless',
        '--norestore',
        '--nofirststartwizard',
        '--nolockcheck',
        '--convert-to', 'pdf',
        '--outdir', CACHE_DIR,
        tempInput,
      ];

      const proc = spawn(sofficePath, args, {
        shell: false,
        windowsHide: true,
        // Inherit less from parent — lighter process
        stdio: ['ignore', 'pipe', 'pipe'],
      });

      let stdout = '';
      let stderr = '';

      proc.stdout.on('data', (data) => { stdout += data.toString(); });
      proc.stderr.on('data', (data) => { stderr += data.toString(); });

      proc.on('error', (err) => {
        fs.promises.unlink(tempInput).catch(() => {});
        reject(new Error(`Failed to start LibreOffice: ${err.message}`));
      });

      proc.on('close', (code) => {
        fs.promises.unlink(tempInput).catch(() => {});

        if (code !== 0) {
          return reject(new Error(`LibreOffice exited with code ${code}\n${stderr}`));
        }

        waitForFile(expectedOutput, 5000)
          .then(resolve)
          .catch(() => {
            reject(new Error(`Conversion completed but PDF not found at ${expectedOutput}\nstdout: ${stdout}\nstderr: ${stderr}`));
          });
      });
    }).catch((err) => {
      reject(new Error(`Failed to copy input file: ${err.message}`));
    });
  });
}

// Pre-warm LibreOffice — spawns a quick headless invocation so the OS
// caches the soffice binary, DLLs, and JVM.  Runs in the background
// at app startup.  The prewarm promise is chained into the conversion
// queue so a real conversion never races with the prewarm process.
function prewarmLibreOffice(sofficePath) {
  if (!sofficePath) return;
  try {
    const proc = spawn(sofficePath, [
      '--headless',
      '--norestore',
      '--nofirststartwizard',
      '--nolockcheck',
      '--version',
    ], {
      shell: false,
      windowsHide: true,
      stdio: 'ignore',
    });
    proc.unref();

    // Block the conversion queue until prewarm finishes (or fails)
    const prewarmDone = new Promise((resolve) => {
      proc.on('close', resolve);
      proc.on('error', resolve);
    });
    conversionQueue = conversionQueue.then(() => prewarmDone);
  } catch (_) {}
}

function cleanupCache() {
  try {
    if (fs.existsSync(CACHE_DIR)) {
      const files = fs.readdirSync(CACHE_DIR);
      for (const file of files) {
        try { fs.unlinkSync(path.join(CACHE_DIR, file)); } catch (_) {}
      }
      try { fs.rmdirSync(CACHE_DIR); } catch (_) {}
    }
  } catch (_) {}
}

module.exports = {
  detectLibreOffice,
  convertToPdf,
  prewarmLibreOffice,
  cleanupCache,
  CACHE_DIR,
};
