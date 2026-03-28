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

function ensureCacheDir() {
  if (!fs.existsSync(CACHE_DIR)) {
    fs.mkdirSync(CACHE_DIR, { recursive: true });
  }
}

function detectLibreOffice(customPath) {
  const envPath = process.env.LIBREOFFICE_PATH;
  if (envPath && fs.existsSync(envPath)) return envPath;
  if (customPath && fs.existsSync(customPath)) return customPath;
  for (const p of SEARCH_PATHS) {
    if (fs.existsSync(p)) return p;
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
  ensureCacheDir();
  const key = getCacheKey(filePath);
  const cachedPath = path.join(CACHE_DIR, `${key}.pdf`);
  if (fs.existsSync(cachedPath)) {
    return cachedPath;
  }
  return null;
}

// Poll for file existence instead of flat delays — checks every 100ms up to 3s
function waitForFile(filePath, timeout) {
  return new Promise((resolve, reject) => {
    // Check immediately
    if (fs.existsSync(filePath)) return resolve(filePath);

    const interval = 100;
    let elapsed = 0;
    const timer = setInterval(() => {
      elapsed += interval;
      if (fs.existsSync(filePath)) {
        clearInterval(timer);
        return resolve(filePath);
      }
      if (elapsed >= timeout) {
        clearInterval(timer);
        reject(new Error(`File not found after ${timeout}ms: ${filePath}`));
      }
    }, interval);
  });
}

// Conversion queue — LibreOffice only supports one headless instance at a time
let conversionQueue = Promise.resolve();

function convertToPdf(sofficePath, filePath) {
  const task = conversionQueue.then(() => doConvertToPdf(sofficePath, filePath));
  // Keep the queue going even if one conversion fails
  conversionQueue = task.catch(() => {});
  return task;
}

function doConvertToPdf(sofficePath, filePath) {
  return new Promise((resolve, reject) => {
    ensureCacheDir();

    // Check cache first
    const cached = getCachedPdf(filePath);
    if (cached) {
      return resolve(cached);
    }

    const cacheKey = getCacheKey(filePath);
    const expectedOutput = path.join(CACHE_DIR, `${cacheKey}.pdf`);

    // Copy input file with cache key name so output PDF gets the right name
    const ext = path.extname(filePath);
    const tempInput = path.join(CACHE_DIR, `${cacheKey}${ext}`);
    fs.copyFileSync(filePath, tempInput);

    const args = [
      '--headless',
      '--norestore',
      '--nofirststartwizard',
      '--convert-to', 'pdf',
      '--outdir', CACHE_DIR,
      tempInput,
    ];

    const proc = spawn(sofficePath, args, {
      shell: false,
      windowsHide: true,
    });

    let stdout = '';
    let stderr = '';

    proc.stdout.on('data', (data) => { stdout += data.toString(); });
    proc.stderr.on('data', (data) => { stderr += data.toString(); });

    proc.on('error', (err) => {
      try { fs.unlinkSync(tempInput); } catch (_) {}
      reject(new Error(`Failed to start LibreOffice: ${err.message}`));
    });

    proc.on('close', (code) => {
      try { fs.unlinkSync(tempInput); } catch (_) {}

      if (code !== 0) {
        return reject(new Error(`LibreOffice exited with code ${code}\n${stderr}`));
      }

      // Poll for the output file (handles Windows file lock delay)
      waitForFile(expectedOutput, 3000)
        .then(resolve)
        .catch(() => {
          reject(new Error(`Conversion completed but PDF not found at ${expectedOutput}\nstdout: ${stdout}\nstderr: ${stderr}`));
        });
    });
  });
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
  cleanupCache,
  CACHE_DIR,
};
