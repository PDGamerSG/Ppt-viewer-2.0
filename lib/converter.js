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
  // Check environment variable first
  const envPath = process.env.LIBREOFFICE_PATH;
  if (envPath && fs.existsSync(envPath)) {
    return envPath;
  }

  // Check custom path from settings
  if (customPath && fs.existsSync(customPath)) {
    return customPath;
  }

  // Scan known paths
  for (const p of SEARCH_PATHS) {
    if (fs.existsSync(p)) {
      return p;
    }
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

function convertToPdf(sofficePath, filePath) {
  return new Promise((resolve, reject) => {
    ensureCacheDir();

    // Check cache first
    const cached = getCachedPdf(filePath);
    if (cached) {
      return resolve(cached);
    }

    const cacheKey = getCacheKey(filePath);
    const expectedOutput = path.join(CACHE_DIR, `${cacheKey}.pdf`);

    // Copy input file to cache with the cache key name so output PDF gets the right name
    const ext = path.extname(filePath);
    const tempInput = path.join(CACHE_DIR, `${cacheKey}${ext}`);
    fs.copyFileSync(filePath, tempInput);

    const args = [
      '--headless',
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

    proc.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    proc.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    proc.on('error', (err) => {
      // Clean up temp input
      try { fs.unlinkSync(tempInput); } catch (_) {}
      reject(new Error(`Failed to start LibreOffice: ${err.message}`));
    });

    proc.on('close', (code) => {
      // Clean up temp input copy
      try { fs.unlinkSync(tempInput); } catch (_) {}

      if (code !== 0) {
        return reject(new Error(`LibreOffice exited with code ${code}\n${stderr}`));
      }

      // 600ms delay for Windows file lock buffer
      setTimeout(() => {
        if (fs.existsSync(expectedOutput)) {
          return resolve(expectedOutput);
        }

        // Retry after another 600ms
        setTimeout(() => {
          if (fs.existsSync(expectedOutput)) {
            return resolve(expectedOutput);
          }
          reject(new Error(`Conversion completed but PDF not found at ${expectedOutput}\nstdout: ${stdout}\nstderr: ${stderr}`));
        }, 600);
      }, 600);
    });
  });
}

function cleanupCache() {
  try {
    if (fs.existsSync(CACHE_DIR)) {
      const files = fs.readdirSync(CACHE_DIR);
      for (const file of files) {
        try {
          fs.unlinkSync(path.join(CACHE_DIR, file));
        } catch (_) {}
      }
      try {
        fs.rmdirSync(CACHE_DIR);
      } catch (_) {}
    }
  } catch (_) {}
}

module.exports = {
  detectLibreOffice,
  convertToPdf,
  cleanupCache,
  CACHE_DIR,
};
