'use strict';
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path   = require('path');
const fs     = require('fs');
const os     = require('os');
const https  = require('https');
const crypto = require('crypto');
const JSZip         = require('jszip');
const Docxtemplater = require('docxtemplater');
const PizZip        = require('pizzip');
const ExcelJS       = require('exceljs');

const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged;

// ── Runtime dir: ~/.bphh_docgen_runtime ──────────────────────────────────────
const RUNTIME_DIR     = path.join(os.homedir(), '.bphh_docgen_runtime');
const CATALOG_RUNTIME = path.join(RUNTIME_DIR, 'catalog.json');
const TEMPLATES_DIR   = path.join(RUNTIME_DIR, 'templates');
const SETTINGS_FILE   = path.join(RUNTIME_DIR, 'settings.json');
const VERSION_FILE    = path.join(RUNTIME_DIR, 'VERSION');
const FINGERPRINT_FILE = path.join(RUNTIME_DIR, 'CATALOG_FINGERPRINT');

const APP_VERSION = '1.1.5';
const DEFAULT_UPDATE_URL = 'https://raw.githubusercontent.com/phongbksneep/toolgendocBPH/main/version.json';

function loadSettings() {
  try { return JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8')); } catch { return {}; }
}

// ── Helper: HTTPS GET returning Buffer ───────────────────────────────────────
function fetchUrl(url) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { timeout: 10000 }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchUrl(res.headers.location).then(resolve).catch(reject);
      }
      if (res.statusCode < 200 || res.statusCode >= 300) {
        return reject(new Error(`HTTP ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('Request timeout')); });
  });
}

function getAssetPath(name) {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, name);
  }
  // Dev mode: electron/main.cjs is at <root>/electron/main.cjs
  // so app root is one level up from __dirname.
  return path.join(__dirname, '..', name);
}

function fileSha256(filePath) {
  if (!fs.existsSync(filePath)) return null;
  const buf = fs.readFileSync(filePath);
  return crypto.createHash('sha256').update(buf).digest('hex');
}

async function ensureRuntime() {
  fs.mkdirSync(RUNTIME_DIR, { recursive: true });

  // Re-extract/copy when version changes, files are missing, or catalog changed
  const runtimeVersion = fs.existsSync(VERSION_FILE)
    ? fs.readFileSync(VERSION_FILE, 'utf-8').trim() : null;
  const runtimeFingerprint = fs.existsSync(FINGERPRINT_FILE)
    ? fs.readFileSync(FINGERPRINT_FILE, 'utf-8').trim() : null;

  const bundlePath = getAssetPath('assets/templates_bundle.zip');
  const fallbackCatalog = getAssetPath('catalog.json');
  const currentFingerprint = fileSha256(bundlePath) || fileSha256(fallbackCatalog);

  const needsExtract = !fs.existsSync(CATALOG_RUNTIME)
    || !fs.existsSync(TEMPLATES_DIR)
    || runtimeVersion !== APP_VERSION
    || (currentFingerprint && runtimeFingerprint !== currentFingerprint);
  if (!needsExtract) return;

  // Remove stale files before fresh extraction
  if (fs.existsSync(TEMPLATES_DIR))   fs.rmSync(TEMPLATES_DIR, { recursive: true, force: true });
  if (fs.existsSync(CATALOG_RUNTIME)) fs.unlinkSync(CATALOG_RUNTIME);

  if (fs.existsSync(bundlePath)) {
    // Preferred path: extract from bundle
    const data = fs.readFileSync(bundlePath);
    const zip = await JSZip.loadAsync(data);
    for (const [relPath, zipEntry] of Object.entries(zip.files)) {
      if (zipEntry.dir) continue;
      const dest = path.join(RUNTIME_DIR, relPath);
      fs.mkdirSync(path.dirname(dest), { recursive: true });
      const buf = await zipEntry.async('nodebuffer');
      fs.writeFileSync(dest, buf);
    }
    fs.writeFileSync(VERSION_FILE, APP_VERSION);
    if (currentFingerprint) fs.writeFileSync(FINGERPRINT_FILE, currentFingerprint);
    return;
  }

  // Backward-compatible fallback for local/dev setups that only have
  // ./templates + ./catalog.json (old workflow, no templates_bundle.zip)
  const fallbackTemplates = getAssetPath('templates');
  if (!fs.existsSync(fallbackCatalog) || !fs.existsSync(fallbackTemplates)) {
    throw new Error('Thiếu templates_bundle.zip: ' + bundlePath);
  }

  fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
  fs.copyFileSync(fallbackCatalog, CATALOG_RUNTIME);
  const files = fs.readdirSync(fallbackTemplates).filter(f => f.toLowerCase().endsWith('.docx'));
  for (const f of files) {
    fs.copyFileSync(path.join(fallbackTemplates, f), path.join(TEMPLATES_DIR, f));
  }
  fs.writeFileSync(VERSION_FILE, APP_VERSION);
  if (currentFingerprint) fs.writeFileSync(FINGERPRINT_FILE, currentFingerprint);
}

function loadCatalog() {
  // Prefer runtime catalog (may be updated)
  if (fs.existsSync(CATALOG_RUNTIME)) {
    return JSON.parse(fs.readFileSync(CATALOG_RUNTIME, 'utf-8'));
  }
  const bundledCatalog = getAssetPath('catalog.json');
  return JSON.parse(fs.readFileSync(bundledCatalog, 'utf-8'));
}

function writeIfExistsOrCreatable(filePath, content) {
  try {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    fs.writeFileSync(filePath, content, 'utf-8');
    return true;
  } catch {
    return false;
  }
}

function getRuntimeDemoPath() {
  return path.join(RUNTIME_DIR, 'project-data.demo.friendly.json');
}

function loadDemoDataObject() {
  const runtimeDemo = getRuntimeDemoPath();
  const bundledDemo = getAssetPath('project-data.demo.friendly.json');
  const p = fs.existsSync(runtimeDemo) ? runtimeDemo : bundledDemo;
  return { path: p, data: JSON.parse(fs.readFileSync(p, 'utf-8')) };
}

function syncDemoLabels(demoObj, key, label) {
  if (!demoObj || typeof demoObj !== 'object') return demoObj;
  const out = { ...demoObj };
  if (Array.isArray(out.global_fields)) {
    out.global_fields = out.global_fields.map(f => f?.key === key ? { ...f, label } : f);
  }
  if (Array.isArray(out.jobs)) {
    out.jobs = out.jobs.map(j => {
      if (!Array.isArray(j?.fields)) return j;
      return {
        ...j,
        fields: j.fields.map(f => f?.key === key ? { ...f, label } : f),
      };
    });
  }
  return out;
}

function healCatalogLabelsFromBundled(catalog) {
  const bundledCatalogPath = getAssetPath('catalog.json');
  if (!fs.existsSync(bundledCatalogPath)) return catalog;

  let bundled;
  try { bundled = JSON.parse(fs.readFileSync(bundledCatalogPath, 'utf-8')); }
  catch { return catalog; }

  let changed = false;
  for (const [key, meta] of Object.entries(catalog || {})) {
    const b = bundled[key];
    if (!b) continue;
    const currentLabel = String(meta?.label || '').trim();
    const currentSample = String(meta?.sample || '').trim();
    const bundledLabel = String(b?.label || '').trim();

    // If runtime label looks like raw sample, prefer curated bundled label.
    if (bundledLabel && (!currentLabel || currentLabel === currentSample)) {
      meta.label = bundledLabel;
      changed = true;
    }
  }

  if (changed) {
    try { fs.writeFileSync(CATALOG_RUNTIME, JSON.stringify(catalog, null, 2), 'utf-8'); } catch { /* ignore */ }
  }
  return catalog;
}

// Remove yellow (or any) w:highlight from all word/*.xml parts in a PizZip
function removeHighlights(zip) {
  const xmlFiles = Object.keys(zip.files).filter(
    f => f.startsWith('word/') && f.endsWith('.xml')
  );
  for (const xmlFile of xmlFiles) {
    const entry = zip.file(xmlFile);
    if (!entry) continue;
    let xml = entry.asText();
    // Self-closing: <w:highlight w:val="yellow"/>
    xml = xml.replace(/<w:highlight[^>]*\/>/g, '');
    zip.file(xmlFile, xml);
  }
}

function detectTemplateDelimitersFromZip(zip) {
  const xmlFiles = Object.keys(zip.files).filter(
    f => f.startsWith('word/') && f.endsWith('.xml')
  );

  let hasDouble = false;
  let hasSingle = false;

  for (const xmlFile of xmlFiles) {
    const entry = zip.file(xmlFile);
    if (!entry) continue;
    const xml = entry.asText();

    // {{KEY}} style
    if (/\{\{\s*[A-Z_]+__[\s\S]*?\}\}/.test(xml)) hasDouble = true;
    // {KEY} style (not {{...}})
    if (/(?<!\{)\{\s*[A-Z_]+__[\s\S]*?\}(?!\})/.test(xml)) hasSingle = true;

    if (hasDouble && hasSingle) break;
  }

  // Prefer explicit double if both appear
  if (hasDouble) return { start: '{{', end: '}}' };
  return { start: '{', end: '}' };
}

// ── IPC handlers ──────────────────────────────────────────────────────────────

ipcMain.handle('get-catalog', async () => {
  await ensureRuntime();
  const cat = loadCatalog();
  return healCatalogLabelsFromBundled(cat);
});

ipcMain.handle('get-templates', async () => {
  await ensureRuntime();
  return fs.readdirSync(TEMPLATES_DIR).filter(f => f.toLowerCase().endsWith('.docx'));
});

// ── Template management ───────────────────────────────────────────────────────

// Helper: escape XML special chars
function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Read template: extract paragraphs as text (preserving {variable} markers)
ipcMain.handle('read-template', async (_event, { templateName }) => {
  await ensureRuntime();
  const tplPath = path.join(TEMPLATES_DIR, templateName);
  if (!fs.existsSync(tplPath)) return { error: 'Template not found: ' + templateName };

  const data = fs.readFileSync(tplPath);
  const zip  = await JSZip.loadAsync(data);
  const docEntry = zip.files['word/document.xml'];
  if (!docEntry) return { error: 'Invalid docx' };

  const docXml = await docEntry.async('text');

  // Merge consecutive text runs within each paragraph, then extract text
  const paragraphs = [];
  const paraRe = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let pm;
  while ((pm = paraRe.exec(docXml)) !== null) {
    const paraXml = pm[0];
    // Extract all <w:t> text
    const bits = [];
    const tRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    let tm;
    while ((tm = tRe.exec(paraXml)) !== null) bits.push(tm[1]);
    const line = bits.join('');
    paragraphs.push(line);
  }
  return { paragraphs };
});

// Save template: create/overwrite a docx from plain text with {variable} syntax
ipcMain.handle('save-template', async (_event, { templateName, content, baseTemplate }) => {
  await ensureRuntime();
  if (!templateName || !templateName.trim()) return { error: 'Tên template không được trống' };
  const name = templateName.endsWith('.docx') ? templateName : templateName + '.docx';

  // Use baseTemplate (or fallback to first available) as structural base for styles
  const baseList = fs.readdirSync(TEMPLATES_DIR).filter(f => f.toLowerCase().endsWith('.docx'));
  const base = baseTemplate && fs.existsSync(path.join(TEMPLATES_DIR, baseTemplate))
    ? baseTemplate
    : (baseList[0] || null);

  if (!base) return { error: 'Không tìm thấy template gốc để copy cấu trúc' };

  const baseData = fs.readFileSync(path.join(TEMPLATES_DIR, base));
  const zip = await JSZip.loadAsync(baseData);

  // Build new document.xml body from lines of text
  const lines = content.split('\n');
  let bodyXml = '<w:body>';
  for (const line of lines) {
    if (line.trim() === '') {
      bodyXml += '<w:p><w:pPr/></w:p>';
      continue;
    }
    // Split line into literal text and {variable} parts
    const parts = line.split(/(\{[^{}]+\})/);
    bodyXml += '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>';
    for (const part of parts) {
      if (!part) continue;
      if (part.startsWith('{') && part.endsWith('}')) {
        // Variable: keep accent style (yellow highlight removed, keep text)
        bodyXml += `<w:r><w:rPr><w:color w:val="C0504D"/></w:rPr><w:t xml:space="preserve">${escapeXml(part)}</w:t></w:r>`;
      } else {
        bodyXml += `<w:r><w:t xml:space="preserve">${escapeXml(part)}</w:t></w:r>`;
      }
    }
    bodyXml += '</w:p>';
  }
  bodyXml += '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>';

  // Replace document.xml body
  let docXml = await zip.files['word/document.xml'].async('text');
  docXml = docXml.replace(/<w:body>[\s\S]*<\/w:body>/, bodyXml);
  zip.file('word/document.xml', docXml);

  const destPath = path.join(TEMPLATES_DIR, name);
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(destPath, buf);

  // Scan for variables in content
  const varRe = /\{([^{}]+)\}/g;
  const vars = new Set();
  let m;
  while ((m = varRe.exec(content)) !== null) vars.add(m[1]);

  // Load catalog, find new vars not yet registered
  const catalog = loadCatalog();
  const newVars = [...vars].filter(v => !catalog[v]);
  for (const v of newVars) {
    // Determine owner from prefix
    let owner = 'GLOBAL';
    if (v.startsWith('LIST_NTCV__'))    owner = 'LIST_NTCV';
    else if (v.startsWith('LIST_VAT_LIEU__')) owner = 'LIST_VAT_LIEU';
    else if (v.startsWith('LIST_YC_NTCV__')) owner = 'LIST_YC_NTCV';
    catalog[v] = { owner, label: v.replace(/^[A-Z_]+__/, '').replace(/_/g, ' '), files: [name], count: 0, sample: '' };
  }
  // Ensure all vars in this template are registered in catalog
  for (const v of vars) {
    if (catalog[v] && Array.isArray(catalog[v].files) && !catalog[v].files.includes(name)) {
      catalog[v].files.push(name);
    }
  }
  // Write updated catalog
  fs.writeFileSync(CATALOG_RUNTIME, JSON.stringify(catalog, null, 2), 'utf-8');

  return { ok: true, name, newVars };
});

// Delete a template file
ipcMain.handle('delete-template', async (_event, { templateName }) => {
  await ensureRuntime();
  const tplPath = path.join(TEMPLATES_DIR, templateName);
  if (!fs.existsSync(tplPath)) return { error: 'Không tìm thấy file' };
  fs.unlinkSync(tplPath);
  return { ok: true };
});

// Import template from user-chosen docx file (copy into runtime templates dir)
ipcMain.handle('import-template', async () => {
  const result = await dialog.showOpenDialog({
    filters: [{ name: 'Word Document', extensions: ['docx'] }],
    properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return null;
  const srcPath = result.filePaths[0];
  const name    = path.basename(srcPath);
  const destPath = path.join(TEMPLATES_DIR, name);
  fs.copyFileSync(srcPath, destPath);

  // Read content for editor pre-fill
  const data = fs.readFileSync(destPath);
  const zip  = await JSZip.loadAsync(data);
  const docEntry = zip.files['word/document.xml'];
  const paragraphs = [];
  if (docEntry) {
    const docXml = await docEntry.async('text');
    const paraRe = /<w:p[ >][\s\S]*?<\/w:p>/g;
    let pm;
    while ((pm = paraRe.exec(docXml)) !== null) {
      const bits = [];
      const tRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
      let tm;
      while ((tm = tRe.exec(pm[0])) !== null) bits.push(tm[1]);
      paragraphs.push(bits.join(''));
    }
  }
  return { name, paragraphs };
});

ipcMain.handle('open-json', async () => {
  const result = await dialog.showOpenDialog({
    filters: [{ name: 'JSON', extensions: ['json'] }],
    properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return null;
  const content = fs.readFileSync(result.filePaths[0], 'utf-8');
  return { path: result.filePaths[0], data: JSON.parse(content) };
});

ipcMain.handle('load-demo-project', async () => {
  try {
    const { path: demoPath, data } = loadDemoDataObject();
    return { path: demoPath, data };
  } catch (e) {
    return { error: e.message };
  }
});

ipcMain.handle('edit-label-local', async (_event, { key, label }) => {
  if (!key || !String(key).trim()) return { error: 'Thiếu key trường' };
  if (!label || !String(label).trim()) return { error: 'Label không được để trống' };

  await ensureRuntime();

  const catalog = loadCatalog();
  if (!catalog[key]) return { error: `Không tồn tại key: ${key}` };

  catalog[key] = { ...catalog[key], label: String(label).trim() };

  // Chỉ sửa local runtime (không tự đồng bộ toàn bộ máy nữa)
  writeIfExistsOrCreatable(CATALOG_RUNTIME, JSON.stringify(catalog, null, 2));

  // Sync demo data local runtime
  const runtimeDemo = getRuntimeDemoPath();
  if (fs.existsSync(runtimeDemo)) {
    try {
      const demoObj = JSON.parse(fs.readFileSync(runtimeDemo, 'utf-8'));
      const updated = syncDemoLabels(demoObj, key, String(label).trim());
      fs.writeFileSync(runtimeDemo, JSON.stringify(updated, null, 2), 'utf-8');
    } catch { /* ignore */ }
  }

  return { ok: true, key, label: String(label).trim() };
});

ipcMain.handle('export-label-pack', async () => {
  await ensureRuntime();

  const version = APP_VERSION;
  const dateStamp = new Date().toISOString().slice(0, 10);

  const catalogText = JSON.stringify(loadCatalog(), null, 2);
  const demoObj = loadDemoDataObject().data;
  const demoText = JSON.stringify(demoObj, null, 2);

  const result = await dialog.showOpenDialog({
    title: 'Chọn thư mục để xuất gói label',
    properties: ['openDirectory', 'createDirectory'],
  });
  if (result.canceled || !result.filePaths?.length) return null;

  const outDir = result.filePaths[0];
  const catName = `catalog-v${version}-${dateStamp}.json`;
  const demoName = `project-data.demo.friendly-v${version}-${dateStamp}.json`;

  const catPath = path.join(outDir, catName);
  const demoPath = path.join(outDir, demoName);

  fs.writeFileSync(catPath, catalogText, 'utf-8');
  fs.writeFileSync(demoPath, demoText, 'utf-8');

  return { ok: true, outDir, files: [catPath, demoPath] };
});

ipcMain.handle('import-label-pack', async () => {
  await ensureRuntime();

  const result = await dialog.showOpenDialog({
    title: 'Chọn 2 file label để import (catalog + demo)',
    filters: [{ name: 'JSON', extensions: ['json'] }],
    properties: ['openFile', 'multiSelections'],
  });
  if (result.canceled || !result.filePaths?.length) return null;

  const files = result.filePaths;
  const catFile = files.find(f => /catalog/i.test(path.basename(f)));
  const demoFile = files.find(f => /project-data\.demo\.friendly/i.test(path.basename(f)) || /demo/i.test(path.basename(f)));

  if (!catFile || !demoFile) {
    return { error: 'Cần chọn đủ 2 file: catalog*.json và project-data.demo.friendly*.json' };
  }

  let catObj, demoObj;
  try {
    catObj = JSON.parse(fs.readFileSync(catFile, 'utf-8'));
    demoObj = JSON.parse(fs.readFileSync(demoFile, 'utf-8'));
  } catch (e) {
    return { error: 'File JSON không hợp lệ: ' + e.message };
  }

  // Ghi đè local runtime
  fs.writeFileSync(CATALOG_RUNTIME, JSON.stringify(catObj, null, 2), 'utf-8');
  fs.writeFileSync(getRuntimeDemoPath(), JSON.stringify(demoObj, null, 2), 'utf-8');

  return { ok: true, imported: [catFile, demoFile] };
});

ipcMain.handle('save-json', async (_event, { filePath, data }) => {
  if (!filePath) {
    const result = await dialog.showSaveDialog({
      filters: [{ name: 'JSON', extensions: ['json'] }],
      defaultPath: 'project-data.json',
    });
    if (result.canceled) return null;
    filePath = result.filePath;
  }
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf-8');
  return filePath;
});

// ── Export blank JSON template ────────────────────────────────────────────────
ipcMain.handle('export-sample-json', async () => {
  await ensureRuntime();
  const catalog = loadCatalog();

  // Build a friendly sample project with all fields empty
  const globalFields = Object.entries(catalog)
    .filter(([, m]) => m.owner === 'GLOBAL')
    .map(([key, m]) => ({ key, label: m.label || key, value: '' }));

  const sample = {
    meta: { version: 4, description: 'File JSON mẫu – điền giá trị vào trường "value"' },
    global_fields: globalFields,
    jobs: [],
  };

  const { filePath, canceled } = await dialog.showSaveDialog({
    title: 'Lưu file JSON mẫu',
    defaultPath: 'project-data-mau.json',
    filters: [{ name: 'JSON', extensions: ['json'] }],
  });
  if (canceled || !filePath) return null;
  fs.writeFileSync(filePath, JSON.stringify(sample, null, 2), 'utf-8');
  return filePath;
});

// ── Export styled Excel template ──────────────────────────────────────────────
ipcMain.handle('export-sample-excel', async () => {
  await ensureRuntime();
  const catalog = loadCatalog();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'BPHH DocGen';
  wb.created = new Date();

  // Color constants
  const NAVY  = '1F3864';
  const BLUE  = '2E75B6';
  const LBLUE = 'D6E4F0';
  const GREEN = '375623';
  const LGREEN= 'E2EFD9';
  const BROWN = '843C0C';
  const LBROWN= 'FCE4D6';
  const PURPLE= '4B0082';
  const LPURPLE='E8DCFA';

  const hdrFont  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: 'Calibri' };
  const keyFont  = { italic: true, color: { argb: 'FF888888' }, size: 8, name: 'Calibri' };
  const dataFont = { size: 9, color: { argb: 'FF1A1A1A' }, name: 'Calibri' };
  const bord = { style: 'thin', color: { argb: 'FFB0C4DE' } };
  const allBorders = { top: bord, left: bord, bottom: bord, right: bord };

  // ── Guide sheet ──
  const wsG = wb.addWorksheet('📖 Hướng dẫn', { properties: { tabColor: { argb: 'FF' + NAVY } } });
  wsG.columns = [{ width: 4 }, { width: 36 }, { width: 56 }];
  const bannerG = wsG.getRow(1);
  bannerG.height = 32;
  wsG.mergeCells('A1:C1');
  Object.assign(wsG.getCell('A1'), {
    value: 'BPHH DocGen – File Nhập Liệu Mẫu',
    font: { bold: true, size: 15, color: { argb: 'FFFFFFFF' }, name: 'Calibri' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  const guides = [
    ['Sheet', 'Nội dung'],
    ['🌐 Thông tin chung',  'Trường dùng chung – điền vào cột "Giá trị"'],
    ['📄 BB Nghiệm thu CV', 'Mỗi hàng = 1 biên bản NTCV'],
    ['🧱 Nghiệm thu VL',    'Mỗi hàng = 1 biên bản NT vật liệu'],
    ['✅ Yêu cầu NTCV',     'Mỗi hàng = 1 yêu cầu NTCV'],
    ['', ''],
    ['Lưu ý', 'Không sửa tên sheet, không xóa hàng tiêu đề'],
    ['Lưu ý', 'Hàng màu xám chứa tên biến – KHÔNG điền vào hàng này'],
    ['Lưu ý', 'Lưu file .xlsx rồi dùng 📥 Import Excel trong app'],
  ];
  guides.forEach(([b, c], i) => {
    const row = wsG.getRow(i + 2);
    row.height = 18;
    const isHdr = b === 'Sheet';
    const isNote = b === 'Lưu ý';
    const bg = isHdr ? 'FF' + BLUE : (isNote ? 'FFFFF8F0' : 'FFF8F8F8');
    wsG.getCell(i + 2, 2).value = b;
    wsG.getCell(i + 2, 2).font = { bold: isHdr, size: 10, color: { argb: isHdr ? 'FFFFFFFF' : 'FF333333' } };
    wsG.getCell(i + 2, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    wsG.getCell(i + 2, 3).value = c;
    wsG.getCell(i + 2, 3).font = { italic: isNote, size: 10, color: { argb: isNote ? 'FF8B0000' : 'FF333333' } };
    wsG.getCell(i + 2, 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    wsG.getCell(i + 2, 3).alignment = { wrapText: true, vertical: 'middle' };
  });

  // ── Global sheet ──
  const wsGlob = wb.addWorksheet('🌐 Thông tin chung', { properties: { tabColor: { argb: 'FF' + BLUE } } });
  wsGlob.columns = [{ width: 42 }, { width: 42 }];
  wsGlob.mergeCells('A1:B1');
  Object.assign(wsGlob.getCell('A1'), {
    value: 'THÔNG TIN CHUNG',
    font: { bold: true, size: 13, color: { argb: 'FFFFFFFF' }, name: 'Calibri' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  wsGlob.getRow(1).height = 28;
  ['Tên trường', 'Giá trị'].forEach((h, i) => {
    Object.assign(wsGlob.getCell(2, i + 1), {
      value: h, font: hdrFont,
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + BLUE } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: allBorders,
    });
  });
  wsGlob.getRow(2).height = 22;
  wsGlob.views = [{ state: 'frozen', ySplit: 2 }];

  const globalFields = Object.entries(catalog).filter(([, m]) => m.owner === 'GLOBAL');
  globalFields.forEach(([key, meta], i) => {
    const row = i + 3;
    const altBg = i % 2 === 0 ? 'FF' + LBLUE : 'FFFFFFFF';
    Object.assign(wsGlob.getCell(row, 1), {
      value: meta.label || key,
      font: { size: 10, color: { argb: 'FF1F2D3D' }, name: 'Calibri' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
      alignment: { horizontal: 'left', vertical: 'middle', wrapText: true },
      border: allBorders,
    });
    wsGlob.getCell(row, 2).border = allBorders;
    wsGlob.getCell(row, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    wsGlob.getRow(row).height = 18;
  });

  // ── Helper: list sheet ──
  function addListSheet(sheetName, owner, tabArgb, hdrArgb, altArgb, tplName) {
    const ws = wb.addWorksheet(sheetName, { properties: { tabColor: { argb: tabArgb } } });
    const fields = Object.entries(catalog).filter(([, m]) => m.owner === owner);
    const ncols = fields.length;
    ws.columns = fields.map(([, m]) => ({ width: Math.max(14, Math.min(28, (m.label || '').length * 1.2)) }));

    // banner
    ws.mergeCells(1, 1, 1, ncols);
    Object.assign(ws.getCell(1, 1), {
      value: sheetName.replace(/^[^\s]+ /, ''),
      font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' }, name: 'Calibri' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } },
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    ws.getRow(1).height = 26;

    // label row
    fields.forEach(([k, m], i) => {
      Object.assign(ws.getCell(2, i + 1), {
        value: m.label || k,
        font: { ...hdrFont, size: 9 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: hdrArgb } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: allBorders,
      });
    });
    ws.getRow(2).height = 34;

    // key row (row 3, light gray – for import parsing)
    fields.forEach(([k], i) => {
      Object.assign(ws.getCell(3, i + 1), {
        value: k,
        font: { ...keyFont },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } },
        alignment: { horizontal: 'center' },
      });
    });
    ws.getRow(3).height = 12;

    // 5 blank data rows
    for (let r = 4; r <= 8; r++) {
      fields.forEach((_, i) => {
        ws.getCell(r, i + 1).border = allBorders;
        ws.getCell(r, i + 1).fill = r % 2 === 0
          ? { type: 'pattern', pattern: 'solid', fgColor: { argb: altArgb } }
          : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      });
      ws.getRow(r).height = 20;
    }
    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  addListSheet('📄 BB Nghiệm thu CV', 'LIST_NTCV',   'FF375623', 'FF375623', 'FF' + LGREEN, 'BB Nghiệm thu công việc.docx');
  addListSheet('🧱 Nghiệm thu VL',    'LIST_VAT_LIEU','FF843C0C', 'FF843C0C', 'FF' + LBROWN, 'Nghiệm thu vật liệu.docx');
  addListSheet('✅ Yêu cầu NTCV',     'LIST_YC_NTCV', 'FF4B0082', 'FF4B0082', 'FF' + LPURPLE,'Yêu cầu nghiệm thu công việc.docx');

  const { filePath, canceled } = await dialog.showSaveDialog({
    title: 'Lưu file Excel mẫu',
    defaultPath: 'project-data-mau.xlsx',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  });
  if (canceled || !filePath) return null;
  await wb.xlsx.writeFile(filePath);
  return filePath;
});

// ── Import Excel → project data ───────────────────────────────────────────────

// ── Export current project data → Excel (round-trip importable) ─────────────
ipcMain.handle('export-current-excel', async (_event, { project }) => {
  const prj = project || {};
  const global = prj.global || {};
  const jobs = Array.isArray(prj.jobs) ? prj.jobs : [];

  const wb = new ExcelJS.Workbook();
  wb.creator = 'BPHH DocGen';
  wb.created = new Date();

  const NAVY  = '1F3864';
  const BLUE  = '2E75B6';
  const LBLUE = 'D6E4F0';
  const GREEN = '375623';
  const LGREEN= 'E2EFD9';
  const BROWN = '843C0C';
  const LBROWN= 'FCE4D6';
  const PURPLE= '4B0082';
  const LPURPLE='E8DCFA';

  const hdrFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: 'Calibri' };
  const bord = { style: 'thin', color: { argb: 'FFB0C4DE' } };
  const allBorders = { top: bord, left: bord, bottom: bord, right: bord };

  // Guide
  const wsG = wb.addWorksheet('📖 Hướng dẫn', { properties: { tabColor: { argb: 'FF' + NAVY } } });
  wsG.columns = [{ width: 4 }, { width: 36 }, { width: 56 }];
  wsG.mergeCells('A1:C1');
  Object.assign(wsG.getCell('A1'), {
    value: 'BPHH DocGen – File dữ liệu export',
    font: { bold: true, size: 15, color: { argb: 'FFFFFFFF' }, name: 'Calibri' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  wsG.getRow(1).height = 32;
  const guides = [
    ['Sheet', 'Nội dung'],
    ['🌐 Thông tin chung',  'Trường dùng chung – cột B là giá trị'],
    ['📄 BB Nghiệm thu CV', 'Mỗi hàng = 1 job đang có dữ liệu'],
    ['🧱 Nghiệm thu VL',    'Mỗi hàng = 1 job đang có dữ liệu'],
    ['✅ Yêu cầu NTCV',     'Mỗi hàng = 1 job đang có dữ liệu'],
    ['', ''],
    ['Lưu ý', 'Có thể gửi file này cho người khác để Import Excel trực tiếp'],
    ['Lưu ý', 'Không sửa tên sheet, không sửa hàng key (hàng xám)'],
    ['Lưu ý', 'Các cột không có dữ liệu có thể để trống'],
  ];
  guides.forEach(([b, c], i) => {
    const row = wsG.getRow(i + 2);
    row.height = 18;
    const isHdr = b === 'Sheet';
    const isNote = b === 'Lưu ý';
    const bg = isHdr ? 'FF' + BLUE : (isNote ? 'FFFFF8F0' : 'FFF8F8F8');
    wsG.getCell(i + 2, 2).value = b;
    wsG.getCell(i + 2, 2).font = { bold: isHdr, size: 10, color: { argb: isHdr ? 'FFFFFFFF' : 'FF333333' } };
    wsG.getCell(i + 2, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    wsG.getCell(i + 2, 3).value = c;
    wsG.getCell(i + 2, 3).font = { italic: isNote, size: 10, color: { argb: isNote ? 'FF8B0000' : 'FF333333' } };
    wsG.getCell(i + 2, 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    wsG.getCell(i + 2, 3).alignment = { wrapText: true, vertical: 'middle' };
  });

  // Global sheet
  const wsGlob = wb.addWorksheet('🌐 Thông tin chung', { properties: { tabColor: { argb: 'FF' + BLUE } } });
  wsGlob.columns = [{ width: 42 }, { width: 42 }];
  wsGlob.mergeCells('A1:B1');
  Object.assign(wsGlob.getCell('A1'), {
    value: 'THÔNG TIN CHUNG',
    font: { bold: true, size: 13, color: { argb: 'FFFFFFFF' }, name: 'Calibri' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  });
  wsGlob.getRow(1).height = 28;
  ['Tên trường', 'Giá trị'].forEach((h, i) => {
    Object.assign(wsGlob.getCell(2, i + 1), {
      value: h, font: hdrFont,
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + BLUE } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: allBorders,
    });
  });
  wsGlob.getRow(2).height = 22;
  wsGlob.views = [{ state: 'frozen', ySplit: 2 }];

  const catalog = loadCatalog();
  const globalFields = Object.entries(catalog).filter(([, m]) => m.owner === 'GLOBAL');
  globalFields.forEach(([key, meta], i) => {
    const row = i + 3;
    const altBg = i % 2 === 0 ? 'FF' + LBLUE : 'FFFFFFFF';
    Object.assign(wsGlob.getCell(row, 1), {
      value: meta.label || key,
      font: { size: 10, color: { argb: 'FF1F2D3D' }, name: 'Calibri' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
      alignment: { horizontal: 'left', vertical: 'middle', wrapText: true },
      border: allBorders,
    });
    wsGlob.getCell(row, 2).value = String(global[key] || '');
    wsGlob.getCell(row, 2).border = allBorders;
    wsGlob.getCell(row, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    wsGlob.getRow(row).height = 18;
  });

  function addListSheetWithData(sheetName, owner, tabArgb, hdrArgb, altArgb, templateName) {
    const ws = wb.addWorksheet(sheetName, { properties: { tabColor: { argb: tabArgb } } });
    const fields = Object.entries(catalog).filter(([, m]) => m.owner === owner);
    const ncols = fields.length;
    ws.columns = fields.map(([, m]) => ({ width: Math.max(14, Math.min(28, (m.label || '').length * 1.2)) }));

    ws.mergeCells(1, 1, 1, ncols);
    Object.assign(ws.getCell(1, 1), {
      value: sheetName.replace(/^[^\s]+ /, ''),
      font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' }, name: 'Calibri' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } },
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    ws.getRow(1).height = 26;

    fields.forEach(([k, m], i) => {
      Object.assign(ws.getCell(2, i + 1), {
        value: m.label || k,
        font: { ...hdrFont, size: 9 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: hdrArgb } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: allBorders,
      });
    });
    ws.getRow(2).height = 34;

    fields.forEach(([k], i) => {
      Object.assign(ws.getCell(3, i + 1), {
        value: k,
        font: { italic: true, color: { argb: 'FF888888' }, size: 8, name: 'Calibri' },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } },
        alignment: { horizontal: 'center' },
      });
    });
    ws.getRow(3).height = 12;

    const rows = jobs.filter(j => j.enabled !== false && j.template === templateName);
    if (!rows.length) {
      for (let r = 4; r <= 8; r++) {
        fields.forEach((_, i) => {
          ws.getCell(r, i + 1).border = allBorders;
          ws.getCell(r, i + 1).fill = r % 2 === 0
            ? { type: 'pattern', pattern: 'solid', fgColor: { argb: altArgb } }
            : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
        });
        ws.getRow(r).height = 20;
      }
    } else {
      rows.forEach((job, idx) => {
        const r = 4 + idx;
        const jf = job.fields || {};
        fields.forEach(([k], cidx) => {
          ws.getCell(r, cidx + 1).value = String(jf[k] || '');
          ws.getCell(r, cidx + 1).border = allBorders;
          ws.getCell(r, cidx + 1).fill = r % 2 === 0
            ? { type: 'pattern', pattern: 'solid', fgColor: { argb: altArgb } }
            : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
          ws.getCell(r, cidx + 1).font = { size: 9, color: { argb: 'FF1A1A1A' }, name: 'Calibri' };
        });
        ws.getRow(r).height = 20;
      });
    }

    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  addListSheetWithData('📄 BB Nghiệm thu CV', 'LIST_NTCV',    'FF375623', 'FF375623', 'FF' + LGREEN,  'BB Nghiệm thu công việc.docx');
  addListSheetWithData('🧱 Nghiệm thu VL',    'LIST_VAT_LIEU', 'FF843C0C', 'FF843C0C', 'FF' + LBROWN,  'Nghiệm thu vật liệu.docx');
  addListSheetWithData('✅ Yêu cầu NTCV',     'LIST_YC_NTCV',  'FF4B0082', 'FF4B0082', 'FF' + LPURPLE, 'Yêu cầu nghiệm thu công việc.docx');

  const { filePath, canceled } = await dialog.showSaveDialog({
    title: 'Lưu file Excel dữ liệu hiện tại',
    defaultPath: 'project-data-export.xlsx',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  });
  if (canceled || !filePath) return null;
  await wb.xlsx.writeFile(filePath);
  return filePath;
});

ipcMain.handle('import-excel', async () => {
  const openResult = await dialog.showOpenDialog({
    title: 'Chọn file Excel nhập liệu',
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile'],
  });
  if (openResult.canceled || !openResult.filePaths.length) return null;

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(openResult.filePaths[0]);

  const catalog = loadCatalog();

  // ── 1. Global fields from "🌐 Thông tin chung" ──────────────────────────────
  const globalFields = {};
  const wsGlob = wb.getWorksheet('🌐 Thông tin chung') || wb.worksheets.find(w => /thông tin chung/i.test(w.name));
  if (wsGlob) {
    wsGlob.eachRow((row, rowNum) => {
      if (rowNum < 3) return;  // skip title + header rows
      const label = String(row.getCell(1).value || '').trim();
      const value = String(row.getCell(2).value || '').trim();
      if (!label || !value) return;
      // Find key by label match in catalog
      const entry = Object.entries(catalog).find(([, m]) =>
        m.owner === 'GLOBAL' && m.label && m.label.toLowerCase().trim() === label.toLowerCase().trim()
      );
      if (entry) globalFields[entry[0]] = value;
    });
  }

  // ── 2. List fields helper: parse a sheet into jobs ─────────────────────────
  function parseListSheet(sheetPattern, templateName) {
    const ws = wb.getWorksheet(sheetPattern) || wb.worksheets.find(w =>
      typeof sheetPattern === 'string' ? w.name.includes(sheetPattern) : sheetPattern.test(w.name)
    );
    if (!ws) return [];

    // Detect key row (row 3 in our template, but also detect by looking for LIST_ prefixed cells)
    let keyRow = 3;
    ws.eachRow((row, rowNum) => {
      if (rowNum > 5) return;
      const firstCell = String(row.getCell(1).value || '');
      if (firstCell.startsWith('LIST_') || firstCell.startsWith('GLOBAL__')) keyRow = rowNum;
    });

    // Build col→key map from key row
    const colKeyMap = {};
    ws.getRow(keyRow).eachCell((cell, colNum) => {
      const val = String(cell.value || '').trim();
      if (val && catalog[val]) colKeyMap[colNum] = val;
    });

    const jobs = [];
    ws.eachRow((row, rowNum) => {
      if (rowNum <= keyRow) return;  // skip header + key rows
      const fields = {};
      let hasData = false;
      row.eachCell((cell, colNum) => {
        const key = colKeyMap[colNum];
        if (!key) return;
        const val = String(cell.value || '').trim();
        if (val) { fields[key] = val; hasData = true; }
      });
      if (hasData) {
        jobs.push({
          template: templateName,
          output:   templateName,
          enabled:  true,
          fields,
        });
      }
    });
    return jobs;
  }

  const allJobs = [
    ...parseListSheet('BB Nghiệm thu CV', 'BB Nghiệm thu công việc.docx'),
    ...parseListSheet('Nghiệm thu VL',    'Nghiệm thu vật liệu.docx'),
    ...parseListSheet('Yêu cầu NTCV',     'Yêu cầu nghiệm thu công việc.docx'),
  ];

  return {
    project: {
      meta: { version: 4, description: 'Nhập từ Excel' },
      global: globalFields,
      jobs: allJobs,
    },
    counts: { global: Object.keys(globalFields).length, jobs: allJobs.length },
  };
});

ipcMain.handle('choose-outdir', async () => {
  const result = await dialog.showOpenDialog({ properties: ['openDirectory', 'createDirectory'] });
  if (result.canceled) return null;
  return result.filePaths[0];
});

ipcMain.handle('generate-docs', async (_event, { jobs, globalFields }) => {
  await ensureRuntime();
  const result = await dialog.showOpenDialog({ properties: ['openDirectory', 'createDirectory'] });
  if (result.canceled) return { canceled: true };
  const outDir = result.filePaths[0];

  let generated = 0, skipped = 0;
  const errors = [];

  for (const job of jobs) {
    if (!job.enabled) { skipped++; continue; }
    const tplPath = path.join(TEMPLATES_DIR, job.template);
    if (!fs.existsSync(tplPath)) { skipped++; errors.push(`Thiếu template: ${job.template}`); continue; }

    try {
      const buf = fs.readFileSync(tplPath);
      const zip = new PizZip(buf);
      const delimiters = detectTemplateDelimitersFromZip(zip);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters,
        nullGetter: () => '',
      });

      // Merge global + job fields; strip highlight is handled by Docxtemplater
      const merged = { ...globalFields, ...(job.fields || {}) };
      doc.render(merged);

      // Remove yellow highlights left in XML by template placeholders
      removeHighlights(doc.getZip());

      const out = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
      const outPath = path.join(outDir, job.output || job.template);
      fs.mkdirSync(path.dirname(outPath), { recursive: true });
      fs.writeFileSync(outPath, out);
      generated++;
    } catch (e) {
      errors.push(`${job.template}: ${e.message}`);
      skipped++;
    }
  }

  return { generated, skipped, outDir, errors };
});

// ── Open path in Finder/Explorer ─────────────────────────────────────────────
ipcMain.handle('open-path', async (_event, target) => {
  // target can be a folder path or a URL (for download links)
  if (target && (target.startsWith('http://') || target.startsWith('https://'))) {
    await shell.openExternal(target);
  } else if (target && fs.existsSync(target)) {
    await shell.openPath(target);
  }
  return null;
});

// ── Check for remote update ───────────────────────────────────────────────────
ipcMain.handle('get-version', () => APP_VERSION);
ipcMain.handle('refresh-runtime-assets', async () => {
  if (fs.existsSync(VERSION_FILE)) fs.unlinkSync(VERSION_FILE);
  if (fs.existsSync(FINGERPRINT_FILE)) fs.unlinkSync(FINGERPRINT_FILE);
  await ensureRuntime();
  return { ok: true };
});
ipcMain.handle('get-platform', () => ({ platform: process.platform, arch: process.arch }));

ipcMain.handle('check-update', async () => {
  const settings  = loadSettings();
  const updateUrl = settings.update_url || DEFAULT_UPDATE_URL;

  try {
    const buf  = await fetchUrl(updateUrl);
    const info = JSON.parse(buf.toString('utf-8'));
    const hasUpdate = info.version && info.version !== APP_VERSION
      && info.version.localeCompare(APP_VERSION, undefined, { numeric: true }) > 0;

    // Pick installer URL for current platform/arch
    let installUrl = null;
    if (process.platform === 'darwin') {
      installUrl = process.arch === 'arm64' ? (info.mac_arm64 || null) : (info.mac_x64 || null);
    } else if (process.platform === 'win32') {
      installUrl = info.win_x64 || null;
    }

    return {
      hasUpdate,
      version:     info.version,
      changelog:   info.changelog || '',
      installUrl,
      catalogUrl:  info.catalog_url || null,
      templatesUrl: info.templates_url || null,
    };
  } catch (e) {
    return { error: e.message };
  }
});

// ── Download installer and open it ───────────────────────────────────────
ipcMain.handle('install-update', async (_event, { url }) => {
  if (!url) return { error: 'No URL' };
  try {
    const fname = decodeURIComponent(path.basename(url.split('?')[0]));
    const dest  = path.join(os.homedir(), 'Downloads', fname);
    const buf   = await fetchUrl(url);
    fs.writeFileSync(dest, buf);
    if (process.platform === 'win32') {
      // spawn NSIS installer detached so app can exit
      require('child_process').spawn(dest, [], { detached: true, stdio: 'ignore' }).unref();
    } else {
      await shell.openPath(dest);  // macOS: Finder opens DMG for drag-install
    }
    return { ok: true, dest };
  } catch (e) {
    return { error: e.message };
  }
});

// ── Autosave (silent persist to disk) ───────────────────────────────────────
const AUTOSAVE_FILE = path.join(RUNTIME_DIR, 'autosave.json');

ipcMain.handle('autosave-load', async () => {
  try {
    if (!fs.existsSync(AUTOSAVE_FILE)) return null;
    return JSON.parse(fs.readFileSync(AUTOSAVE_FILE, 'utf-8'));
  } catch { return null; }
});

ipcMain.handle('autosave-save', async (_event, payload) => {
  try {
    fs.mkdirSync(RUNTIME_DIR, { recursive: true });
    fs.writeFileSync(AUTOSAVE_FILE, JSON.stringify(payload), 'utf-8');
  } catch { /* ignore */ }
});

ipcMain.handle('autosave-clear', async () => {
  try { if (fs.existsSync(AUTOSAVE_FILE)) fs.unlinkSync(AUTOSAVE_FILE); } catch { /* ignore */ }
});

// ── Download & apply remote catalog / templates ───────────────────────────────
ipcMain.handle('download-update', async (_event, { type, url }) => {
  if (!url) return { error: 'URL không được cung cấp' };
  try {
    const buf = await fetchUrl(url);
    if (type === 'catalog') {
      fs.mkdirSync(RUNTIME_DIR, { recursive: true });
      fs.writeFileSync(CATALOG_RUNTIME, buf);
      return { ok: true };
    }
    if (type === 'templates') {
      fs.mkdirSync(RUNTIME_DIR, { recursive: true });
      const zip = await JSZip.loadAsync(buf);
      for (const [relPath, entry] of Object.entries(zip.files)) {
        if (entry.dir) continue;
        const dest = path.join(RUNTIME_DIR, relPath);
        fs.mkdirSync(path.dirname(dest), { recursive: true });
        fs.writeFileSync(dest, await entry.async('nodebuffer'));
      }
      return { ok: true };
    }
    return { error: 'Loại cập nhật không hợp lệ' };
  } catch (e) {
    return { error: e.message };
  }
});

// ── Window ────────────────────────────────────────────────────────────────────

function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1200,
    minHeight: 700,
    title: 'BPHH DocGen',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (isDev) {
    win.loadURL('http://localhost:5173');
    win.webContents.openDevTools({ mode: 'detach' });
  } else {
    win.loadFile(path.join(__dirname, '../dist/index.html'));
  }
}

app.whenReady().then(async () => {
  // Startup guard: always check runtime fingerprint against bundled assets.
  // If bundle/catalog changed, ensureRuntime() will force re-extract automatically.
  await ensureRuntime();
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
