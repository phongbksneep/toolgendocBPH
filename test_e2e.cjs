'use strict';
/**
 * E2E test for Electron DOCX generation logic (no Electron runtime needed).
 * Tests: catalog load, template extraction, DOCX render, highlight removal, output validity.
 * Run: node test_e2e.cjs
 */

const path = require('path');
const fs   = require('fs');
const os   = require('os');

const JSZip         = require('jszip');
const Docxtemplater = require('docxtemplater');
const PizZip        = require('pizzip');

// ── paths (same as main.cjs dev mode) ────────────────────────────────────────
const ROOT         = __dirname;
const BUNDLE_PATH  = path.join(ROOT, 'assets', 'templates_bundle.zip');
const CATALOG_PATH = path.join(ROOT, 'catalog.json');
const OUT_DIR      = path.join(os.tmpdir(), 'bphh_e2e_out');

// Keep track of pass/fail
let passed = 0, failed = 0;
function ok(name)       { console.log(`  ✅  ${name}`); passed++; }
function fail(name, e)  { console.error(`  ❌  ${name}: ${e}`); failed++; }

// ── helpers (mirrors main.cjs) ───────────────────────────────────────────────
function removeHighlights(zip) {
  const xmlFiles = Object.keys(zip.files).filter(
    f => f.startsWith('word/') && f.endsWith('.xml')
  );
  for (const xmlFile of xmlFiles) {
    const entry = zip.file(xmlFile);
    if (!entry) continue;
    let xml = entry.asText();
    xml = xml.replace(/<w:highlight[^>]*\/>/g, '');
    zip.file(xmlFile, xml);
  }
}

async function run() {
  console.log('=== BPHH DocGen – E2E test ===\n');

  // ── 1. catalog.json load ──────────────────────────────────────────────────
  let catalog;
  try {
    catalog = JSON.parse(fs.readFileSync(CATALOG_PATH, 'utf-8'));
    const count = Object.keys(catalog).length;
    ok(`catalog.json loaded – ${count} fields`);
    if (count < 50) fail('catalog field count', `only ${count} fields, expected ≥50`);
    else ok('catalog field count ≥ 50');
  } catch (e) {
    fail('catalog.json load', e.message);
    return;
  }

  // ── 2. templates_bundle.zip extraction ───────────────────────────────────
  let templates;
  try {
    const data  = fs.readFileSync(BUNDLE_PATH);
    const jszip = await JSZip.loadAsync(data);
    fs.mkdirSync(OUT_DIR, { recursive: true });
    const templateFiles = [];
    for (const [relPath, entry] of Object.entries(jszip.files)) {
      if (entry.dir) continue;
      const dest = path.join(OUT_DIR, relPath);
      fs.mkdirSync(path.dirname(dest), { recursive: true });
      const buf = await entry.async('nodebuffer');
      fs.writeFileSync(dest, buf);
      if (relPath.startsWith('templates/') && relPath.endsWith('.docx')) {
        templateFiles.push(relPath);
      }
    }
    templates = templateFiles;
    ok(`templates_bundle.zip extracted – ${templateFiles.length} templates`);
  } catch (e) {
    fail('templates_bundle.zip extraction', e.message);
    return;
  }

  // ── 3. catalog has labels for all fields ─────────────────────────────────
  let unlabeled = 0;
  for (const [key, meta] of Object.entries(catalog)) {
    if (!meta.label) unlabeled++;
  }
  if (unlabeled === 0) ok('All catalog fields have labels');
  else fail(`${unlabeled} catalog fields missing label`, '');

  // ── 4. Generate DOCX for each template ───────────────────────────────────
  // Build test data: fill every catalog key with a visible marker
  const testData = {};
  for (const key of Object.keys(catalog)) {
    testData[key] = `[${key}]`;
  }

  let genOk = 0, genFail = 0;
  const docxOutDir = path.join(OUT_DIR, 'generated');
  fs.mkdirSync(docxOutDir, { recursive: true });

  for (const tplRel of templates) {
    const tplPath = path.join(OUT_DIR, tplRel);
    const outName = path.basename(tplRel);
    const outPath = path.join(docxOutDir, outName);
    try {
      const buf = fs.readFileSync(tplPath);
      const zip = new PizZip(buf);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '{{', end: '}}' },
        nullGetter: () => '',
      });
      doc.render(testData);
      removeHighlights(doc.getZip());
      const out = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
      fs.writeFileSync(outPath, out);
      genOk++;
    } catch (e) {
      console.error(`     ↳ render error in ${outName}: ${e.message}`);
      genFail++;
    }
  }

  if (genFail === 0) ok(`All ${genOk} templates rendered without error`);
  else fail(`${genFail} templates failed to render`, `${genOk} ok`);

  // ── 5. Verify no leftover {{...}} placeholders ────────────────────────────
  let unreplaced = 0;
  const generatedFiles = fs.readdirSync(docxOutDir).filter(f => f.endsWith('.docx'));
  for (const fname of generatedFiles) {
    const fpath = path.join(docxOutDir, fname);
    const zip = new PizZip(fs.readFileSync(fpath));
    const docXml = zip.file('word/document.xml');
    if (!docXml) continue;
    const xml = docXml.asText();
    const leftover = [...xml.matchAll(/\{\{([^}]+)\}\}/g)].map(m => m[1]);
    if (leftover.length) {
      console.error(`     ↳ ${fname}: unreplaced tags: ${[...new Set(leftover)].slice(0, 5).join(', ')}`);
      unreplaced++;
    }
  }
  if (unreplaced === 0) ok('No unreplaced {{...}} tags in any output DOCX');
  else fail(`${unreplaced} files still have unreplaced placeholders`, '');

  // ── 6. Verify no w:highlight in output ───────────────────────────────────
  let highlighted = 0;
  for (const fname of generatedFiles) {
    const fpath = path.join(docxOutDir, fname);
    const zip  = new PizZip(fs.readFileSync(fpath));
    const xmlFiles = Object.keys(zip.files).filter(f => f.startsWith('word/') && f.endsWith('.xml'));
    for (const xf of xmlFiles) {
      const xml = zip.file(xf)?.asText() || '';
      if (/<w:highlight/.test(xml)) {
        console.error(`     ↳ ${fname}/${xf} still has w:highlight!`);
        highlighted++;
        break;
      }
    }
  }
  if (highlighted === 0) ok('No w:highlight in any output DOCX');
  else fail(`${highlighted} files still contain w:highlight`, '');

  // ── 7. Output file sizes are reasonable ──────────────────────────────────
  let tooSmall = 0;
  for (const fname of generatedFiles) {
    const sz = fs.statSync(path.join(docxOutDir, fname)).size;
    if (sz < 5000) { console.error(`     ↳ ${fname} too small: ${sz} bytes`); tooSmall++; }
  }
  if (tooSmall === 0) ok(`All ${generatedFiles.length} output files have reasonable size`);
  else fail(`${tooSmall} output files suspiciously small`, '');

  // ── Summary ───────────────────────────────────────────────────────────────
  console.log(`\n${'─'.repeat(40)}`);
  console.log(`Results: ${passed} passed, ${failed} failed`);
  console.log(`Output files: ${OUT_DIR}/generated/`);
  if (failed > 0) process.exit(1);
}

run().catch(e => { console.error('FATAL:', e); process.exit(1); });
