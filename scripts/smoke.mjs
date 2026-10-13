#!/usr/bin/env node
// End-to-end smoke test for the code paths behind `render_document`.
//
// Run it in a NORMAL terminal, outside the DSH sandbox: the engine starts a
// native child process with piped stdio, which the sandbox denies with EPERM
// (that denial is the reason this plugin exists). Inside a DSH session, call
// the `render_document` tool instead.
//
//   node scripts/smoke.mjs <document> [--engine libreoffice|word] [--pages 1,2]
//                          [--dpi 144] [--keep-pdf] [--json]
//
// The document's own directory acts as the workspace, so the workspace guard
// behaves exactly as it does in a session.

import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { findKit, renderImages, describeKitError } from '../lib/kit.mjs';
import { exportPdf, probeWord } from '../lib/word.mjs';
import {
  allocateOutputDir,
  parseDpi,
  parsePages,
  pickEngine,
  resolveInWorkspace,
  shapeResult,
} from '../lib/paths.mjs';

function parseArgs(argv) {
  const options = { engine: 'libreoffice', pages: undefined, dpi: undefined, keepPdf: false, json: false, source: undefined };
  for (let index = 0; index < argv.length; index++) {
    const arg = argv[index];
    if (arg === '--engine') options.engine = argv[++index];
    else if (arg === '--pages') options.pages = argv[++index];
    else if (arg === '--dpi') options.dpi = Number(argv[++index]);
    else if (arg === '--keep-pdf') options.keepPdf = true;
    else if (arg === '--json') options.json = true;
    else if (!arg.startsWith('--') && !options.source) options.source = arg;
    else throw new Error(`unexpected argument: ${arg}`);
  }
  return options;
}

const options = parseArgs(process.argv.slice(2));
if (!options.source) {
  console.error('usage: node scripts/smoke.mjs <document> [--engine word] [--pages 1,2] [--dpi 144] [--keep-pdf] [--json]');
  process.exit(2);
}

const workspace = path.dirname(path.resolve(options.source));
const sourcePath = resolveInWorkspace(workspace, options.source, 'source');
if (!existsSync(sourcePath)) throw new Error(`source not found: ${sourcePath}`);

const pages = parsePages(options.pages);
const dpi = parseDpi(options.dpi);
const kit = findKit();
const wordProbe = options.engine === 'word' ? await probeWord({ refresh: true }) : { available: false };
const { engine } = pickEngine({
  requested: options.engine,
  extension: path.extname(sourcePath).slice(1).toLowerCase(),
  wordAvailable: wordProbe.available,
  kitAvailable: Boolean(kit),
});

const outputDir = allocateOutputDir(workspace, `preview-${path.basename(sourcePath, path.extname(sourcePath))}`);
const started = Date.now();
let scratch;
try {
  let inputPath = sourcePath;
  let wordPdf;
  if (engine === 'word') {
    scratch = await mkdtemp(path.join(tmpdir(), 'dsh-render-document-smoke-'));
    wordPdf = path.join(scratch, `${path.basename(sourcePath, path.extname(sourcePath))}.pdf`);
    const exported = await exportPdf({ sourcePath, targetPath: wordPdf });
    console.error(`word: exported ${exported.pages} page(s) with Word ${exported.version ?? '?'}`);
    inputPath = wordPdf;
  }

  const manifest = await renderImages({ apiPath: kit.apiPath, inputPath, outputDir, pages, dpi });
  const result = shapeResult({ manifest, engine, outputDir, manifestPath: path.join(outputDir, 'manifest.json') });
  if (options.keepPdf && wordPdf) {
    const kept = path.join(outputDir, `${path.basename(sourcePath, path.extname(sourcePath))}.pdf`);
    await copyFile(wordPdf, kept);
    result.pdfPath = kept;
  }
  result.elapsedMs = Date.now() - started;

  if (options.json) console.log(JSON.stringify(result, null, 2));
  else {
    console.log(`ok: ${result.pageCount} page(s) via ${result.engine} in ${result.elapsedMs} ms`);
    console.log(`    output: ${result.outputDir}`);
    for (const image of result.images) console.log(`    page ${image.page}: ${image.path} (${image.width}x${image.height})`);
    if (result.missingFonts.length) console.log(`    missing fonts: ${result.missingFonts.join(', ')}`);
  }
} catch (error) {
  console.error(`smoke failed: ${describeKitError(error)}`);
  process.exitCode = 1;
} finally {
  if (scratch) await rm(scratch, { recursive: true, force: true }).catch(() => {});
}
