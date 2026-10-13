import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';

import {
  DEFAULT_DPI,
  allocateOutputDir,
  isInside,
  parseDpi,
  parsePages,
  pickEngine,
  resolveInWorkspace,
  shapeResult,
} from '../lib/paths.mjs';

test('parsePages accepts the documented selections', () => {
  assert.equal(parsePages(undefined), 'all');
  assert.equal(parsePages(''), 'all');
  assert.equal(parsePages('all'), 'all');
  assert.equal(parsePages('ALL'), 'all');
  assert.deepEqual(parsePages('3'), [3]);
  assert.deepEqual(parsePages('1,3-5'), [1, 3, 4, 5]);
  assert.deepEqual(parsePages('2, 4'), [2, 4]);
  assert.deepEqual(parsePages('5-5'), [5]);
  assert.deepEqual(parsePages('3,1,3'), [1, 3], 'duplicates collapse and the result is sorted');
});

test('parsePages rejects malformed selections', () => {
  assert.throws(() => parsePages('0'), /below page 1/);
  assert.throws(() => parsePages('5-3'), /ends before it starts/);
  assert.throws(() => parsePages('abc'), /cannot parse/);
  assert.throws(() => parsePages('1-2000'), /more than 1000 pages/);
});

test('parseDpi validates the engine range', () => {
  assert.equal(parseDpi(undefined), DEFAULT_DPI);
  assert.equal(parseDpi(24), 24);
  assert.equal(parseDpi(600), 600);
  assert.throws(() => parseDpi(23), /between 24 and 600/);
  assert.throws(() => parseDpi(601), /between 24 and 600/);
  assert.throws(() => parseDpi(144.5), /between 24 and 600/);
  assert.throws(() => parseDpi('high'), /between 24 and 600/);
});

test('isInside keeps paths in the workspace', () => {
  const root = path.resolve('C:/work/project');
  assert.equal(isInside(root, path.join(root, 'doc.docx')), true);
  assert.equal(isInside(root, path.join(root, 'nested', 'deep', 'a.png')), true);
  assert.equal(isInside(root, root), true);
  assert.equal(isInside(root, path.resolve('C:/work/other/doc.docx')), false);
  assert.equal(isInside(root, path.resolve('C:/work/project/../escape.docx')), false);
});

test('resolveInWorkspace refuses escapes and empty input', () => {
  const cwd = path.resolve('C:/work/project');
  assert.equal(resolveInWorkspace(cwd, 'doc.docx'), path.join(cwd, 'doc.docx'));
  assert.equal(resolveInWorkspace(cwd, path.join(cwd, 'sub', 'x.xlsx')), path.join(cwd, 'sub', 'x.xlsx'));
  assert.throws(() => resolveInWorkspace(cwd, '../secret.docx', 'source'), /outside the session workspace/);
  assert.throws(() => resolveInWorkspace(cwd, 'C:/Windows/System32/clock.avi', 'source'), /outside the session workspace/);
  assert.throws(() => resolveInWorkspace(cwd, '   ', 'source'), /non-empty path/);
});

test('allocateOutputDir returns the first free name', () => {
  const cwd = path.resolve('C:/work/project');
  const taken = new Set([path.join(cwd, 'preview-report'), path.join(cwd, 'preview-report-2')]);
  const free = allocateOutputDir(cwd, 'preview-report', (candidate) => taken.has(candidate));
  assert.equal(free, path.join(cwd, 'preview-report-3'));

  const first = allocateOutputDir(cwd, 'preview-report', () => false);
  assert.equal(first, path.join(cwd, 'preview-report'));
});

test('allocateOutputDir sanitizes names and always returns something usable', () => {
  const cwd = path.resolve('C:/work/project');
  const sanitized = allocateOutputDir(cwd, 'preview:a*b?c', () => false);
  assert.equal(path.dirname(sanitized), cwd);
  assert.equal(path.basename(sanitized).includes(':'), false);
  const fallback = allocateOutputDir(cwd, '...', () => false);
  assert.equal(path.basename(fallback), 'preview');
});

test('allocateOutputDir gives up instead of looping forever', () => {
  const cwd = path.resolve('C:/work/project');
  assert.throws(() => allocateOutputDir(cwd, 'preview', () => true), /Cannot allocate an output directory/);
});

test('pickEngine defaults to LibreOffice and guards the Word engine', () => {
  assert.deepEqual(
    pickEngine({ requested: 'libreoffice', extension: 'docx', wordAvailable: true, kitAvailable: true }),
    { engine: 'libreoffice' },
  );
  assert.deepEqual(
    pickEngine({ requested: 'word', extension: 'docx', wordAvailable: true, kitAvailable: true }),
    { engine: 'word' },
  );
  assert.throws(
    () => pickEngine({ requested: 'word', extension: 'docx', wordAvailable: false, kitAvailable: true }),
    /Word is not installed/,
  );
  assert.throws(
    () => pickEngine({ requested: 'word', extension: 'xlsx', wordAvailable: true, kitAvailable: true }),
    /cannot open "\.xlsx"/,
  );
  assert.throws(
    () => pickEngine({ requested: 'libreoffice', extension: 'docx', wordAvailable: false, kitAvailable: false }),
    /engine was not found/,
  );
  assert.throws(
    () => pickEngine({ requested: 'pandoc', extension: 'docx', wordAvailable: true, kitAvailable: true }),
    /expected "libreoffice" or "word"/,
  );
});

test('shapeResult mirrors the engine manifest', () => {
  const manifest = {
    schemaVersion: 1,
    backend: 'native',
    rasterEngine: 'pdfium',
    dpi: 144,
    pageCount: 2,
    images: [
      { page: 1, path: 'C:/work/project/preview-x/page-0001.png', width: 1191, height: 1684, byteLength: 10 },
      { path: 'C:/work/project/preview-x/page-0002.png', width: 1191, height: 1684 },
    ],
    missingFonts: ['Calibri'],
  };
  const result = shapeResult({
    manifest,
    engine: 'word',
    outputDir: 'C:/work/project/preview-x',
    manifestPath: 'C:/work/project/preview-x/manifest.json',
  });

  assert.equal(result.status, 'ready');
  assert.equal(result.engine, 'word');
  assert.equal(result.rasterEngine, 'pdfium');
  assert.equal(result.pageCount, 2);
  assert.deepEqual(result.images[0], {
    page: 1,
    path: 'C:/work/project/preview-x/page-0001.png',
    width: 1191,
    height: 1684,
  });
  assert.equal(result.images[1].page, 2, 'a missing page number falls back to the image position');
  assert.deepEqual(result.missingFonts, ['Calibri']);
  assert.equal('pdfPath' in result, false);
});

test('shapeResult tolerates an empty manifest and carries a kept PDF', () => {
  const result = shapeResult({
    manifest: undefined,
    engine: 'libreoffice',
    outputDir: 'C:/work/project/preview-y',
    manifestPath: 'C:/work/project/preview-y/manifest.json',
    pdfPath: 'C:/work/project/preview-y/report.pdf',
  });
  assert.equal(result.pageCount, 0);
  assert.deepEqual(result.images, []);
  assert.deepEqual(result.missingFonts, []);
  assert.equal(result.pdfPath, 'C:/work/project/preview-y/report.pdf');
});
