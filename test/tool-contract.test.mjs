import assert from 'node:assert/strict';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { apply, inject, name } from '../lib/index.js';

// A workspace-like temporary directory stands in for the session workspace.
const workspace = await mkdtemp(path.join(tmpdir(), 'render-document-test-'));
await writeFile(path.join(workspace, 'report.docx'), 'placeholder');
await writeFile(path.join(workspace, 'sheet.xlsx'), 'placeholder');
await writeFile(path.join(workspace, 'notes.txt'), 'placeholder');
await mkdir(path.join(workspace, 'taken'));

const registered = [];
apply({ tools: { register: (definition) => registered.push(definition) } });
const tool = registered[0];
const exec = { agent: { session: { header: { cwd: workspace } } } };

test('the plugin exports a Cordis plugin that registers render_document', () => {
  assert.equal(name, 'tool-render-document');
  assert.deepEqual(inject, ['tools']);
  assert.equal(registered.length, 1);
  assert.equal(tool.name, 'render_document');
  assert.equal(typeof tool.execute, 'function');
  assert.equal(tool.parameters.source.required, true);
  assert.equal(tool.output.schema.properties.status.required, true);
  assert.match(tool.description, /word/i);
});

test('execute requires an agent session', async () => {
  await assert.rejects(() => tool.execute({ source: 'report.docx' }, {}), /requires an agent Session/);
});

test('execute refuses paths outside the workspace', async () => {
  await assert.rejects(() => tool.execute({ source: '../escape.docx' }, exec), /outside the session workspace/);
  await assert.rejects(() => tool.execute({ source: 'C:/Windows/win.ini' }, exec), /outside the session workspace/);
});

test('execute refuses missing files and unsupported types', async () => {
  await assert.rejects(() => tool.execute({ source: 'missing.docx' }, exec), /file not found/);
  await assert.rejects(() => tool.execute({ source: 'notes.txt' }, exec), /cannot be rendered/);
});

test('execute validates pages and dpi before touching any engine', async () => {
  await assert.rejects(() => tool.execute({ source: 'report.docx', pages: '0' }, exec), /below page 1/);
  await assert.rejects(() => tool.execute({ source: 'report.docx', pages: 'nope' }, exec), /cannot parse/);
  await assert.rejects(() => tool.execute({ source: 'report.docx', dpi: 5 }, exec), /between 24 and 600/);
});

test('execute refuses an unknown engine and a Word request for spreadsheets', async () => {
  await assert.rejects(() => tool.execute({ source: 'report.docx', engine: 'pandoc' }, exec), /expected "libreoffice" or "word"/);
  // The Word check runs before any probe, so this needs no Microsoft Word.
  await assert.rejects(() => tool.execute({ source: 'sheet.xlsx', engine: 'word' }, exec), /cannot open "\.xlsx"/);
});

test('execute refuses to reuse an existing output directory', async () => {
  await assert.rejects(
    () => tool.execute({ source: 'report.docx', outputDir: 'taken' }, exec),
    /already exists/,
  );
  await assert.rejects(
    () => tool.execute({ source: 'report.docx', outputDir: '../outside' }, exec),
    /outside the session workspace/,
  );
});

test('output.render explains a skipped result', () => {
  const parts = tool.output.render({}, { status: 'skipped', engine: 'libreoffice', note: 'no engine here' });
  assert.equal(parts.length, 1);
  assert.equal(parts[0].type, 'text');
  assert.match(parts[0].text, /skipped: no engine here/);
});

test('output.render lists pages, fonts and the PDF', () => {
  const parts = tool.output.render({}, {
    status: 'ready',
    engine: 'word',
    pageCount: 1,
    outputDir: 'preview-report',
    pdfPath: 'preview-report/report.pdf',
    missingFonts: ['Calibri'],
    images: [{ page: 1, path: 'preview-report/page-0001.png', width: 1191, height: 1684 }],
  });
  const text = parts[0].text;
  assert.match(text, /Rendered 1 page\(s\) with the word engine/);
  assert.match(text, /page 1: preview-report\/page-0001\.png/);
  assert.match(text, /Missing fonts.*Calibri/);
  assert.match(text, /PDF: preview-report\/report\.pdf/);
  assert.match(text, /read_image/);
});
