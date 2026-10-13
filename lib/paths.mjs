// Pure helpers: page selection, workspace containment, output-directory
// allocation and engine choice. Kept dependency-free so `node --test` exercises
// them without the LibreOffice engine or a DSH installation.

import { existsSync } from 'node:fs';
import path from 'node:path';

/** Extensions the bundled LibreOffice engine can rasterize (see `capabilities --json`). */
export const RENDERABLE_EXTENSIONS = new Set([
  'doc', 'docx', 'docm', 'odt',
  'xls', 'xlsx', 'xlsm', 'ods',
  'ppt', 'pptx', 'pptm', 'odp',
  'pdf',
]);

/** Extensions Microsoft Word opens for the print-accurate engine. */
export const WORD_EXTENSIONS = new Set(['doc', 'docx', 'docm', 'rtf']);

export const MIN_DPI = 24;
export const MAX_DPI = 600;
export const DEFAULT_DPI = 144;

/**
 * Parse a page selection into `"all"` or a sorted array of one-based numbers.
 * Accepts "all", "", undefined, "3", "1,3-5", "2, 4".
 * @param {unknown} value raw `pages` argument.
 * @returns {"all" | number[]}
 */
export function parsePages(value) {
  if (value === undefined || value === null) return 'all';
  const text = String(value).trim().toLowerCase();
  if (text === '' || text === 'all') return 'all';

  const selected = new Set();
  for (const rawChunk of text.split(',')) {
    const chunk = rawChunk.trim();
    if (chunk === '') continue;

    const range = /^(\d+)\s*-\s*(\d+)$/.exec(chunk);
    if (range) {
      const from = Number(range[1]);
      const to = Number(range[2]);
      if (from < 1) throw new Error(`pages: "${chunk}" starts below page 1`);
      if (to < from) throw new Error(`pages: "${chunk}" ends before it starts`);
      if (to - from > 999) throw new Error(`pages: "${chunk}" spans more than 1000 pages`);
      for (let page = from; page <= to; page++) selected.add(page);
      continue;
    }

    if (/^\d+$/.test(chunk)) {
      const page = Number(chunk);
      if (page < 1) throw new Error(`pages: "${chunk}" is below page 1`);
      selected.add(page);
      continue;
    }

    throw new Error(`pages: cannot parse "${chunk}"; use one-based numbers like "1,3-5" or "all"`);
  }

  if (selected.size === 0) return 'all';
  return [...selected].sort((a, b) => a - b);
}

/**
 * Validate a dpi value.
 * @param {unknown} value raw `dpi` argument.
 * @returns {number}
 */
export function parseDpi(value) {
  if (value === undefined || value === null) return DEFAULT_DPI;
  const dpi = Number(value);
  if (!Number.isInteger(dpi) || dpi < MIN_DPI || dpi > MAX_DPI) {
    throw new Error(`dpi: expected an integer between ${MIN_DPI} and ${MAX_DPI}, got ${JSON.stringify(value)}`);
  }
  return dpi;
}

/**
 * Is `target` the same path as `root` or inside it?
 * Comparison is case-insensitive on Windows.
 * @param {string} root
 * @param {string} target
 */
export function isInside(root, target) {
  const from = path.resolve(root);
  const to = path.resolve(target);
  const relative = path.relative(from, to);
  if (relative === '') return true;
  if (relative.startsWith('..') || path.isAbsolute(relative)) return false;
  return true;
}

/**
 * Resolve a caller-supplied path against the session workspace and refuse
 * anything outside it. The tool runs in the DSH host process with the user's
 * full rights, so this guard is the only thing keeping it inside the workspace.
 * @param {string} cwd session working directory (workspace root)
 * @param {string} candidate relative or absolute path
 * @param {string} label argument name used in the error message
 */
export function resolveInWorkspace(cwd, candidate, label = 'path') {
  if (typeof candidate !== 'string' || candidate.trim() === '') {
    throw new Error(`${label}: expected a non-empty path`);
  }
  const absolute = path.resolve(cwd, candidate);
  if (!isInside(cwd, absolute)) {
    throw new Error(`${label}: "${candidate}" is outside the session workspace (${cwd})`);
  }
  return absolute;
}

/**
 * Pick a fresh output directory inside the workspace: `<base>`, then
 * `<base>-2`, `<base>-3`, ... The engine refuses to write into an existing
 * directory, so the caller must always receive a free name.
 * @param {string} cwd workspace root
 * @param {string} baseName directory name stem, e.g. "preview-report"
 * @param {(candidate: string) => boolean} isTaken existence probe (injectable for tests)
 */
export function allocateOutputDir(cwd, baseName, isTaken = existsSync) {
  const safeBase = baseName.replace(/[\\/:*?"<>|]+/g, '-').replace(/^[.\s]+|[.\s]+$/g, '') || 'preview';
  for (let attempt = 1; attempt <= 50; attempt++) {
    const name = attempt === 1 ? safeBase : `${safeBase}-${attempt}`;
    const candidate = path.join(cwd, name);
    if (!isTaken(candidate)) return candidate;
  }
  throw new Error(`Cannot allocate an output directory for "${baseName}" inside ${cwd}`);
}

/**
 * Decide which engine serves a request, and whether visual checking is possible
 * at all. `libreoffice` is the default because the bundled engine is always
 * present; `word` is requested explicitly when print-accurate pagination
 * matters.
 * @param {{ requested?: string, extension: string, wordAvailable: boolean, kitAvailable: boolean }} input
 */
export function pickEngine({ requested = 'libreoffice', extension, wordAvailable, kitAvailable }) {
  if (requested === 'libreoffice') {
    if (!kitAvailable) throw new Error('engine "libreoffice": the bundled LibreOffice engine was not found');
    return { engine: 'libreoffice' };
  }
  if (requested === 'word') {
    if (!WORD_EXTENSIONS.has(extension)) {
      throw new Error(`engine "word": Microsoft Word cannot open ".${extension}"; use engine "libreoffice"`);
    }
    if (!wordAvailable) throw new Error('engine "word": Microsoft Word is not installed on this machine');
    if (!kitAvailable) throw new Error('engine "word": the bundled LibreOffice engine is needed to rasterize the exported PDF');
    return { engine: 'word' };
  }
  throw new Error(`engine: expected "libreoffice" or "word", got ${JSON.stringify(requested)}`);
}

/**
 * Shape the tool result from the engine manifest.
 * @param {{ manifest: any, engine: string, outputDir: string, manifestPath: string, pdfPath?: string }} input
 */
export function shapeResult({ manifest, engine, outputDir, manifestPath, pdfPath }) {
  const images = Array.isArray(manifest?.images) ? manifest.images : [];
  return {
    status: 'ready',
    engine,
    rasterEngine: manifest?.rasterEngine,
    pageCount: Number(manifest?.pageCount ?? images.length),
    dpi: manifest?.dpi,
    outputDir,
    manifestPath,
    images: images.map((image, index) => ({
      page: Number(image?.page ?? index + 1),
      path: String(image?.path ?? ''),
      width: Number(image?.width ?? 0),
      height: Number(image?.height ?? 0),
    })),
    missingFonts: Array.isArray(manifest?.missingFonts) ? manifest.missingFonts.map(String) : [],
    ...(pdfPath ? { pdfPath } : {}),
  };
}
