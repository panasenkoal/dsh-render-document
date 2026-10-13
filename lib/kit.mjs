// Locate the bundled LibreOffice Kit and render pages through its Node API.
//
// The plugin runs inside the DSH host process, which is not confined by the
// sandbox, so the kit may start its native engine with piped stdio. That is
// exactly what fails for commands DSH spawns under `workspace-write`
// (`spawn EPERM`), and why this tool exists.

import { existsSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const API_RELATIVE = path.join('libreoffice-kit', 'lib', 'index.js');
const MAX_SEARCH_DEPTH = 6;

function looksLikeKit(root) {
  return typeof root === 'string' && root !== '' && existsSync(path.join(root, API_RELATIVE));
}

function searchForKit(root, depth = MAX_SEARCH_DEPTH, seen = new Set()) {
  if (depth < 0 || typeof root !== 'string' || root === '') return undefined;
  let resolved;
  try {
    resolved = path.resolve(root);
  } catch {
    return undefined;
  }
  if (seen.has(resolved) || !existsSync(resolved)) return undefined;
  seen.add(resolved);

  if (looksLikeKit(resolved)) return resolved;
  const nested = path.join(resolved, 'node_modules', '@deepseek-ai');
  if (looksLikeKit(nested)) return nested;

  let entries;
  try {
    entries = readdirSync(resolved, { withFileTypes: true });
  } catch {
    return undefined;
  }
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (entry.name === 'node_modules' || entry.name === 'dsh-runtimes' || entry.name === 'dependencies' || entry.name.startsWith('.')) {
      const found = searchForKit(path.join(resolved, entry.name), depth - 1, seen);
      if (found) return found;
      continue;
    }
    if (/^(@deepseek-ai|libreoffice-kit|runtime)/.test(entry.name)) {
      const found = searchForKit(path.join(resolved, entry.name), depth - 1, seen);
      if (found) return found;
    }
  }
  return undefined;
}

/**
 * Find the `@deepseek-ai` root that contains `libreoffice-kit`.
 * Order: explicit override, the packaged application, extra runtimes under
 * `$DSH_HOME`, then a bounded search from the executable's directory.
 * @param {{ env?: NodeJS.ProcessEnv, execPath?: string, resourcesPath?: string }} [options]
 * @returns {{ apiPath: string, kitchenRoot: string } | undefined}
 */
export function findKit({ env = process.env, execPath = process.execPath, resourcesPath = process.resourcesPath } = {}) {
  const candidates = [];

  // 1. Explicit override. Two conventions are accepted: the `@deepseek-ai`
  //    directory itself, or the native engine package inside it (which is what
  //    the `DSH_OFFICE_KIT` scripts use), in which case the sibling API package
  //    is what we need.
  if (env.DSH_OFFICE_KIT) {
    candidates.push(env.DSH_OFFICE_KIT);
    candidates.push(path.dirname(env.DSH_OFFICE_KIT));
    candidates.push(path.join(env.DSH_OFFICE_KIT, 'node_modules', '@deepseek-ai'));
  }

  // 2. Packaged desktop application: <app>/resources/app.asar.unpacked/dsh/node_modules/@deepseek-ai
  if (typeof resourcesPath === 'string' && resourcesPath !== '') {
    candidates.push(path.join(resourcesPath, 'app.asar.unpacked', 'dsh', 'node_modules', '@deepseek-ai'));
  }
  candidates.push(path.join(path.dirname(execPath), 'resources', 'app.asar.unpacked', 'dsh', 'node_modules', '@deepseek-ai'));

  // 2b. Conventional per-user and machine-wide installations, so a standalone
  //     run (smoke script, CLI tooling) finds the engine without the override.
  for (const root of [
    env.LOCALAPPDATA && path.join(env.LOCALAPPDATA, 'Programs'),
    env.ProgramFiles,
    env['ProgramFiles(x86)'],
    env.ProgramW6432,
  ]) {
    if (root) candidates.push(path.join(root, 'DeepSeek Harness', 'resources', 'app.asar.unpacked', 'dsh', 'node_modules', '@deepseek-ai'));
  }

  // 3. Extra runtimes under DSH_HOME.
  const home = env.DSH_HOME || (env.USERPROFILE ? path.join(env.USERPROFILE, '.dsh') : '');
  if (home) candidates.push(path.join(home, 'dsh-runtimes'));

  for (const candidate of candidates) {
    if (looksLikeKit(candidate)) return { apiPath: path.join(candidate, API_RELATIVE), kitchenRoot: candidate };
  }
  for (const candidate of candidates) {
    const found = searchForKit(candidate);
    if (found) return { apiPath: path.join(found, API_RELATIVE), kitchenRoot: found };
  }

  // 4. Last resort: bounded search next to the running executable.
  const found = searchForKit(path.dirname(execPath));
  if (found) return { apiPath: path.join(found, API_RELATIVE), kitchenRoot: found };
  return undefined;
}

/** Load the kit module for an `apiPath` returned by `findKit`. */
export async function loadKit(apiPath) {
  const module = await import(pathToFileURL(apiPath).href);
  if (typeof module.createConverter !== 'function') {
    throw new Error(`The LibreOffice Kit at ${apiPath} does not export createConverter()`);
  }
  return module;
}

/**
 * Render document pages to PNG files in a fresh directory.
 * @param {{ apiPath: string, inputPath: string, outputDir: string, pages: "all"|number[], dpi: number, timeoutMs?: number, signal?: AbortSignal }} request
 * @returns {Promise<any>} the engine manifest (also written to `<outputDir>/manifest.json`)
 */
export async function renderImages({ apiPath, inputPath, outputDir, pages, dpi, timeoutMs = 300_000, signal }) {
  const kit = await loadKit(apiPath);
  const converter = await kit.createConverter({ timeoutMs });
  try {
    /** @type {Record<string, unknown>} */
    const request = { inputPath, outputDir, dpi };
    if (pages !== 'all') request.pages = pages;
    return await converter.renderImages(request, signal);
  } finally {
    await converter.dispose().catch(() => {});
  }
}

/** Human-readable message for a kit failure. */
export function describeKitError(error) {
  const code = error?.code;
  const message = error?.message ?? String(error);
  switch (code) {
    case 'unsupported-format':
      return `the engine does not support this format (${message})`;
    case 'invalid-document':
      return `the document could not be read (${message})`;
    case 'input-too-large':
      return `the document exceeds the engine input limit (${message})`;
    case 'output-too-large':
      return `the result exceeds the engine output limits; render fewer pages or lower dpi (${message})`;
    case 'invalid-output':
      return `the engine refused the output location (${message})`;
    case 'timeout':
      return `rendering timed out (${message})`;
    case 'unavailable':
      return `the LibreOffice engine is unavailable in this installation (${message})`;
    default:
      return message;
  }
}

/** Diagnostic used by the test script: does a path look like a usable kit root? */
export function probeKit(kitchenRoot) {
  if (!looksLikeKit(kitchenRoot)) return { ok: false, reason: 'libreoffice-kit/lib/index.js not found' };
  try {
    const engineRoot = path.join(kitchenRoot, 'libreoffice-kit-win32-x64');
    return { ok: true, kitchenRoot, enginePresent: existsSync(engineRoot), engineRoot: statSync(kitchenRoot).isDirectory() ? engineRoot : undefined };
  } catch (error) {
    return { ok: false, reason: String(error?.message ?? error) };
  }
}
