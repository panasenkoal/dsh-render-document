// Microsoft Word engine: probe for Word and export a document to PDF through
// the PowerShell helper next to this file.
//
// Word needs its own profile and TEMP, which the DSH sandbox denies to confined
// children. This module is only ever called from the host process, so the helper
// starts with the user's normal token.

import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HELPER = fileURLToPath(new URL('./word-export.ps1', import.meta.url));
const PROBE_TTL_MS = 5 * 60 * 1000;

/** Windows PowerShell 5.1, the interpreter whose .NET runtime drives Office COM reliably. */
function resolvePowerShell({ env = process.env } = {}) {
  const systemRoot = env.SystemRoot || env.windir || 'C:\\Windows';
  const bundled = path.join(systemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
  return existsSync(bundled) ? bundled : 'powershell.exe';
}

function runHelper({ args, timeoutMs, signal, powershell = resolvePowerShell() }) {
  return new Promise((resolve, reject) => {
    const child = spawn(powershell, ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', HELPER, ...args], {
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    });

    let stdout = '';
    let stderr = '';
    let settled = false;

    const finish = (error, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener?.('abort', onAbort);
      if (error) reject(error);
      else resolve(value);
    };

    const onAbort = () => {
      child.kill();
      finish(new Error('Word export cancelled'));
    };

    const timer = timeoutMs > 0
      ? setTimeout(() => {
        child.kill();
        finish(new Error(`Word did not finish within ${Math.round(timeoutMs / 1000)} s`));
      }, timeoutMs)
      : undefined;

    signal?.addEventListener?.('abort', onAbort, { once: true });

    child.stdout.on('data', (chunk) => { stdout += chunk.toString('utf8'); });
    child.stderr.on('data', (chunk) => { stderr += chunk.toString('utf8'); });
    child.on('error', (error) => finish(error));
    child.on('close', (code) => {
      const lines = stdout.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
      const last = lines.at(-1);
      let parsed;
      if (last && (last.startsWith('{') || last.startsWith('['))) {
        try { parsed = JSON.parse(last); } catch { parsed = undefined; }
      }
      if (!parsed) {
        finish(new Error(`Word helper exited with code ${code}${stderr ? `: ${stderr.trim()}` : ''}`));
        return;
      }
      finish(undefined, { code, parsed, stderr });
    });
  });
}

/**
 * Is Microsoft Word installed? Cached for five minutes per process.
 * @param {{ timeoutMs?: number, signal?: AbortSignal, env?: NodeJS.ProcessEnv, refresh?: boolean }} [options]
 * @returns {Promise<{ available: boolean, word?: string, error?: string }>}
 */
export async function probeWord({ timeoutMs = 60_000, signal, env = process.env, refresh = false } = {}) {
  if (!refresh && probeCache && Date.now() - probeCache.at < PROBE_TTL_MS) return probeCache.value;
  let value;
  try {
    const { parsed } = await runHelper({ args: ['-Check'], timeoutMs, signal, powershell: resolvePowerShell({ env }) });
    value = parsed.ok
      ? { available: true, word: String(parsed.word ?? '') }
      : { available: false, error: String(parsed.error ?? 'Microsoft Word is not installed') };
  } catch (error) {
    value = { available: false, error: String(error?.message ?? error) };
  }
  probeCache = { at: Date.now(), value };
  return value;
}

/** Test seam: forget the cached probe result. */
export function resetWordProbe() {
  probeCache = undefined;
}

let probeCache;

/**
 * Export a document to PDF with Word.
 * @param {{ sourcePath: string, targetPath: string, timeoutMs?: number, signal?: AbortSignal, env?: NodeJS.ProcessEnv }} request
 * @returns {Promise<{ pages: number, pdfPath: string, version?: string }>}
 */
export async function exportPdf({ sourcePath, targetPath, timeoutMs = 300_000, signal, env = process.env }) {
  const { parsed } = await runHelper({
    args: ['-Source', sourcePath, '-Target', targetPath],
    timeoutMs,
    signal,
    powershell: resolvePowerShell({ env }),
  });
  if (!parsed.ok) throw new Error(`Microsoft Word could not export the document: ${parsed.error}`);
  return {
    pages: Number(parsed.pages ?? 0),
    pdfPath: String(parsed.pdf ?? targetPath),
    version: parsed.version ? String(parsed.version) : undefined,
  };
}

/** The helper path, exposed for diagnostics and tests. */
export const wordHelperPath = HELPER;
