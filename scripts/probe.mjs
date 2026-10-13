#!/usr/bin/env node
// Diagnostic: report where the plugin finds its engine and whether Word is
// reachable. Spawns no child processes except the Word probe, so it is safe to
// run anywhere — but a full render only works outside the DSH sandbox.
//
//   node scripts/probe.mjs            # locate the kit
//   node scripts/probe.mjs --word     # also probe Microsoft Word

import { existsSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

import { findKit } from '../lib/kit.mjs';

const args = process.argv.slice(2);
const kit = findKit();

console.log('dsh-render-document probe');
console.log(`  node            ${process.version}`);
console.log(`  platform        ${process.platform} ${process.arch}`);
console.log(`  execPath        ${process.execPath}`);
console.log(`  DSH_HOME        ${process.env.DSH_HOME ?? '(unset)'}`);
console.log(`  DSH_OFFICE_KIT  ${process.env.DSH_OFFICE_KIT ?? '(unset)'}`);

if (!kit) {
  console.log('  kit             NOT FOUND');
  console.log('  hint: set DSH_OFFICE_KIT to the @deepseek-ai root that contains libreoffice-kit');
} else {
  console.log(`  kit root        ${kit.kitchenRoot}`);
  console.log(`  kit api         ${kit.apiPath}`);
  const engineRoot = path.join(kit.kitchenRoot, 'libreoffice-kit-win32-x64');
  console.log(`  native engine   ${existsSync(engineRoot) ? engineRoot : 'NOT FOUND (expected on Windows)'}`);
  try {
    const kitModule = await import(pathToFileURL(kit.apiPath).href);
    const runtime = await kitModule.discoverRuntime?.();
    console.log(`  kit runtime     ${JSON.stringify(runtime ?? kitModule.ENGINE_VERSION ?? 'unknown')}`);
  } catch (error) {
    console.log(`  kit runtime     failed to load: ${error.message}`);
  }
}

if (args.includes('--word')) {
  const { probeWord } = await import('../lib/word.mjs');
  const probe = await probeWord({ refresh: true });
  console.log(`  word            ${probe.available ? probe.word : `unavailable (${probe.error})`}`);
}
