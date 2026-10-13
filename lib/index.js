// dsh-render-document — a DeepSeek Harness plugin that registers the
// `render_document` tool.
//
// Why a host tool: DSH confines every command it spawns (restricted token, Low
// integrity, write access only inside the session workspace and a private temp
// directory). The bundled LibreOffice Kit CLI starts its engine with piped
// stdio, which the Windows sandbox denies — `spawn EPERM` — so each layout check
// escalates the sandbox and asks the user for approval. Microsoft Word cannot
// run confined at all: it fails to create its working files. Tool code runs in
// the DSH host process, outside that confinement, which is what makes a single
// unapproved call possible.
//
// Security model (see README): `source` and `outputDir` must resolve inside the
// session working directory, output directories are never overwritten, paths are
// passed to child processes as an argument array without a shell, no network
// access, and every child process is killed when the call is cancelled.

import { copyFile, mkdtemp, rm } from 'node:fs/promises';
import { existsSync as existsSyncSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

import { renderImages, findKit, describeKitError } from './kit.mjs';
import { exportPdf, probeWord } from './word.mjs';
import {
  RENDERABLE_EXTENSIONS,
  WORD_EXTENSIONS,
  allocateOutputDir,
  parseDpi,
  parsePages,
  pickEngine,
  resolveInWorkspace,
  shapeResult,
} from './paths.mjs';

/**
 * The registry accepts a plain definition object, so a missing typing helper
 * (an older or newer dsh-tools) degrades to identity instead of failing to load.
 */
let defineTool = (definition) => definition;
try {
  const tools = await import('@deepseek-ai/dsh-tools');
  if (typeof tools.defineTool === 'function') defineTool = tools.defineTool;
} catch {
  // Typing only; the registration below still carries the full definition.
}

/** Stable Loader identity. */
export const name = 'tool-render-document';

/** Services used by the tool. */
export const inject = ['tools'];

const TOOL_DESCRIPTION = [
  'Render document pages to PNG images for a visual layout check, without sandbox escalation.',
  'Supports .docx/.doc/.odt, .xlsx/.xls/.ods, .pptx/.ppt/.odp and .pdf.',
  'engine "libreoffice" (default) uses the bundled engine and always works.',
  'engine "word" exports the document through Microsoft Word first, so page breaks and field positions match what Word prints — use it for print forms and pagination questions.',
  'Pages are one-based; pass "1,3-5" or omit for all pages. The output directory must not exist yet and is created inside the workspace.',
  'After a "ready" result, inspect the returned images (read_image) before reporting on layout. A "skipped" result means this deployment cannot render and visual checking must end.',
].join(' ');

/** Register the render_document tool. */
export function apply(ctx) {
  ctx.tools.register(defineTool({
    name: 'render_document',
    description: TOOL_DESCRIPTION,
    parameters: {
      source: {
        type: 'string',
        required: true,
        description: 'Document to render: a path inside the session workspace, absolute or relative to it.',
      },
      pages: {
        type: 'string',
        description: 'One-based pages to render, for example "1", "1,3-5". Omit or pass "all" for every page.',
      },
      dpi: {
        type: 'integer',
        description: 'Raster resolution between 24 and 600. Default 144.',
      },
      engine: {
        type: 'string',
        description: 'Rendering engine: "libreoffice" (default) or "word" for Word-accurate pagination of .docx/.doc files.',
      },
      outputDir: {
        type: 'string',
        description: 'New directory for the PNG pages, relative to the workspace or absolute inside it. Omit to auto-name "preview-<document>".',
      },
      keepPdf: {
        type: 'boolean',
        description: 'With engine "word", keep the exported PDF next to the images for printing or delivery.',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          status: {
            type: 'string',
            required: true,
            description: '"ready" when images were produced, "skipped" when this deployment has no renderer.',
          },
          engine: {
            type: 'string',
            required: true,
            description: 'Engine that produced the pages: "libreoffice" or "word".',
          },
          note: {
            type: 'string',
            description: 'Why rendering was skipped.',
          },
          rasterEngine: {
            type: 'string',
            description: 'Rasterizer inside the engine: "libreoffice" for Office documents, "pdfium" for PDF input.',
          },
          pageCount: {
            type: 'integer',
            description: 'Page or slide count reported by the engine.',
          },
          dpi: {
            type: 'integer',
            description: 'Resolution actually used.',
          },
          outputDir: {
            type: 'string',
            description: 'Directory holding the PNG pages and manifest.json.',
          },
          manifestPath: {
            type: 'string',
            description: 'The engine manifest with image paths, sizes and font diagnostics.',
          },
          pdfPath: {
            type: 'string',
            description: 'Kept PDF, when keepPdf was requested with engine "word".',
          },
          missingFonts: {
            type: 'array',
            description: 'Font families the document asks for but the machine lacks, so layout may shift.',
            items: { type: 'string' },
          },
          images: {
            type: 'array',
            description: 'Rendered PNG pages in page order.',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                page: { type: 'integer', required: true, description: 'One-based page number.' },
                path: { type: 'string', required: true, description: 'PNG file path; open it with read_image.' },
                width: { type: 'integer', description: 'Pixel width.' },
                height: { type: 'integer', description: 'Pixel height.' },
              },
            },
          },
        },
      },
      render: (_args, value) => {
        if (value.status === 'skipped') {
          return [{ type: 'text', text: `render_document skipped: ${value.note ?? 'no renderer is available in this deployment'}` }];
        }
        const lines = [
          `Rendered ${value.pageCount} page(s) with the ${value.engine} engine into ${value.outputDir}`,
        ];
        for (const image of value.images ?? []) {
          lines.push(`- page ${image.page}: ${image.path} (${image.width}x${image.height})`);
        }
        if (value.pdfPath) lines.push(`PDF: ${value.pdfPath}`);
        if (value.missingFonts?.length) {
          lines.push(`Missing fonts (layout may differ from the authoring machine): ${value.missingFonts.join(', ')}`);
        }
        lines.push('Inspect these pages with read_image before reporting on layout.');
        return [{ type: 'text', text: lines.join('\n') }];
      },
    },
    async execute(args, exec) {
      const cwd = exec?.agent?.session?.header?.cwd;
      if (!cwd) throw new Error('render_document requires an agent Session with a working directory');

      const sourcePath = resolveInWorkspace(cwd, args.source, 'source');
      if (!existsSyncSync(sourcePath)) throw new Error(`source: file not found: ${args.source}`);
      const extension = path.extname(sourcePath).slice(1).toLowerCase();
      if (!RENDERABLE_EXTENSIONS.has(extension)) {
        throw new Error(`source: ".${extension}" cannot be rendered; supported: ${[...RENDERABLE_EXTENSIONS].join(', ')}`);
      }

      const pages = parsePages(args.pages);
      const dpi = parseDpi(args.dpi);
      const requestedEngine = args.engine ?? 'libreoffice';
      const timeoutMs = 300_000;

      const outputDir = args.outputDir
        ? resolveInWorkspace(cwd, args.outputDir, 'outputDir')
        : allocateOutputDir(cwd, `preview-${path.basename(sourcePath, path.extname(sourcePath))}`);
      if (existsSyncSync(outputDir)) {
        throw new Error(`outputDir: "${path.relative(cwd, outputDir)}" already exists; the engine never overwrites an existing directory, so choose a new name`);
      }

      const kit = findKit();
      if (requestedEngine === 'libreoffice' && !kit) {
        return {
          status: 'skipped',
          engine: 'libreoffice',
          note: 'the bundled LibreOffice engine was not found in this installation, so visual layout checking is unavailable',
        };
      }

      // Probe Word only when it could actually serve this request: the probe
      // starts a PowerShell process.
      const wordProbe = requestedEngine === 'word' && WORD_EXTENSIONS.has(extension)
        ? await probeWord({ signal: exec.signal })
        : { available: false };
      const { engine } = pickEngine({
        requested: requestedEngine,
        extension,
        wordAvailable: wordProbe.available,
        kitAvailable: Boolean(kit),
      });

      let scratch;
      try {
        let inputPath = sourcePath;
        let wordPdfPath;
        if (engine === 'word') {
          scratch = await mkdtemp(path.join(tmpdir(), 'dsh-render-document-'));
          wordPdfPath = path.join(scratch, `${path.basename(sourcePath, path.extname(sourcePath))}.pdf`);
          await exportPdf({ sourcePath, targetPath: wordPdfPath, timeoutMs, signal: exec.signal });
          inputPath = wordPdfPath;
        }

        const manifest = await renderImages({
          apiPath: kit.apiPath,
          inputPath,
          outputDir,
          pages,
          dpi,
          timeoutMs,
          signal: exec.signal,
        });

        const result = shapeResult({
          manifest,
          engine,
          outputDir,
          manifestPath: path.join(outputDir, 'manifest.json'),
        });

        if (args.keepPdf && wordPdfPath) {
          const keptPdf = path.join(outputDir, `${path.basename(sourcePath, path.extname(sourcePath))}.pdf`);
          await copyFile(wordPdfPath, keptPdf);
          result.pdfPath = keptPdf;
        }
        return result;
      } catch (error) {
        if (engine === 'word') {
          throw new Error(`render_document failed with the Word engine: ${error?.message ?? error}`);
        }
        throw new Error(`render_document failed: ${describeKitError(error)}`);
      } finally {
        if (scratch) await rm(scratch, { recursive: true, force: true }).catch(() => {});
      }
    },
  }));
}
