# dsh-render-document

A DeepSeek Harness plugin that adds the **`render_document`** tool: it rasterizes Office
documents and PDFs into PNG pages so the agent can check layout — **without sandbox
escalation and without approval prompts**.

Russian version: [README.md](README.md).

## Why it exists

DSH runs every command it spawns inside a sandbox: a restricted, low-integrity token that may
write only inside the session workspace and a private temporary directory. Two consequences
turned layout checking into a stream of approval prompts:

1. **The bundled LibreOffice Kit CLI always fails under the sandbox** with `spawn EPERM`: it
   starts its engine with piped stdio, which the sandbox denies. The agent then does the only
   thing the documentation allows — escalates to `danger-full-access` — and the user gets an
   approval prompt for every render. In one machine's session logs that was **47 of 72**
   approval requests.
2. **Microsoft Word cannot run under the sandbox at all**: it inherits the restricted token,
   cannot create its working files, and shows "Word could not create the work file. Check the
   TEMP environment variable." The automation call then hangs or fails with `0x800A1772`, and
   the `Document` object never reaches the caller.

Plugin code runs inside the DSH host process — **outside** that confinement. That is what lets
the tool start an engine with the user's normal token, once, with no approval.

## Install

```sh
dsh plugin --profile <profile> add github:<owner>/dsh-render-document
# then restart DSH
```

Installing host code requires full access once (a DSH plugin-manager constraint, not this
plugin's). After the restart the `render_document` tool is available to the agent in every
workspace.

## Usage

The agent calls the tool, but the contract is worth knowing:

| Argument | Type | Meaning |
|---|---|---|
| `source` | string, required | Document inside the session workspace: `.docx/.doc/.odt`, `.xlsx/.xls/.ods`, `.pptx/.ppt/.odp`, `.pdf` |
| `pages` | string | One-based pages: `"1"`, `"1,3-5"`. All pages by default |
| `dpi` | integer | 24…600, default 144 |
| `engine` | string | `"libreoffice"` (default) or `"word"` |
| `outputDir` | string | Fresh directory for the PNGs inside the workspace; defaults to `preview-<document>` |
| `keepPdf` | boolean | With `engine: "word"`, also keep the exported PDF next to the pages |

Result:

```json
{
  "status": "ready",
  "engine": "word",
  "rasterEngine": "pdfium",
  "pageCount": 2,
  "dpi": 144,
  "outputDir": "C:/work/preview-report",
  "manifestPath": "C:/work/preview-report/manifest.json",
  "images": [
    { "page": 1, "path": "C:/work/preview-report/page-0001.png", "width": 1191, "height": 1684 },
    { "page": 2, "path": "C:/work/preview-report/page-0002.png", "width": 1191, "height": 1684 }
  ],
  "missingFonts": []
}
```

`status: "skipped"` means this installation has no renderer at all: visual checking must end
instead of looking for workarounds. The bundled `office-docx` / `office-xlsx` / `office-pptx`
skills describe the tool exactly this way.

## Engines

| | `libreoffice` (default) | `word` |
|---|---|---|
| Renders | .docx/.doc/.odt, .xlsx/.xls/.ods, .pptx/.ppt/.odp, .pdf | .docx/.doc (+ .rtf); the PDF is then rasterized by the same engine |
| Pagination | as LibreOffice computes it | **as Microsoft Word computes it** |
| Availability | always (engine ships with DSH) | only when Microsoft Word is installed |
| Speed | faster | slower (Word start plus export) |

Rule of thumb: use `libreoffice` for ordinary layout checks, and `word` for print forms where
the page break position and field placement must match Word. On one real document (a
single-page registration form) the two renders differ by **3% of pixels** across the whole
page — a LibreOffice preview does not show what Word will print.

## Security model

The tool deliberately runs outside the DSH sandbox — otherwise it would be pointless. What
that means, and how it is bounded:

- **Paths.** `source` and `outputDir` resolve against the session workspace, and any path
  outside it is rejected. Absolute paths elsewhere are refused exactly like `../`.
- **No arbitrary code.** The tool accepts no commands, scripts, or engine flags: a document,
  pages, a dpi, and a choice of two engines.
- **Arguments are passed as an array**, never through a shell, so shell metacharacters cannot
  be injected.
- **Existing directories are never overwritten.** The engine creates a fresh directory in
  exclusive mode; a taken name is an error rather than a silent overwrite.
- **No network access.** Neither the engine nor the Word helper talks to the network.
- **Child processes are governed.** Word starts hidden, the document opens read-only, and a
  Word instance the plugin did not create (one the user had open) is never closed. Cancelling
  the call kills the child.
- **One child process runs with full user rights** — the thin Word wrapper
  (`lib/word-export.ps1`); its entire source lives in this repository.

## Requirements

- Node.js ≥ 22.19.0 (as bundled with DSH).
- The LibreOffice Kit engine, which ships with DSH; its location is discovered automatically
  (`DSH_OFFICE_KIT`, the unpacked application runtime, `$DSH_HOME/dsh-runtimes`).
- For `engine: "word"`: Microsoft Word plus the matching Microsoft Visual C++ v14
  Redistributable.

## Diagnostics

```sh
node scripts/probe.mjs --word   # where the engine was found, and whether Word is reachable
```

## Development

```sh
node test/paths.test.mjs          # page selection, path guard, engine choice
node test/tool-contract.test.mjs  # tool contract and result rendering
node scripts/smoke.mjs <document>            # end-to-end through LibreOffice
node scripts/smoke.mjs <document> --engine word --keep-pdf
```

`npm test` uses the standard runner (`node --test`). Inside a DSH session that runner cannot
work — it spawns test files with piped stdio, which the sandbox denies (`spawn EPERM`) — so run
the test files directly, as above. The smoke script must also run in a normal terminal: the
engine starts a native child with piped stdio.

## Limitations

- Whole pages only: worksheet regions need the engine's separate `sheet`/`range` options.
- The engine's PNG tile protocol is private and may change with a Kit update, which would
  require a plugin update.
- `libreoffice` pagination is not Word pagination; use `engine: "word"` for print forms.
- Word COM is a single-user resource: concurrent calls are serialized by the engine, not by
  this plugin.

## License

MIT — see [LICENSE](LICENSE).
