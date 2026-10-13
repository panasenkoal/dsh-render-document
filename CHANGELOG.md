# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-10-13

### Added

- `render_document` tool for the DeepSeek Harness, registered from a host plugin.
- LibreOffice engine: renders DOCX/XLSX/PPTX/ODT/ODS/ODP and PDF pages to PNG through
  the bundled `@deepseek-ai/libreoffice-kit` Node API.
- Word engine: exports DOCX to PDF through Microsoft Word COM automation, then
  rasterizes the PDF, for print-accurate pagination.
- Workspace guard: `source` and `outputDir` must stay inside the session working
  directory; output directories are never overwritten.
- `status: "skipped"` result when no engine is available in this deployment.
- Unit tests for page selection, output-directory allocation and engine selection.

[Unreleased]: https://github.com/panasenkoal/dsh-render-document/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/panasenkoal/dsh-render-document/releases/tag/v0.1.0
