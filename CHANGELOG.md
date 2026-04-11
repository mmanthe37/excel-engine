# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] — 2025-07-15

### Added

- **Formula Recalculation** — LibreOffice headless recalculation with graceful fallback (`excel_engine/recalc.py`)
- **Formula Error Scanning** — Detects #REF!, #DIV/0!, #VALUE!, #NAME?, #NULL!, #NUM!, #N/A across all sheets
- **Financial Model Presets** — IB-standard color coding and number formatting (`excel_engine/presets/financial.py`)
- **Go MCP Server** — High-performance compiled MCP bridge using `go-sdk v0.2.0` (`go-mcp-server/`)
- `recalculate_formulas` and `recalc_timeout` config options in EngineConfig
- `formula_errors` field in EngineResult with summary output
- `verify_formula_errors()` and `count_formulas()` in WorkbookVerifier
- GUI displays formula error summary after execution
- 46 new tests for recalculation, error scanning, and financial presets (604+ total)

### Changed

- Version bumped to 1.1.0
- Engine pipeline now has 5 phases: SCAN → GROUP → EXECUTE → VERIFY → RECALCULATE
- Distribution formats expanded from 6 to 7 (added Go MCP Server)
- README updated with Go MCP Server, recalc, and financial presets documentation
- Project structure updated to reflect new modules

## [1.0.0] — 2025-07-14

### Added

- Deep SAM instruction parser with step splitting and context extraction
- Error recovery system with classification, configurable retry, and exponential backoff
- Chart automation for scatter, area, combo, and sparklines
- Strengthened TaskExtractor patterns and WorkbookVerifier checks
- End-to-end SAM assignment simulation tests
- Integration tests with sample workbook fixtures

### Changed

- Deprecated Copilot Extension in favor of MCP server
- Fixed PyPI metadata and badge for publishing
- Use pip3 instead of pip for macOS compatibility

### Fixed

- README documentation now matches actual API, CLI, and MCP tool interfaces
- Python version badge corrected to 3.10+ (matching pyproject.toml)
- Project structure diagram regenerated from actual filesystem

## [0.1.0] — 2025-07-14

### Added

- 6-layer automation architecture (openpyxl → xlwings → AppleScript → System Events → VBA → PyAutoGUI)
- Section-based instruction planner with checkpoint decomposition
- CLI application (`excel-engine run`, `plan`, `verify`)
- MCP Server integration for use with AI assistants
- Copilot CLI Extension and Plugin distribution support
- Checkpoint verification system
- Initial documentation (architecture, layers, troubleshooting)
- GitHub Actions CI/CD workflows (test matrix, PyPI publish)

[1.1.0]: https://github.com/mmanthe37/excel-engine/releases/tag/v1.1.0
[1.0.0]: https://github.com/mmanthe37/excel-engine/releases/tag/v1.0.0
[0.1.0]: https://github.com/mmanthe37/excel-engine/releases/tag/v0.1.0
