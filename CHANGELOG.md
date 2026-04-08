# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[1.0.0]: https://github.com/mmanthe37/excel-engine/releases/tag/v1.0.0
[0.1.0]: https://github.com/mmanthe37/excel-engine/releases/tag/v0.1.0
