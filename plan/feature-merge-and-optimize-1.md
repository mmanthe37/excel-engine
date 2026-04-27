---
goal: Merge DEV unique features into CURRENT, undo harmful changes, add new user features, rebuild .app
version: 1.0
date_created: 2026-04-27
last_updated: 2026-04-27
owner: Michael Manthe
status: 'Planned'
tags: [feature, refactor, merge, architecture, optimization]
---

# Introduction

![Status: Planned](https://img.shields.io/badge/status-Planned-blue)

Merge the best of two diverged excel-engine repositories (CURRENT at `~/excel-engine` and DEV at `~/Dev/excel-engine`) into a single, optimized engine. Undo recent changes that removed valuable layer dispatch mappings, integrate DEV-only features (copilot-studio-connector, manage_slicer, validation methods, openpyxl sort), add three new user-facing features (additional resource file uploads, multi-format instruction uploads, optional .xlsm conversion), and rebuild the macOS .app bundle from the merged codebase.

## 1. Requirements & Constraints

- **REQ-001**: Final merged engine must pass all existing tests (901+ in CURRENT, 918+ in DEV) with zero failures
- **REQ-002**: All CURRENT corruption safeguards (OOXML validation, backup, restore, UNSAFE_CHARS_PATTERN) must be preserved
- **REQ-003**: All DEV unique features (copilot-studio-connector, manage_slicer, _validate_task_targets, _sort_openpyxl_range) must be integrated
- **REQ-004**: Users must be able to upload additional resource/data files (.xlsx, .pdf, .docx, .zip, .txt, images) alongside the main workbook
- **REQ-005**: Users must be able to upload instructions in .docx, .pdf, .txt, .zip, .rtf formats
- **REQ-006**: When VBA macros are required, engine must prompt/suggest .xlsm conversion rather than converting automatically
- **REQ-007**: Engine must never corrupt or destroy user's original input file (backup before any modification)
- **REQ-008**: Merged .app bundle must include all fixes and be rebuilt via PyInstaller
- **SEC-001**: No secrets, credentials, or user data in committed code
- **CON-001**: CURRENT repo (`~/excel-engine`) is the canonical repo — all merges target this directory
- **CON-002**: DEV repo (`~/Dev/excel-engine`) is read-only source — no commits to DEV
- **CON-003**: Python 3.13 is the target runtime (CURRENT venv); DEV's Python 3.14 is not used
- **CON-004**: .xlsx format must remain the default output — .xlsm only when explicitly opted-in by user
- **GUD-001**: Prefer openpyxl (Layer 1) for operations it can handle — it's fastest and cross-platform
- **GUD-002**: Each phase must be independently testable before proceeding to the next
- **PAT-001**: Layer cascade pattern: attempt fastest layer first, escalate on failure
- **PAT-002**: All file I/O must go through PathHandler for safe path handling

## 2. Implementation Steps

### Phase 1: Undo Harmful Changes & Restore Layer Dispatch Mappings

- GOAL-001: Restore the 9 TaskType layer dispatch entries that were incorrectly removed from CURRENT's `config.py` as "unreachable", plus adopt DEV's VBA-before-SYSTEM_EVENTS layer ordering

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-001 | Restore `OPENPYXL` to `THREE_D_REFERENCE` mapping in `config.py` TASK_LAYER_MAP (DEV has `[OPENPYXL, XLWINGS, APPLESCRIPT]`, CURRENT has `[XLWINGS, APPLESCRIPT]`) — openpyxl CAN write cross-sheet formula strings | | |
| TASK-002 | Restore `OPENPYXL` to `EXTERNAL_REFERENCE` mapping (same pattern as TASK-001) — openpyxl can write external reference formula strings | | |
| TASK-003 | Restore `XLWINGS` to `TABLE_CREATE` mapping — xlwings can create Excel tables via live API as fallback | | |
| TASK-004 | Restore `XLWINGS` to `PAGE_BREAK` mapping — xlwings API supports page break insertion | | |
| TASK-005 | Restore `OPENPYXL` to `SORT` mapping + add `APPLESCRIPT` — DEV has `[OPENPYXL, XLWINGS, APPLESCRIPT]`, enables new `_sort_openpyxl_range()` method | | |
| TASK-006 | Restore `APPLESCRIPT` to `SUBTOTAL` mapping — AppleScript can drive Excel subtotals as fallback | | |
| TASK-007 | Restore `APPLESCRIPT` to `GOAL_SEEK` mapping — AppleScript can drive Excel Goal Seek dialog | | |
| TASK-008 | Add `VBA` to `SLICER` mapping — DEV has `[VBA, SYSTEM_EVENTS]`, supports new `manage_slicer()` VBA method | | |
| TASK-009 | Restore `APPLESCRIPT` to `SHEET_COPY` mapping — AppleScript `copy_worksheet` exists as fallback | | |
| TASK-010 | Update `layer_order` in config.py to place VBA before SYSTEM_EVENTS (DEV ordering) — VBA is more reliable than System Events UI automation | | |
| TASK-011 | Run full test suite to verify no regressions from restored mappings | | |

### Phase 2: Merge Safe DEV Additions (No Conflicts)

- GOAL-002: Integrate DEV-only features that have zero conflict with CURRENT code: copilot-studio-connector module, manage_slicer() method, copilot_excel_run.sh improvement, README additions

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-012 | Copy `copilot-studio-connector/` directory from DEV to CURRENT (1,084 LOC, 11 files: apiDefinition.swagger.json, apiProperties.json, script.csx, tools/*.py, server/*.py, resources/*.py) | | |
| TASK-013 | Add `manage_slicer()` method from DEV's `vba_layer.py` (lines 326-456, ~130 lines) to CURRENT's `vba_layer.py` — handles slicer creation, positioning, item selection via VBA | | |
| TASK-014 | Update `copilot_excel_run.sh` with DEV's RUN_ARGS conditional check (lines 91-97) — prevents empty array expansion when no extra args passed | | |
| TASK-015 | Add DEV's "Copilot Autonomous Completion Protocol" section to README.md (lines 224-255, ~32 lines) — documents 6-step autonomous execution protocol for Copilot sessions | | |
| TASK-016 | Run full test suite to verify no regressions | | |

### Phase 3: Merge Complex DEV Changes (engine.py Reconciliation)

- GOAL-003: Integrate DEV's engine.py improvements while preserving CURRENT's corruption safeguards. This requires careful reconciliation of conflicting approaches.

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-017 | Add `_validate_task_targets()` method from DEV (lines 1297-1381, ~84 lines) to CURRENT's engine.py — validates sheet/cell/range/table targets before dispatch. Must integrate AFTER `_prepare_workbook()` in the execution flow | | |
| TASK-018 | Add `_validate_cell_reference()` static method from DEV (lines 1384-1391) — validates A1-style cell references using openpyxl's `coordinate_to_tuple()` | | |
| TASK-019 | Add `_validate_range_reference()` static method from DEV (lines 1392-1408) — validates A1 range references, handles multi-part comma-separated ranges | | |
| TASK-020 | Add `_sort_openpyxl_range()` method from DEV (lines 1447-1515, ~68 lines) — in-place openpyxl sorting with nested `_resolve_col_index()` helper. Required by TASK-005's config mapping change | | |
| TASK-021 | Add `_finalize_xlwings_copy()` method from DEV (lines 1409-1420) — manages xlwings desktop copy lifecycle. Must be reconciled with CURRENT's `_cleanup()` desktop copy handling (lines 1319-1328) | | |
| TASK-022 | Verify `_prepare_workbook()` (CURRENT-only, lines 204-232) is preserved and NOT replaced by DEV's removal — this is a critical corruption safeguard | | |
| TASK-023 | Verify `_restore_backup()` (CURRENT-only, lines 1296-1302) is preserved — DEV removed this; it must stay | | |
| TASK-024 | Wire `_validate_task_targets()` call into `_execute_task()` (line ~559) — call validation BEFORE layer dispatch, AFTER `_prepare_workbook()` | | |
| TASK-025 | Wire `_sort_openpyxl_range()` into `_exec_openpyxl()` dispatch for `TaskType.SORT` — add handler case in the openpyxl dispatcher | | |
| TASK-026 | Run full test suite + add tests for new validation methods and sort implementation | | |

### Phase 4: New Feature — Additional Resource/Data File Uploads

- GOAL-004: Allow users to upload additional resource files (companion .xlsx workbooks, PDFs, images, .docx, .zip, .txt) that are placed in the same working directory as the main workbook during execution

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-027 | Add `resource_files: Optional[list[Path]]` parameter to `ExcelEngine.run()` and `ExcelEngine.execute()` method signatures | | |
| TASK-028 | Add `resource_files` parameter to `EngineConfig` dataclass (default: empty list) | | |
| TASK-029 | Implement resource file staging in `_prepare_workbook()`: copy resource files to the same temp directory as the workbook so relative references resolve correctly | | |
| TASK-030 | Update `gui/app.py`: add Streamlit `file_uploader` widget for "Additional Resource/Data Files" with `accept_multiple_files=True` and type filter `["xlsx", "xlsm", "xls", "csv", "pdf", "docx", "txt", "zip", "png", "jpg", "jpeg", "gif"]` | | |
| TASK-031 | Update `cli.py`: add `--resources` CLI argument accepting one or more file paths, passed through to `engine.run()` | | |
| TASK-032 | Update `mcp-server/server.py`: add `resource_files` parameter to `complete_assignment` and `execute_openpyxl` / `execute_live` tools | | |
| TASK-033 | Handle ZIP resource files: if a `.zip` is uploaded as a resource, extract its contents to the working directory | | |
| TASK-034 | Add tests for resource file staging, ZIP extraction, and cleanup | | |

### Phase 5: New Feature — Multi-Format Instruction Upload

- GOAL-005: Ensure all instruction file formats are supported end-to-end from GUI upload through parsing. Add ZIP container support for instructions.

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-035 | Add `.zip` to `SUPPORTED_EXTENSIONS` in `instruction_parser.py` (line 33) and implement `_parse_zip()` method — extract and concatenate .txt/.docx/.pdf contents from ZIP | | |
| TASK-036 | Update `gui/app.py` instruction uploader `type` filter from `["txt", "docx", "pdf", "rtf"]` to `["txt", "docx", "pdf", "rtf", "zip"]` | | |
| TASK-037 | Verify `cli.py` instruction argument accepts any file path (no extension filter) — should already work since parser handles dispatch | | |
| TASK-038 | Add tests for ZIP instruction parsing (single .txt in ZIP, multiple files in ZIP, nested .docx in ZIP) | | |

### Phase 6: New Feature — Optional .xlsm Conversion for VBA Macros

- GOAL-006: When VBA macros are required for a task (e.g., PivotTables, Slicers), detect the need and present the user with an option to convert to .xlsm format. Never convert automatically.

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-039 | Add `vba_conversion_policy` to `EngineConfig`: enum with values `"ask"` (default), `"always"`, `"never"` — controls .xlsm conversion behavior | | |
| TASK-040 | Implement `_check_vba_requirement()` in engine.py: scans planned tasks for VBA-layer-only TaskTypes (PIVOT_TABLE, PIVOT_CHART, SLICER). Returns `True` if any VBA tasks found AND workbook is `.xlsx` (not already `.xlsm`) | | |
| TASK-041 | Implement `_prompt_xlsm_conversion()` in engine.py: when `vba_conversion_policy="ask"` and VBA tasks detected, emit a callback event/return a prompt object. In CLI mode, print prompt and read stdin. In GUI mode, show Streamlit dialog | | |
| TASK-042 | Implement `_convert_to_xlsm()` in engine.py: rename/copy `.xlsx` → `.xlsm` using openpyxl (save with `keep_vba=True` if VBA data exists, otherwise just change extension). Update internal workbook path reference. Backup original first. | | |
| TASK-043 | Update `gui/app.py`: add conversion dialog using `st.dialog` or `st.warning` + `st.button` when VBA tasks detected — "This assignment requires VBA macros. Convert to .xlsm? [Yes] [No, skip VBA tasks]" | | |
| TASK-044 | Update `cli.py`: add `--xlsm-policy` flag (`ask`/`always`/`never`) mapping to `EngineConfig.vba_conversion_policy` | | |
| TASK-045 | Ensure `.xlsm` output is handled in download flow: GUI download button reflects actual file extension; CLI output preserves extension | | |
| TASK-046 | Add tests for VBA detection, conversion prompt, .xlsm save, and policy enforcement | | |

### Phase 7: Cleanup & Optimization

- GOAL-007: Remove stale code, reconcile test suites, ensure no dead code or broken references remain

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-047 | Merge DEV's unique test file(s) into CURRENT: `test_extraction_false_positives.py` (if different from `test_extraction_fixes.py`) and any other DEV-only test functions | | |
| TASK-048 | Audit all dispatch handlers in engine.py `_exec_openpyxl`, `_exec_xlwings`, etc. for unreachable code paths — every handler must have a corresponding config.py mapping and vice versa | | |
| TASK-049 | Verify xlwings_layer.py methods: CURRENT has 37 methods (12 recently added). Confirm each method is callable from engine.py dispatch and has a corresponding config mapping | | |
| TASK-050 | Run `pytest --cov` to measure code coverage; identify untested code paths in new features | | |
| TASK-051 | Update `pyproject.toml` if new dependencies are needed (e.g., FastAPI/Starlette for copilot-studio-connector) — add as optional dependency group `[copilot-studio]` | | |

### Phase 8: Rebuild macOS .app Bundle

- GOAL-008: Rebuild the standalone macOS application from the merged, tested CURRENT codebase using PyInstaller

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-052 | Copy DEV's `gui/build_app.py` PyInstaller spec/script to CURRENT (if not already present) — this is the .app build pipeline | | |
| TASK-053 | Install PyInstaller in CURRENT venv: `pip install pyinstaller` | | |
| TASK-054 | Run PyInstaller build: `python gui/build_app.py` (or equivalent) — produces `dist/Excel Engine.app` | | |
| TASK-055 | Verify .app launches correctly: double-click or `open "dist/Excel Engine.app"` — should open Streamlit GUI in browser | | |
| TASK-056 | Verify .app includes all merged code: check `Contents/Resources/` for updated engine.py, path_handler.py, etc. | | |
| TASK-057 | Code-sign .app if `sign_app.command` script exists — run it to sign the bundle for macOS Gatekeeper | | |

### Phase 9: Final Integration Testing

- GOAL-009: End-to-end validation of the merged engine with real workbook(s)

| Task | Description | Completed | Date |
|------|-------------|-----------|------|
| TASK-058 | Run full test suite: `pytest -q` — must pass 100% (target: 920+ tests) | | |
| TASK-059 | Test with a real SAM assignment file: verify backup creation, task execution, OOXML validation, no corruption | | |
| TASK-060 | Test resource file upload via GUI: upload main .xlsx + companion .xlsx + .pdf resource | | |
| TASK-061 | Test instruction ZIP upload: create ZIP containing .docx instructions, upload via GUI, verify parsing | | |
| TASK-062 | Test .xlsm conversion prompt: run assignment with PivotTable tasks, verify prompt appears, test both accept/reject paths | | |
| TASK-063 | Test .app bundle end-to-end: launch app, upload files, run assignment, download result | | |

## 3. Alternatives

- **ALT-001**: Cherry-pick DEV commits via `git cherry-pick` instead of manual file merge — rejected because DEV has uncommitted dirty changes (13 modified files) that aren't in any commit, and the two repos have diverged significantly with conflicting changes in the same files
- **ALT-002**: Make DEV the canonical repo and backport CURRENT's safeguards — rejected because CURRENT is clean (committed, pushed), has corruption safeguards, and is 2 commits ahead
- **ALT-003**: Use `git merge` with a remote pointing to DEV — rejected because both repos track the same GitHub remote but have diverged locally; a git merge would create complex conflicts in engine.py, config.py, and xlwings_layer.py that are harder to resolve than manual integration
- **ALT-004**: Skip .xlsm conversion feature, just fail gracefully on VBA tasks with .xlsx — rejected per user requirement REQ-006 to offer the option
- **ALT-005**: Auto-convert to .xlsm when VBA needed — rejected per user requirement that conversion must be user-opted, not automatic

## 4. Dependencies

- **DEP-001**: DEV repo at `/Users/michaelmanthejr/Dev/excel-engine` must remain accessible for file copying during merge
- **DEP-002**: `openpyxl>=3.1.0` — core dependency, already installed
- **DEP-003**: `xlwings>=0.30.0` — live Excel bridge, already installed
- **DEP-004**: `pyautogui>=0.9.54` — UI automation, already installed
- **DEP-005**: `python-docx>=0.8.11` — .docx parsing, already installed
- **DEP-006**: `pdfplumber>=0.9.0` — .pdf parsing, already installed
- **DEP-007**: `PyInstaller>=6.0` — .app bundle building, needs installation for Phase 8
- **DEP-008**: `fastapi` + `starlette` + `uvicorn` — for copilot-studio-connector, add as optional dep group
- **DEP-009**: Python 3.13 venv at `/Users/michaelmanthejr/excel-engine/.venv` — must remain healthy
- **DEP-010**: Microsoft Excel for Mac 365 — required for live layer testing (Phases 2-6)

## 5. Files

- **FILE-001**: `excel_engine/config.py` — Restore 9 layer dispatch mappings, reorder layer_order (Phase 1)
- **FILE-002**: `excel_engine/engine.py` — Add 5 DEV methods, wire validation + sort, add VBA detection + .xlsm conversion (Phases 3, 6)
- **FILE-003**: `excel_engine/layers/vba_layer.py` — Add `manage_slicer()` ~130 lines (Phase 2)
- **FILE-004**: `excel_engine/layers/xlwings_layer.py` — Verify 37 methods intact, no changes expected (Phase 7 audit)
- **FILE-005**: `excel_engine/utils/path_handler.py` — No changes; verify safeguards preserved (Phase 7 audit)
- **FILE-006**: `excel_engine/parsers/instruction_parser.py` — Add `.zip` support + `_parse_zip()` method (Phase 5)
- **FILE-007**: `gui/app.py` — Add resource file uploader, expand instruction formats, add .xlsm dialog (Phases 4, 5, 6)
- **FILE-008**: `excel_engine/cli.py` — Add `--resources` and `--xlsm-policy` flags (Phases 4, 6)
- **FILE-009**: `mcp-server/server.py` — Add `resource_files` parameter to tools (Phase 4)
- **FILE-010**: `copilot_excel_run.sh` — RUN_ARGS conditional fix (Phase 2)
- **FILE-011**: `README.md` — Add Copilot Autonomous Completion Protocol section (Phase 2)
- **FILE-012**: `copilot-studio-connector/` — New directory, 11 files (Phase 2)
- **FILE-013**: `pyproject.toml` — Add optional dependency groups for copilot-studio, pyinstaller (Phase 7)
- **FILE-014**: `gui/build_app.py` — Copy from DEV for .app building (Phase 8)
- **FILE-015**: `tests/` — New tests for validation, sort, resources, ZIP parsing, .xlsm conversion (Phases 3-6, 9)

## 6. Testing

- **TEST-001**: Verify all 9 restored config.py mappings have corresponding dispatch handlers (extend `test_engine_dispatch.py`)
- **TEST-002**: Test `manage_slicer()` VBA method with mock AppleScript/VBA execution
- **TEST-003**: Test `_validate_task_targets()` with valid/invalid sheet names, cell refs, ranges
- **TEST-004**: Test `_validate_cell_reference()` with valid (A1, ZZ999) and invalid (123, AAA) inputs
- **TEST-005**: Test `_validate_range_reference()` with single ranges, multi-part comma ranges, invalid formats
- **TEST-006**: Test `_sort_openpyxl_range()` with numeric data, string data, multi-key sorting, empty ranges
- **TEST-007**: Test resource file staging: files copied to working dir, accessible during execution, cleaned up after
- **TEST-008**: Test ZIP resource extraction: nested files extracted, non-archive files left as-is
- **TEST-009**: Test `_parse_zip()` instruction parsing: single .txt, multiple files, nested .docx, empty ZIP
- **TEST-010**: Test `_check_vba_requirement()`: detects PivotTable/Slicer tasks, ignores non-VBA tasks
- **TEST-011**: Test `_prompt_xlsm_conversion()`: returns correct prompt object in programmatic mode
- **TEST-012**: Test `_convert_to_xlsm()`: backup created, extension changed, openpyxl re-save works, OOXML validation passes
- **TEST-013**: Test VBA conversion policy: "ask" triggers prompt, "always" auto-converts, "never" skips VBA tasks
- **TEST-014**: Full regression suite: all 901+ existing CURRENT tests must still pass after each phase
- **TEST-015**: End-to-end real file test: run actual SAM assignment with backup verification

## 7. Risks & Assumptions

- **RISK-001**: Merging DEV's engine.py changes (~205 lines) into CURRENT's engine.py (~1,346 lines) may introduce subtle execution flow bugs — mitigated by running full test suite after Phase 3
- **RISK-002**: DEV's removal of `_prepare_workbook()` and `_restore_backup()` suggests a different execution model; we must NOT adopt that removal — CURRENT's corruption safeguards are non-negotiable
- **RISK-003**: xlwings_layer.py methods in CURRENT (37 methods) may overlap with DEV's approach of moving operations to openpyxl — keeping both provides more fallback options, but increases maintenance surface
- **RISK-004**: copilot-studio-connector adds FastAPI/Starlette dependencies — keeping as optional group prevents bloating the core install
- **RISK-005**: PyInstaller .app rebuild may fail if Python 3.13 compatibility issues exist with PyInstaller — mitigated by testing build before final release
- **RISK-006**: .xlsm conversion may break SAM fingerprinting (`config.py:213` has `sam_fingerprint_protected=True`) — mitigated by making conversion opt-in only, never automatic
- **RISK-007**: Resource file staging could consume significant disk space for large files — mitigated by temp directory cleanup in `_cleanup()`
- **ASSUMPTION-001**: DEV repo at `~/Dev/excel-engine` will remain accessible throughout the merge process
- **ASSUMPTION-002**: User will re-download a clean SAM assignment file for end-to-end testing (previous file was corrupted)
- **ASSUMPTION-003**: The 13 uncommitted modified files in DEV contain the latest experimental features — we'll use DEV's current working tree state, not just its committed state
- **ASSUMPTION-004**: Microsoft Excel for Mac 365 is installed and accessible for live layer testing

## 8. Related Specifications / Further Reading

- [Excel Engine README](https://github.com/mmanthe37/excel-engine/blob/main/README.md)
- [Architecture Documentation](docs/architecture.md)
- [Layer Documentation](docs/layers.md)
- [Troubleshooting Guide](docs/troubleshooting.md)
- [CONTRIBUTING.md](CONTRIBUTING.md)
- [CHANGELOG.md](CHANGELOG.md)
- [openpyxl Documentation](https://openpyxl.readthedocs.io/)
- [xlwings Documentation](https://docs.xlwings.org/)
- [PyInstaller Documentation](https://pyinstaller.org/en/stable/)
- [OOXML Standard (ISO/IEC 29500)](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
