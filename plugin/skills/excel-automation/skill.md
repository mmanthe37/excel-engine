# Excel Automation Skill

## Description

Provides comprehensive knowledge and patterns for automating Microsoft Excel on macOS using the Excel Engine's 6-layer architecture. This skill covers offline file manipulation (openpyxl), live API bridging (xlwings), AppleScript commands, System Events ribbon control, VBA injection via the Visual Basic Editor, and last-resort PyAutoGUI desktop control.

## When to Trigger

- User asks to complete an Excel workbook task (formulas, formatting, charts, tables)
- User needs to automate repetitive Excel operations on macOS
- User encounters Excel automation errors (repair dialogs, VBA failures, System Events timeouts)
- User wants to create PivotTables, slicers, or PivotCharts on Mac Excel
- User needs to sort, filter, or apply subtotals via automation

## Key Knowledge

### Layer Selection

Always use the **lowest-numbered layer** that can accomplish the task:

1. **openpyxl** — offline, most reliable, handles ~60% of tasks
2. **xlwings** — live Excel API, needed for tables and real-time verification
3. **AppleScript** — Excel-specific (sort, filter, freeze, hyperlinks, save)
4. **System Events** — ribbon/dialog UI (subtotals, slicer styling, contextual tabs)
5. **VBA via VBE** — PivotTables, slicers, PivotCharts (clipboard injection only)
6. **PyAutoGUI** — screenshot-click loop, last resort only

### Critical Anti-Patterns

- **NEVER** rename or move SAM workbooks (fingerprint protection)
- **NEVER** use `osascript "run VB macro"` — it silently fails on Mac
- **NEVER** hardcode System Events group indexes — they change per ribbon tab
- **NEVER** assume GUI operations succeeded — always verify with screenshot
- **NEVER** write monolithic VBA >100 lines — split into sub-procedures
- **ALWAYS** add `CleanExistingObjects` before creating PivotTables/Slicers
- **ALWAYS** divide Retina screenshot coordinates by 2 for PyAutoGUI

### ExcelEngine API

```python
from excel_engine import ExcelEngine
from excel_engine.config import EngineConfig

engine = ExcelEngine(EngineConfig(verify_after_each_section=True))
result = engine.run(workbook=Path("file.xlsx"), instructions=Path("instructions.docx"))
```

### Supported Task Types

`formula`, `table_create`, `table_style`, `table_total_row`, `calculated_column`,
`formatting`, `conditional_format`, `number_format`, `alignment`, `column_width`,
`row_height`, `freeze_panes`, `autofilter`, `sort`, `subtotal`, `chart_bar`,
`chart_line`, `chart_pie`, `named_range`, `data_validation`, `slicer`,
`pivot_table`, `pivot_chart`, `sheet_create`, `sheet_rename`, `cell_value`,
`merge_cells`, `border`, `fill`, `font`, `save`, `print_settings`
