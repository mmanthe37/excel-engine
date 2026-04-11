"""
Configuration, constants, and enums for the Excel Engine.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, IntEnum
from pathlib import Path
from typing import Optional


class Layer(IntEnum):
    """Execution layers in priority order."""
    OPENPYXL = 1
    XLWINGS = 2
    APPLESCRIPT = 3
    SYSTEM_EVENTS = 4
    VBA = 5
    PYAUTOGUI = 6


class TaskType(Enum):
    """Recognized task types that the engine can execute."""
    # ── Data Entry & Formulas ──
    FORMULA = "formula"
    CELL_VALUE = "cell_value"
    TEXT_FUNCTION = "text_function"          # CONCAT, LEFT, MID, RIGHT, UPPER, etc.
    LOOKUP_FUNCTION = "lookup_function"      # XLOOKUP, VLOOKUP, HLOOKUP, INDEX/MATCH
    FILTER_FUNCTION = "filter_function"      # FILTER dynamic array function
    SORT_FUNCTION = "sort_function"          # SORT dynamic array function
    UNIQUE_FUNCTION = "unique_function"      # UNIQUE dynamic array function
    THREE_D_REFERENCE = "three_d_reference"  # cross-sheet formulas like =SUM(Sheet1:Sheet3!A1)
    EXTERNAL_REFERENCE = "external_reference"  # links to other workbooks

    # ── Tables ──
    TABLE_CREATE = "table_create"
    TABLE_STYLE = "table_style"
    TABLE_TOTAL_ROW = "table_total_row"
    CALCULATED_COLUMN = "calculated_column"

    # ── Formatting ──
    FORMATTING = "formatting"
    CONDITIONAL_FORMAT = "conditional_format"
    NUMBER_FORMAT = "number_format"
    ALIGNMENT = "alignment"
    COLUMN_WIDTH = "column_width"
    ROW_HEIGHT = "row_height"
    FONT = "font"
    FILL = "fill"
    BORDER = "border"
    MERGE_CELLS = "merge_cells"
    TAB_COLOR = "tab_color"

    # ── View & Layout ──
    FREEZE_PANES = "freeze_panes"
    SPLIT_PANES = "split_panes"
    PAGE_BREAK = "page_break"
    PRINT_SETTINGS = "print_settings"

    # ── Data Tools ──
    AUTOFILTER = "autofilter"
    ADVANCED_FILTER = "advanced_filter"
    SORT = "sort"
    SUBTOTAL = "subtotal"
    DATA_VALIDATION = "data_validation"
    GOAL_SEEK = "goal_seek"

    # ── Charts ──
    CHART_BAR = "chart_bar"
    CHART_LINE = "chart_line"
    CHART_PIE = "chart_pie"
    CHART_SCATTER = "chart_scatter"
    CHART_AREA = "chart_area"
    CHART_COMBO = "chart_combo"
    CHART_HISTOGRAM = "chart_histogram"
    SPARKLINE = "sparkline"

    # ── Ranges & References ──
    NAMED_RANGE = "named_range"
    HYPERLINK = "hyperlink"

    # ── Advanced Features ──
    SLICER = "slicer"
    PIVOT_TABLE = "pivot_table"
    PIVOT_CHART = "pivot_chart"

    # ── Sheet Operations ──
    SHEET_CREATE = "sheet_create"
    SHEET_RENAME = "sheet_rename"
    SHEET_MOVE = "sheet_move"
    SHEET_COPY = "sheet_copy"

    # ── File Operations ──
    SAVE = "save"
    SAVE_AS = "save_as"


# Maps each TaskType to the preferred layer(s) that can handle it,
# with fallback order.
TASK_LAYER_MAP: dict[TaskType, list[Layer]] = {
    # ── Data Entry & Formulas ──
    TaskType.FORMULA:             [Layer.OPENPYXL, Layer.XLWINGS, Layer.APPLESCRIPT],
    TaskType.CELL_VALUE:          [Layer.OPENPYXL, Layer.XLWINGS, Layer.APPLESCRIPT],
    TaskType.TEXT_FUNCTION:       [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.LOOKUP_FUNCTION:     [Layer.OPENPYXL, Layer.XLWINGS, Layer.APPLESCRIPT],    # openpyxl can write the formula text
    TaskType.FILTER_FUNCTION:     [Layer.XLWINGS, Layer.APPLESCRIPT],    # dynamic array — live only
    TaskType.SORT_FUNCTION:       [Layer.XLWINGS, Layer.APPLESCRIPT],    # dynamic array — live only
    TaskType.UNIQUE_FUNCTION:     [Layer.XLWINGS, Layer.APPLESCRIPT],    # dynamic array — live only
    TaskType.THREE_D_REFERENCE:   [Layer.XLWINGS, Layer.APPLESCRIPT],    # cross-sheet — live preferred
    TaskType.EXTERNAL_REFERENCE:  [Layer.XLWINGS, Layer.APPLESCRIPT],    # external links — live only

    # ── Tables ──
    TaskType.TABLE_CREATE:        [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.TABLE_STYLE:         [Layer.OPENPYXL, Layer.SYSTEM_EVENTS],
    TaskType.TABLE_TOTAL_ROW:     [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.CALCULATED_COLUMN:   [Layer.XLWINGS, Layer.APPLESCRIPT],    # structural refs need LIVE

    # ── Formatting ──
    TaskType.FORMATTING:          [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.CONDITIONAL_FORMAT:  [Layer.OPENPYXL],
    TaskType.NUMBER_FORMAT:       [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.ALIGNMENT:           [Layer.OPENPYXL],
    TaskType.COLUMN_WIDTH:        [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.ROW_HEIGHT:          [Layer.OPENPYXL],
    TaskType.FONT:                [Layer.OPENPYXL],
    TaskType.FILL:                [Layer.OPENPYXL],
    TaskType.BORDER:              [Layer.OPENPYXL],
    TaskType.MERGE_CELLS:         [Layer.OPENPYXL],
    TaskType.TAB_COLOR:           [Layer.OPENPYXL, Layer.XLWINGS],

    # ── View & Layout ──
    TaskType.FREEZE_PANES:        [Layer.OPENPYXL, Layer.APPLESCRIPT],
    TaskType.SPLIT_PANES:         [Layer.XLWINGS],
    TaskType.PAGE_BREAK:          [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.PRINT_SETTINGS:      [Layer.OPENPYXL],

    # ── Data Tools ──
    TaskType.AUTOFILTER:          [Layer.OPENPYXL, Layer.APPLESCRIPT],
    TaskType.ADVANCED_FILTER:     [Layer.XLWINGS],
    TaskType.SORT:                [Layer.APPLESCRIPT, Layer.XLWINGS],
    TaskType.SUBTOTAL:            [Layer.XLWINGS, Layer.APPLESCRIPT],
    TaskType.DATA_VALIDATION:     [Layer.OPENPYXL],
    TaskType.GOAL_SEEK:           [Layer.XLWINGS, Layer.APPLESCRIPT],    # live Excel required

    # ── Charts ──
    TaskType.CHART_BAR:           [Layer.OPENPYXL],
    TaskType.CHART_LINE:          [Layer.OPENPYXL],
    TaskType.CHART_PIE:           [Layer.OPENPYXL],
    TaskType.CHART_SCATTER:       [Layer.OPENPYXL, Layer.XLWINGS, Layer.SYSTEM_EVENTS],
    TaskType.CHART_AREA:          [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.CHART_COMBO:         [Layer.OPENPYXL, Layer.XLWINGS, Layer.SYSTEM_EVENTS],  # openpyxl merges bar+line
    TaskType.CHART_HISTOGRAM:     [Layer.SYSTEM_EVENTS],                 # cx:chart — must use UI
    TaskType.SPARKLINE:           [Layer.XLWINGS, Layer.SYSTEM_EVENTS],  # sparklines need live API

    # ── Ranges & References ──
    TaskType.NAMED_RANGE:         [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.HYPERLINK:           [Layer.OPENPYXL, Layer.XLWINGS],

    # ── Advanced Features ──
    TaskType.SLICER:              [Layer.SYSTEM_EVENTS],                 # must use ribbon UI
    TaskType.PIVOT_TABLE:         [Layer.VBA],
    TaskType.PIVOT_CHART:         [Layer.VBA],

    # ── Sheet Operations ──
    TaskType.SHEET_CREATE:        [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.SHEET_RENAME:        [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.SHEET_MOVE:          [Layer.OPENPYXL, Layer.XLWINGS],
    TaskType.SHEET_COPY:          [Layer.OPENPYXL, Layer.APPLESCRIPT, Layer.XLWINGS],    # openpyxl can copy_worksheet

    # ── File Operations ──
    TaskType.SAVE:                [Layer.APPLESCRIPT, Layer.XLWINGS],
    TaskType.SAVE_AS:             [Layer.APPLESCRIPT],
}
# Note: Layer.PYAUTOGUI is available as a last-resort fallback but no TaskTypes
# map to it by default. It can be added to layer_order for pixel-level automation.


@dataclass
class EngineConfig:
    """Runtime configuration for the Excel Engine."""

    # Timeouts (seconds)
    scan_timeout: float = 120.0
    section_timeout: float = 300.0
    applescript_timeout: float = 30.0
    ui_automation_timeout: float = 60.0
    vba_execution_timeout: float = 45.0

    # Retries
    max_retries: int = 3
    retry_delay: float = 2.0

    # Layer preferences — lower index = higher priority
    layer_order: list[Layer] = field(
        default_factory=lambda: [
            Layer.OPENPYXL,
            Layer.XLWINGS,
            Layer.APPLESCRIPT,
            Layer.SYSTEM_EVENTS,
            Layer.VBA,
            Layer.PYAUTOGUI,
        ]
    )

    # Paths
    desktop_path: Path = field(
        default_factory=lambda: Path.home() / "Desktop"
    )
    working_dir: Optional[Path] = None

    # SAM-specific
    sam_fingerprint_protected: bool = True  # never rename/move/save-as SAM files
    copy_to_desktop_for_xlwings: bool = True  # avoid colon-in-path bug

    # UI automation
    retina_display: bool = True  # divide physical coords by 2 for PyAutoGUI
    ui_delay_between_actions: float = 0.5

    # VBA
    vba_split_threshold: int = 50  # lines per sub to avoid OOM Error 7

    # Verification
    verify_after_each_section: bool = True

    # Circuit breaker for layer cascade
    circuit_breaker_threshold: int = 5
    circuit_breaker_reset_seconds: int = 300

    # Parallel execution (opt-in)
    parallel_execution: bool = False
    max_workers: int = 4

    # Formula recalculation (requires LibreOffice — gracefully skipped if absent)
    recalculate_formulas: bool = True
    recalc_timeout: int = 30

    def get_layers_for_task(self, task_type: TaskType) -> list[Layer]:
        """Return the ordered list of layers that can handle a given task type."""
        candidates = TASK_LAYER_MAP.get(task_type, [])
        # Only consider layers present in layer_order to avoid ValueError
        available = [l for l in candidates if l in self.layer_order]
        return sorted(available, key=lambda l: self.layer_order.index(l))
