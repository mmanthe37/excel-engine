"""
Layer 3: AppleScript — Excel-specific AppleScript commands.

Handles multi-level sort, autofilter, freeze panes, sheet navigation,
cell formulas, shape positioning, save, and save-as.

All commands are executed via subprocess osascript with proper escaping.

KEY RULES:
  - Multi-level sort: key1 ... order1 ... key2 ... order2 ... header header yes
  - Subtotal enum values: xlSum = -4157, xlAverage = -4106
  - Save-as xlsx: file format Excel XML file format
  - Always wrap in tell application "Microsoft Excel"
"""

from __future__ import annotations

import logging
from typing import Optional

from excel_engine.utils.mac_utils import MacUtils

logger = logging.getLogger(__name__)


class AppleScriptLayer:
    """Layer 3 — Excel-specific AppleScript automation."""

    def __init__(self, timeout: float = 30.0) -> None:
        self.timeout = timeout

    def _run(self, script: str) -> str:
        """Execute an AppleScript command string."""
        return MacUtils.run_applescript(script, timeout=self.timeout)

    def _tell_excel(self, commands: str) -> str:
        """Wrap commands in tell application 'Microsoft Excel'."""
        script = f'tell application "Microsoft Excel"\n{commands}\nend tell'
        return self._run(script)

    @staticmethod
    def _escape(value: str) -> str:
        """Escape a string for AppleScript embedding."""
        return value.replace("\\", "\\\\").replace('"', '\\"')

    # ── Sheet Navigation ──

    def activate_sheet(self, sheet_name: str) -> None:
        """Activate a specific worksheet."""
        safe_name = self._escape(sheet_name)
        self._tell_excel(
            f'activate object sheet "{safe_name}" of workbook 1'
        )
        logger.info("Activated sheet: %s", sheet_name)

    def get_active_sheet_name(self) -> str:
        """Get the name of the currently active sheet."""
        return self._tell_excel("get name of active sheet")

    # ── Cell Operations ──

    def set_cell_value(self, cell: str, value: str, sheet: Optional[str] = None) -> None:
        """Set a cell's value via AppleScript."""
        if sheet:
            self.activate_sheet(sheet)
        safe_val = self._escape(str(value))
        self._tell_excel(
            f'set value of range "{cell}" of active sheet to "{safe_val}"'
        )

    def set_cell_formula(self, cell: str, formula: str, sheet: Optional[str] = None) -> None:
        """
        Set a cell's formula via AppleScript.
        Supports structural references like =CONCAT([@LAST],", ",[@FIRST]).
        """
        if sheet:
            self.activate_sheet(sheet)
        safe_formula = self._escape(formula)
        self._tell_excel(
            f'set formula of range "{cell}" of active sheet to "{safe_formula}"'
        )
        logger.debug("AppleScript formula %s = %s", cell, formula)

    def get_cell_value(self, cell: str, sheet: Optional[str] = None) -> str:
        """Read a cell's value via AppleScript."""
        if sheet:
            self.activate_sheet(sheet)
        return self._tell_excel(
            f'get value of range "{cell}" of active sheet'
        )

    # ── Sorting ──

    def sort_range(
        self,
        range_str: str,
        keys: list[dict],
        sheet: Optional[str] = None,
    ) -> None:
        """
        Multi-level sort via AppleScript.

        keys: list of dicts with 'range' and 'order' ('ascending'/'descending')
        Example:
            keys=[
                {"range": "A1:A50", "order": "ascending"},
                {"range": "B1:B50", "order": "descending"},
            ]
        """
        if sheet:
            self.activate_sheet(sheet)

        key_parts = []
        for i, key in enumerate(keys, 1):
            key_range = self._escape(key["range"])
            order = key.get("order", "ascending")
            order_enum = (
                "sort ascending" if order == "ascending" else "sort descending"
            )
            key_parts.append(f"key{i} range \"{key_range}\" order{i} {order_enum}")

        keys_str = " ".join(key_parts)
        cmd = f'sort range "{self._escape(range_str)}" of active sheet {keys_str} header header yes'
        self._tell_excel(cmd)
        logger.info("Sorted %s with %d keys", range_str, len(keys))

    # ── AutoFilter ──

    def set_autofilter(
        self,
        range_str: str,
        field: int,
        criteria: str,
        sheet: Optional[str] = None,
    ) -> None:
        """
        Apply autofilter to a range.
        field: 1-based column index within the range.
        criteria: filter value (e.g., ">100", "Sales", etc.)
        """
        if sheet:
            self.activate_sheet(sheet)
        safe_criteria = self._escape(criteria)
        self._tell_excel(
            f'autofilter range "{range_str}" of active sheet '
            f'field {field} criteria1 "{safe_criteria}"'
        )
        logger.info("Autofilter on %s, field %d, criteria '%s'", range_str, field, criteria)

    def clear_autofilter(self, sheet: Optional[str] = None) -> None:
        """Clear all autofilters on the active sheet."""
        if sheet:
            self.activate_sheet(sheet)
        self._tell_excel(
            'if auto filter mode of active sheet then\n'
            '    autofilter range (used range of active sheet)\n'
            'end if'
        )

    # ── Freeze Panes ──

    def freeze_panes(self, cell: str, sheet: Optional[str] = None) -> None:
        """
        Freeze panes at a cell. First selects the cell, then sets freeze.
        e.g., cell='A2' freezes row 1.
        """
        if sheet:
            self.activate_sheet(sheet)
        self._tell_excel(
            f'select range "{self._escape(cell)}" of active sheet\n'
            f'set freeze panes of active window to true'
        )
        logger.info("Freeze panes at %s", cell)

    def unfreeze_panes(self) -> None:
        """Remove freeze panes."""
        self._tell_excel("set freeze panes of active window to false")

    # ── Save ──

    def save(self) -> None:
        """Save the active workbook."""
        self._tell_excel("save workbook 1")
        logger.info("Saved workbook via AppleScript")

    def save_as_xlsx(self, path: str, filename: str) -> None:
        """
        Save-as .xlsx format.
        path: POSIX directory path (e.g., /Users/me/Desktop)
        filename: file name without extension
        """
        safe_path = self._escape(path)
        safe_name = self._escape(filename)
        full_path = f"{safe_path}/{safe_name}.xlsx"
        self._tell_excel(
            f'save workbook 1 as filename "{full_path}" '
            f'file format Excel XML file format'
        )
        logger.info("Saved as xlsx: %s", full_path)

    # ── Shape Positioning (for slicers etc.) ──

    def set_shape_position(
        self,
        shape_name: str,
        left: Optional[float] = None,
        top: Optional[float] = None,
        width: Optional[float] = None,
        height: Optional[float] = None,
    ) -> None:
        """Position a shape (slicer, chart, etc.) on the active sheet."""
        safe_name = self._escape(shape_name)
        commands = []
        if left is not None:
            commands.append(
                f'set left position of shape "{safe_name}" of active sheet to {left}'
            )
        if top is not None:
            commands.append(
                f'set top of shape "{safe_name}" of active sheet to {top}'
            )
        if width is not None:
            commands.append(
                f'set width of shape "{safe_name}" of active sheet to {width}'
            )
        if height is not None:
            commands.append(
                f'set height of shape "{safe_name}" of active sheet to {height}'
            )
        if commands:
            self._tell_excel("\n".join(commands))
            logger.info("Positioned shape '%s'", shape_name)

    # ── Worksheet Info ──

    def get_sheet_names(self) -> list[str]:
        """Get all sheet names in the active workbook."""
        result = self._tell_excel("get name of every sheet of workbook 1")
        return [s.strip() for s in result.split(",")]

    def get_used_range(self, sheet: Optional[str] = None) -> str:
        """Get the used range address of a sheet."""
        if sheet:
            self.activate_sheet(sheet)
        return self._tell_excel(
            "get address of used range of active sheet"
        )

    def get_workbook_name(self) -> str:
        """Get the name of the active workbook."""
        return self._tell_excel("get name of workbook 1")

    # ── Selection / Navigation ──

    def select_range(self, range_str: str, sheet: Optional[str] = None) -> None:
        """Select a range."""
        if sheet:
            self.activate_sheet(sheet)
        self._tell_excel(f'select range "{self._escape(range_str)}" of active sheet')

    def go_to_cell(self, cell: str, sheet: Optional[str] = None) -> None:
        """Navigate to a specific cell."""
        if sheet:
            self.activate_sheet(sheet)
        self._tell_excel(f'select range "{self._escape(cell)}" of active sheet')

    # ── Miscellaneous ──

    def calculate(self) -> None:
        """Force recalculation."""
        self._tell_excel("calculate")

    def screen_updating(self, enabled: bool) -> None:
        """Toggle screen updating for performance."""
        val = "true" if enabled else "false"
        self._tell_excel(f"set screen updating to {val}")
