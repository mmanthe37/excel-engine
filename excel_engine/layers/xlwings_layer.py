"""
Layer 2: xlwings — Live Excel API bridge.

Handles live cell read/write, named range operations, table properties,
range-level operations (subtotals), split panes, and advanced filter.

KEY RULES:
  - Colons in paths break xlwings — use PathHandler.safe_copy_for_xlwings()
  - Excel must be running for xlwings to work
  - subtotal via rng.api.subtotal()
  - split panes via app.api.active_window.split.set(True)
  - advanced filter via rng.api.advanced_filter(action=2, ...)
"""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger(__name__)

try:
    import xlwings as xw
except ImportError:
    xw = None  # type: ignore[assignment]
    logger.warning("xlwings not installed — Layer 2 unavailable")


class XlwingsLayer:
    """Layer 2 — live Excel API bridge via xlwings."""

    def __init__(self) -> None:
        self._app: Optional[Any] = None
        self._wb: Optional[Any] = None
        self._path: Optional[Path] = None

    @property
    def available(self) -> bool:
        return xw is not None

    def _require_xlwings(self) -> None:
        if not self.available:
            raise RuntimeError("xlwings is not installed")

    # ── Connection ──

    def connect(self, path: Optional[Path] = None) -> None:
        """
        Connect to Excel. If path is given, open/connect to that workbook.
        If no path, connect to the active workbook.
        """
        self._require_xlwings()
        if path:
            self._path = Path(path)
            self._wb = xw.Book(str(self._path))
        else:
            self._wb = xw.books.active
        self._app = self._wb.app
        logger.info("Connected to workbook: %s", self._wb.name)

    def disconnect(self) -> None:
        """Disconnect without closing Excel."""
        self._wb = None
        self._app = None

    def _ws(self, sheet_name: Optional[str] = None):
        """Get a sheet object."""
        if not self._wb:
            raise RuntimeError("Not connected to a workbook")
        if sheet_name:
            return self._wb.sheets[sheet_name]
        return self._wb.sheets.active

    # ── Cell Operations ──

    def read_cell(self, cell: str, sheet: Optional[str] = None) -> Any:
        """Read a cell value from the live workbook."""
        ws = self._ws(sheet)
        return ws.range(cell).value

    def write_cell(self, cell: str, value: Any, sheet: Optional[str] = None) -> None:
        """Write a value to a cell in the live workbook."""
        ws = self._ws(sheet)
        ws.range(cell).value = value

    def read_range(self, range_str: str, sheet: Optional[str] = None) -> list[list[Any]]:
        """Read a range of values as a 2D list."""
        ws = self._ws(sheet)
        return ws.range(range_str).value

    def write_range(
        self, start_cell: str, data: list[list[Any]], sheet: Optional[str] = None
    ) -> None:
        """Write a 2D array to a range."""
        ws = self._ws(sheet)
        ws.range(start_cell).value = data

    def set_formula(self, cell: str, formula: str, sheet: Optional[str] = None) -> None:
        """Set a formula on a live cell — supports structural references."""
        ws = self._ws(sheet)
        if not formula.startswith("="):
            formula = f"={formula}"
        ws.range(cell).formula = formula
        logger.debug("Live formula %s = %s", cell, formula)

    # ── Named Ranges ──

    def get_named_range_value(self, name: str) -> Any:
        """Read the value of a workbook-scoped named range."""
        if not self._wb:
            raise RuntimeError("Not connected")
        return self._wb.names[name].refers_to_range.value

    def set_named_range_value(self, name: str, value: Any) -> None:
        """Set the value of a workbook-scoped named range."""
        if not self._wb:
            raise RuntimeError("Not connected")
        self._wb.names[name].refers_to_range.value = value

    # ── Table Operations ──

    def get_table_range(self, table_name: str, sheet: str) -> str:
        """Get the address of a table by name."""
        ws = self._ws(sheet)
        for table in ws.api.list_objects():
            if table.name() == table_name:
                rng = table.range()
                return rng.get()
        raise ValueError(f"Table '{table_name}' not found on sheet '{sheet}'")

    def set_table_style(self, table_name: str, style: str, sheet: str) -> None:
        """Set the style of an existing table."""
        ws = self._ws(sheet)
        for table in ws.api.list_objects():
            if table.name() == table_name:
                table.table_style.set(style)
                logger.info("Set table '%s' style to %s", table_name, style)
                return
        raise ValueError(f"Table '{table_name}' not found")

    # ── Subtotals ──

    def add_subtotal(
        self,
        range_str: str,
        group_by: int,
        function: int,
        total_list: list[int],
        replace: bool = True,
        page_breaks: bool = False,
        summary_below: bool = True,
        sheet: Optional[str] = None,
    ) -> None:
        """
        Add subtotals to a range via the xlwings API.

        Args:
            range_str: e.g. "A1:G50"
            group_by: 1-based column number to group by
            function: xlSum=-4157, xlAverage=-4106, etc.
            total_list: list of 1-based column numbers to total
            replace: replace existing subtotals
            page_breaks: add page breaks between groups
            summary_below: put summary below data
        """
        ws = self._ws(sheet)
        rng = ws.range(range_str).api
        rng.subtotal(
            group_by,
            function,
            total_list,
            replace,
            page_breaks,
            summary_below,
        )
        logger.info("Added subtotals to %s, group by col %d", range_str, group_by)

    # ── Split Panes ──

    def split_panes(self, row: int, col: int, sheet: Optional[str] = None) -> None:
        """
        Split the window at a given row/col position.
        Must select the split cell first.
        """
        ws = self._ws(sheet)
        cell_ref = f"{chr(64 + col)}{row}"
        ws.range(cell_ref).select()
        self._app.api.active_window.split.set(True)
        logger.info("Split panes at row %d, col %d", row, col)

    def remove_split_panes(self) -> None:
        """Remove split panes from the active window."""
        if self._app:
            self._app.api.active_window.split.set(False)

    # ── Advanced Filter ──

    def advanced_filter(
        self,
        list_range: str,
        criteria_range: str,
        copy_to_range: Optional[str] = None,
        unique: bool = False,
        sheet: Optional[str] = None,
    ) -> None:
        """
        Apply an advanced filter.
        action=1: filter in place, action=2: copy to another location.
        """
        ws = self._ws(sheet)
        list_rng = ws.range(list_range).api
        criteria_rng = ws.range(criteria_range).api

        if copy_to_range:
            copy_rng = ws.range(copy_to_range).api
            list_rng.advanced_filter(
                action=2,
                criteria_range=criteria_rng,
                copy_to_range=copy_rng,
                unique=unique,
            )
            logger.info("Advanced filter: copy to %s", copy_to_range)
        else:
            list_rng.advanced_filter(
                action=1,
                criteria_range=criteria_rng,
                unique=unique,
            )
            logger.info("Advanced filter: in place")

    # ── Column Width ──

    def autofit_columns(self, range_str: str, sheet: Optional[str] = None) -> None:
        """Autofit column widths for a range."""
        ws = self._ws(sheet)
        ws.range(range_str).columns.autofit()

    def set_column_width(
        self, column: str, width: float, sheet: Optional[str] = None
    ) -> None:
        """Set a column's width."""
        ws = self._ws(sheet)
        ws.range(f"{column}:{column}").column_width = width

    # ── Number Format ──

    def set_number_format(
        self, range_str: str, fmt: str, sheet: Optional[str] = None
    ) -> None:
        """Apply a number format to a live range."""
        ws = self._ws(sheet)
        ws.range(range_str).number_format = fmt

    # ── Sheet Navigation ──

    def activate_sheet(self, sheet_name: str) -> None:
        """Activate (select) a sheet."""
        if not self._wb:
            raise RuntimeError("Not connected")
        self._wb.sheets[sheet_name].activate()
        time.sleep(0.3)

    # ── Save ──

    def save(self) -> None:
        """Save the active workbook."""
        if self._wb:
            self._wb.save()
            logger.info("Saved workbook via xlwings")
