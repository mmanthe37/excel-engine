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
        from openpyxl.utils import get_column_letter
        cell_ref = f"{get_column_letter(col)}{row}"
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

    # ── Charts ──

    def add_chart(
        self,
        chart_type: str,
        data_range: str,
        sheet: Optional[str] = None,
        title: str = "",
        anchor: str = "E2",
        width: float = 450,
        height: float = 300,
        secondary_axis_series: Optional[list[int]] = None,
    ) -> None:
        """
        Add a chart using xlwings live API.

        chart_type: xlwings constant name — 'scatter', 'area', 'line',
                    'bar_clustered', 'column_clustered', etc.
        secondary_axis_series: list of 0-based series indices to put on secondary axis.
        """
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")

        ws = self._ws(sheet)

        # Map friendly names to xlwings chart type strings
        _CHART_MAP = {
            "scatter": "xy_scatter",
            "scatter_lines": "xy_scatter_lines",
            "scatter_smooth": "xy_scatter_smooth",
            "area": "area",
            "area_stacked": "area_stacked",
            "line": "line",
            "line_markers": "line_markers",
            "bar": "bar_clustered",
            "column": "column_clustered",
            "pie": "pie",
            "combo": "column_clustered",  # combo starts as column, add line series on secondary
        }

        xl_type = _CHART_MAP.get(chart_type, chart_type)

        chart = ws.charts.add(
            left=ws.range(anchor).left,
            top=ws.range(anchor).top,
            width=width,
            height=height,
        )
        chart.set_source_data(ws.range(data_range))
        chart.chart_type = xl_type

        if title:
            chart.api.has_title = True
            chart.api.chart_title.text = title

        # Move specified series to secondary axis
        if secondary_axis_series:
            for idx in secondary_axis_series:
                try:
                    series = chart.api.series_collection(idx + 1)  # 1-based
                    series.axis_group = 2  # xlSecondary
                except Exception:
                    logger.warning("Could not set series %d to secondary axis", idx)

        logger.info(
            "Added %s chart '%s' at %s via xlwings",
            chart_type, title, anchor,
        )

    def add_sparkline(
        self,
        data_range: str,
        location_range: str,
        sparkline_type: str = "line",
        sheet: Optional[str] = None,
    ) -> None:
        """
        Add sparklines using xlwings API.

        sparkline_type: 'line' (6), 'column' (7), 'win_loss' (8)
        Uses SparklineGroups.Add via the COM/Apple Events API.
        """
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")

        ws = self._ws(sheet)

        _SPARK_TYPE = {
            "line": 6,       # xlSparkLine
            "column": 7,     # xlSparkColumn
            "win_loss": 8,   # xlSparkColumnStacked100
        }
        spark_type_id = _SPARK_TYPE.get(sparkline_type, 6)

        try:
            loc = ws.range(location_range)
            loc.api.sparkline_groups.add(
                spark_type_id,
                ws.range(data_range).api,
            )
            logger.info(
                "Added %s sparkline: data=%s, location=%s via xlwings",
                sparkline_type, data_range, location_range,
            )
        except Exception as e:
            logger.warning("xlwings sparkline failed (%s) — fallback to System Events", e)
            raise

    # ── Save ──

    def save(self) -> None:
        """Save the active workbook."""
        if self._wb:
            self._wb.save()
            logger.info("Saved workbook via xlwings")

    # ── Named Range Creation ──

    def create_named_range(
        self, name: str, refers_to: str, sheet: Optional[str] = None,
    ) -> None:
        """Create a workbook-scoped named range."""
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")
        ws = self._ws(sheet)
        ref_range = ws.range(refers_to)
        self._wb.names.add(name, f"={ws.name}!{ref_range.address}")
        logger.info("Created named range '%s' → %s", name, refers_to)

    # ── Hyperlinks ──

    def add_hyperlink(
        self,
        cell: str,
        url: str,
        display_text: Optional[str] = None,
        sheet: Optional[str] = None,
    ) -> None:
        """Add a hyperlink to a cell."""
        self._require_xlwings()
        ws = self._ws(sheet)
        rng = ws.range(cell)
        rng.add_hyperlink(url, text_to_display=display_text or url)
        logger.info("Added hyperlink at %s → %s", cell, url)

    # ── Sheet Operations ──

    def add_sheet(self, name: str, before: Optional[str] = None, after: Optional[str] = None) -> None:
        """Add a new sheet."""
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")
        kwargs: dict[str, Any] = {"name": name}
        if before:
            kwargs["before"] = self._wb.sheets[before]
        elif after:
            kwargs["after"] = self._wb.sheets[after]
        self._wb.sheets.add(**kwargs)
        logger.info("Added sheet '%s'", name)

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """Rename a sheet."""
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")
        self._wb.sheets[old_name].name = new_name
        logger.info("Renamed sheet '%s' → '%s'", old_name, new_name)

    def move_sheet(
        self, sheet_name: str, before: Optional[str] = None, after: Optional[str] = None,
    ) -> None:
        """Move a sheet before/after another sheet."""
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")
        sheet = self._wb.sheets[sheet_name]
        if before:
            sheet.api.move(before=self._wb.sheets[before].api)
        elif after:
            sheet.api.move(after=self._wb.sheets[after].api)
        logger.info("Moved sheet '%s'", sheet_name)

    def copy_sheet(self, sheet_name: str, new_name: Optional[str] = None) -> None:
        """Copy a sheet within the same workbook."""
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")
        sheet = self._wb.sheets[sheet_name]
        sheet.api.copy(after=sheet.api)
        if new_name:
            self._wb.sheets[sheet.name + " (2)"].name = new_name
        logger.info("Copied sheet '%s'", sheet_name)

    # ── Goal Seek ──

    def goal_seek(
        self,
        target_cell: str,
        goal_value: float,
        changing_cell: str,
        sheet: Optional[str] = None,
    ) -> None:
        """Run Goal Seek on a target cell."""
        self._require_xlwings()
        ws = self._ws(sheet)
        target_rng = ws.range(target_cell)
        changing_rng = ws.range(changing_cell)
        target_rng.api.goal_seek(goal=goal_value, changing_cell=changing_rng.api)
        logger.info("Goal Seek: %s → %s (changing %s)", target_cell, goal_value, changing_cell)

    # ── Sort ──

    def sort_range(
        self, range_str: str, keys: list[dict], sheet: Optional[str] = None,
    ) -> None:
        """Sort a range using xlwings API."""
        self._require_xlwings()
        ws = self._ws(sheet)
        rng = ws.range(range_str)
        if keys:
            key1_ref = ws.range(keys[0].get("cell", "A1"))
            order = 1 if keys[0].get("ascending", True) else 2
            rng.api.sort(key1=key1_ref.api, order1=order, header=1)
        else:
            rng.api.sort(key1=rng.api, order1=1, header=1)
        logger.info("Sorted range %s", range_str)

    # ── Tab Color ──

    def set_tab_color(self, sheet_name: str, color: str) -> None:
        """Set the tab color of a sheet (hex string like '#FF0000')."""
        self._require_xlwings()
        if not self._wb:
            raise RuntimeError("Not connected")
        sheet = self._wb.sheets[sheet_name]
        # Convert hex to RGB int for macOS Excel COM
        hex_clean = color.lstrip("#")
        r, g, b = int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16)
        sheet.api.tab.color.set(r + g * 256 + b * 65536)
        logger.info("Set tab color of '%s' to %s", sheet_name, color)

    # ── Table Total Row ──

    def set_table_total_row(
        self, table_name: str, show: bool = True, sheet: Optional[str] = None,
    ) -> None:
        """Show or hide the Total Row on a table."""
        self._require_xlwings()
        ws = self._ws(sheet or "Sheet1")
        for table in ws.api.list_objects():
            if table.name() == table_name:
                table.show_totals.set(show)
                logger.info("Table '%s' total row: %s", table_name, show)
                return
        raise ValueError(f"Table '{table_name}' not found")

    # ── Generic Formatting ──

    def apply_formatting(
        self, ref: str, params: dict, sheet: Optional[str] = None,
    ) -> None:
        """Apply formatting from task params to a range via xlwings."""
        self._require_xlwings()
        ws = self._ws(sheet)
        rng = ws.range(ref)
        if "bold" in params:
            rng.font.bold = params["bold"]
        if "italic" in params:
            rng.font.italic = params["italic"]
        if "font_size" in params:
            rng.font.size = params["font_size"]
        if "font_color" in params:
            rng.font.color = params["font_color"]
        if "fill_color" in params:
            rng.color = params["fill_color"]
        if "number_format" in params:
            rng.number_format = params["number_format"]
        logger.info("Applied formatting to %s", ref)
