"""
Layer 5: VBA via VBE — Clipboard injection for complex automation.

Handles PivotTables, PivotCharts, and complex operations that require VBA.

METHOD:
  1. Write VBA code to clipboard via pbcopy
  2. Open VBE: Option+F11
  3. Select all: Ctrl+A (or Cmd+A on Mac VBE)
  4. Paste: Cmd+V
  5. Run: F5
  6. Close VBE: Cmd+Q or Option+F11 again

KEY RULES:
  - Always include CleanExistingObjects sub
  - Split into small subs to prevent OOM Error 7
  - Set obj = Nothing after every PivotCache/PivotTable creation
  - VBA split threshold default: 50 lines per sub
"""

from __future__ import annotations

import logging
import time
import textwrap
from pathlib import Path
from typing import Optional

from excel_engine.utils.mac_utils import MacUtils

logger = logging.getLogger(__name__)


class VBALayer:
    """Layer 5 — VBA code execution via clipboard injection into VBE."""

    def __init__(
        self,
        execution_timeout: float = 45.0,
        split_threshold: int = 50,
    ) -> None:
        self.execution_timeout = execution_timeout
        self.split_threshold = split_threshold

    @staticmethod
    def _escape_vba(s: str) -> str:
        """Escape a string for VBA string literals (double quotes → double-double quotes)."""
        return s.replace('"', '""')

    # ── Core Execution ──

    def execute_vba(self, code: str) -> None:
        """
        Execute VBA code by injecting it into the VBE via clipboard.
        Steps: pbcopy → Option+F11 → Cmd+A → Cmd+V → F5
        """
        MacUtils.activate_excel()
        time.sleep(0.5)

        MacUtils.clipboard_copy(code)
        logger.debug("VBA code copied to clipboard (%d chars)", len(code))

        # Open VBE: Option + F11
        MacUtils.run_applescript(
            'tell application "System Events"\n'
            '    key code 98 using option down\n'  # F11 = key code 98
            'end tell'
        )
        time.sleep(2)

        # Select all existing code: Cmd+A
        MacUtils.run_applescript(
            'tell application "System Events"\n'
            '    keystroke "a" using command down\n'
            'end tell'
        )
        time.sleep(0.3)

        # Paste: Cmd+V
        MacUtils.run_applescript(
            'tell application "System Events"\n'
            '    keystroke "v" using command down\n'
            'end tell'
        )
        time.sleep(0.5)

        # Run: F5
        MacUtils.run_applescript(
            'tell application "System Events"\n'
            '    key code 96\n'  # F5 = key code 96
            'end tell'
        )
        # TODO: Replace fixed 45s sleep with polling-based VBA completion detection
        # e.g., check for a sentinel cell value that VBA sets when done
        time.sleep(self.execution_timeout)

        # Close VBE: Option + F11
        MacUtils.run_applescript(
            'tell application "System Events"\n'
            '    key code 98 using option down\n'
            'end tell'
        )
        time.sleep(0.5)
        logger.info("VBA code executed")

    def execute_vba_file(self, path: Path) -> None:
        """Execute VBA code from a file."""
        code = Path(path).read_text(encoding="utf-8")
        self.execute_vba(code)

    # ── PivotTable Generation ──

    def create_pivot_table(
        self,
        source_sheet: str,
        source_range: str,
        dest_sheet: str,
        dest_cell: str = "A3",
        table_name: str = "PivotTable1",
        row_fields: Optional[list[str]] = None,
        column_fields: Optional[list[str]] = None,
        value_fields: Optional[list[dict]] = None,
        filter_fields: Optional[list[str]] = None,
    ) -> None:
        """
        Generate and execute VBA to create a PivotTable.

        value_fields: list of dicts with 'name', 'function' (xlSum, xlCount, etc.),
                      and optional 'number_format'.
        """
        row_fields = row_fields or []
        column_fields = column_fields or []
        value_fields = value_fields or []
        filter_fields = filter_fields or []

        # Build the VBA code
        row_code = self._pivot_field_code(row_fields, "xlRowField")
        col_code = self._pivot_field_code(column_fields, "xlColumnField")
        filter_code = self._pivot_field_code(filter_fields, "xlPageField")
        value_code = self._pivot_value_code(value_fields)

        safe_source_sheet = self._escape_vba(source_sheet)
        safe_source_range = self._escape_vba(source_range)
        safe_dest_sheet = self._escape_vba(dest_sheet)
        safe_dest_cell = self._escape_vba(dest_cell)
        safe_table_name = self._escape_vba(table_name)

        vba = textwrap.dedent(f"""\
            Sub CleanExistingObjects()
                Dim ws As Worksheet
                Set ws = Sheets("{safe_dest_sheet}")
                Dim pt As PivotTable
                For Each pt In ws.PivotTables
                    If pt.Name = "{safe_table_name}" Then
                        pt.TableRange2.Clear
                    End If
                Next pt
                Set ws = Nothing
            End Sub

            Sub CreatePivotTable()
                CleanExistingObjects

                Dim srcWs As Worksheet
                Set srcWs = Sheets("{safe_source_sheet}")

                Dim srcRange As Range
                Set srcRange = srcWs.Range("{safe_source_range}")

                Dim pc As PivotCache
                Set pc = ActiveWorkbook.PivotCaches.Create( _
                    SourceType:=xlDatabase, _
                    SourceData:=srcRange)

                Dim destWs As Worksheet
                Set destWs = Sheets("{safe_dest_sheet}")

                Dim pt As PivotTable
                Set pt = pc.CreatePivotTable( _
                    TableDestination:=destWs.Range("{safe_dest_cell}"), _
                    TableName:="{safe_table_name}")

                {row_code}
                {col_code}
                {filter_code}
                {value_code}

                ' Clean up to avoid OOM Error 7
                Set pt = Nothing
                Set pc = Nothing
                Set srcRange = Nothing
                Set srcWs = Nothing
                Set destWs = Nothing
            End Sub

            Sub Main()
                CreatePivotTable
            End Sub
        """)

        self.execute_vba(vba)
        logger.info("Created PivotTable '%s' on '%s'", table_name, dest_sheet)

    def create_pivot_chart(
        self,
        pivot_table_name: str,
        chart_type: str = "xlColumnClustered",
        chart_title: str = "",
        dest_sheet: Optional[str] = None,
    ) -> None:
        """Generate and execute VBA to create a PivotChart from an existing PivotTable."""
        safe_pivot = self._escape_vba(pivot_table_name)
        safe_title = self._escape_vba(chart_title) if chart_title else ""
        safe_dest = self._escape_vba(dest_sheet) if dest_sheet else ""
        dest_line = f'Set ws = Sheets("{safe_dest}")' if dest_sheet else "Set ws = ActiveSheet"
        title_line = f'ch.Chart.HasTitle = True\n                ch.Chart.ChartTitle.Text = "{safe_title}"' if chart_title else ""

        vba = textwrap.dedent(f"""\
            Sub CreatePivotChart()
                Dim ws As Worksheet
                {dest_line}

                Dim pt As PivotTable
                Set pt = ws.PivotTables("{safe_pivot}")

                Dim ch As ChartObject
                Set ch = ws.ChartObjects.Add( _
                    Left:=300, Top:=50, Width:=450, Height:=300)

                ch.Chart.SetSourceData pt.TableRange1
                ch.Chart.ChartType = {chart_type}
                {title_line}

                Set ch = Nothing
                Set pt = Nothing
                Set ws = Nothing
            End Sub

            Sub Main()
                CreatePivotChart
            End Sub
        """)

        self.execute_vba(vba)
        logger.info("Created PivotChart for '%s'", pivot_table_name)

    # ── Helpers ──

    @staticmethod
    def _pivot_field_code(fields: list[str], orientation: str) -> str:
        """Generate VBA to add row/column/filter fields."""
        lines = []
        for f in fields:
            safe_f = f.replace('"', '""')
            lines.append(
                f'With pt.PivotFields("{safe_f}")\n'
                f'    .Orientation = {orientation}\n'
                f'End With'
            )
        return "\n                ".join(lines)

    @staticmethod
    def _pivot_value_code(value_fields: list[dict]) -> str:
        """Generate VBA to add value fields with functions and number formats."""
        lines = []
        for vf in value_fields:
            name = vf["name"].replace('"', '""')
            func = vf.get("function", "xlSum")
            nf = vf.get("number_format", "").replace('"', '""')
            block = (
                f'With pt.PivotFields("{name}")\n'
                f'    .Orientation = xlDataField\n'
                f'    .Function = {func}\n'
            )
            if nf:
                block += f'    .NumberFormat = "{nf}"\n'
            block += "End With"
            lines.append(block)
        return "\n                ".join(lines)

    def generate_clean_sub(self, object_type: str = "PivotTable") -> str:
        """Generate a CleanExistingObjects VBA sub."""
        return textwrap.dedent(f"""\
            Sub CleanExistingObjects()
                On Error Resume Next
                Dim obj As {object_type}
                ' Clean up existing objects
                On Error GoTo 0
            End Sub
        """)
