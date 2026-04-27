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

SENTINEL_SHEET = "__excel_engine_sentinel"
POLL_INTERVAL = 0.5


class VBALayer:
    """Layer 5 — VBA code execution via clipboard injection into VBE."""

    def __init__(
        self,
        execution_timeout: float = 45.0,
        split_threshold: int = 50,
        poll_interval: float = POLL_INTERVAL,
        sentinel_checker=None,
    ) -> None:
        self.execution_timeout = execution_timeout
        self.split_threshold = split_threshold
        self.poll_interval = poll_interval
        self._sentinel_checker = sentinel_checker

    @staticmethod
    def _escape_vba(s: str) -> str:
        """Escape a string for VBA string literals (double quotes → double-double quotes)."""
        return s.replace('"', '""')

    @staticmethod
    def inject_sentinel(code: str) -> str:
        """Append VBA lines that create a sentinel sheet and write "DONE" to A1.

        The sentinel is placed just before ``End Sub`` in the last subroutine so
        it fires after all real work completes.
        """
        sentinel_lines = (
            f'\n    On Error Resume Next\n'
            f'    Dim __sentWs As Worksheet\n'
            f'    Set __sentWs = Nothing\n'
            f'    Set __sentWs = ActiveWorkbook.Sheets("{SENTINEL_SHEET}")\n'
            f'    If __sentWs Is Nothing Then\n'
            f'        Set __sentWs = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))\n'
            f'        __sentWs.Name = "{SENTINEL_SHEET}"\n'
            f'    End If\n'
            f'    __sentWs.Range("A1").Value = "DONE"\n'
            f'    Set __sentWs = Nothing\n'
            f'    On Error GoTo 0\n'
        )
        # Insert before the last 'End Sub'
        last_end = code.rfind("End Sub")
        if last_end == -1:
            return code + sentinel_lines + "\nEnd Sub\n"
        return code[:last_end] + sentinel_lines + code[last_end:]

    def _check_sentinel(self) -> bool:
        """Check whether the sentinel cell has been set to DONE.

        Uses the injected ``sentinel_checker`` callable if provided (for
        testing), otherwise tries xlwings/AppleScript.
        """
        if self._sentinel_checker is not None:
            return self._sentinel_checker()
        try:
            import xlwings as xw
            wb = xw.books.active
            ws = wb.sheets[SENTINEL_SHEET]
            return ws.range("A1").value == "DONE"
        except Exception:
            return False

    def _cleanup_sentinel(self) -> None:
        """Delete the sentinel sheet after execution completes."""
        try:
            import xlwings as xw
            wb = xw.books.active
            app = wb.app
            app.display_alerts = False
            wb.sheets[SENTINEL_SHEET].delete()
            app.display_alerts = True
        except Exception:
            pass

    # ── Core Execution ──

    def execute_vba(self, code: str) -> None:
        """
        Execute VBA code by injecting it into the VBE via clipboard.
        Steps: inject sentinel → pbcopy → Option+F11 → Cmd+A → Cmd+V → F5 → poll for completion.
        """
        code = self.inject_sentinel(code)

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

        # Poll for sentinel "DONE" instead of fixed sleep
        deadline = time.time() + self.execution_timeout
        completed = False
        while time.time() < deadline:
            try:
                if self._check_sentinel():
                    completed = True
                    logger.info("VBA sentinel detected — macro completed early")
                    break
            except Exception:
                pass
            time.sleep(self.poll_interval)

        if not completed:
            logger.warning("VBA execution timed out after %.1fs", self.execution_timeout)

        # Clean up sentinel sheet
        self._cleanup_sentinel()

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

    def manage_slicer(
        self,
        field: str = "Product",
        target_range: Optional[str] = None,
        selection: Optional[str] = None,
        dest_sheet: Optional[str] = None,
    ) -> None:
        """Create/configure a slicer for the first available PivotTable."""
        safe_field = self._escape_vba(field or "Product")
        safe_selection = self._escape_vba(selection) if selection else ""
        safe_dest = self._escape_vba(dest_sheet) if dest_sheet else ""

        if dest_sheet:
            sheet_block = textwrap.dedent(
                f"""\
                Set ws = Nothing
                On Error Resume Next
                Set ws = Sheets("{safe_dest}")
                On Error GoTo 0
                """
            )
        else:
            sheet_block = textwrap.dedent(
                """\
                Set ws = Nothing
                Dim candidateWs As Worksheet
                For Each candidateWs In ActiveWorkbook.Worksheets
                    If candidateWs.PivotTables.Count > 0 Then
                        Set ws = candidateWs
                        Exit For
                    End If
                Next candidateWs
                """
            )

        position_block = ""
        if target_range and ":" in target_range:
            start_cell, end_cell = target_range.split(":", 1)
            position_block = textwrap.dedent(
                f"""\
                    Dim r1 As Range
                    Dim r2 As Range
                    Set r1 = ws.Range("{self._escape_vba(start_cell)}")
                    Set r2 = ws.Range("{self._escape_vba(end_cell)}")
                    sl.Left = r1.Left
                    sl.Top = r1.Top
                    sl.Width = (r2.Left + r2.Width) - r1.Left
                    sl.Height = (r2.Top + r2.Height) - r1.Top
                """
            )

        selection_block = ""
        if safe_selection:
            selection_block = textwrap.dedent(
                f"""\
                    Dim si As SlicerItem
                    sc.ClearManualFilter
                    For Each si In sc.SlicerItems
                        si.Selected = (LCase(si.Name) = LCase("{safe_selection}"))
                    Next si
                """
            )

        vba = textwrap.dedent(
            f"""\
            Sub ManageSlicer()
                Dim ws As Worksheet
                {sheet_block}
                If ws Is Nothing Then
                    Err.Raise vbObjectError + 510, "ManageSlicer", "No worksheet with PivotTable found."
                End If
                If ws.PivotTables.Count = 0 Then
                    Err.Raise vbObjectError + 511, "ManageSlicer", "PivotTable required before slicer creation."
                End If

                Dim pt As PivotTable
                Set pt = ws.PivotTables(1)

                Dim sc As SlicerCache
                Set sc = Nothing

                Dim existingCache As SlicerCache
                For Each existingCache In ActiveWorkbook.SlicerCaches
                    If LCase(existingCache.SourceName) = LCase("{safe_field}") Then
                        Set sc = existingCache
                        Exit For
                    End If
                Next existingCache

                If sc Is Nothing Then
                    On Error Resume Next
                    Set sc = ActiveWorkbook.SlicerCaches.Add2(pt, "{safe_field}")
                    If sc Is Nothing Then
                        Set sc = ActiveWorkbook.SlicerCaches.Add(pt, "{safe_field}")
                    End If
                    On Error GoTo 0
                End If

                If sc Is Nothing Then
                    Err.Raise vbObjectError + 512, "ManageSlicer", "Unable to create slicer cache for field '{safe_field}'."
                End If

                Dim sl As Slicer
                If sc.Slicers.Count > 0 Then
                    Set sl = sc.Slicers(1)
                Else
                    Set sl = sc.Slicers.Add(ws, , "{safe_field} Slicer", "{safe_field}", 220, 80, 180, 220)
                End If

                {position_block}
                {selection_block}

                Set sl = Nothing
                Set sc = Nothing
                Set pt = Nothing
                Set ws = Nothing
            End Sub

            Sub Main()
                ManageSlicer
            End Sub
            """
        )

        self.execute_vba(vba)
        logger.info(
            "Managed slicer field=%s target_range=%s selection=%s",
            field,
            target_range,
            selection,
        )

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
