"""
Layer 4: System Events — Ribbon and dialog UI automation.

Handles ribbon navigation, slicer creation, slicer tab configuration,
histogram chart insertion, Format Data Series pane, subtotals dialog,
and table styles gallery.

All operations are done via AppleScript talking to System Events
to drive the Excel UI.

KEY RULES:
  - Ribbon: tab group → radio buttons → scroll area → groups → buttons
  - Slicer: Insert tab → Slicer button → checkbox dialog → OK
  - Histogram: Insert > Chart > Histogram menu path
  - Always add delays between UI actions for reliability
"""

from __future__ import annotations

import logging
import time
from typing import Optional

from excel_engine.utils.mac_utils import MacUtils

logger = logging.getLogger(__name__)


class SystemEventsLayer:
    """Layer 4 — ribbon and dialog UI automation via System Events."""

    def __init__(self, delay: float = 0.5) -> None:
        self.delay = delay

    def _run(self, script: str) -> str:
        """Execute a System Events AppleScript."""
        return MacUtils.run_applescript(script, timeout=60.0)

    def _activate_excel(self) -> None:
        """Bring Excel to the front."""
        MacUtils.activate_excel()
        time.sleep(self.delay)

    def _click_ribbon_tab(self, tab_name: str) -> str:
        """
        Generate AppleScript to click a ribbon tab.
        Returns the System Events script fragment.
        """
        return (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click the ribbon tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "{tab_name}"\n'
            f'        end tell\n'
            f'        delay {self.delay}\n'
            f'    end tell\n'
            f'end tell'
        )

    # ── Slicer Creation ──

    def insert_slicer(self, field_names: list[str]) -> None:
        """
        Insert a slicer using the ribbon UI.
        1. Click Insert tab
        2. Click Slicer button
        3. Check the specified field checkboxes in the dialog
        4. Click OK
        """
        self._activate_excel()

        checkboxes_script = "\n".join(
            f'                        click checkbox "{name}" of scroll area 1 of sheet 1'
            for name in field_names
        )

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click Insert tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "Insert"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        -- Click Slicer button in the Filters group\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "Filters"\n'
            f'                click button "Slicer"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'        delay 1.5\n'
            f'        -- Check fields in the slicer dialog\n'
            f'        tell window 1\n'
            f'{checkboxes_script}\n'
            f'            delay {self.delay}\n'
            f'            click button "OK"\n'
            f'        end tell\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        time.sleep(1)
        logger.info("Inserted slicer for fields: %s", field_names)

    def configure_slicer(
        self,
        columns: Optional[int] = None,
        caption: Optional[str] = None,
        width: Optional[float] = None,
        height: Optional[float] = None,
    ) -> None:
        """
        Configure a selected slicer via the Slicer tab in the ribbon.
        Uses AXIncrementor elements for numeric values.
        """
        self._activate_excel()

        commands = []
        commands.append(
            f'tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'    click radio button "Slicer"\n'
            f'end tell\n'
            f'delay {self.delay * 2}'
        )

        if columns is not None:
            commands.append(
                f'tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
                f'    tell group "Buttons"\n'
                f'        set value of text field "Columns:" to "{columns}"\n'
                f'    end tell\n'
                f'end tell'
            )

        if caption is not None:
            commands.append(
                f'tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
                f'    tell group "Slicer"\n'
                f'        set value of text field "Caption:" to "{caption}"\n'
                f'    end tell\n'
                f'end tell'
            )

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        {"".join(commands)}\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        logger.info("Configured slicer: cols=%s, caption=%s", columns, caption)

    # ── Histogram Chart ──

    def insert_histogram(self, data_range: Optional[str] = None) -> None:
        """
        Insert a histogram chart via the ribbon UI.
        Histogram is a cx:chart type that cannot be created with openpyxl.
        Path: Insert tab → Charts group → Histogram button.
        """
        self._activate_excel()

        if data_range:
            MacUtils.run_applescript(
                f'tell application "Microsoft Excel"\n'
                f'    select range "{data_range}" of active sheet\n'
                f'end tell'
            )
            time.sleep(self.delay)

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click Insert tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "Insert"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        -- Click the statistical chart button (Histogram)\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "Charts"\n'
            f'                click menu button "Insert Statistic Chart"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'        delay {self.delay}\n'
            f'        -- Click Histogram in the dropdown\n'
            f'        click menu item "Histogram" of menu 1 of menu button "Insert Statistic Chart" of group "Charts" of scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        time.sleep(2)
        logger.info("Inserted histogram chart")

    def configure_histogram_bins(
        self,
        bin_width: Optional[float] = None,
        overflow: Optional[float] = None,
        underflow: Optional[float] = None,
    ) -> None:
        """
        Configure histogram bins via the Format Data Series pane.
        Must double-click the histogram bars first to open the pane.
        """
        self._activate_excel()

        commands = []
        if bin_width is not None:
            commands.append(
                f'set value of text field "Bin width" of group 1 of scroll area 1 of window "Format Data Series" to "{bin_width}"'
            )
        if overflow is not None:
            commands.append(
                f'click checkbox "Overflow bin" of group 1 of scroll area 1 of window "Format Data Series"\n'
                f'set value of text field "Overflow bin" of group 1 of scroll area 1 of window "Format Data Series" to "{overflow}"'
            )
        if underflow is not None:
            commands.append(
                f'click checkbox "Underflow bin" of group 1 of scroll area 1 of window "Format Data Series"\n'
                f'set value of text field "Underflow bin" of group 1 of scroll area 1 of window "Format Data Series" to "{underflow}"'
            )

        commands_str = "\n        ".join(commands)
        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        {commands_str}\n'
            f'    end tell\n'
            f'end tell'
        )
        if commands:
            self._run(script)
            logger.info(
                "Configured histogram: bin_width=%s, overflow=%s, underflow=%s",
                bin_width, overflow, underflow,
            )

    # ── Scatter Chart ──

    def insert_scatter_chart(self, data_range: Optional[str] = None) -> None:
        """
        Insert a scatter (XY) chart via the ribbon UI.
        Path: Insert tab → Charts group → Insert Scatter (X, Y) → Scatter
        """
        self._activate_excel()

        if data_range:
            MacUtils.run_applescript(
                f'tell application "Microsoft Excel"\n'
                f'    select range "{data_range}" of active sheet\n'
                f'end tell'
            )
            time.sleep(self.delay)

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click Insert tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "Insert"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        -- Click the scatter chart menu button\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "Charts"\n'
            f'                click menu button "Insert Scatter (X, Y) or Bubble Chart"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'        delay {self.delay}\n'
            f'        -- Click Scatter in the dropdown\n'
            f'        click menu item "Scatter" of menu 1 of menu button '
            f'"Insert Scatter (X, Y) or Bubble Chart" of group "Charts" '
            f'of scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        time.sleep(2)
        logger.info("Inserted scatter chart via ribbon")

    # ── Combo Chart ──

    def insert_combo_chart(
        self,
        data_range: Optional[str] = None,
        secondary_axis: bool = True,
    ) -> None:
        """
        Insert a combo chart via the ribbon UI.
        Path: Insert tab → Charts group → Insert Combo Chart → Clustered Column – Line on Secondary Axis
        """
        self._activate_excel()

        if data_range:
            MacUtils.run_applescript(
                f'tell application "Microsoft Excel"\n'
                f'    select range "{data_range}" of active sheet\n'
                f'end tell'
            )
            time.sleep(self.delay)

        chart_item = (
            "Clustered Column - Line on Secondary Axis"
            if secondary_axis
            else "Clustered Column - Line"
        )

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click Insert tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "Insert"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        -- Click the combo chart menu button\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "Charts"\n'
            f'                click menu button "Insert Combo Chart"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'        delay {self.delay}\n'
            f'        -- Click the desired combo type\n'
            f'        click menu item "{chart_item}" of menu 1 of menu button '
            f'"Insert Combo Chart" of group "Charts" '
            f'of scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        time.sleep(2)
        logger.info("Inserted combo chart (secondary_axis=%s) via ribbon", secondary_axis)

    # ── Sparklines ──

    def insert_sparkline(
        self,
        data_range: str,
        location_range: str,
        sparkline_type: str = "Line",
    ) -> None:
        """
        Insert sparklines via the ribbon UI.
        Path: Insert tab → Sparklines group → Line/Column/Win-Loss button

        data_range:     Source data range (e.g., "B2:M2")
        location_range: Where to place the sparkline (e.g., "N2")
        sparkline_type: "Line", "Column", or "Win/Loss"
        """
        self._activate_excel()

        # First select the location cell(s)
        MacUtils.run_applescript(
            f'tell application "Microsoft Excel"\n'
            f'    select range "{location_range}" of active sheet\n'
            f'end tell'
        )
        time.sleep(self.delay)

        # Map type to ribbon button name
        button_name = {
            "Line": "Line",
            "Column": "Column",
            "Win/Loss": "Win/Loss",
        }.get(sparkline_type, "Line")

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click Insert tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "Insert"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        -- Click sparkline type button in Sparklines group\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "Sparklines"\n'
            f'                click button "{button_name}"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'        delay 1.5\n'
            f'        -- The Create Sparklines dialog opens; set data range\n'
            f'        tell window 1\n'
            f'            set value of text field 1 to "{data_range}"\n'
            f'            delay {self.delay}\n'
            f'            click button "OK"\n'
            f'        end tell\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        time.sleep(1.5)
        logger.info(
            "Inserted %s sparkline: data=%s, location=%s",
            sparkline_type, data_range, location_range,
        )

    # ── Table Style Gallery ──

    def apply_table_style_via_ribbon(self, style_name: str) -> None:
        """
        Apply a table style via the Table Design tab's style gallery.
        The table must be selected first.
        """
        self._activate_excel()

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click Table Design tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "Table Design"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        -- Look for the style in the gallery\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "Table Styles"\n'
            f'                click button "{style_name}"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        logger.info("Applied table style '%s' via ribbon", style_name)

    # ── Subtotals Dialog ──

    def open_subtotals_dialog(self) -> None:
        """
        Open the Subtotals dialog via Data tab → Subtotal button.
        """
        self._activate_excel()

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        -- Click Data tab\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "Data"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        -- Click Subtotal button\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "Outline"\n'
            f'                click button "Subtotal"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        time.sleep(1)
        logger.info("Opened Subtotals dialog")

    # ── Generic Ribbon Helpers ──

    def click_ribbon_button(self, tab: str, group: str, button: str) -> None:
        """Click any ribbon button by tab name, group name, and button name."""
        self._activate_excel()

        script = (
            f'tell application "System Events"\n'
            f'    tell process "Microsoft Excel"\n'
            f'        set frontmost to true\n'
            f'        delay {self.delay}\n'
            f'        tell tab group 1 of group 1 of tool bar 1 of window 1\n'
            f'            click radio button "{tab}"\n'
            f'        end tell\n'
            f'        delay {self.delay * 2}\n'
            f'        tell scroll area 1 of group 1 of tool bar 1 of window 1\n'
            f'            tell group "{group}"\n'
            f'                click button "{button}"\n'
            f'            end tell\n'
            f'        end tell\n'
            f'    end tell\n'
            f'end tell'
        )
        self._run(script)
        logger.info("Clicked ribbon: %s → %s → %s", tab, group, button)

    def press_key(self, key_code: int, modifiers: Optional[list[str]] = None) -> None:
        """Press a key with optional modifiers (command, shift, option, control)."""
        mod_str = ""
        if modifiers:
            mod_str = " using {" + ", ".join(f"{m} down" for m in modifiers) + "}"
        script = (
            f'tell application "System Events"\n'
            f'    key code {key_code}{mod_str}\n'
            f'end tell'
        )
        self._run(script)

    def type_text(self, text: str) -> None:
        """Type text via System Events keystroke."""
        safe_text = text.replace("\\", "\\\\").replace('"', '\\"')
        script = (
            f'tell application "System Events"\n'
            f'    keystroke "{safe_text}"\n'
            f'end tell'
        )
        self._run(script)
