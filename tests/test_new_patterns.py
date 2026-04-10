"""
Tests for newly added regex patterns in the TaskExtractor.

Covers P0 (hide/unhide, format-as-bold, distant-verb-function, filter-data, autofit),
P1 (text-wrapping, remove-duplicates, chart-modifications, headers/footers, cell-styles),
P2 (change-chart-type, if-format-conditional, enter-SUM-formula, bare-numbers, short-lines).
"""

import pytest

from excel_engine.parsers.task_extractor import TaskExtractor
from excel_engine.config import TaskType


@pytest.fixture
def ext():
    return TaskExtractor()


# ── P0-1: Hide/Unhide rows & columns ──

class TestHideUnhide:
    def test_hide_columns(self, ext):
        tasks = ext.extract("Hide columns B and C")
        assert any(t.task_type == TaskType.COLUMN_WIDTH for t in tasks)

    def test_unhide_column(self, ext):
        tasks = ext.extract("Unhide column D in the worksheet")
        assert any(t.task_type == TaskType.COLUMN_WIDTH for t in tasks)

    def test_show_columns(self, ext):
        tasks = ext.extract("Show the columns that were hidden")
        assert any(t.task_type == TaskType.COLUMN_WIDTH for t in tasks)

    def test_hide_rows(self, ext):
        tasks = ext.extract("Hide rows 5 through 10")
        assert any(t.task_type == TaskType.ROW_HEIGHT for t in tasks)

    def test_unhide_rows(self, ext):
        tasks = ext.extract("Unhide the rows in the range")
        assert any(t.task_type == TaskType.ROW_HEIGHT for t in tasks)


# ── P0-2: "Format X as Bold" ──

class TestFormatAsBold:
    def test_format_as_bold(self, ext):
        tasks = ext.extract("Format cell A1 as bold")
        assert any(t.task_type == TaskType.FONT for t in tasks)

    def test_make_as_italic(self, ext):
        tasks = ext.extract("Make the header as italic")
        assert any(t.task_type == TaskType.FONT for t in tasks)

    def test_bare_as_underline(self, ext):
        tasks = ext.extract("Set the title as underline in the header row")
        assert any(t.task_type == TaskType.FONT for t in tasks)

    def test_as_strikethrough(self, ext):
        tasks = ext.extract("Mark completed items as strikethrough")
        assert any(t.task_type == TaskType.FONT for t in tasks)


# ── P0-3: "using the X function" ──

class TestDistantVerbFunction:
    def test_using_the_average_function(self, ext):
        tasks = ext.extract("Calculate the result using the AVERAGE function")
        assert any(t.task_type == TaskType.FORMULA for t in tasks)

    def test_today_function(self, ext):
        tasks = ext.extract("Insert the TODAY function in cell A1")
        assert any(t.task_type == TaskType.FORMULA for t in tasks)

    def test_now_function(self, ext):
        tasks = ext.extract("Use the NOW function to display the current date and time")
        assert any(t.task_type == TaskType.FORMULA for t in tasks)


# ── P0-4: "Filter the data" → AUTOFILTER ──

class TestFilterData:
    def test_filter_the_data(self, ext):
        tasks = ext.extract("Filter the data to show only sales above 1000")
        assert any(t.task_type == TaskType.AUTOFILTER for t in tasks)

    def test_filter_records(self, ext):
        tasks = ext.extract("Filter the records by department")
        assert any(t.task_type == TaskType.AUTOFILTER for t in tasks)

    def test_filter_table(self, ext):
        tasks = ext.extract("Filter the table to show active employees")
        assert any(t.task_type == TaskType.AUTOFILTER for t in tasks)


# ── P0-5: AutoFit standalone ──

class TestAutoFit:
    def test_autofit_columns(self, ext):
        tasks = ext.extract("AutoFit all the columns in the worksheet")
        assert any(t.task_type == TaskType.COLUMN_WIDTH for t in tasks)

    def test_autofit_bare(self, ext):
        tasks = ext.extract("Select the range and autofit")
        assert any(t.task_type == TaskType.COLUMN_WIDTH for t in tasks)

    def test_auto_fit_column(self, ext):
        tasks = ext.extract("Auto fit column A to the contents")
        assert any(t.task_type == TaskType.COLUMN_WIDTH for t in tasks)


# ── P1-6: Text wrapping ──

class TestTextWrapping:
    def test_text_wrapping(self, ext):
        tasks = ext.extract("Apply text wrapping to the header cells")
        assert any(t.task_type == TaskType.ALIGNMENT for t in tasks)

    def test_enable_wrapping(self, ext):
        tasks = ext.extract("Enable wrapping for cells A1:A10")
        assert any(t.task_type == TaskType.ALIGNMENT for t in tasks)

    def test_turn_on_text_wrapping(self, ext):
        tasks = ext.extract("Turn on text wrapping in the description column")
        assert any(t.task_type == TaskType.ALIGNMENT for t in tasks)

    def test_disable_wrapping(self, ext):
        tasks = ext.extract("Disable text wrapping in cells B1:B5")
        assert any(t.task_type == TaskType.ALIGNMENT for t in tasks)


# ── P1-7: Remove duplicates ──

class TestRemoveDuplicates:
    def test_remove_duplicates(self, ext):
        tasks = ext.extract("Remove duplicates from the list of names")
        assert any(t.task_type == TaskType.DATA_VALIDATION for t in tasks)

    def test_delete_the_duplicates(self, ext):
        tasks = ext.extract("Delete the duplicates in column A")
        assert any(t.task_type == TaskType.DATA_VALIDATION for t in tasks)

    def test_eliminate_duplicate(self, ext):
        tasks = ext.extract("Eliminate duplicate entries from the range")
        assert any(t.task_type == TaskType.DATA_VALIDATION for t in tasks)


# ── P1-8: Chart modifications ──

class TestChartModifications:
    def test_add_chart_title(self, ext):
        tasks = ext.extract("Add a chart title to the bar chart")
        assert any(t.task_type == TaskType.CHART_BAR for t in tasks)

    def test_change_trendline(self, ext):
        tasks = ext.extract("Add a trendline to the chart")
        assert any(t.task_type == TaskType.CHART_BAR for t in tasks)

    def test_format_data_labels(self, ext):
        tasks = ext.extract("Format data labels on the chart")
        assert any(t.task_type == TaskType.CHART_BAR for t in tasks)

    def test_remove_legend(self, ext):
        tasks = ext.extract("Remove the legend from the chart")
        assert any(t.task_type == TaskType.CHART_BAR for t in tasks)


# ── P1-9: Headers/footers for printing ──

class TestPrintHeadersFooters:
    def test_add_header(self, ext):
        tasks = ext.extract("Add a header to the printed page")
        assert any(t.task_type == TaskType.PRINT_SETTINGS for t in tasks)

    def test_insert_footer(self, ext):
        tasks = ext.extract("Insert a footer with page numbers")
        assert any(t.task_type == TaskType.PRINT_SETTINGS for t in tasks)

    def test_create_headers(self, ext):
        tasks = ext.extract("Create headers for the printed report")
        assert any(t.task_type == TaskType.PRINT_SETTINGS for t in tasks)


# ── P1-10: Cell style names ──

class TestCellStyleNames:
    def test_apply_heading_style(self, ext):
        tasks = ext.extract("Apply the Heading 1 style to cell A1")
        assert any(t.task_type == TaskType.FONT for t in tasks)

    def test_use_title_style(self, ext):
        tasks = ext.extract("Use the Title style for the header row")
        assert any(t.task_type == TaskType.FONT for t in tasks)


# ── P2-11: Change chart type ──

class TestChangeChartType:
    def test_change_chart_type_to_line(self, ext):
        tasks = ext.extract("Change the chart type to line")
        assert any(t.task_type == TaskType.CHART_LINE for t in tasks)

    def test_switch_chart_to_pie(self, ext):
        tasks = ext.extract("Switch the chart type to pie")
        assert any(t.task_type == TaskType.CHART_PIE for t in tasks)

    def test_change_chart_to_bar(self, ext):
        tasks = ext.extract("Change the chart to bar")
        assert any(t.task_type == TaskType.CHART_BAR for t in tasks)


# ── P2-12: "If total exceeds 1000, format in red" ──

class TestIfFormatConditional:
    def test_if_total_exceeds_format_red(self, ext):
        tasks = ext.extract("If the total exceeds 1000, format it in red")
        assert any(t.task_type == TaskType.CONDITIONAL_FORMAT for t in tasks)

    def test_if_value_above_format_as_bold(self, ext):
        tasks = ext.extract("If the value is above 500, format as bold")
        assert any(t.task_type == TaskType.CONDITIONAL_FORMAT for t in tasks)


# ── P2-13: "enter a SUM formula" ──

class TestEnterNamedFormula:
    def test_enter_sum_formula(self, ext):
        tasks = ext.extract("Enter a SUM formula in cell B10 to total the column")
        assert any(t.task_type == TaskType.FORMULA for t in tasks)

    def test_create_average_formula(self, ext):
        tasks = ext.extract("Create an AVERAGE formula for the grades")
        assert any(t.task_type == TaskType.FORMULA for t in tasks)

    def test_add_countif_formula(self, ext):
        tasks = ext.extract("Add a COUNTIF formula to count active employees")
        assert any(t.task_type == TaskType.FORMULA for t in tasks)


# ── P2-14: Bare number extraction ──

class TestBareNumbers:
    def test_enter_number_in_cell(self, ext):
        tasks = ext.extract("Enter 100 in cell A1")
        cell_tasks = [t for t in tasks if t.task_type == TaskType.CELL_VALUE]
        assert len(cell_tasks) > 0
        assert cell_tasks[0].value == "100"

    def test_type_number_in_cell(self, ext):
        tasks = ext.extract("Type 42.5 in cell B2")
        cell_tasks = [t for t in tasks if t.task_type == TaskType.CELL_VALUE]
        assert len(cell_tasks) > 0
        assert cell_tasks[0].value == "42.5"

    def test_put_number_in_cell(self, ext):
        tasks = ext.extract("Put 1,500 in cell C3")
        cell_tasks = [t for t in tasks if t.task_type == TaskType.CELL_VALUE]
        assert len(cell_tasks) > 0
        assert cell_tasks[0].value == "1,500"


# ── P2-15: Minimum line length threshold ──

class TestShortInstructions:
    def test_short_bold_instruction(self, ext):
        tasks = ext.extract("Bold the A1.")
        assert len(tasks) > 0

    def test_short_wrap_instruction(self, ext):
        tasks = ext.extract("Wrap text.")
        assert len(tasks) > 0
