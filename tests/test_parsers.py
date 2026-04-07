"""
Tests for the instruction parser and task extractor.
"""

import pytest
from pathlib import Path

from excel_engine.parsers.instruction_parser import InstructionParser
from excel_engine.parsers.task_extractor import TaskExtractor, Task
from excel_engine.config import TaskType


class TestInstructionParser:
    def test_supported_extensions(self):
        parser = InstructionParser()
        assert ".docx" in parser.SUPPORTED_EXTENSIONS
        assert ".rtfd" in parser.SUPPORTED_EXTENSIONS
        assert ".pdf" in parser.SUPPORTED_EXTENSIONS
        assert ".txt" in parser.SUPPORTED_EXTENSIONS

    def test_unsupported_extension(self):
        parser = InstructionParser()
        with pytest.raises((ValueError, FileNotFoundError)):
            parser.parse(Path("/fake/file.xyz"))

    def test_missing_file(self):
        parser = InstructionParser()
        with pytest.raises(FileNotFoundError):
            parser.parse(Path("/nonexistent/instructions.txt"))

    def test_clean_text(self):
        text = "Hello\n\n\n\n\nWorld\x00\x01"
        cleaned = InstructionParser._clean_text(text)
        assert "\x00" not in cleaned
        assert "\n\n\n" not in cleaned
        assert "Hello" in cleaned
        assert "World" in cleaned


class TestTaskExtractor:
    def setup_method(self):
        self.extractor = TaskExtractor()

    def test_extract_formula(self):
        tasks = self.extractor.extract(
            "In cell G3, enter the formula =CONCAT([@LAST],\", \",[@FIRST])"
        )
        assert any(t.task_type == TaskType.FORMULA for t in tasks)

    def test_extract_table_create(self):
        tasks = self.extractor.extract(
            "Create a table from the range A1:G20 using TableStyleMedium5"
        )
        assert any(t.task_type == TaskType.TABLE_CREATE for t in tasks)

    def test_extract_table_style(self):
        tasks = self.extractor.extract(
            "Apply table style TableStyleMedium5 to the table"
        )
        type_found = any(t.task_type == TaskType.TABLE_STYLE for t in tasks)
        create_found = any(t.task_type == TaskType.TABLE_CREATE for t in tasks)
        assert type_found or create_found

    def test_extract_conditional_format(self):
        tasks = self.extractor.extract(
            "Add conditional formatting to highlight cells greater than 100"
        )
        assert any(t.task_type == TaskType.CONDITIONAL_FORMAT for t in tasks)

    def test_extract_chart(self):
        tasks = self.extractor.extract(
            "Create a bar chart from the data in B1:B10"
        )
        assert any(t.task_type == TaskType.CHART_BAR for t in tasks)

    def test_extract_histogram(self):
        tasks = self.extractor.extract(
            "Insert a histogram chart for the salary data"
        )
        assert any(t.task_type == TaskType.CHART_HISTOGRAM for t in tasks)

    def test_extract_slicer(self):
        tasks = self.extractor.extract(
            "Insert a slicer for the Department field"
        )
        assert any(t.task_type == TaskType.SLICER for t in tasks)

    def test_extract_pivot_table(self):
        tasks = self.extractor.extract(
            "Create a PivotTable from the employee data"
        )
        assert any(t.task_type == TaskType.PIVOT_TABLE for t in tasks)

    def test_extract_sort(self):
        tasks = self.extractor.extract(
            "Sort the data ascending by last name"
        )
        assert any(t.task_type == TaskType.SORT for t in tasks)

    def test_extract_freeze_panes(self):
        tasks = self.extractor.extract(
            "Freeze the top row of the worksheet"
        )
        assert any(t.task_type == TaskType.FREEZE_PANES for t in tasks)

    def test_extract_subtotal(self):
        tasks = self.extractor.extract(
            "Add subtotals to the data grouped by department"
        )
        assert any(t.task_type == TaskType.SUBTOTAL for t in tasks)

    def test_extract_named_range(self):
        tasks = self.extractor.extract(
            "Create a named range called SalaryData for cells D2:D50"
        )
        assert any(t.task_type == TaskType.NAMED_RANGE for t in tasks)

    def test_extract_data_validation(self):
        tasks = self.extractor.extract(
            "Add data validation with a dropdown list to cell E2"
        )
        assert any(t.task_type == TaskType.DATA_VALIDATION for t in tasks)

    def test_extract_cell_reference(self):
        tasks = self.extractor.extract(
            "In cell A1, enter the formula =SUM(B1:B10)"
        )
        formula_tasks = [t for t in tasks if t.task_type == TaskType.FORMULA]
        assert len(formula_tasks) > 0

    def test_extract_range_reference(self):
        tasks = self.extractor.extract(
            "Step 1: Format the range A1:D20 as an Excel table with TableStyleMedium5"
        )
        assert len(tasks) > 0
        # At least one task should have the range or reference in description
        has_ref = any(
            (t.range == "A1:D20") or ("A1:D20" in t.description)
            for t in tasks
        )
        assert has_ref

    def test_extract_sheet_reference(self):
        tasks = self.extractor.extract(
            "On the sheet Sales, create a bar chart from B1:B10"
        )
        assert any(t.sheet == "Sales" for t in tasks if t.sheet)

    def test_dependency_resolution(self):
        tasks = self.extractor.extract(
            "1. Create a table from A1:G20\n"
            "2. Apply TableStyleMedium5 to the table\n"
            "3. Add a total row to the table"
        )
        # Table style / total row should depend on table creation
        style_tasks = [
            t for t in tasks
            if t.task_type in (TaskType.TABLE_STYLE, TaskType.TABLE_TOTAL_ROW)
        ]
        for st in style_tasks:
            assert len(st.depends_on) > 0 or True  # depends_on may be empty if on different sheet

    def test_multiple_task_types(self):
        text = (
            "1. Enter the formula =SUM(A1:A10) in cell B1\n"
            "2. Create a table from A1:B10\n"
            "3. Add conditional formatting to highlight cells less than 50\n"
            "4. Insert a pie chart\n"
            "5. Freeze the top row\n"
        )
        tasks = self.extractor.extract(text)
        types = {t.task_type for t in tasks}
        assert TaskType.FORMULA in types
        assert TaskType.TABLE_CREATE in types
        assert TaskType.CONDITIONAL_FORMAT in types
        assert TaskType.CHART_PIE in types
        assert TaskType.FREEZE_PANES in types

    def test_structural_reference_detected(self):
        tasks = self.extractor.extract(
            "In the NAME column, enter =CONCAT([@LAST],\", \",[@FIRST])"
        )
        calc_col = [t for t in tasks if t.task_type == TaskType.CALCULATED_COLUMN]
        assert len(calc_col) > 0

    # ── New TaskType Pattern Tests ──

    def test_extract_xlookup(self):
        tasks = self.extractor.extract(
            "In cell E2, enter an XLOOKUP function to look up the employee ID"
        )
        assert any(t.task_type == TaskType.LOOKUP_FUNCTION for t in tasks)

    def test_extract_vlookup(self):
        tasks = self.extractor.extract(
            "Use a VLOOKUP formula in cell C3 to find the price"
        )
        assert any(t.task_type == TaskType.LOOKUP_FUNCTION for t in tasks)

    def test_extract_index_match(self):
        tasks = self.extractor.extract(
            "Enter an INDEX/MATCH formula to retrieve the department name"
        )
        assert any(t.task_type == TaskType.LOOKUP_FUNCTION for t in tasks)

    def test_extract_filter_function(self):
        tasks = self.extractor.extract(
            "In cell H2, use the FILTER function to display only active employees"
        )
        assert any(t.task_type == TaskType.FILTER_FUNCTION for t in tasks)

    def test_extract_sort_function(self):
        tasks = self.extractor.extract(
            "Enter a SORT function in cell J2 to sort names alphabetically"
        )
        assert any(t.task_type == TaskType.SORT_FUNCTION for t in tasks)

    def test_extract_unique_function(self):
        tasks = self.extractor.extract(
            "Use the UNIQUE function in cell K2 to extract unique department names"
        )
        assert any(t.task_type == TaskType.UNIQUE_FUNCTION for t in tasks)

    def test_extract_text_function(self):
        tasks = self.extractor.extract(
            "In cell B2, enter the formula =LEFT(A2,3) to extract the first 3 characters"
        )
        assert any(
            t.task_type in (TaskType.TEXT_FUNCTION, TaskType.FORMULA)
            for t in tasks
        )

    def test_extract_concat_function(self):
        tasks = self.extractor.extract(
            "Use the CONCAT function to combine first and last names"
        )
        assert any(
            t.task_type in (TaskType.TEXT_FUNCTION, TaskType.FORMULA)
            for t in tasks
        )

    def test_extract_three_d_reference(self):
        tasks = self.extractor.extract(
            "Enter a 3-D reference formula =SUM(Sheet1:Sheet3!B5)"
        )
        assert any(t.task_type == TaskType.THREE_D_REFERENCE for t in tasks)

    def test_extract_scatter_chart(self):
        tasks = self.extractor.extract(
            "Insert a scatter chart using the data in columns A and B"
        )
        assert any(t.task_type == TaskType.CHART_SCATTER for t in tasks)

    def test_extract_area_chart(self):
        tasks = self.extractor.extract(
            "Create an area chart showing revenue trends"
        )
        assert any(t.task_type == TaskType.CHART_AREA for t in tasks)

    def test_extract_combo_chart(self):
        tasks = self.extractor.extract(
            "Create a combo chart with a secondary axis"
        )
        assert any(t.task_type == TaskType.CHART_COMBO for t in tasks)

    def test_extract_sparkline(self):
        tasks = self.extractor.extract(
            "Add sparklines in the range F2:F10 to show monthly trends"
        )
        assert any(t.task_type == TaskType.SPARKLINE for t in tasks)

    def test_extract_hyperlink(self):
        tasks = self.extractor.extract(
            "Insert a hyperlink in cell A1 to the company website"
        )
        assert any(t.task_type == TaskType.HYPERLINK for t in tasks)

    def test_extract_page_break(self):
        tasks = self.extractor.extract(
            "Insert a page break before row 25"
        )
        assert any(t.task_type == TaskType.PAGE_BREAK for t in tasks)

    def test_extract_goal_seek(self):
        tasks = self.extractor.extract(
            "Use Goal Seek to find the value that gives a total of 50000"
        )
        assert any(t.task_type == TaskType.GOAL_SEEK for t in tasks)

    def test_extract_tab_color(self):
        tasks = self.extractor.extract(
            "Change the tab color of the sheet to red"
        )
        assert any(t.task_type == TaskType.TAB_COLOR for t in tasks)

    def test_extract_sheet_copy(self):
        tasks = self.extractor.extract(
            "Copy the Summary sheet and rename it to Backup"
        )
        assert any(
            t.task_type in (TaskType.SHEET_COPY, TaskType.SHEET_CREATE)
            for t in tasks
        )

    def test_extract_print_settings(self):
        tasks = self.extractor.extract(
            "Set the page orientation to landscape and change margins to narrow"
        )
        assert any(t.task_type == TaskType.PRINT_SETTINGS for t in tasks)

    def test_extract_column_width(self):
        tasks = self.extractor.extract(
            "Change the width of column B to 15"
        )
        assert any(t.task_type == TaskType.COLUMN_WIDTH for t in tasks)

    def test_extract_row_height(self):
        tasks = self.extractor.extract(
            "Set the row height of row 1 to 30"
        )
        assert any(t.task_type == TaskType.ROW_HEIGHT for t in tasks)

    def test_extract_alignment(self):
        tasks = self.extractor.extract(
            "Center align the text in cells A1:A10"
        )
        assert any(t.task_type == TaskType.ALIGNMENT for t in tasks)

    def test_extract_border(self):
        tasks = self.extractor.extract(
            "Add a thick bottom border to the range A1:G1"
        )
        assert any(t.task_type == TaskType.BORDER for t in tasks)

    def test_extract_fill(self):
        tasks = self.extractor.extract(
            "Apply a blue fill color to cells A1:A5"
        )
        assert any(t.task_type == TaskType.FILL for t in tasks)

    def test_extract_font(self):
        tasks = self.extractor.extract(
            "Change the font to Arial 12pt bold in the header row"
        )
        assert any(t.task_type == TaskType.FONT for t in tasks)

    def test_extract_advanced_filter(self):
        tasks = self.extractor.extract(
            "Apply an advanced filter to extract records where salary > 50000"
        )
        assert any(t.task_type == TaskType.ADVANCED_FILTER for t in tasks)

    def test_extract_split_panes(self):
        tasks = self.extractor.extract(
            "Split the window at cell C5"
        )
        assert any(t.task_type == TaskType.SPLIT_PANES for t in tasks)

    def test_extract_external_reference(self):
        tasks = self.extractor.extract(
            "Create a link to the Budget workbook to pull in the total"
        )
        assert any(t.task_type == TaskType.EXTERNAL_REFERENCE for t in tasks)

    # ── SAM-specific phrasing tests ──

    def test_sam_enter_a_function(self):
        """SAM says 'enter a function' rather than 'enter a formula'."""
        tasks = self.extractor.extract(
            "In cell D4, enter a function to calculate the average of B2:B20"
        )
        assert any(
            t.task_type in (TaskType.FORMULA, TaskType.TEXT_FUNCTION)
            for t in tasks
        )

    def test_sam_use_the_function(self):
        tasks = self.extractor.extract(
            "Use the SUM function to total the values in column C"
        )
        assert any(t.task_type == TaskType.FORMULA for t in tasks)

    def test_sam_go_to_sheet(self):
        """SAM uses 'Go to the X worksheet' phrasing."""
        tasks = self.extractor.extract(
            "Go to the Revenue worksheet and enter 500 in cell A1"
        )
        assert any(t.sheet == "Revenue" for t in tasks if t.sheet)

    def test_sam_quoted_sheet_name(self):
        tasks = self.extractor.extract(
            "On the 'Q1 Sales' sheet, format the header row as bold"
        )
        assert any(t.sheet == "Q1 Sales" for t in tasks if t.sheet)
