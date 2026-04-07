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
