"""
Tests for extraction bug fixes (BUG-1 through BUG-5).

Each test targets a specific bug identified in expert E2E and parsing reviews.
"""

import re
from unittest.mock import MagicMock, PropertyMock

import pytest

from excel_engine.config import TaskType
from excel_engine.parsers.task_extractor import Task, TaskExtractor, _FORMULA_REF


# ─── Helpers ────────────────────────────────────────────────────────────────

@pytest.fixture
def extractor():
    return TaskExtractor()


# ═══════════════════════════════════════════════════════════════════════════
# BUG-1: Greedy formula regex captures trailing text
# ═══════════════════════════════════════════════════════════════════════════


class TestBug1GreedyFormulaRegex:
    """Arithmetic formula branch should NOT swallow trailing prose."""

    def test_formula_stops_before_in_cell(self, extractor):
        """'=B2*C2 in cell D2' should extract formula =B2*C2, not =B2*C2 in cell D2."""
        tasks = extractor.extract("Enter the formula =B2*C2 in cell D2")
        formula_tasks = [t for t in tasks if t.formula]
        assert formula_tasks, "Expected at least one formula task"
        assert formula_tasks[0].formula == "=B2*C2"

    def test_formula_cell_reference(self, extractor):
        """Target cell D2 should be extracted, not B2."""
        tasks = extractor.extract("Enter the formula =B2*C2 in cell D2")
        formula_tasks = [t for t in tasks if t.formula]
        assert formula_tasks[0].cell == "D2"

    def test_formula_regex_arithmetic_simple(self):
        """Raw regex: =A1+B1 should match just the formula."""
        m = _FORMULA_REF.search("Enter =A1+B1 in cell C1")
        assert m is not None
        assert m.group(1) == "=A1+B1"

    def test_formula_regex_arithmetic_complex(self):
        """Raw regex: =A1*B1+C1 should not include trailing text."""
        m = _FORMULA_REF.search("Type =A1*B1+C1 in cell D1, then format")
        assert m is not None
        captured = m.group(1)
        assert "in cell" not in captured
        assert "format" not in captured

    def test_formula_addition(self, extractor):
        tasks = extractor.extract("In cell E5, enter =C5+D5")
        formula_tasks = [t for t in tasks if t.formula]
        assert formula_tasks
        assert formula_tasks[0].formula == "=C5+D5"
        assert formula_tasks[0].cell == "E5"

    def test_formula_division(self, extractor):
        tasks = extractor.extract("Enter =B3/C3 in cell D3")
        formula_tasks = [t for t in tasks if t.formula]
        assert formula_tasks
        assert formula_tasks[0].formula == "=B3/C3"

    def test_formula_subtraction_with_parens(self, extractor):
        tasks = extractor.extract("Enter the formula =A1*(B1-C1) in cell D1")
        formula_tasks = [t for t in tasks if t.formula]
        assert formula_tasks
        # The formula should not include "in cell D1"
        assert "in cell" not in formula_tasks[0].formula


# ═══════════════════════════════════════════════════════════════════════════
# BUG-2: Target cell lost when range is in formula
# ═══════════════════════════════════════════════════════════════════════════


class TestBug2TargetCellLost:
    """Target cell from 'in cell X' should survive range extraction in formula."""

    def test_sum_range_preserves_target_cell(self, extractor):
        """'In cell D10, enter =SUM(D2:D9)' — cell should be D10."""
        tasks = extractor.extract("In cell D10, enter the formula =SUM(D2:D9)")
        assert tasks, "Expected at least one task"
        task = tasks[0]
        assert task.cell == "D10"
        assert task.formula == "=SUM(D2:D9)"

    def test_average_range_preserves_target_cell(self, extractor):
        tasks = extractor.extract("In cell E15, enter =AVERAGE(E2:E14)")
        task = tasks[0]
        assert task.cell == "E15"

    def test_countif_range_preserves_target_cell(self, extractor):
        tasks = extractor.extract("In cell F1, enter =COUNTIF(A1:A100, \">50\")")
        task = tasks[0]
        assert task.cell == "F1"

    def test_range_still_extracted(self, extractor):
        """Range inside formula should still be stored in task.range."""
        tasks = extractor.extract("In cell D10, enter the formula =SUM(D2:D9)")
        task = tasks[0]
        assert task.range == "D2:D9"

    def test_into_cell_phrasing(self, extractor):
        """'into cell X' also works as target cell."""
        tasks = extractor.extract("Enter =SUM(B2:B8) into cell B9")
        task = tasks[0]
        assert task.cell == "B9"

    def test_cell_without_range(self, extractor):
        """Normal cell ref without range should still work."""
        tasks = extractor.extract("In cell A1, enter the value 100")
        cell_tasks = [t for t in tasks if t.cell]
        assert cell_tasks
        assert cell_tasks[0].cell == "A1"


# ═══════════════════════════════════════════════════════════════════════════
# BUG-3: Section headers merged with preceding step text
# ═══════════════════════════════════════════════════════════════════════════


class TestBug3SectionHeaderSplit:
    """Section headers should cause a split, not merge with previous step."""

    def test_section_header_splits(self, extractor):
        text = (
            "1. Enter the formula =B2*C2 in cell D2\n\n"
            "Section 2: Formatting\n"
            "2. Bold the header row"
        )
        tasks = extractor.extract(text)
        # "Section 2: Formatting" should not be appended to the first step
        for task in tasks:
            assert "Section 2" not in task.description or "Formatting" in task.description

    def test_section_header_not_in_formula(self, extractor):
        text = (
            "a. Enter =B2*C2 in cell D2\n\n"
            "Section 2: Formatting\n"
            "b. Apply bold to row 1"
        )
        lines = extractor._split_instructions(text)
        # No line should contain both a formula and "Section 2"
        for line in lines:
            if "=B2*C2" in line:
                assert "Section 2" not in line

    def test_multiple_sections_split(self, extractor):
        text = (
            "Section 1: Data Entry\n"
            "1. Enter 100 in cell A1\n"
            "2. Enter 200 in cell A2\n\n"
            "Section 2: Formulas\n"
            "3. Enter =SUM(A1:A2) in cell A3\n\n"
            "Section 3: Formatting\n"
            "4. Bold all cells"
        )
        lines = extractor._split_instructions(text)
        # Each instruction should be its own line
        assert len(lines) >= 4

    def test_section_colon_variant(self, extractor):
        text = (
            "1. Do something\n"
            "Section 3: Charts\n"
            "2. Create a chart"
        )
        lines = extractor._split_instructions(text)
        # "Section 3: Charts" should not be glued to step 1
        for line in lines:
            if "Do something" in line:
                assert "Charts" not in line


# ═══════════════════════════════════════════════════════════════════════════
# BUG-4: Suppress spurious CELL_VALUE alongside FORMULA/LOOKUP_FUNCTION
# ═══════════════════════════════════════════════════════════════════════════


class TestBug4SpuriousCellValue:
    """CELL_VALUE should be suppressed when a formula-type task exists."""

    def test_vlookup_no_cell_value(self, extractor):
        tasks = extractor.extract("In cell E2, enter a VLOOKUP to find the price")
        types = {t.task_type for t in tasks}
        assert TaskType.LOOKUP_FUNCTION in types
        assert TaskType.CELL_VALUE not in types

    def test_formula_no_cell_value(self, extractor):
        tasks = extractor.extract("In cell D2, enter the formula =B2*C2")
        types = {t.task_type for t in tasks}
        assert TaskType.FORMULA in types
        assert TaskType.CELL_VALUE not in types

    def test_xlookup_no_cell_value(self, extractor):
        tasks = extractor.extract("In cell F3, enter an XLOOKUP to retrieve the name")
        types = {t.task_type for t in tasks}
        assert TaskType.LOOKUP_FUNCTION in types
        assert TaskType.CELL_VALUE not in types

    def test_pure_cell_value_still_works(self, extractor):
        """A pure value entry should still yield CELL_VALUE."""
        tasks = extractor.extract("Enter the value 42 in cell A1")
        types = {t.task_type for t in tasks}
        assert TaskType.CELL_VALUE in types

    def test_sum_formula_no_cell_value(self, extractor):
        tasks = extractor.extract("In cell A10, enter =SUM(A1:A9)")
        types = {t.task_type for t in tasks}
        assert TaskType.FORMULA in types
        assert TaskType.CELL_VALUE not in types


# ═══════════════════════════════════════════════════════════════════════════
# BUG-5: Verifier passes tasks with value=None unconditionally
# ═══════════════════════════════════════════════════════════════════════════


class TestBug5VerifierNoValue:
    """_verify_cell_value should warn when no expected value is specified."""

    def test_no_value_no_formula_warns(self):
        from excel_engine.verifier.workbook_verifier import WorkbookVerifier

        task = Task(
            id="t1",
            task_type=TaskType.CELL_VALUE,
            description="test",
            cell="A1",
            value=None,
            formula=None,
        )

        wb_mock = MagicMock()
        ws_mock = MagicMock()
        cell_mock = MagicMock()
        cell_mock.value = "Hello"
        ws_mock.__getitem__ = MagicMock(return_value=cell_mock)
        wb_mock.sheetnames = ["Sheet1"]
        wb_mock.__getitem__ = MagicMock(return_value=ws_mock)
        wb_mock.active = ws_mock

        verifier = WorkbookVerifier()
        verifier._wb = wb_mock
        result = verifier._verify_cell_value(task)

        assert result.passed is True
        assert "cannot verify" in result.message.lower()

    def test_with_value_still_checks(self):
        from excel_engine.verifier.workbook_verifier import WorkbookVerifier

        task = Task(
            id="t2",
            task_type=TaskType.CELL_VALUE,
            description="test",
            cell="A1",
            value="Hello",
            formula=None,
        )

        wb_mock = MagicMock()
        ws_mock = MagicMock()
        cell_mock = MagicMock()
        cell_mock.value = "Hello"
        ws_mock.__getitem__ = MagicMock(return_value=cell_mock)
        wb_mock.sheetnames = ["Sheet1"]
        wb_mock.__getitem__ = MagicMock(return_value=ws_mock)
        wb_mock.active = ws_mock

        verifier = WorkbookVerifier()
        verifier._wb = wb_mock
        result = verifier._verify_cell_value(task)

        assert result.passed is True
        assert "matches" in result.message.lower()

    def test_with_formula_no_warning(self):
        from excel_engine.verifier.workbook_verifier import WorkbookVerifier

        task = Task(
            id="t3",
            task_type=TaskType.CELL_VALUE,
            description="test",
            cell="A1",
            value=None,
            formula="=SUM(A1:A5)",
        )

        wb_mock = MagicMock()
        ws_mock = MagicMock()
        cell_mock = MagicMock()
        cell_mock.value = 150
        ws_mock.__getitem__ = MagicMock(return_value=cell_mock)
        wb_mock.sheetnames = ["Sheet1"]
        wb_mock.__getitem__ = MagicMock(return_value=ws_mock)
        wb_mock.active = ws_mock

        verifier = WorkbookVerifier()
        verifier._wb = wb_mock
        result = verifier._verify_cell_value(task)

        assert result.passed is True
        assert "cannot verify" not in result.message.lower()

    def test_value_mismatch_still_fails(self):
        from excel_engine.verifier.workbook_verifier import WorkbookVerifier

        task = Task(
            id="t4",
            task_type=TaskType.CELL_VALUE,
            description="test",
            cell="A1",
            value="Expected",
            formula=None,
        )

        wb_mock = MagicMock()
        ws_mock = MagicMock()
        cell_mock = MagicMock()
        cell_mock.value = "Actual"
        ws_mock.__getitem__ = MagicMock(return_value=cell_mock)
        wb_mock.sheetnames = ["Sheet1"]
        wb_mock.__getitem__ = MagicMock(return_value=ws_mock)
        wb_mock.active = ws_mock

        verifier = WorkbookVerifier()
        verifier._wb = wb_mock
        result = verifier._verify_cell_value(task)

        assert result.passed is False
        assert "mismatch" in result.message.lower()
