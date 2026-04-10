"""
Negative tests for TaskExtractor — verify graceful handling of edge cases.

Covers empty input, gibberish, non-instruction text, very short strings,
non-actionable text, and malformed cell references.
"""

import pytest

from excel_engine.parsers.task_extractor import TaskExtractor


@pytest.fixture
def extractor():
    return TaskExtractor()


# ══════════════════════════════════════════════════════════════════
# Empty / trivial input
# ══════════════════════════════════════════════════════════════════

class TestEmptyInput:
    def test_empty_string(self, extractor):
        tasks = extractor.extract("")
        assert tasks == []

    def test_whitespace_only(self, extractor):
        tasks = extractor.extract("   \n\t\n  ")
        assert tasks == []

    def test_single_space(self, extractor):
        tasks = extractor.extract(" ")
        assert tasks == []


# ══════════════════════════════════════════════════════════════════
# Non-instruction / gibberish text
# ══════════════════════════════════════════════════════════════════

class TestGibberish:
    def test_random_prose(self, extractor):
        tasks = extractor.extract("The weather is nice today. I like dogs.")
        assert len(tasks) == 0

    def test_lorem_ipsum(self, extractor):
        tasks = extractor.extract(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
        )
        assert len(tasks) == 0

    def test_numbers_only(self, extractor):
        tasks = extractor.extract("12345 67890 11111")
        assert len(tasks) == 0

    def test_special_characters(self, extractor):
        tasks = extractor.extract("!@#$%^&*(){}[]<>~`|\\")
        assert len(tasks) == 0

    def test_non_excel_instructions(self, extractor):
        tasks = extractor.extract(
            "Please open the door. Turn off the lights. Water the plants."
        )
        assert len(tasks) == 0


# ══════════════════════════════════════════════════════════════════
# Very short strings below meaningful threshold
# ══════════════════════════════════════════════════════════════════

class TestShortStrings:
    def test_single_word(self, extractor):
        tasks = extractor.extract("hello")
        assert len(tasks) == 0

    def test_two_characters(self, extractor):
        tasks = extractor.extract("ab")
        assert len(tasks) == 0

    def test_single_letter(self, extractor):
        tasks = extractor.extract("A")
        assert len(tasks) == 0


# ══════════════════════════════════════════════════════════════════
# Text with no actionable Excel commands
# ══════════════════════════════════════════════════════════════════

class TestNoActionableCommands:
    def test_discussion_about_excel(self, extractor):
        tasks = extractor.extract(
            "Excel is a powerful spreadsheet application made by Microsoft. "
            "It was first released in 1985 and has grown to become the most "
            "popular spreadsheet software in the world."
        )
        assert len(tasks) == 0

    def test_question_about_excel(self, extractor):
        tasks = extractor.extract("What is the best way to learn Excel?")
        assert len(tasks) == 0

    def test_generic_business_text(self, extractor):
        tasks = extractor.extract(
            "The quarterly report is due next Friday. "
            "Please ensure all departments submit their data on time."
        )
        assert len(tasks) == 0


# ══════════════════════════════════════════════════════════════════
# Malformed cell references — should not crash
# ══════════════════════════════════════════════════════════════════

class TestMalformedReferences:
    def test_cell_ref_like_but_invalid(self, extractor):
        # Should not crash, may or may not extract tasks
        tasks = extractor.extract("Enter value in cell ZZZZZ99999999")
        assert isinstance(tasks, list)

    def test_missing_cell_ref(self, extractor):
        tasks = extractor.extract("Enter value in cell")
        assert isinstance(tasks, list)

    def test_backwards_ref(self, extractor):
        tasks = extractor.extract("Enter 100 in cell 1A")
        assert isinstance(tasks, list)

    def test_unicode_in_cell_ref(self, extractor):
        tasks = extractor.extract("Enter value in cell Ä1")
        assert isinstance(tasks, list)

    def test_extremely_long_input(self, extractor):
        long_text = "Enter value in cell A1. " * 1000
        tasks = extractor.extract(long_text)
        assert isinstance(tasks, list)
        # Should handle repetitive input without crashing
        assert len(tasks) >= 0  # may or may not extract — just no crash

    def test_null_bytes_in_input(self, extractor):
        tasks = extractor.extract("Enter value\x00in cell A1")
        assert isinstance(tasks, list)

    def test_newlines_only(self, extractor):
        tasks = extractor.extract("\n\n\n\n\n")
        assert tasks == []
