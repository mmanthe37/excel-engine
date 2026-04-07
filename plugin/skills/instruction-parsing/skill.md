# Instruction Parsing Skill

## Description

Provides knowledge and patterns for extracting structured task lists from SAM (Skills Assessment Manager) assignment instruction files in various formats (.docx, .rtfd, .pdf, .txt). Covers text extraction, step identification, task classification, dependency detection, and execution plan generation.

## When to Trigger

- User provides an instruction file for an Excel assignment
- User needs to parse a .docx, .rtfd, .pdf, or .txt instruction file
- User wants to understand what steps an assignment requires before executing
- User asks to "read the instructions" or "parse the assignment"

## Key Knowledge

### Format Detection and Extraction

| Format | Method |
|--------|--------|
| `.docx` | `python-docx`: `Document(path).paragraphs` |
| `.rtfd` / `.rtf` | `textutil -convert txt path -stdout` |
| `.pdf` | `pdfplumber` or `pdftotext path -` |
| `.txt` | Direct file read |

### SAM Instruction Structure

SAM instructions typically follow this pattern:
1. Header with module name and workbook filename
2. Numbered steps (sometimes with sub-steps a, b, c)
3. Each step specifies: target sheet, cell/range, operation, expected value
4. Steps may reference previous steps ("using the table you created in step 3")

### Task Extraction Rules

- Parse each numbered step into a `Task` object with: id, type, sheet, cell/range, value/formula, params
- Detect implicit dependencies (e.g., "format the table" depends on "create the table")
- Identify which automation layer each task requires
- Group tasks by worksheet for efficient execution
- Flag ambiguous instructions for user clarification

### Using the Engine Parser

```python
from excel_engine.parsers.instruction_parser import InstructionParser
from excel_engine.parsers.task_extractor import TaskExtractor

parser = InstructionParser()
text = parser.parse(Path("instructions.docx"))

extractor = TaskExtractor()
tasks = extractor.extract(text)

for task in tasks:
    print(f"{task.id}: {task.task_type.value} on {task.sheet}!{task.cell}")
```

### Common SAM Instruction Patterns

- "In cell B5, enter a formula using the SUM function..." → `TaskType.FORMULA`
- "Format the range B2:B10 as Currency..." → `TaskType.NUMBER_FORMAT`
- "Create an Excel table from the range A1:F20..." → `TaskType.TABLE_CREATE`
- "Apply the Table Style Medium 9..." → `TaskType.TABLE_STYLE`
- "Sort the data ascending by column B..." → `TaskType.SORT`
- "Add a PivotTable on a new worksheet..." → `TaskType.PIVOT_TABLE`
- "Insert a slicer for the Category field..." → `TaskType.SLICER`
- "Freeze panes at row 2..." → `TaskType.FREEZE_PANES`
