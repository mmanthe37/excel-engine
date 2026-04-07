# Contributing to Excel Engine

Thank you for your interest in contributing! This guide will help you get set up.

## Development Setup

### Prerequisites

- macOS 13.0+ (Ventura or later)
- Microsoft Excel for Mac 365
- Python 3.11+
- Git

### Getting Started

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/excel-engine.git
cd excel-engine

# Create a virtual environment
python3 -m venv venv
source venv/bin/activate

# Install in editable mode with dev dependencies
pip install -e ".[dev]"
```

### macOS Permissions

Grant your terminal app the following permissions in **System Settings → Privacy & Security**:

- **Accessibility** — required for UI automation
- **Screen Recording** — required for screenshot-based detection

## Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=excel_engine

# Run a specific test file
pytest tests/test_planner.py

# Run tests matching a pattern
pytest -k "test_openpyxl"
```

> **Note:** Some integration tests require Excel to be running. These are marked with `@pytest.mark.integration` and skipped by default. Run them with:
>
> ```bash
> pytest -m integration
> ```

## Code Style

### Type Hints

All functions must include type annotations:

```python
def set_cell_value(sheet: str, cell: str, value: str | float) -> bool:
    """Set a cell value in the active workbook."""
    ...
```

### Docstrings

Use Google-style docstrings:

```python
def execute_checkpoint(self, checkpoint: str) -> CheckpointResult:
    """Execute a single checkpoint instruction.

    Args:
        checkpoint: The instruction text to execute.

    Returns:
        A CheckpointResult indicating success or failure.

    Raises:
        LayerExhaustionError: If all 6 layers fail to complete the task.
    """
```

### Linting and Formatting

```bash
# Lint with ruff
ruff check .

# Format with ruff
ruff format .

# Type check with mypy
mypy excel_engine/
```

## Pull Request Process

1. **Fork** the repository and create a branch from `main`.
2. **Name your branch** descriptively: `fix/retina-scaling`, `feat/pivot-table-layer`, etc.
3. **Write tests** for any new functionality.
4. **Ensure all checks pass**: `pytest`, `ruff check .`, `mypy excel_engine/`.
5. **Update documentation** if you change public APIs or add features.
6. **Open a PR** with a clear description of the change and any relevant context.

### Commit Messages

Use conventional commit style:

```
feat: add PivotTable creation via VBA layer
fix: correct Retina coordinate scaling in PyAutoGUI layer
docs: add troubleshooting guide for VBProject access
test: add integration tests for AppleScript layer
```

## Architecture Notes

When adding or modifying a layer, keep these principles in mind:

- **Layers are ordered by cost**: prefer lower-numbered layers (faster, less invasive).
- **Each layer must be independently testable** with mock fixtures.
- **The engine escalates automatically**: don't add cross-layer dependencies.
- **macOS-specific code** should be isolated in layer implementations, not in the planner or engine core.

## Questions?

Open an issue with the `question` label and we'll help you out.
