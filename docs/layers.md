# Automation Layers

Excel Engine uses a 6-layer cascading architecture. Each layer provides different capabilities and trade-offs.

---

## Layer 1: openpyxl — Offline File Manipulation

**Module:** `excel_engine.layers.openpyxl_layer`

The fastest and most reliable layer. Operates directly on `.xlsx` files without requiring Excel to be running.

### Capabilities

- Read/write cell values and formulas
- Cell formatting (font, fill, border, alignment, number format)
- Named ranges and defined names
- Sheet creation, deletion, renaming, reordering
- Row/column insertion, deletion, resizing
- Data validation rules
- Conditional formatting (basic)
- Table creation and styling
- Image insertion

### Limitations

- Cannot evaluate formulas (writes formula strings, not computed values)
- Cannot create PivotTables
- Cannot interact with Excel UI
- Changes require saving and reopening the file in Excel

### Usage Pattern

```python
from openpyxl import load_workbook

wb = load_workbook("workbook.xlsx")
ws = wb["Sheet1"]
ws["B4"] = "=SUM(B5:B10)"
wb.save("workbook.xlsx")
```

---

## Layer 2: xlwings — Live Excel Bridge

**Module:** `excel_engine.layers.xlwings_layer`

Communicates with a running Excel instance via Apple Events. Provides live read/write access with formula evaluation.

### Capabilities

- Everything openpyxl can do, plus:
- Live formula evaluation (read computed values)
- Real-time cell updates visible in Excel
- Chart creation and modification
- Sheet-level operations with undo support
- AutoFit column widths

### Limitations

- Requires Excel to be running
- Slower than openpyxl for bulk operations
- Some operations may trigger Excel recalculation delays
- Limited PivotTable support

### Usage Pattern

```python
import xlwings as xw

wb = xw.Book("workbook.xlsx")
ws = wb.sheets["Sheet1"]
ws.range("B4").value = "=SUM(B5:B10)"
```

---

## Layer 3: AppleScript — Excel-Specific Operations

**Module:** `excel_engine.layers.applescript_layer`

Executes AppleScript commands targeting the Microsoft Excel application. Useful for operations that have direct AppleScript dictionary support.

### Capabilities

- Save, Save As, Close operations
- Print and page setup configuration
- Worksheet protection/unprotection
- Window management (zoom, freeze panes, split)
- Excel-specific properties not exposed via xlwings

### Limitations

- macOS only
- Slower than direct API calls
- Limited error reporting (AppleScript errors are often vague)
- Cannot interact with dialogs or the ribbon

### Usage Pattern

```python
import subprocess

script = '''
tell application "Microsoft Excel"
    set value of cell "B4" of active sheet to "=SUM(B5:B10)"
end tell
'''
subprocess.run(["osascript", "-e", script], capture_output=True)
```

---

## Layer 4: System Events — Ribbon & Dialog UI Automation

**Module:** `excel_engine.layers.system_events_layer`

Uses macOS System Events to automate the Excel UI — clicking ribbon buttons, navigating menus, and interacting with dialogs.

### Capabilities

- Ribbon button clicks (Home, Insert, Page Layout, etc.)
- Menu bar navigation (File → Print, etc.)
- Dialog box interaction (text fields, checkboxes, dropdowns, OK/Cancel)
- Tab switching in multi-tab dialogs
- Keyboard shortcuts

### Limitations

- Fragile — UI changes between Excel versions can break scripts
- Requires Accessibility permission
- Cannot read cell data (use other layers for that)
- Timing-sensitive (needs delays between UI actions)

### Usage Pattern

```python
import subprocess

script = '''
tell application "System Events"
    tell process "Microsoft Excel"
        -- Click the Insert tab
        click radio button "Insert" of tab group 1 of group 2 of toolbar 1 of window 1
        -- Click PivotTable button
        click button "PivotTable" of group 1 of toolbar 1 of window 1
    end tell
end tell
'''
subprocess.run(["osascript", "-e", script], capture_output=True)
```

---

## Layer 5: VBA via VBE — Complex Automation

**Module:** `excel_engine.layers.vba_layer`

Injects and executes VBA macros through the Visual Basic Editor. Required for operations that have no AppleScript or xlwings equivalent.

### Capabilities

- PivotTable creation and modification
- Advanced chart customization
- Complex conditional formatting
- Custom functions and macros
- Workbook events
- Any operation expressible in VBA

### Limitations

- Requires "Trust access to the VBA project object model" to be enabled
- File must be saved as `.xlsm` (or VBA injected at runtime)
- Slower execution due to VBE interaction overhead
- VBA errors can crash Excel if not handled carefully

### Prerequisites

Enable VBA project access in Excel:
1. **Excel → Settings → Security → Trust Center**
2. Check **"Trust access to the VBA project object model"**

### Usage Pattern

```python
import xlwings as xw

wb = xw.Book("workbook.xlsm")
vba_code = '''
Sub CreatePivot()
    ' VBA code for PivotTable creation
End Sub
'''
wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)
wb.macro("CreatePivot").run()
```

---

## Layer 6: PyAutoGUI — Last-Resort Desktop Control

**Module:** `excel_engine.layers.pyautogui_layer`

Pixel-level mouse and keyboard automation. Used only when no higher-level layer can accomplish the task.

### Capabilities

- Mouse clicks at absolute or relative coordinates
- Keyboard input and hotkeys
- Screenshot capture and image matching (template matching)
- Drag and drop operations
- Any operation a human could perform

### Limitations

- Requires Screen Recording permission
- Retina display coordinates need scaling (÷2)
- Extremely fragile — any change in screen layout breaks it
- Cannot run headless
- Slowest layer

### Retina Display Handling

On Retina displays, PyAutoGUI coordinates are in logical pixels, but screenshots are in physical pixels:

```python
import pyautogui

# Screenshots are 2x resolution on Retina
screenshot = pyautogui.screenshot()
# If template matching returns (x, y) from the screenshot,
# divide by 2 for the click coordinate
pyautogui.click(x // 2, y // 2)
```

### Usage Pattern

```python
import pyautogui

# Locate a UI element by image
location = pyautogui.locateOnScreen("insert_tab.png", confidence=0.9)
if location:
    pyautogui.click(pyautogui.center(location))
```

---

## Layer Escalation Summary

```
Task arrives
    │
    ├─ Can openpyxl handle it?       → Layer 1 (fastest)
    ├─ Can xlwings handle it?        → Layer 2
    ├─ Can AppleScript handle it?    → Layer 3
    ├─ Can System Events handle it?  → Layer 4
    ├─ Can VBA handle it?            → Layer 5
    ├─ Can PyAutoGUI handle it?      → Layer 6 (last resort)
    └─ Nothing works                 → LayerExhaustionError
```
