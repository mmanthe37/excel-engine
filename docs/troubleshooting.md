# Troubleshooting

Common issues and their solutions when running Excel Engine.

---

## VBProject Object Model Is Inaccessible

**Symptom:** Layer 5 (VBA) fails with an error like:

```
AttributeError: VBProject is not accessible
```

or

```
Runtime Error: Programmatic access to Visual Basic project is not trusted
```

**Cause:** Excel's security settings block external access to the VBA project object model by default.

**Solution:**

1. Open Microsoft Excel
2. Go to **Excel → Settings** (⌘ + ,)
3. Click **Security** (or **Trust Center**)
4. Check **"Trust access to the VBA project object model"**
5. Restart Excel

> **Note:** This setting resets when Excel updates. If VBA suddenly stops working after an update, re-enable this setting.

---

## Retina Display Coordinate Mismatch

**Symptom:** Layer 6 (PyAutoGUI) clicks in the wrong location — usually offset by exactly 2x from where it should click.

**Cause:** macOS Retina displays use a 2x scaling factor. PyAutoGUI's `screenshot()` returns an image at physical resolution (e.g., 5120×2880), but `click()` expects logical coordinates (e.g., 2560×1440).

**Solution:**

Excel Engine handles this automatically, but if you're writing custom PyAutoGUI code:

```python
import pyautogui

screenshot = pyautogui.screenshot()
# Template match returns physical pixel coordinates
match = pyautogui.locate(template, screenshot)
if match:
    center = pyautogui.center(match)
    # Divide by Retina scale factor for click coordinates
    pyautogui.click(center.x // 2, center.y // 2)
```

To detect the scale factor programmatically:

```python
import subprocess
result = subprocess.run(
    ["system_profiler", "SPDisplaysDataType"],
    capture_output=True, text=True
)
is_retina = "Retina" in result.stdout
scale = 2 if is_retina else 1
```

---

## macOS Path Colon vs. Slash

**Symptom:** AppleScript file operations fail with "file not found" errors even though the file exists.

**Cause:** AppleScript uses the legacy HFS path format with colons (`:`) as separators, while Python and the shell use POSIX paths with slashes (`/`).

| Format | Example |
|--------|---------|
| POSIX | `/Users/you/Documents/workbook.xlsx` |
| HFS | `Macintosh HD:Users:you:Documents:workbook.xlsx` |

**Solution:**

Convert POSIX paths to HFS format in AppleScript:

```applescript
set posixPath to "/Users/you/Documents/workbook.xlsx"
set hfsPath to POSIX file posixPath as text
```

Excel Engine handles this conversion automatically in Layer 3.

---

## Accessibility Permission Denied

**Symptom:** Layer 4 (System Events) fails silently or raises:

```
System Events got an error: not allowed assistive access
```

**Cause:** The terminal app running Excel Engine hasn't been granted Accessibility permissions.

**Solution:**

1. Open **System Settings → Privacy & Security → Accessibility**
2. Click the lock to make changes
3. Add your terminal app (Terminal.app, iTerm2, VS Code, etc.)
4. Toggle it **ON**
5. Restart the terminal app

> **Tip:** If using VS Code's integrated terminal, you need to add **Visual Studio Code** to the Accessibility list, not Terminal.app.

---

## Screen Recording Permission Denied

**Symptom:** PyAutoGUI's `screenshot()` returns a blank or all-black image.

**Cause:** The terminal app hasn't been granted Screen Recording permissions.

**Solution:**

1. Open **System Settings → Privacy & Security → Screen Recording**
2. Add your terminal app
3. Toggle it **ON**
4. Restart the terminal app

---

## Excel Not Responding / Hung

**Symptom:** The engine hangs waiting for Excel to respond. Layer 2 (xlwings) or Layer 3 (AppleScript) calls time out.

**Cause:** Excel is displaying a modal dialog (e.g., "Do you want to save?", "Enable macros?", a crash recovery dialog) that blocks all automation commands.

**Solution:**

1. **Check for dialogs** — Switch to Excel and dismiss any open dialogs.
2. **Force quit and relaunch:**
   ```bash
   osascript -e 'tell application "Microsoft Excel" to quit saving no'
   sleep 2
   open -a "Microsoft Excel"
   ```
3. **Prevent save prompts** — Excel Engine saves before each section to avoid "unsaved changes" dialogs.

---

## openpyxl Cannot Open `.xlsm` Files With VBA

**Symptom:** Opening an `.xlsm` file with openpyxl and saving it strips all VBA macros.

**Cause:** openpyxl does not preserve VBA projects by default.

**Solution:**

Use `keep_vba=True` when loading:

```python
from openpyxl import load_workbook

wb = load_workbook("workbook.xlsm", keep_vba=True)
# ... make changes ...
wb.save("workbook.xlsm")
```

> **Note:** Even with `keep_vba=True`, some complex VBA projects may not round-trip perfectly. For VBA-heavy files, prefer Layer 2 (xlwings) or Layer 5 (VBA injection).

---

## xlwings Cannot Find Excel

**Symptom:**

```
xlwings.XlwingsError: Cannot find Excel application
```

**Cause:** Excel is not installed, or xlwings cannot locate it.

**Solution:**

1. Verify Excel is installed: `ls /Applications/Microsoft\ Excel.app`
2. Try opening Excel manually: `open -a "Microsoft Excel"`
3. If using a virtual environment, ensure xlwings is installed in it: `pip install xlwings`

---

## Formula Shows as Text Instead of Evaluating

**Symptom:** After writing a formula with openpyxl, Excel displays the formula string (e.g., `=SUM(B5:B10)`) instead of the computed value.

**Cause:** The cell's number format is set to "Text" (`@`), which prevents formula evaluation.

**Solution:**

Set the number format to "General" before writing the formula:

```python
from openpyxl import load_workbook

wb = load_workbook("workbook.xlsx")
ws = wb.active
ws["B4"].number_format = "General"
ws["B4"] = "=SUM(B5:B10)"
wb.save("workbook.xlsx")
```

---

## PivotTable Creation Fails

**Symptom:** Layer 5 (VBA) PivotTable creation macro raises a runtime error.

**Common Causes:**

1. **Source range is invalid** — Ensure the data range includes headers and has no completely blank rows/columns in the middle.
2. **PivotTable name conflict** — A PivotTable with the same name already exists. Use a unique name.
3. **Destination conflict** — The destination cell overlaps with existing data.

**Solution:**

```vba
' Check if PivotTable already exists and delete it
Dim pt As PivotTable
For Each pt In ws.PivotTables
    If pt.Name = "SalesPivot" Then
        pt.TableRange2.Clear
        Exit For
    End If
Next pt
```

---

## Slow Execution / Long Delays

**Symptom:** The engine takes much longer than expected, especially with many checkpoints.

**Possible Causes:**

1. **Excel recalculation** — Large workbooks recalculate after every change.
2. **Layer escalation** — If lower layers consistently fail, time is wasted on retries.
3. **Screen recording overhead** — PyAutoGUI screenshots are expensive.

**Solutions:**

- Disable automatic calculation during bulk operations:
  ```python
  import xlwings as xw
  app = xw.apps.active
  app.calculation = "manual"
  # ... do work ...
  app.calculation = "automatic"
  ```
- Review logs to see which layers are being attempted and optimize the layer selection strategy.
- Avoid unnecessary escalation to Layer 6 by ensuring Layers 1–5 cover the operation.
