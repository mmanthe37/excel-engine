# Excel Engine — User Guide

## What is Excel Engine?

Excel Engine is a tool that automatically completes Excel assignments for you. Give it your Excel file and instructions, and it does the work.

**Works on:** macOS, Windows, and Linux

---

## How to Use (The Easy Way)

### Step 1: Download
Download the Excel Engine folder from [GitHub](https://github.com/mmanthe37/excel-engine).
Click the green **"Code"** button → **"Download ZIP"** → Unzip the folder.

### Step 2: Open the App

| Your Computer | What to Double-Click |
|---------------|---------------------|
| **Mac** | Open the `gui` folder → double-click **`launch.command`** |
| **Windows** | Open the `gui` folder → double-click **`launch.bat`** |
| **Linux** | Open the `gui` folder → double-click **`launch.sh`** (or right-click → "Run as Program") |

> **First time only:** The app will set itself up automatically. This takes 1-2 minutes. After that, it opens instantly.

### Step 3: Use the App

1. **Upload your Excel file** — Click "Browse files" under "Excel Workbook" and select your `.xlsx` file
2. **Upload your instructions** — Click "Browse files" under "Instructions" and select your instruction file (`.txt`, `.docx`, or `.pdf`)  
   *Or* just paste the instructions into the text box
3. **Click "🚀 Run Excel Engine"**
4. **Download your completed file** — Click the download button when it's done!

---

## Troubleshooting

### "Python is not installed"
The app needs Python to run. Download it free from [python.org/downloads](https://www.python.org/downloads/).

**Windows users:** When installing Python, make sure to check ✅ **"Add Python to PATH"** at the bottom of the installer!

### The app won't open / "Apple could not verify" (Mac)
- **Mac standalone .app:** macOS Gatekeeper blocks unsigned apps. To open it:
  1. **Right-click** (or Control-click) the app → click **"Open"**
  2. Click **"Open"** again in the confirmation dialog
  3. It will open normally from now on
  - *Alternative:* Go to **System Settings → Privacy & Security** → click **"Open Anyway"** next to the blocked app message
- **Mac launcher:** You may need to right-click → "Open" the first time
- **Windows:** If Windows Defender blocks it, click "More info" → "Run anyway"
- **Linux:** You may need to make it executable: right-click → Properties → Permissions → "Allow executing"

### Something went wrong during execution
- Try switching to "Standard Mode (works everywhere)" in the sidebar
- Make sure your Excel file is a `.xlsx` file (not `.xls` or `.csv`)
- Check that your instruction file clearly describes what to do

---

## Standalone App (No Python Needed)

Download pre-built versions from the [Releases](https://github.com/mmanthe37/excel-engine/releases) page:
- **Mac:** `ExcelEngine-macOS.zip` — unzip, then **right-click → Open** the first time (macOS blocks unsigned apps)
- **Windows:** `ExcelEngine-Windows.zip` — unzip and double-click `Excel Engine.exe` (click "Run anyway" if prompted)
- **Linux:** `ExcelEngine-Linux.tar.gz` — extract and run `./Excel Engine`

---

## Need Help?

Open an issue on [GitHub](https://github.com/mmanthe37/excel-engine/issues) or check the [main documentation](../README.md).
