# Excel Engine — Chrome Extension

Manifest V3 Chrome extension for Excel Engine automation.

## Features

- **File Upload** — Select .xlsx workbooks and instruction files
- **Resource Files** — Attach additional data files
- **Text Instructions** — Type instructions directly
- **Engine Options** — Layer selection, dry run mode
- **Settings** — Configurable server URL and API key
- **Real-time Feedback** — Progress bar and result display

## Installation (Developer Mode)

1. Open Chrome → `chrome://extensions/`
2. Enable "Developer mode" (top right toggle)
3. Click "Load unpacked"
4. Select the `chrome-extension/` folder
5. The Excel Engine icon appears in your toolbar

## Usage

1. Click the Excel Engine icon in Chrome toolbar
2. Select your .xlsx workbook
3. Select instruction file or type instructions
4. (Optional) Add resource files and adjust options
5. Click "Run Engine"
6. View results

## Server Requirement

The extension communicates with a local Excel Engine MCP server:

```bash
# Start the server
cd copilot-studio-connector/server
python app.py  # http://localhost:8080
```

Configure the server URL in Settings (link at bottom of popup).

## Chrome Web Store Publishing

To publish to the Chrome Web Store:

1. Create a developer account at https://chrome.google.com/webstore/devconsole
2. Pay the one-time $5 developer fee
3. Zip the extension:
   ```bash
   cd chrome-extension
   zip -r ../dist/excel-engine-chrome.zip . -x ".*"
   ```
4. Upload the zip at the developer console
5. Fill in store listing details and submit for review

## Permissions

- `storage` — Save server URL and API key settings
- `http://localhost:8080/*` — Communicate with local MCP server
