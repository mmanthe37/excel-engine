# Excel Engine — iOS App

Native SwiftUI iOS client for Excel Engine.

## Architecture

The iOS app is a **client** that communicates with your Excel Engine MCP server backend. The Python engine runs on a server (your Mac, cloud, etc.), and the iOS app sends requests via the MCP HTTP protocol.

```
┌──────────────┐     HTTP/JSON-RPC     ┌──────────────────┐
│  iOS App     │ ───────────────────→  │  MCP Server      │
│  (SwiftUI)   │ ←─────────────────── │  (Python/Go)     │
└──────────────┘                       └──────────────────┘
```

## Features

- **File Picker** — Select .xlsx workbooks from Files app
- **Multi-format Instructions** — Upload .docx, .pdf, .txt, or type directly
- **Resource Files** — Attach additional data files
- **Real-time Progress** — Watch engine processing status
- **Share Results** — Share completed workbooks via iOS share sheet
- **Configurable** — Server URL, API key, layer selection, dry run mode

## Building

### Prerequisites
- Xcode 15+ (or Xcode 16 beta)
- iOS 16.0+ deployment target
- Apple Developer account (for TestFlight)

### Build for Simulator
```bash
cd ios-app
xcodegen generate
xcodebuild -scheme ExcelEngine -sdk iphonesimulator -configuration Release
```

### Build for TestFlight

1. **Open in Xcode:**
   ```bash
   cd ios-app
   xcodegen generate  # if ExcelEngine.xcodeproj doesn't exist
   open ExcelEngine.xcodeproj
   ```

2. **Fix code signing:**
   - Select the ExcelEngine target → Signing & Capabilities
   - Enable "Automatically manage signing"
   - Select your team: "Michael Manthe (25QSUNYFC9)"
   - Xcode will create/download provisioning profiles

3. **Fix certificate trust (if needed):**
   - Open Keychain Access → login keychain
   - Find "Apple Development: michaelamanthe2@gmail.com"
   - Double-click → Trust → set "When using this certificate" to "Always Trust"
   - Or reset: `security trust-settings-import -d /System/Library/Security/Certificates.pem`

4. **Archive:**
   - Product → Archive (or Cmd+Shift+B with Release scheme)
   - Select "Distribute App" → "TestFlight Internal Only" or "App Store Connect"

5. **Upload to TestFlight:**
   - From the Archive organizer, click "Distribute App"
   - Choose "App Store Connect" → Upload
   - Or use command line:
     ```bash
     xcodebuild -scheme ExcelEngine -sdk iphoneos archive -archivePath build/ExcelEngine.xcarchive
     xcodebuild -exportArchive -archivePath build/ExcelEngine.xcarchive -exportPath build/ -exportOptionsPlist exportOptions.plist
     xcrun altool --upload-app -f build/ExcelEngine.ipa -u YOUR_APPLE_ID -p APP_SPECIFIC_PASSWORD
     ```

6. **App Store Connect:**
   - Go to https://appstoreconnect.apple.com
   - The build will appear under TestFlight after processing (~15 min)
   - Add test users and enable internal/external testing

## Server Setup

The iOS app needs a running Excel Engine server to connect to:

```bash
# Option 1: Copilot Studio Connector (recommended)
cd copilot-studio-connector/server
pip install -r requirements.txt
python app.py  # runs on http://localhost:8080

# Option 2: MCP Server
cd mcp-server
python server.py
```

For remote access, expose the server via:
- **ngrok**: `ngrok http 8080`
- **Tailscale**: Access your Mac's IP from your phone
- **Cloud deployment**: Deploy the connector to any cloud provider

## Configuration

In the iOS app Settings (gear icon):
- **Server URL**: Your MCP server endpoint (e.g., `http://192.168.1.x:8080` for local)
- **API Key**: Optional, matches `EXCEL_ENGINE_API_KEY` on server
- **Max Layer**: 1-6 (Layer 1 works everywhere, Layers 2-6 need macOS with Excel)
- **Dry Run**: Plan without executing
- **Recalculate**: Use LibreOffice for formula recalculation

## Bundle ID

`com.michaelmanthe.excel-engine`

Team ID: `25QSUNYFC9`
