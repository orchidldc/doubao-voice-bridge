# Feishu Voice Bridge

Windows local bridge for sending new text from a Feishu/Lark Docx document into the currently selected desktop input field.

This public repository is a distribution and documentation repository. It intentionally does not contain the application source code, private configuration, logs, screenshots with personal data, or any Feishu/Lark credentials.

## What It Does

- Reads plain text from one configured Feishu/Lark Docx document.
- Uses the document content at startup as a baseline.
- Detects only text appended after the baseline.
- Pastes the new text into the foreground Windows input target after you press `F8`.
- Supports Feishu/Lark CLI user login by default, so the app does not store your account token.

The tool does not perform speech recognition. You can use any mobile input method or dictation tool to write into the configured Feishu/Lark document.

## Download

Download the Windows x64 release package from the GitHub Releases page:

```text
FeishuVoiceBridge-v0.1.0-windows-x64.zip
```

The release package should contain:

- `feishu_voice_bridge.exe`
- `config.example.json`
- `README-USER.md`
- `THIRD_PARTY_NOTICES.txt`
- `DISCLAIMER.txt`
- `start.bat`

Do not download or run files from unofficial mirrors.

## Quick Start

1. Install and log in to `lark-cli`.
2. Extract the release zip into a local folder.
3. Copy `config.example.json` to `config.json`.
4. Edit `config.json` and set `doc_url` to your own Feishu/Lark Docx link.
5. Double-click `feishu_voice_bridge.exe`, or run `start.bat`.
6. Put the cursor in your target input field.
7. Press `F8` to start listening, `F9` to pause, `F10` to reset the baseline, and `F12` to exit.

CLI login example:

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

## Security Model

The recommended mode is:

```json
{
  "auth_mode": "lark_cli"
}
```

In this mode the app calls your local `lark-cli` and reads as the logged-in user. The app does not ask you to paste account tokens into its config file.

The alternate `app` mode reads `FEISHU_APP_ID` and `FEISHU_APP_SECRET` from environment variables. Do not put secrets in `config.json`, GitHub issues, screenshots, logs, or chat messages.

## Privacy

The app is designed to run locally on Windows. It reads the configured Feishu/Lark document through Feishu/Lark APIs or CLI, compares text locally, and pastes new text into your selected application.

The public repository does not include telemetry, analytics configuration, or hosted backend code.

## Important Limits

- This tool only supports Feishu/Lark Docx document links or document IDs.
- It does not support old Doc, Wiki, Sheet, Base, or browser DOM scraping.
- Some applications reject synthetic paste/input events. In that case, try another target input field or run the app with appropriate Windows permissions.
- Do not use this tool in password fields, browser address bars, payment forms, or any sensitive input field.

## Documentation

- [Quick Start](docs/quick-start.md)
- [Feishu/Lark Setup](docs/setup-feishu.md)
- [Troubleshooting](docs/troubleshooting.md)
- [Design Notes](docs/design.md)
- [Security Policy](SECURITY.md)
- [Privacy Notice](PRIVACY.md)
- [Disclaimer](DISCLAIMER.md)

## License / Rights Notice

No open-source license is granted for the application itself in this public repository. See [NOTICE](NOTICE).

