# Quick Start

## 1. Prepare Feishu/Lark CLI

Install `lark-cli`, then log in as the account that can read your target Docx document:

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

## 2. Prepare Config

Extract the release zip, then start the GUI:

```powershell
.\DouBaoVoiceBridge.exe
```

Fill in your Feishu/Lark Docx URL in the GUI and click "Save config". The saved private config file is named:

```text
doubao_voice_bridge_config.json
```

Do not commit or publish this file.

## 3. Run

Double-click:

```text
DouBaoVoiceBridge.exe
```

## 4. Use

- Listening starts automatically by default. If you stopped it manually, click "Start listening".
- Put the cursor in the target input field.
- Press `F8` to capture the baseline.
- Continue writing into the configured Feishu/Lark document from your phone.
- New appended text will be pasted into the target input field.
- Press `F9` to pause, `F10` to reset, `F12` to exit.

## 5. Optional Cleanup

The GUI includes a "clear document body" action. It removes body blocks from the configured Feishu/Lark Docx document and keeps the page title.

Use it only after confirming the configured document is the dedicated bridge document.
