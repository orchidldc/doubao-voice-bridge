# DouBao Voice Bridge

[![中文说明](https://img.shields.io/badge/README-中文-red)](README.md)

A Windows GUI bridge for DouBao-style cross-screen voice input. Speak on your phone, write recognized text into a Feishu/Lark Docx document, and paste newly appended text into the current Windows input target.

This public repository is for documentation and binary distribution. It intentionally does not include source code, private config, logs, screenshots, real document links, or account secrets.

## Download

Download the Windows x64 release from GitHub Releases:

[Download DouBao Voice Bridge](https://github.com/orchidldc/doubao-voice-bridge/releases/latest)

Current package:

```text
DouBaoVoiceBridge-v0.2.0-windows-x64.zip
```

## What It Does

The app uses a Feishu/Lark Docx document as a local bridge:

1. Open a dedicated Feishu/Lark Docx document on your phone.
2. Use DouBao or any mobile dictation/input method to write text into that document.
3. Run DouBao Voice Bridge on Windows.
4. Put the cursor in the target input field.
5. Press `F8` to capture the baseline.
6. The app pastes only text appended after `F8`.

It does not perform speech recognition, call DouBao APIs, scrape browser DOM, take screenshots, or upload local data.

## GUI

Since `v0.2.0`, the main program is a GUI app:

- save config,
- check Feishu/Lark connection,
- read once,
- test paste,
- start listening,
- stop listening,
- open config,
- open logs.

The underlying bridge CLI is embedded in the GUI binary resources, so users do not need to manage multiple executables.

## Quick Start

1. Extract `DouBaoVoiceBridge-v0.2.0-windows-x64.zip`.
2. Double-click `DouBaoVoiceBridge.exe`.
3. Fill in your Feishu/Lark Docx URL.
4. Make sure `lark-cli` is logged in:

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

5. Click "Save config".
6. Click "Check connection".
7. Click "Start listening".
8. Put the cursor in the target input field.
9. Press `F8` to capture the baseline.
10. Continue writing into the Feishu/Lark document from your phone.

## Hotkeys

- `F8`: start and capture baseline
- `F9`: pause
- `F10`: reset baseline
- `F12`: exit the bridge process

## Safety

Do not use this tool in password fields, browser address bars, payment forms, admin consoles, or sensitive input targets.

Do not publish your real config file, Feishu/Lark document URL, tokens, App Secret, logs, or private screenshots.

## License / Rights

No open-source license is granted for the application source code or compiled application through this repository. See [NOTICE](NOTICE).

