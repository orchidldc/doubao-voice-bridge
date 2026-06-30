# DouBao Voice Bridge

[![中文说明](https://img.shields.io/badge/README-中文-red)](README.md)

A Windows GUI bridge for DouBao-style cross-screen voice input. Speak on your phone, write recognized text into a Feishu/Lark Docx document, and paste newly appended text into the current Windows input target.

This public repository includes reviewable source code, build scripts, documentation, and binary releases. It intentionally does not include private config, logs, screenshots, real document links, or account secrets.

## Download

Download the Windows x64 release from GitHub Releases:

[Download DouBao Voice Bridge](https://github.com/orchidldc/doubao-voice-bridge/releases/latest)

Current package:

```text
DouBaoVoiceBridge-v0.3.1-windows-x64.zip
```

## What It Does

The app uses a Feishu/Lark Docx document as a local bridge:

1. Open a dedicated Feishu/Lark Docx document on your phone.
2. Use DouBao or any mobile dictation/input method to write text into that document.
3. Run DouBao Voice Bridge on Windows.
4. Put the cursor in the target input field.
5. The app captures the startup baseline automatically.
6. The app pastes only text appended after startup or after the next `F8`.

It does not perform speech recognition, call DouBao APIs, scrape browser DOM, take screenshots, or upload local data.

## GUI

Since `v0.2.0`, the main program is a GUI app:

- save config,
- check Feishu/Lark connection,
- read once,
- test paste,
- clear document body while keeping the Feishu/Lark page title,
- start listening,
- stop listening,
- open config,
- open logs.

The underlying bridge CLI is embedded in the GUI binary resources, so users do not need to manage multiple executables. The current GUI starts the bridge process automatically on launch, captures the initial baseline automatically, and cleans up its background bridge process tree when the GUI closes.

## Quick Start

1. Extract `DouBaoVoiceBridge-v0.3.1-windows-x64.zip`.
2. Double-click `DouBaoVoiceBridge.exe`.
3. Fill in your Feishu/Lark Docx URL.
4. Make sure `lark-cli` is logged in:

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

5. Click "Save config".
6. Click "Check connection".
7. Listening starts automatically by default and captures the initial baseline automatically. If you stopped it manually, click "Start listening".
8. Put the cursor in the target input field.
9. Press `F8` only when you want to start again and capture a new baseline.
10. Continue writing into the Feishu/Lark document from your phone.

## Hotkeys

- `F8`: start and capture baseline
- `F9`: pause
- `F10`: reset baseline
- `F12`: exit the bridge process

## iOS / Multi-Device Editing Compatibility

On some mobile clients, especially older iPhones or cross-device editing sessions, Feishu/Lark `raw_content` may lightly rewrite the document tail or earlier paragraph text. `v0.2.1` adds three append-recovery layers:

- trailing-whitespace tolerance,
- tail-anchor append recovery,
- near-tail context append recovery.

Real middle-of-document edits, deletions, or unmatched context still reset the baseline instead of pasting guessed text.

## Build From Source

On Windows:

```powershell
python -m pip install -r requirements.txt
.\build_gui.bat
```

The GUI binary is generated at:

```text
dist\DouBaoVoiceBridge.exe
```

## Image Bridge

The image bridge is opt-in. When enabled, the tool detects newly inserted images from `lark-cli docs +fetch`, downloads them to a local temporary directory, and writes image data to the Windows clipboard.

Default image settings:

```json
{
  "enable_image_bridge": false,
  "image": {
    "enabled": false,
    "insert_mode": "clipboard_bitmap",
    "allow_file_drop_fallback": false
  }
}
```

`clipboard_bitmap` is the default mode for web AI input boxes, Word, chat boxes, and editors that support normal image paste. It writes CF_DIB/Bitmap data only and does not put the image path or `xxx.jpg` filename into the text clipboard.

`clipboard_file` is explicit file mode. Use it only when you manually set `image.insert_mode` to `clipboard_file`. It writes a FileDropList for applications that support file paste/upload, but some input boxes interpret that as a filename, so it is not the default.

`auto` is also bitmap-first. It does not fall back to FileDropList unless `allow_file_drop_fallback` is set to `true`.

If a target application cannot accept clipboard images, the tool will not paste the filename as a workaround. Use manual upload or explicit `clipboard_file` mode instead. For diagnostics:

```powershell
DouBaoVoiceBridge.exe --config config.json --debug-image-pipeline
DouBaoVoiceBridge.exe --inspect-clipboard
DouBaoVoiceBridge.exe --test-image-output --file C:\path\to\image.jpg --mode clipboard_bitmap
DouBaoVoiceBridge.exe --test-image-output --file C:\path\to\image.jpg --mode clipboard_file
```

## Safety

Do not use this tool in password fields, browser address bars, payment forms, admin consoles, or sensitive input targets.

Do not publish your real config file, Feishu/Lark document URL, tokens, App Secret, logs, or private screenshots.

## License / Rights

No open-source license is granted for the application source code or compiled application through this repository. See [NOTICE](NOTICE).
