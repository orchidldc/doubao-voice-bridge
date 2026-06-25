# Troubleshooting

## The Window Opens And Immediately Closes

Use the current release package. On first run, the app should create `config.json` from `config.example.json` and wait for input instead of closing immediately.

If it still closes, run from PowerShell:

```powershell
.\DouBaoVoiceBridge.exe
```

## `lark-cli` Not Found

Install `lark-cli`, add it to PATH, or set:

```json
{
  "lark_cli_path": "C:\\path\\to\\lark-cli.exe"
}
```

## CLI Not Logged In

Run:

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

## Nothing Is Pasted

Check:

- The cursor is inside a normal text input field.
- The target app accepts paste/input events.
- The app captured the startup baseline, or you pressed `F8` after putting the cursor in the target.
- The Feishu/Lark document text was appended after startup or after the latest `F8`.
- The foreground window has not changed from the captured target.

For apps such as VS Code, terminals, browsers, and Electron apps, focus behavior can vary. Start with Notepad to verify the paste path.

## Window Changed Warnings

This is a safety check. The app keeps pending text and waits until the target window is valid again.

If you want looser matching, adjust:

```json
{
  "target_window_mode": "process"
}
```

Use `any` only if you accept the risk of pasting into the wrong foreground window.

## iPhone Or Another Client Causes "Non-Append Edit" Warnings

Upgrade to the current `v0.2.1` release. It handles several Feishu/Lark `raw_content` rewrite cases:

- trailing whitespace shifts,
- old document tail still matching after earlier text changes,
- older iPhone / iOS clients lightly rewriting the near-tail paragraph while appending text.

If the warning still appears, reset the baseline with `F10`, then press `F8` again after placing the cursor in the target input field.

## Background Bridge Process Still Running

The current GUI stops the bridge process tree when you close the window and also cleans up known bridge process names on startup.

If an older build left an elevated background process behind, launch the current `DouBaoVoiceBridge.exe` once as administrator, then close it normally.
