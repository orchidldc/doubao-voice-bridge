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
- You pressed `F8` after putting the cursor in the target.
- The Feishu/Lark document text was appended after `F8`.
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
