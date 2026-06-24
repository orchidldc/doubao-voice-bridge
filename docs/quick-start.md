# Quick Start

## 1. Prepare Feishu/Lark CLI

Install `lark-cli`, then log in as the account that can read your target Docx document:

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

## 2. Prepare Config

Extract the release zip, then copy:

```powershell
Copy-Item config.example.json config.json
```

Edit `config.json`:

```json
{
  "auth_mode": "lark_cli",
  "doc_url": "https://xxx.feishu.cn/docx/xxxxxxxx"
}
```

Use your own Docx link. Do not commit this file.

## 3. Run

Double-click:

```text
feishu_voice_bridge.exe
```

Or run:

```powershell
.\start.bat
```

## 4. Use

- Put the cursor in the target input field.
- Press `F8` to start and capture the baseline.
- Continue writing into the configured Feishu/Lark document from your phone.
- New appended text will be pasted into the target input field.
- Press `F9` to pause, `F10` to reset, `F12` to exit.

