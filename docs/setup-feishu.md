# Feishu/Lark Setup

## Recommended: CLI User Login

The recommended mode is `auth_mode=lark_cli`.

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

If the executable cannot find `lark-cli`, set the full path:

```json
{
  "lark_cli_path": "C:\\path\\to\\lark-cli.exe"
}
```

If you use profiles:

```json
{
  "lark_cli_profile": "your-profile-name"
}
```

## Alternate: App Credential Mode

Only use this if you understand Feishu/Lark app permissions.

Set environment variables:

```powershell
$env:FEISHU_APP_ID = "cli_xxxxxxxxxxxxx"
$env:FEISHU_APP_SECRET = "replace_with_your_secret"
```

Then set:

```json
{
  "auth_mode": "app"
}
```

Do not put `FEISHU_APP_SECRET` in `config.json`.

## Document Link

Supported:

```text
https://xxx.feishu.cn/docx/xxxxxxxx
```

Not supported in this release:

- Wiki links
- Sheet links
- Base links
- old Doc links
- arbitrary browser pages

