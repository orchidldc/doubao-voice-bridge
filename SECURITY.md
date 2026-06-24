# Security Policy

## Supported Version

Only the latest published Windows release is supported for security fixes.

## Report A Security Issue

Do not post secrets, document URLs, tokens, screenshots with private data, or personal information in public GitHub issues.

For a private report, contact the repository owner through the private channel listed on the GitHub profile or release page.

## Sensitive Data Rules

Never share:

- `FEISHU_APP_SECRET`
- Feishu/Lark access tokens or refresh tokens
- Real `config.json`
- Real Feishu/Lark document URLs
- Logs containing document titles or window titles
- Screenshots containing private text

## Local Execution Risks

This app can paste text into the active Windows input target. You are responsible for choosing the correct target before pressing `F8`.

Do not use it in password fields, address bars, payment pages, admin consoles, or any place where accidental paste could cause harm.

