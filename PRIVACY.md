# Privacy Notice

This tool is a local Windows utility.

## Data Access

The app reads plain text from the single Feishu/Lark Docx document configured in `config.json`.

Depending on your configuration, access is performed through either:

- local `lark-cli` user authentication, or
- Feishu/Lark app credentials stored in environment variables.

## Local Processing

Text comparison, baseline tracking, and paste preparation happen locally on your computer.

The app is not designed to send document text to a separate backend service.

## Logs

Runtime logs are intended for debugging and should only contain status, lengths, hashes, target-window metadata, and errors. Do not publish logs without reviewing them.

## User Responsibility

You are responsible for:

- selecting the correct document,
- selecting the correct Windows input target,
- protecting credentials and local config files,
- complying with your organization's Feishu/Lark and data handling policies.

