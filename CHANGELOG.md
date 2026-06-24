# Changelog

## v0.2.0 - 2026-06-24

- Renamed the public product to DouBao Voice Bridge.
- Added a focused Windows GUI app for the DouBao cross-screen input workflow.
- Embedded the bridge CLI into the GUI binary resources.
- Changed the release package to expose only one user-facing executable: `DouBaoVoiceBridge.exe`.
- Changed the default README to Chinese and added an English README button.
- Added smart leading-newline handling for voice-input fragments.

## v0.1.0 - 2026-06-24

- Added Windows one-file executable packaging.
- Added Feishu/Lark CLI user authentication mode.
- Added config validation and non-crashing first-run behavior.
- Added Docx raw text polling and baseline delta paste workflow.
- Added hotkeys: `F8` start, `F9` pause, `F10` reset baseline, `F12` exit.
- Added window-target safety modes: `locked`, `process`, and `any`.
- Added trailing whitespace tolerance for Feishu/Lark raw content changes.
- Added public release documentation without source code or private config.
