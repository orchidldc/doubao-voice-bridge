# Changelog

## v0.2.1 - 2026-06-25

- Added tail-anchor append detection for multi-account Feishu/Lark collaborative editing.
- Fixed false "non-append edit" resets when Feishu/Lark `raw_content` slightly rewrites earlier paragraph text but the old document tail still matches.
- Added near-tail context append recovery for older iPhone / iOS editing cases where the old document tail is lightly rewritten.
- Accepted the `needs_refresh` lark-cli user-auth state so a refreshable local login no longer makes the GUI bridge exit immediately.
- Added GUI "clear document body" action while preserving the Feishu/Lark page title.
- Made GUI bridge listening start automatically by default.
- Made the bridge automatically capture the initial baseline when watch mode starts, so opening the GUI no longer requires pressing `F8` once.
- Made GUI shutdown clean up the bridge background process tree.
- Replaced the app icon with a high-resolution Clash-style blue `D` icon.
- Added reviewable source code, build script, PyInstaller spec, requirements, and icon assets to the public repository.
- Kept strict fallback behavior for real middle edits, deletions, or unmatched context.

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
