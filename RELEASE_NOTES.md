# Release Notes - v0.2.0

DouBao Voice Bridge now ships as a focused Windows GUI app.

## Highlights

- Public product name changed to DouBao Voice Bridge.
- New GUI for config, connection checks, single read, paste test, and start/stop listening.
- Only one user-facing executable in the release package: `DouBaoVoiceBridge.exe`.
- Embedded bridge CLI resource, so users do not need to manage multiple exe files.
- Smart leading-newline handling reduces unwanted blank lines before pasted voice fragments.
- Default README is Chinese, with an English README button.

## Upgrade Notes

If you used `v0.1.0`, switch to the new GUI executable and keep your private config out of GitHub.

Never upload your real config file to GitHub.

## Known Limitations

- Only Docx raw text is supported.
- Wiki links are not supported by this binary.
- Some desktop apps may reject synthetic paste events.
- Global hotkeys may fail on locked-down Windows environments.
