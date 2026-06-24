# Release Notes - v0.1.0

Initial Windows x64 public binary release.

## Highlights

- One-file Windows executable.
- Feishu/Lark CLI login mode.
- Local baseline diff and paste workflow.
- Safer first-run behavior when `config.json` is missing.
- Window focus checks to reduce accidental paste risk.

## Upgrade Notes

If you used an earlier local build, replace only the executable and keep your private `config.json` outside the public repository.

Never upload your real `config.json` to GitHub.

## Known Limitations

- Only Docx raw text is supported.
- Wiki links are not supported by this binary.
- Some desktop apps may reject synthetic paste events.
- Global hotkeys may fail on locked-down Windows environments.

