# Release Notes - v0.2.1

Patch release for multi-account Feishu/Lark collaborative editing.

## Highlights

- Added tail-anchor append detection.
- When another account edits the same document and Feishu/Lark rewrites earlier `raw_content`, the bridge can still paste text appended after the old document tail.
- Real middle edits, deletions, or unmatched tails still reset the baseline instead of pasting guessed text.

## Upgrade Notes

If you use a second phone or another Feishu/Lark account to edit the same bridge document, upgrade to this version.

Never upload your real config file to GitHub.

## Known Limitations

- Only Docx raw text is supported.
- Wiki links are not supported by this binary.
- Some desktop apps may reject synthetic paste events.
- Global hotkeys may fail on locked-down Windows environments.
