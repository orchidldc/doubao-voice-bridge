# Release Notes - v0.2.1

Stable GUI patch release for DouBao Voice Bridge.

## Highlights

- Improved iOS and multi-device append detection when Feishu/Lark `raw_content` lightly rewrites the document tail.
- Added near-tail context append recovery in addition to tail-anchor detection.
- Added a GUI action to clear document body blocks while keeping the Feishu/Lark page title.
- The GUI now starts the bridge process automatically by default.
- The bridge now captures the startup baseline automatically, so opening the GUI no longer requires pressing `F8` once.
- Closing the GUI now cleans up the bridge background process tree.
- A local lark-cli login in the `needs_refresh` state is now accepted instead of being treated as logged out.
- Replaced the app icon with a high-resolution Clash-style blue `D` icon.
- Published reviewable source code and build scripts in the repository.

## Upgrade Notes

Upgrade to this version if you use an iPhone, multiple mobile devices, or multiple Feishu/Lark sessions to edit the same bridge document.

Never upload your real config file to GitHub.

## Known Limitations

- Only Docx raw text is supported.
- Wiki links are not supported by this binary.
- Some desktop apps may reject synthetic paste events.
- Global hotkeys may fail on locked-down Windows environments.
