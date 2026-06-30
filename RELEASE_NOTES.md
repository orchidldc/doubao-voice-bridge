# Release Notes - v0.3.1

Image paste repair release for DouBao Voice Bridge.

## Highlights

- Added image bridge support for Feishu/Lark Docx documents.
- Newly inserted document images can now be downloaded through `lark-cli` and pasted into the current Windows target as real bitmap clipboard data.
- Fixed the bug where image filenames such as `xxx.jpg` or `xxx.jpeg` could be pasted as text instead of the image.
- The text delta pipeline now strips media spans before comparing baselines, even when image bridge mode is disabled.
- Default image mode is `clipboard_bitmap`, with FileDrop fallback disabled by default.
- Added a GUI toggle for enabling image cross-screen insertion.
- Added diagnostics for image parsing and Windows clipboard formats.

## Upgrade Notes

Use this version if you want image cross-screen insertion, or if older builds pasted image filenames instead of images.

For image insertion, enable the image bridge in the GUI or set:

```json
{
  "enable_image_bridge": true,
  "image": {
    "enabled": true,
    "insert_mode": "clipboard_bitmap",
    "allow_file_drop_fallback": false
  }
}
```

Existing text-only workflows continue to work. Image filenames are no longer allowed to enter the text paste path.

Never upload your real config file to GitHub.

## Known Limitations

- Image bridge requires `auth_mode=lark_cli`.
- Only Feishu/Lark Docx image tags visible through `lark-cli docs +fetch` are supported.
- Some target apps may reject bitmap clipboard paste; use manual upload or explicit `clipboard_file` mode for those apps.
- Wiki links are not supported by this binary.
- Global hotkeys may fail on locked-down Windows environments.
