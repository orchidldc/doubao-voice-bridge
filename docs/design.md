# Design Notes

## Goal

The tool bridges text already recognized by a mobile input method into a Windows desktop input target.

It intentionally avoids browser automation, DOM scraping, screenshots, OCR, and speech recognition.

## Flow

```text
Mobile input method
        |
        v
Feishu/Lark Docx document
        |
        v
Local Windows executable
        |
        v
Current desktop input target
```

## Baseline Delta

When the user presses `F8`, the app reads the configured document and stores its current text as a baseline.

During polling, only newly appended text is prepared for paste. If the existing document content changes in a non-append way, the app resets the baseline instead of guessing.

## Target Safety

The app captures foreground window metadata at `F8` time. Before pasting, it checks the current foreground window against the selected mode:

- `locked`: same window only.
- `process`: same process is accepted.
- `any`: current foreground window is accepted.

The safest default is `locked`.

