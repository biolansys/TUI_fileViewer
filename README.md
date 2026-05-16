# Files Viewer TUI

Textual TUI app to browse, view, and edit:
- CSV (`.csv`, `.tsv`)
- Parquet (`.parquet`, `.parq`)
- ORC (`.orc`)
- Avro (`.avro`)
- Arrow/Feather (`.arrow`, `.feather`)
- Excel (`.xlsx`, `.xls`)
- JSON (`.json`)
- XML (`.xml`)
- HTML (`.html`, `.htm`)
- Markdown (`.md`, `.markdown`)
- PDF (`.pdf`)
- Images (`.gif`, `.png`, `.jpg`, `.jpeg`, `.bmp`, `.webp`, `.tif`, `.tiff`)

## Run

```bash
pip install -e .
files-viewer
```

## Controls

- `Ctrl+O`: focus file tree
- `Backspace`: move explorer to parent directory
- `Alt+Left`: previous Excel sheet
- `Alt+Right`: next Excel sheet
- `Ctrl+R`: open current HTML file in default browser
- `Ctrl+T`: convert current table to another format
- `Ctrl+S`: save current file
- `e`: edit selected table cell
- `q`: quit

## Notes

- Table files are shown in a grid and can be edited cell-by-cell.
- JSON is shown as pretty text in `Text` and as hierarchy in `JSON Tree`.
- For multi-sheet Excel files, the app loads the first sheet and lets you switch sheets with `Alt+Left/Alt+Right`.
- PDF files are shown as extracted text preview (bounded pages/chars).
- HTML files are rendered into semantic text preview (headings, lists, links, tables).
- Markdown files are shown as raw source in `Text` and rendered in `Markdown` tab.
- Image files render in an ANSI-art preview in the `Image` tab.
- Conversion supports: `csv`, `tsv`, `parquet`, `parq`, `arrow`, `feather`, `orc`, `json`, `avro`.