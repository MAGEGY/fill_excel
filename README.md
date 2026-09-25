# Fill Excel from Documents

A client-side web app that fills any Excel sheet from photos/scans of documents
(passports, iqamas, IDs, …) using an AI vision model.

## How it works

1. **AI engine** — pick a provider:
   - **Google Gemini** (default; free key at [aistudio.google.com](https://aistudio.google.com))
   - **OpenAI-compatible** endpoint (any `chat/completions` API that accepts image input)
   - **Offline OCR** (Tesseract.js, no key — works for standard passport/iqama fields)
2. **Excel** — upload any `.xlsx`. The header row is auto-detected; every column
   becomes a field to fill. Uncheck columns you want to skip. `No.`/serial
   columns are auto-numbered.
2. **Excel** — upload any `.xlsx`. The header row is auto-detected (merged
   banner/title rows are ignored); every column becomes a field to fill.
   Uncheck columns you want to skip. `No.`/serial columns keep any prefilled
   numbers and continue them for new rows.
3. **Rows** — press **+ Add row** per record. Inside each row press the small
   **+** to scan a document image. Extracted values fill the row's fields.
   A photo that shows several people/lines returns one record per person —
   extra records automatically create new rows. Once a row is full, the next
   scanned document starts a new row automatically. Fields still empty are
   flagged **missing** — scan another document for the same row, or press
   **ignore**.
4. **Preview & download** — Preview renders the sheet as it will look after
   download: the original title/header rows (with merged cells), the existing
   data rows, then the new rows highlighted. New rows are **appended after the
   existing data** (first empty row) — existing rows are never overwritten.

## Robustness

- Templates saved by WPS Office embed drawings with default-namespace XML that
  crashes ExcelJS; the app detects this and reloads the file without the
  embedded images so the data still loads.

## Output formatting

- Extracted cells use the template's own font per column (copied from existing
  data cells, else the header font; never smaller than 12) — no more tiny 11pt
  text in tall rows.
- Dates are written as real dates formatted `dd/mm/yyyy`.
- Cells containing Arabic are right-aligned (RTL); Latin text left-aligned (LTR).
- ID/phone columns are stored as text (leading zeros preserved).
- Font: Arial 11.

## Run it

Serve the folder with any static server, e.g.:

```bash
python -m http.server 8765
# open http://localhost:8765
```

Everything runs in the browser — no backend. Your API key stays in
`localStorage`; files never leave your machine except the image bytes sent
to the chosen AI provider.

## Files

- `index.html` — UI
- `app.js` — Excel parsing (ExcelJS), AI/OCR extraction, styled export
- `styles.css`
