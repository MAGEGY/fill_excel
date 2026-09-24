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
3. **Rows** — press **+ Add row** per record. Inside each row press the small
   **+** to scan a document image. Extracted values fill the row's fields.
   Fields still empty are flagged **missing** — scan another document for the
   same row, or press **ignore** on that field.
4. **Preview & download** — check the table, then download the filled workbook.

## Output formatting

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
