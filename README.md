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
3. **Rows** — press **+ Add row** per record (or just scan into the first
   card). Inside each row press the small **+** to scan document images —
   select as many as you like at once. Scans are **grouped by entity**, for
   any kind of sheet and any kind of document: records that share an
   ID-like value (passport, iqama, national ID, plate, invoice no, …) or a
   matching name merge into the same row; the AI also tags every record with
   who/what it is about, so sheets without name/ID columns group correctly
   too. Only an unrecognized entity opens a new row — e.g. 20 pictures of 18
   students produce 18 rows. Phone numbers are never used for matching
   (siblings share a parent's mobile). A photo that shows several people
   returns one record per person. Fields still empty are flagged
   **missing** — scan another document for that entity, or press **ignore**.
4. **Preview & download** — Preview renders the sheet as it will look after
   download: the original title/header rows (with merged cells), the existing
   data rows, then the new rows highlighted. New rows are **appended after the
   existing data** (first empty row) — existing rows are never overwritten.

## Robustness

- **Free-tier quota.** Gemini's free tier allows only ~15–20 requests per
  minute (plus a daily cap). The app sends up to 3 pictures per request,
  spaces requests ≥4 s apart, waits the server-suggested delay and retries on
  `429`, and on a persistent quota error falls back to sibling models
  (`gemini-2.5-flash-lite`, `gemini-2.0-flash`, …) that have separate quota
  buckets. Pictures that still fail are outlined red — just add them again.
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
