# Studentlist Web App

This web app lets you clean and export student Excel lists directly in the browser.

## What it does

- Uploads an `.xlsx` file (first worksheet is used)
- Detects the header row and loads all columns
- Lets you choose which columns to keep
- Lets you filter rows by `Studium` or `Hörerstatus`
- Shows a live preview count of rows after filtering
- Exports a new Excel file with:
  - selected columns
  - filtered rows
  - two export modes:
    - **Studierendenzentriert**: merges duplicate students and combines multiple `Studium` values
    - **Statistikzentriert**: keeps separate rows and replaces `Matrikelnummer` with random short IDs
  - adjusted column widths for readability
  - Excel auto-filter enabled so columns are sortable/filterable

## Privacy

All processing happens in your browser. No file upload to a server is required.

## Run locally

```bash
npm install
npm run dev
```

Then open the local Vite URL shown in the terminal.

## Build

```bash
npm run build
```
