# PWA Data Extractor

Desktop utility for extracting measurements from PWA PDF reports and exporting them to Excel.

## Current app

The repository now uses a `PySide6` desktop UI with the existing PWA parsing and export logic moved into a reusable backend.

## Workflow

1. Add one or more PWA PDF reports.
2. Choose the Excel export path.
3. Process the reports locally.
4. Review patients with more than two entries and adjust the selected pair if needed.
5. Export the workbook.

## Output workbook

The export preserves the original workbook structure:

- `All Data`
- `Kept Data`
- `Averaged Data`

## Run locally

```powershell
.\.venv\Scripts\python.exe -m pip install -r .\requirements.txt
.\.venv\Scripts\python.exe .\pwa_extractor.py
```

## Build executable

```powershell
.\.venv\Scripts\pyinstaller.exe .\pwa_extractor.spec
```

## Project layout

- `pwa_extractor.py`: launcher entry point
- `app.py`: PySide6 user interface
- `backend.py`: PDF parsing, pairing logic, and Excel export

## Notes

- Detailed reports are processed normally.
- Clinical reports and unrecognized PDFs are still surfaced as special rows in the export.
- The current pairing behavior is preserved: matching is based on peripheral systolic pressure by default.

## Contact

- Email: `thomaswhart28@gmail.com`
