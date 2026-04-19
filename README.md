# PWA Data Extractor

Desktop utility for extracting measurements from PWA detailed report PDFs, reviewing repeated measurements, and exporting averaged patient data to Excel.

## What This Is For

This tool is used when you have one or more PWA detailed report PDFs and want to turn them into a structured Excel workbook.

The main use case is not just raw extraction. The app is designed to:

- extract measurement fields from PWA detailed reports
- group repeated entries by patient
- choose the best two measurements to keep for averaging
- let you manually review patients with more than two entries
- export both the detailed rows and the final averaged patient rows

This is useful when a patient has multiple PWA recordings and you want a cleaner final dataset built from the best matching pair.

## Expected Input

Use PWA **Detailed Reports** in PDF format.

The app will also detect:

- **Clinical reports**
- **unrecognized PDFs**

Those are still surfaced in the export as special rows so they are visible, but they are not treated as normal detailed-report data for averaging.

## Core Workflow

1. Add one or more PWA detailed report PDFs.
2. Choose where the Excel workbook should be saved.
3. Process the PDFs locally.
4. Review any patients with more than two entries.
5. Export the workbook.

## How Patient Averaging Works

If a patient has exactly two valid entries, that pair is used automatically.

If a patient has more than two valid entries, the app preselects an automatic pair and sends that patient to the **Multi-entry review** tab.

In review, you choose exactly two rows to keep for export. The app also shows:

- the currently selected pair
- absolute differences between the selected rows
- pair alerts when the selected measurements differ beyond the configured threshold

## Auto-Pairing Logic

The current pairing behavior is preserved from the original tool.

By default, automatic pairing is based on:

- **Peripheral Systolic Pressure (mmHg)**

When a patient has more than two entries, the app checks all possible 2-row combinations and selects the closest pair based on the configured analysis mode.

## Output Workbook

The export preserves the workbook structure used by the original tool:

- `All Data`
- `Kept Data`
- `Averaged Data`

### Sheet Summary

`All Data`

- all parsed detailed-report rows
- special rows for clinical or unrecognized PDFs
- review status and pairing context

`Kept Data`

- the rows that were selected to be kept for averaging

`Averaged Data`

- one averaged row per patient based on the final selected pair
- pair-difference and alert-related fields used for downstream review

## Review Features

The desktop UI includes:

- overview of processed files and review counts
- multi-entry review queue
- pair selection with keep checkboxes
- absolute-difference display for the selected pair
- configurable highlight thresholds
- configurable pair-alert threshold
- PDF preview for the selected source file

## Threshold Settings

The settings panel lets you adjust:

- green highlight threshold
- yellow highlight threshold
- pair-alert threshold

These settings affect the review display and alert behavior for the current session.

## Run Locally

```powershell
.\.venv\Scripts\python.exe -m pip install -r .\requirements.txt
.\.venv\Scripts\python.exe .\pwa_extractor.py
```

## Build Executable

```powershell
.\.venv\Scripts\pyinstaller.exe .\pwa_extractor.spec
```

## Project Layout

- `pwa_extractor.py`: launcher entry point
- `app.py`: PySide6 desktop user interface
- `backend.py`: PDF parsing, pairing logic, and Excel export
- `pwa_extractor.spec`: PyInstaller build spec

## Notes

- Processing is local.
- Detailed reports are the intended source files.
- Clinical reports and unrecognized PDFs remain visible in export as special rows.
- The app is built to preserve the existing extraction and pairing behavior while improving review workflow and usability.

## Contact

- Email: `thomaswhart28@gmail.com`
