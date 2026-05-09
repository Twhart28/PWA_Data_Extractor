# PWA Data Extractor

Desktop utility for extracting PWA PDF reports, reviewing repeated measurements, and exporting an Excel workbook with raw, kept, averaged, and skipped-file data.

## What This App Does

The app processes PWA **Detailed** or **Clinical** report PDFs locally. It extracts measurements, groups files by subject/visit/timepoint, chooses the best two repeated measurements for averaging, and exports a structured Excel workbook.

Use the **Report type** selector before processing:

- **Detailed**: full PWA detailed-report field set.
- **Clinical**: shorter clinical-report field set. Empty detailed-only columns are hidden from the app tables and Excel export.

Files that do not match the selected report type are not averaged. They are listed in the export's `Skipped Files` sheet.

## Import Workflow

1. Add PDF files by dragging them into the source panel or browsing.
2. Choose **Detailed** or **Clinical** report mode.
3. Choose how filenames should be grouped: subject only, subject + timepoint, subject + visit, or subject + visit + timepoint.
4. Optionally paste a custom filename regex.
5. Choose the output workbook path.
6. Process PDFs.
7. Review flagged patients.
8. Export the workbook.

You can remove selected files from the import list or clear the list entirely before processing.

## Filename Grouping

The app groups rows by the parsed **Patient ID**. Patient ID is built from filename parts according to the selected grouping mode.

The built-in parser handles common filename patterns and preserves study prefixes, such as `IAS003`.

Examples:

- `IAS003 PWA1.pdf` -> subject `IAS003`
- `IAS003_T2 PWA1.pdf` -> subject `IAS003`, timepoint `2`

Suffixes such as `PWA1`, `PWA2`, `Report1`, or `Run1` are treated as measurement/report suffixes, not timepoints, unless an explicit timepoint token is present.

## Custom Filename Regex

Leave the regex field blank to use the built-in parser.

For unusual filename formats, use **Copy regex prompt**. Paste that prompt into an AI with your example filenames. The prompt asks the AI to first determine the expected subject, visit, and timepoint outputs, ask follow-up questions if needed, and only then return a Python regex.

The final regex must use Python named capture groups:

- `(?P<subject>...)` is required.
- `(?P<visit>...)` is optional.
- `(?P<timepoint>...)` is optional.

The app matches the regex against the filename without the `.pdf` extension.

## Review Logic

If a subject has exactly two valid entries, that pair is selected automatically.

If a subject has more than two valid entries, the app preselects the closest automatic pair and sends the subject to review. In review, choose exactly two measurements to keep.

Subjects with exactly two entries can also be sent to review when their pair differences exceed the alert threshold.

Review tools include:

- selected-pair difference boxes
- green/yellow/red pair-difference highlights
- keep buttons for choosing exactly two rows
- reset to automatic pairing
- PDF viewing from table rows
- confirm-pair tracking

## Pair-Difference Thresholds

The import screen has two pair-difference thresholds:

- **Green up to**: differences at or below this value are green.
- **Alert above**: differences above this value are red and are flagged for review.

Values between those two thresholds are yellow.

The default alert threshold is `6.0 mmHg`.

## Export Workbook

The export contains four sheets:

`All Data`

- all parsed rows for the selected report mode
- wrong-type and unrecognized rows as visible skipped/special rows
- review status for each row

`Kept Data`

- only rows selected for averaging

`Averaged Data`

- one row per averaged subject
- averaged measurement values
- pair-difference columns
- pair-alert fields

`Skipped Files`

- subjects with only one uploaded file
- files skipped because they were the wrong report type
- unrecognized PDFs
- a reason column explaining why each file was skipped

## Clinical Mode Export

When **Clinical** mode is selected, the app hides detailed-only columns in:

- `All Data`
- `Kept Data`
- `Averaged Data`

This keeps the Clinical workbook focused on fields available from Clinical reports.

## Local Processing

All PDF parsing and Excel export happen locally on this computer. The app does not upload PDF contents.

## Run Locally

```powershell
.\.venv\Scripts\python.exe -m pip install -r .\requirements.txt
.\.venv\Scripts\python.exe .\pwa_extractor.py
```

## Build Portable Executable

```powershell
.\.venv\Scripts\pyinstaller.exe .\pwa_extractor.spec
```

## Build Installer

```powershell
.\build_release.ps1
```

That produces:

- `dist\pwa_extractor.exe`
- `release\PWA_Data_Extractor_Setup.exe`

## Project Files

- `pwa_extractor.py`: launcher entry point
- `app.py`: PySide6 desktop user interface
- `backend.py`: PDF parsing, pairing logic, and Excel export
- `README.md`: in-app help shown from the top-right `?` button
- `pwa_extractor.spec`: PyInstaller build spec
- `pwa_extractor_installer.iss`: Inno Setup installer script
- `build_release.ps1`: release build script

## Contact

- Email: `thomaswhart28@gmail.com`
