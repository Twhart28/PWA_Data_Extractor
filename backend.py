from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl.styles import Alignment

# Analysis mode selector. Choose 1 to use combined peripheral SYS/DIA/MEAN
# matching or 2 to match only on peripheral systolic pressure.
ANALYSIS_MODE = 2

APP_DIR = Path(__file__).resolve().parent
APP_TITLE = "PWA Data Extractor"
APP_SUBTITLE = (
    "Process PWA detailed reports, review multi-entry patients, and export the "
    "same Excel workbook structure."
)
APP_ICON_PATH = APP_DIR / "App_Logo.ico"
README_PATH = APP_DIR / "README.md"
CONTACT_EMAIL = "thomaswhart28@gmail.com"
REPOSITORY_URL = "https://github.com/Twhart28/PWA_Data_Extractor"

COLUMNS = [
    "Source File",
    "Patient ID",
    "Scanned ID",
    "Scan Date",
    "Scan Time",
    "Record #",
    "Analyed",
    "Date of Birth",
    "Age",
    "Gender",
    "Height (m)",
    "# of Pulses",
    "Pulse Height",
    "Pulse Height Variation (%)",
    "Diastolic Variation (%)",
    "Shape Deviation (%)",
    "Pulse Length Variation (%)",
    "Overall Quality (%)",
    "Peripheral Systolic Pressure (mmHg)",
    "Peripheral Diastolic Pressure (mmHg)",
    "Peripheral Pulse Pressure (mmHg)",
    "Peripheral Mean Pressure (mmHg)",
    "Aortic Systolic Pressure (mmHg)",
    "Aortic Diastolic Pressure (mmHg)",
    "Aortic Pulse Pressure (mmHg)",
    "Heart Rate (bpm)",
    "Pulse Pressure Amplification (%)",
    "Period (ms)",
    "Ejection Duration (ms)",
    "Ejection Duration (%)",
    "Aortic T2 (ms)",
    "P1 Height (mmHg)",
    "Aortic Augmentation (mmHg)",
    "Aortic AIx AP/PP(%)",
    "Aortic AIx P2/P1(%)",
    "Aortic AIx AP/PP @ HR75 (%)",
    "Buckberg SEVR (%)",
    "PTI Systolic (mmHg.s/min)",
    "PTI Diastolic (mmHg.s/min)",
    "End Systolic Pressure (mmHg)",
    "MAP Systolic (mmHg)",
    "MAP Diastolic (mmHg)",
]

EXTRA_COLUMNS = ["Source Path"]
ALL_DATA_COLUMNS = [*COLUMNS, *EXTRA_COLUMNS]
DETAILED_REPORT_MARKER = "PWA Detailed Report"
CLINICAL_REPORT_MARKER = "PWA Clinical Report"
CLINICAL_REPORT_MESSAGE = (
    "Recognized as a Clinical Report, only upload the Detailed Reports"
)
UNRECOGNIZED_REPORT_MESSAGE = "Not recognized as a PWA Detailed Report"
ANALYSIS_FIELDS_BY_MODE: dict[int, list[str]] = {
    1: [
        "Peripheral Systolic Pressure (mmHg)",
        "Peripheral Diastolic Pressure (mmHg)",
        "Peripheral Mean Pressure (mmHg)",
    ],
    2: ["Peripheral Systolic Pressure (mmHg)"],
}
PAIR_DIFF_SOURCE_FIELDS = [
    "Peripheral Systolic Pressure (mmHg)",
    "Peripheral Diastolic Pressure (mmHg)",
    "Peripheral Mean Pressure (mmHg)",
    "Aortic Systolic Pressure (mmHg)",
    "Aortic Diastolic Pressure (mmHg)",
]
PAIR_DIFF_EXPORT_COLUMNS = {
    "Peripheral Systolic Pressure (mmHg)": "Pair Diff Peripheral Systolic (mmHg)",
    "Peripheral Diastolic Pressure (mmHg)": "Pair Diff Peripheral Diastolic (mmHg)",
    "Peripheral Mean Pressure (mmHg)": "Pair Diff Peripheral Mean (mmHg)",
    "Aortic Systolic Pressure (mmHg)": "Pair Diff Aortic Systolic (mmHg)",
    "Aortic Diastolic Pressure (mmHg)": "Pair Diff Aortic Diastolic (mmHg)",
}

README_FALLBACK_TEXT = """# PWA Data Extractor

Convert one or many PWA analysis PDFs into a structured Excel workbook.

## Workflow

1. Add one or more PWA PDF reports.
2. Choose where the Excel workbook should be saved.
3. Process the reports locally.
4. Review patients with more than two entries and adjust pairings if needed.
5. Export the workbook.

## Output

The export contains:

- **All Data**
- **Kept Data**
- **Averaged Data**
"""


@dataclass
class AnalysisBundle:
    dataframe: pd.DataFrame
    special_row_mask: pd.Series
    analyzed_df: pd.DataFrame
    kept_indices: set[int]
    used_pairs: dict[str, tuple[int, int]]
    manual_patients: list[str]


def load_readme_text() -> str:
    if README_PATH.exists():
        return README_PATH.read_text(encoding="utf-8")
    return README_FALLBACK_TEXT


def default_output_path() -> Path:
    timestamp = datetime.now().strftime("%m-%d-%y %H-%M")
    return Path.home() / "Downloads" / f"PWA Export ({timestamp}).xlsx"


def extract_text(pdf_path: Path) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = [page.extract_text() or "" for page in pdf.pages]
    return "\n".join(pages_text)


def _search(pattern: str, text: str) -> str | None:
    match = re.search(pattern, text, flags=re.IGNORECASE)
    return match.group(1) if match else None


def _to_number(value: str) -> int | float | str:
    normalized = value.strip()
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", normalized):
        return float(normalized) if "." in normalized else int(normalized)
    return value


def _extract_scan_datetime(text: str) -> tuple[str | None, str | None]:
    date_time_match = None
    for date_time_match in re.finditer(
        r"([0-9]{2}/[0-9]{2}/[0-9]{4})\s+([0-9]{2}:[0-9]{2}(?::[0-9]{2})?)",
        text,
    ):
        pass
    if date_time_match:
        return date_time_match.group(1), date_time_match.group(2)
    return None, None


def derive_patient_id(pdf_path: Path) -> str:
    stem = pdf_path.stem.strip()
    if not stem:
        return ""

    first_break_index: int | None = None
    for index, char in enumerate(stem):
        if char in (" ", "_"):
            first_break_index = index
            break

    if first_break_index is None:
        return stem

    trailing_token = stem[first_break_index + 1 :].split(" ", 1)[0].split("_", 1)[0]
    if re.fullmatch(r"T\d+", trailing_token, flags=re.IGNORECASE):
        second_break_index = None
        for index in range(first_break_index + 1, len(stem)):
            if stem[index] in (" ", "_") and index > first_break_index + 1:
                second_break_index = index
                break
        if second_break_index is None:
            return stem
        return stem[:second_break_index].strip()

    return stem[:first_break_index].strip()


def parse_report_text(text: str) -> dict[str, object]:
    normalized = re.sub(r"\s+", " ", text)

    patient_id = _search(r"Patient ID:\s*(\S+)", normalized)
    dob = _search(r"Date Of Birth:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", normalized)
    scan_date, scan_time = _extract_scan_datetime(normalized)

    age_gender_match = re.search(
        r"Age, Gender:\s*([0-9]+),\s*([A-Za-z]+)",
        normalized,
        flags=re.IGNORECASE,
    )
    age = age_gender_match.group(1) if age_gender_match else None
    gender = age_gender_match.group(2) if age_gender_match else None

    height_cm = _search(r"Height:\s*([0-9.]+)\s*cm", normalized)
    height_m = round(float(height_cm) / 100, 2) if height_cm else None

    pulses = _search(r"Number Of Pulses:\s*([0-9]+)", normalized)

    heart_rate_period = re.search(
        r"Heart Rate, Period:\s*([0-9.]+)\s*bpm,\s*([0-9.]+)\s*ms",
        normalized,
        flags=re.IGNORECASE,
    )
    heart_rate = heart_rate_period.group(1) if heart_rate_period else None
    period = heart_rate_period.group(2) if heart_rate_period else None

    ejection_match = re.search(
        r"Ejection Duration \(ED\):\s*([0-9.]+)\s*ms,\s*([0-9.]+)\s*%",
        normalized,
        flags=re.IGNORECASE,
    )
    ejection_ms = ejection_match.group(1) if ejection_match else None
    ejection_pct = ejection_match.group(2) if ejection_match else None

    aortic_t2 = _search(r"Aortic T2:\s*([0-9.]+)\s*ms", normalized)
    p1_height = _search(r"P1 Height.*?:\s*([0-9.]+)\s*mmHg", normalized)
    aortic_augmentation = _search(
        r"Aortic Augmentation.*?:\s*([-+]?[0-9.]+)\s*mmHg",
        normalized,
    )

    aix_match = re.search(
        r"Aortic AIx \(AP/PP, P2/P1\):\s*([-+]?[0-9.]+)\s*%,\s*([-+]?[0-9.]+)\s*%",
        normalized,
        flags=re.IGNORECASE,
    )
    aortic_aix_ap_pp = aix_match.group(1) if aix_match else None
    aortic_aix_p2_p1 = aix_match.group(2) if aix_match else None

    aix_hr75 = _search(
        r"Aortic AIx \(AP/PP\) @HR75:\s*([-+]?[0-9.]+)\s*%",
        normalized,
    )
    buckberg = _search(r"Buckberg SEVR:\s*([0-9.]+)\s*%", normalized)

    pti_match = re.search(
        r"PTI \(Systole, Diastole\):\s*([0-9.]+),\s*([0-9.]+)\s*mmHg\.s/min",
        normalized,
        flags=re.IGNORECASE,
    )
    pti_systolic = pti_match.group(1) if pti_match else None
    pti_diastolic = pti_match.group(2) if pti_match else None

    end_systolic_pressure = _search(
        r"End Systolic Pressure:\s*([0-9.]+)\s*mmHg",
        normalized,
    )

    map_match = re.search(
        r"MAP \(Systole, Diastole\):\s*([0-9.]+),\s*([0-9.]+)\s*mmHg",
        normalized,
        flags=re.IGNORECASE,
    )
    map_systolic = map_match.group(1) if map_match else None
    map_diastolic = map_match.group(2) if map_match else None

    pulse_height = _search(r"Pulse Height:\s*([0-9.]+)", normalized)
    pulse_height_variation = _search(
        r"Pulse Height Variation:\s*([0-9.]+)\s*%",
        normalized,
    )
    diastolic_variation = _search(
        r"Diastolic Variation:\s*([0-9.]+)\s*%",
        normalized,
    )
    shape_deviation = _search(r"Shape Deviation:\s*([0-9.]+)\s*%", normalized)
    pulse_length_variation = _search(
        r"Pulse Length Variation:\s*([0-9.]+)\s*%",
        normalized,
    )
    overall_quality = _search(r"Overall Quality:\s*([0-9.]+)\s*%", normalized)

    amplification = _search(r"PP Amplification:\s*([0-9.]+)\s*%", normalized)

    brachial_match = re.search(
        r"Brachial SYS/DIA:\s*([0-9.]+)/([0-9.]+)",
        normalized,
        flags=re.IGNORECASE,
    )
    peripheral_sys = brachial_match.group(1) if brachial_match else None
    peripheral_dia = brachial_match.group(2) if brachial_match else None

    aortic_sys = None
    aortic_dia = None
    peripheral_pp = None
    aortic_pp = None
    peripheral_mean = None
    table_heart_rate = None

    sp_match = re.search(r"SP\s+([0-9.]+)\s+([0-9.]+)", normalized, flags=re.IGNORECASE)
    if sp_match:
        peripheral_sys = peripheral_sys or sp_match.group(1)
        aortic_sys = sp_match.group(2)

    dp_match = re.search(r"DP\s+([0-9.]+)\s+([0-9.]+)", normalized, flags=re.IGNORECASE)
    if dp_match:
        peripheral_dia = peripheral_dia or dp_match.group(1)
        aortic_dia = dp_match.group(2)

    pp_match = re.search(r"PP\s+([0-9.]+)\s+([0-9.]+)", normalized, flags=re.IGNORECASE)
    if pp_match:
        peripheral_pp = pp_match.group(1)
        aortic_pp = pp_match.group(2)

    map_hr_match = re.search(
        r"MAP HR\s+([0-9.]+)\s+([0-9.]+)",
        normalized,
        flags=re.IGNORECASE,
    )
    if map_hr_match:
        peripheral_mean = map_hr_match.group(1)
        table_heart_rate = map_hr_match.group(2)

    if peripheral_sys and peripheral_dia and peripheral_pp is None:
        try:
            peripheral_pp = str(float(peripheral_sys) - float(peripheral_dia))
        except ValueError:
            peripheral_pp = None

    if aortic_sys and aortic_dia and aortic_pp is None:
        try:
            aortic_pp = str(float(aortic_sys) - float(aortic_dia))
        except ValueError:
            aortic_pp = None

    heart_rate = heart_rate or table_heart_rate

    record = {
        "Scanned ID": patient_id,
        "Scan Date": scan_date,
        "Scan Time": scan_time,
        "Date of Birth": dob,
        "Age": age,
        "Gender": gender,
        "Height (m)": height_m,
        "# of Pulses": pulses,
        "Pulse Height": pulse_height,
        "Pulse Height Variation (%)": pulse_height_variation,
        "Diastolic Variation (%)": diastolic_variation,
        "Shape Deviation (%)": shape_deviation,
        "Pulse Length Variation (%)": pulse_length_variation,
        "Overall Quality (%)": overall_quality,
        "Peripheral Systolic Pressure (mmHg)": peripheral_sys,
        "Peripheral Diastolic Pressure (mmHg)": peripheral_dia,
        "Peripheral Pulse Pressure (mmHg)": peripheral_pp,
        "Peripheral Mean Pressure (mmHg)": peripheral_mean,
        "Aortic Systolic Pressure (mmHg)": aortic_sys,
        "Aortic Diastolic Pressure (mmHg)": aortic_dia,
        "Aortic Pulse Pressure (mmHg)": aortic_pp,
        "Heart Rate (bpm)": heart_rate,
        "Pulse Pressure Amplification (%)": amplification,
        "Period (ms)": period,
        "Ejection Duration (ms)": ejection_ms,
        "Ejection Duration (%)": ejection_pct,
        "Aortic T2 (ms)": aortic_t2,
        "P1 Height (mmHg)": p1_height,
        "Aortic Augmentation (mmHg)": aortic_augmentation,
        "Aortic AIx AP/PP(%)": aortic_aix_ap_pp,
        "Aortic AIx P2/P1(%)": aortic_aix_p2_p1,
        "Aortic AIx AP/PP @ HR75 (%)": aix_hr75,
        "Buckberg SEVR (%)": buckberg,
        "PTI Systolic (mmHg.s/min)": pti_systolic,
        "PTI Diastolic (mmHg.s/min)": pti_diastolic,
        "End Systolic Pressure (mmHg)": end_systolic_pressure,
        "MAP Systolic (mmHg)": map_systolic,
        "MAP Diastolic (mmHg)": map_diastolic,
    }

    for key, value in record.items():
        if isinstance(value, str):
            record[key] = _to_number(value)
    return record


def detect_report_type(text: str) -> str:
    normalized = text.lower()
    if DETAILED_REPORT_MARKER.lower() in normalized:
        return "detailed"
    if CLINICAL_REPORT_MARKER.lower() in normalized:
        return "clinical"
    return "unrecognized"


def empty_record(message: str, pdf_path: Path) -> dict[str, object]:
    record: dict[str, object] = {column: None for column in COLUMNS}
    record["Source File"] = pdf_path.name
    record["Source Path"] = str(pdf_path)
    record["Patient ID"] = message
    return record


def process_pdf(pdf_path: Path) -> dict[str, object]:
    text = extract_text(pdf_path)
    report_type = detect_report_type(text)

    if report_type == "detailed":
        data = parse_report_text(text)
        data["Source File"] = pdf_path.name
        data["Source Path"] = str(pdf_path)
        data["Patient ID"] = derive_patient_id(pdf_path)
        return data

    if report_type == "clinical":
        return empty_record(CLINICAL_REPORT_MESSAGE, pdf_path)

    return empty_record(UNRECOGNIZED_REPORT_MESSAGE, pdf_path)


def prepare_dataframe(records: list[dict[str, object]]) -> tuple[pd.DataFrame, pd.Series]:
    df = pd.DataFrame(records)

    for column in ALL_DATA_COLUMNS:
        if column not in df.columns:
            df[column] = None

    df = df[ALL_DATA_COLUMNS]
    df["Special Row"] = df["Patient ID"].isin(
        {CLINICAL_REPORT_MESSAGE, UNRECOGNIZED_REPORT_MESSAGE}
    )
    df.loc[df["Special Row"], COLUMNS[2:]] = None

    df.sort_values(
        by=["Special Row", "Patient ID", "Scan Date", "Scan Time"],
        inplace=True,
    )

    special_rows = df["Special Row"]
    regular_df = df.loc[~special_rows].drop_duplicates(
        subset=["Patient ID", "Scan Time", "PTI Diastolic (mmHg.s/min)"],
        keep="first",
    )
    df = pd.concat([regular_df, df.loc[special_rows]], ignore_index=True)

    df.sort_values(
        by=["Special Row", "Patient ID", "Scan Date", "Scan Time"],
        inplace=True,
        ignore_index=True,
    )

    special_row_mask = df["Special Row"].copy()
    df["Record #"] = None
    valid_rows = ~df["Special Row"]
    df.loc[valid_rows, "Record #"] = (
        df[valid_rows].groupby("Patient ID").cumcount() + 1
    )

    return df, special_row_mask


def closest_pair_indices(
    df: pd.DataFrame,
    fields: list[str],
) -> tuple[int, int] | None:
    if len(df) < 2:
        return None

    systolic_only = fields == ["Peripheral Systolic Pressure (mmHg)"]
    diastolic_values = (
        pd.to_numeric(df["Peripheral Diastolic Pressure (mmHg)"], errors="coerce")
        if systolic_only and "Peripheral Diastolic Pressure (mmHg)" in df
        else None
    )

    min_distance = float("inf")
    min_diastolic_diff = float("inf")
    closest_pair: tuple[int, int] | None = None

    for i, idx_i in enumerate(df.index[:-1]):
        for idx_j in df.index[i + 1 :]:
            diff = df.loc[idx_i, fields] - df.loc[idx_j, fields]
            distance = (diff.pow(2).sum()) ** 0.5
            diastolic_diff = float("inf")
            if systolic_only and diastolic_values is not None:
                diastolic_diff = diastolic_values.loc[idx_i] - diastolic_values.loc[idx_j]
                diastolic_diff = (
                    abs(diastolic_diff) if pd.notna(diastolic_diff) else float("inf")
                )

            if distance < min_distance:
                min_distance = distance
                min_diastolic_diff = diastolic_diff
                closest_pair = (idx_i, idx_j)
            elif distance == min_distance and systolic_only:
                if diastolic_diff < min_diastolic_diff:
                    min_diastolic_diff = diastolic_diff
                    closest_pair = (idx_i, idx_j)

    return closest_pair


def average_pair_rows(
    pair_df: pd.DataFrame,
    excluded_fields: set[str],
) -> dict[str, object]:
    averaged: dict[str, object] = {}
    for column in pair_df.columns:
        if column in excluded_fields:
            continue
        if column == "Patient ID":
            averaged[column] = pair_df[column].iloc[0]
            continue

        numeric_values = pd.to_numeric(pair_df[column], errors="coerce")
        if numeric_values.notna().any():
            averaged[column] = numeric_values.mean()
        else:
            non_null = pair_df[column].dropna()
            averaged[column] = non_null.iloc[0] if not non_null.empty else None

    return averaged


def calculate_pair_differences(pair_df: pd.DataFrame) -> dict[str, float | None]:
    differences: dict[str, float | None] = {}
    for source_field, export_column in PAIR_DIFF_EXPORT_COLUMNS.items():
        if source_field not in pair_df.columns or len(pair_df.index) < 2:
            differences[export_column] = None
            continue

        numeric_values = pd.to_numeric(pair_df[source_field], errors="coerce")
        if len(numeric_values.index) < 2 or numeric_values.isna().any():
            differences[export_column] = None
            continue

        differences[export_column] = abs(
            float(numeric_values.iloc[0]) - float(numeric_values.iloc[1])
        )
    return differences


def pair_alert_triggered(
    pair_df: pd.DataFrame,
    threshold: float,
) -> bool:
    numeric_fields = [
        "Peripheral Systolic Pressure (mmHg)",
        "Peripheral Diastolic Pressure (mmHg)",
    ]
    for field in numeric_fields:
        if field not in pair_df.columns or len(pair_df.index) < 2:
            continue
        numeric_values = pd.to_numeric(pair_df[field], errors="coerce")
        if len(numeric_values.index) < 2 or numeric_values.isna().any():
            continue
        if abs(float(numeric_values.iloc[0]) - float(numeric_values.iloc[1])) > threshold:
            return True
    return False


def build_analyzed_data(
    df: pd.DataFrame,
    mode: int,
    manual_pairs: dict[str, tuple[int, int]] | None = None,
    pair_alert_threshold: float = 5.0,
) -> tuple[pd.DataFrame, set[int], dict[str, tuple[int, int]]]:
    analysis_fields = ANALYSIS_FIELDS_BY_MODE.get(mode, ANALYSIS_FIELDS_BY_MODE[1])

    numeric_df = df.copy()
    for field in analysis_fields:
        numeric_df[field] = pd.to_numeric(numeric_df[field], errors="coerce")

    analyzed_records: list[dict[str, object]] = []
    kept_indices: set[int] = set()
    used_pairs: dict[str, tuple[int, int]] = {}
    excluded_fields = {
        "Source File",
        "Scanned ID",
        "Scan Date",
        "Scan Time",
        "Analyed",
        "Record #",
        "Source Path",
    }

    manual_pairs = manual_pairs or {}

    for patient_id, group in numeric_df.groupby("Patient ID"):
        valid_group = group.dropna(subset=analysis_fields)
        pair: tuple[int, int] | None = manual_pairs.get(patient_id)

        if not pair or not all(index in valid_group.index for index in pair):
            pair = closest_pair_indices(valid_group, analysis_fields)
        if pair is None:
            continue

        pair_df = df.loc[list(pair)]
        averaged_record = average_pair_rows(pair_df, excluded_fields)
        averaged_record.update(calculate_pair_differences(pair_df))
        averaged_record["Patient Entry Count"] = len(group)
        averaged_record["Pair Alert Threshold (mmHg)"] = pair_alert_threshold
        averaged_record["Pair Alert"] = (
            "Yes" if pair_alert_triggered(pair_df, pair_alert_threshold) else "No"
        )
        averaged_record["Patient ID"] = patient_id
        analyzed_records.append(averaged_record)
        kept_indices.update(pair)
        used_pairs[patient_id] = pair

    return pd.DataFrame(analyzed_records), kept_indices, used_pairs


def build_analysis(
    records: list[dict[str, object]],
    manual_pairs: dict[str, tuple[int, int]] | None = None,
    mode: int = ANALYSIS_MODE,
    pair_alert_threshold: float = 5.0,
) -> AnalysisBundle:
    dataframe, special_row_mask = prepare_dataframe(records)
    analyzed_df, kept_indices, used_pairs = build_analyzed_data(
        dataframe,
        mode,
        manual_pairs,
        pair_alert_threshold=pair_alert_threshold,
    )

    manual_patients = [
        patient_id
        for patient_id, group in dataframe.loc[dataframe["Special Row"] != True].groupby("Patient ID")
        if len(group) > 2
    ]

    return AnalysisBundle(
        dataframe=dataframe,
        special_row_mask=special_row_mask,
        analyzed_df=analyzed_df,
        kept_indices=kept_indices,
        used_pairs=used_pairs,
        manual_patients=manual_patients,
    )


def patient_entry_counts(df: pd.DataFrame) -> dict[str, int]:
    regular_rows = df.loc[df["Special Row"] != True]
    return {
        patient_id: int(len(group))
        for patient_id, group in regular_rows.groupby("Patient ID")
    }


def display_dataframe(bundle: AnalysisBundle) -> pd.DataFrame:
    frame = bundle.dataframe.copy()
    frame["Analyed"] = "No"
    if bundle.kept_indices:
        frame.loc[frame.index.isin(bundle.kept_indices), "Analyed"] = "Yes"
    return frame


def patient_rows(df: pd.DataFrame, patient_id: str) -> pd.DataFrame:
    return df.loc[(df["Patient ID"] == patient_id) & (df["Special Row"] != True)]


def initial_manual_pairs(
    df: pd.DataFrame,
    auto_pairs: dict[str, tuple[int, int]],
    manual_patients: list[str],
) -> dict[str, list[int]]:
    pairs: dict[str, list[int]] = {}
    for patient_id in manual_patients:
        auto_pair = list(auto_pairs.get(patient_id, ()))
        patient_frame = patient_rows(df, patient_id)
        fallback = list(patient_frame.index[:2])
        pairs[patient_id] = auto_pair[:2] if len(auto_pair) == 2 else fallback
    return pairs


def data_sheet_path(data_sheet_folder: Path | None, patient_id: str) -> Path | None:
    if data_sheet_folder is None or not data_sheet_folder.exists():
        return None

    subject_prefix = re.split(r"[ _]", patient_id, maxsplit=1)[0].lower()
    for candidate in sorted(data_sheet_folder.glob("*.pdf")):
        if candidate.stem.lower().startswith(subject_prefix):
            return candidate
    return None


def format_value(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}".rstrip("0").rstrip(".")
    return str(value)


def format_pressure_triplet(sys: object, dia: object, mean: object) -> str:
    if pd.isna(sys) and pd.isna(dia) and pd.isna(mean):
        return "—"

    parts: list[str] = []
    if not pd.isna(sys) or not pd.isna(dia):
        left = format_value(sys) or "—"
        right = format_value(dia) or "—"
        parts.append(f"{left}/{right}")
    if not pd.isna(mean):
        parts.append(f"MAP {format_value(mean)}")
    if not parts:
        return "—"
    if len(parts) == 1:
        return parts[0]
    return f"{parts[0]} ({parts[1]})"


def record_status(patient_id: object) -> str:
    value = str(patient_id or "")
    if value == CLINICAL_REPORT_MESSAGE:
        return "Clinical report"
    if value == UNRECOGNIZED_REPORT_MESSAGE:
        return "Unrecognized"
    return "Detailed report"


def save_to_excel(
    records: list[dict[str, object]],
    output_path: Path,
    manual_pairs: dict[str, tuple[int, int]] | None = None,
    mode: int = ANALYSIS_MODE,
    pair_alert_threshold: float = 5.0,
) -> int:
    bundle = build_analysis(
        records,
        manual_pairs=manual_pairs,
        mode=mode,
        pair_alert_threshold=pair_alert_threshold,
    )
    df = display_dataframe(bundle)

    kept_df = df[df["Analyed"] == "Yes"].copy()
    averaged_df = bundle.analyzed_df.drop(columns=["Record #"], errors="ignore").copy()

    date_columns = ["Scan Date", "Date of Birth"]

    def normalize_dates(frame: pd.DataFrame) -> pd.DataFrame:
        for date_column in date_columns:
            if date_column not in frame.columns:
                continue

            parsed_dates = pd.to_datetime(
                frame[date_column],
                errors="coerce",
                dayfirst=True,
            )
            frame.loc[:, date_column] = parsed_dates
        return frame

    df = normalize_dates(df)
    kept_df = normalize_dates(kept_df)
    averaged_df = normalize_dates(averaged_df)

    def strip_aux_columns(frame: pd.DataFrame) -> pd.DataFrame:
        return frame.drop(columns=["Special Row", *EXTRA_COLUMNS], errors="ignore")

    df_to_save = strip_aux_columns(df.copy())
    kept_df_to_save = strip_aux_columns(kept_df.copy())
    averaged_df_to_save = strip_aux_columns(averaged_df.copy())

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_to_save.to_excel(writer, sheet_name="All Data", index=False)
        kept_df_to_save.to_excel(writer, sheet_name="Kept Data", index=False)
        averaged_df_to_save.to_excel(writer, sheet_name="Averaged Data", index=False)

        header_alignment = Alignment(horizontal="left")
        center_alignment = Alignment(horizontal="center")
        sheet_frames = {
            "All Data": df_to_save,
            "Kept Data": kept_df_to_save,
            "Averaged Data": averaged_df_to_save,
        }

        for sheet_name, frame in sheet_frames.items():
            sheet = writer.book[sheet_name]

            for first_column in sheet.iter_cols(min_col=1, max_col=1):
                for cell in first_column:
                    cell.alignment = center_alignment

            if sheet.max_column > 1:
                for first_column in sheet.iter_cols(min_col=2, max_col=2):
                    for cell in first_column:
                        cell.alignment = header_alignment

            if sheet_name == "All Data":
                for row_index, is_special in bundle.special_row_mask.items():
                    if not is_special:
                        continue

                    patient_cell = sheet.cell(row=row_index + 2, column=2)
                    patient_cell.alignment = header_alignment

            for date_column in date_columns:
                if date_column not in frame.columns:
                    continue

                date_col_index = frame.columns.get_loc(date_column) + 1

                for column_cells in sheet.iter_cols(
                    min_col=date_col_index,
                    max_col=date_col_index,
                    min_row=2,
                    max_row=sheet.max_row,
                ):
                    for date_cell in column_cells:
                        date_cell.number_format = "MM/DD/YY"

    return len(df)
