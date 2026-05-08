from __future__ import annotations

import json
import sys
from pathlib import Path

from backend import (
    GROUP_MODE_SUBJECT,
    REPORT_MODE_DETAILED,
    build_analysis,
    default_output_path,
    display_dataframe,
    format_value,
    pair_alert_triggered,
    patient_entry_counts,
    patient_rows,
    process_pdf,
)


def empty_snapshot() -> dict[str, object]:
    return {
        "importState": {
            "reportMode": REPORT_MODE_DETAILED,
            "groupingMode": GROUP_MODE_SUBJECT,
            "files": [],
            "outputPath": str(default_output_path()),
            "greenMax": 3.0,
            "yellowMax": 6.0,
            "pairAlertThreshold": 5.0,
        },
        "reviewPatients": [],
        "summary": {
            "processedFiles": 0,
            "reviewPatients": 0,
            "averagedPatients": 0,
            "specialRows": 0,
        },
    }


def _review_patients(records: list[dict[str, object]], pair_alert_threshold: float) -> list[dict[str, object]]:
    bundle = build_analysis(records, pair_alert_threshold=pair_alert_threshold)
    entry_counts = patient_entry_counts(bundle.dataframe)

    alert_review_patients: list[str] = []
    for patient_id, pair in bundle.used_pairs.items():
      if entry_counts.get(patient_id) != 2:
          continue
      pair_df = bundle.dataframe.loc[list(pair)]
      if pair_alert_triggered(pair_df, pair_alert_threshold):
          alert_review_patients.append(patient_id)

    review_patients = [
        *bundle.manual_patients,
        *[patient_id for patient_id in alert_review_patients if patient_id not in bundle.manual_patients],
    ]

    result: list[dict[str, object]] = []
    for patient_id in review_patients:
        rows = patient_rows(bundle.dataframe, patient_id)
        auto_pair = set(bundle.used_pairs.get(patient_id, ()))
        row_models = []
        selected_files: list[str] = []
        for frame_index, row in rows.iterrows():
            if frame_index in auto_pair and row.get("Source File"):
                selected_files.append(str(row.get("Source File")))
            row_models.append(
                {
                    "keep": frame_index in auto_pair,
                    "sys": row.get("Peripheral Systolic Pressure (mmHg)"),
                    "dia": row.get("Peripheral Diastolic Pressure (mmHg)"),
                    "map": row.get("Peripheral Mean Pressure (mmHg)"),
                    "aorticSys": row.get("Aortic Systolic Pressure (mmHg)"),
                    "aorticDia": row.get("Aortic Diastolic Pressure (mmHg)"),
                    "sourceFile": row.get("Source File"),
                    "pairMethod": "Auto" if frame_index in auto_pair else "Manual",
                }
            )

        diff_values = {
            "sys": 0,
            "dia": 0,
            "map": 0,
            "aorticSys": 0,
            "aorticDia": 0,
        }
        if len(auto_pair) == 2:
            pair_df = bundle.dataframe.loc[list(auto_pair)]
            numeric_fields = {
                "sys": "Peripheral Systolic Pressure (mmHg)",
                "dia": "Peripheral Diastolic Pressure (mmHg)",
                "map": "Peripheral Mean Pressure (mmHg)",
                "aorticSys": "Aortic Systolic Pressure (mmHg)",
                "aorticDia": "Aortic Diastolic Pressure (mmHg)",
            }
            for key, field_name in numeric_fields.items():
                values = pair_df[field_name].tolist()
                if len(values) == 2 and values[0] is not None and values[1] is not None:
                    diff_values[key] = abs(float(values[0]) - float(values[1]))

        if patient_id in bundle.manual_patients and patient_id in alert_review_patients:
            reason = "multi_entry_pair_alert"
        elif patient_id in bundle.manual_patients:
            reason = "multi_entry"
        else:
            reason = "pair_alert"

        result.append(
            {
                "id": patient_id,
                "reason": reason,
                "selectedCount": len(auto_pair),
                "approved": False,
                "selectedFiles": selected_files,
                "currentPairDiffs": diff_values,
                "rows": row_models,
            }
        )

    return result


def process_payload(payload: dict[str, object]) -> dict[str, object]:
    report_mode = str(payload.get("reportMode") or REPORT_MODE_DETAILED)
    grouping_mode = str(payload.get("groupingMode") or GROUP_MODE_SUBJECT)
    pair_alert_threshold = float(payload.get("pairAlertThreshold") or 5.0)
    files = [Path(path) for path in payload.get("files", [])]

    records = [
        process_pdf(pdf_path, report_mode=report_mode, grouping_mode=grouping_mode)
        for pdf_path in files
    ]
    bundle = build_analysis(records, pair_alert_threshold=pair_alert_threshold)
    display_df = display_dataframe(bundle)
    review_patients = _review_patients(records, pair_alert_threshold)

    return {
        "importState": {
            "reportMode": report_mode,
            "groupingMode": grouping_mode,
            "files": [str(path) for path in files],
            "outputPath": str(payload.get("outputPath") or default_output_path()),
            "greenMax": float(payload.get("greenMax") or 3.0),
            "yellowMax": float(payload.get("yellowMax") or 6.0),
            "pairAlertThreshold": pair_alert_threshold,
        },
        "reviewPatients": review_patients,
        "summary": {
            "processedFiles": len(files),
            "reviewPatients": len(review_patients),
            "averagedPatients": len(bundle.analyzed_df),
            "specialRows": int((display_df["Special Row"] == True).sum()),
        },
    }


def main() -> int:
    command = sys.argv[1] if len(sys.argv) > 1 else "snapshot"

    if command == "snapshot":
        print(json.dumps(empty_snapshot()))
        return 0

    if command == "process":
        payload = json.load(sys.stdin)
        print(json.dumps(process_payload(payload)))
        return 0

    raise SystemExit(f"Unknown command: {command}")


if __name__ == "__main__":
    raise SystemExit(main())
