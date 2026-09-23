"""
src/structural_reports.py

Traffic_Data - Structural Reports
=================================

Purpose
-------
Save and display Structural Scan results.

Responsibilities
----------------
    - CSV output
    - JSON output
    - brand_names.json
    - console summaries

This module does NOT:
    - read Excel
    - detect headers
    - detect totals
    - clean data
    - modify raw files
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd

from src.structural_models import (
    BrandReference,
    WorkbookScanResult,
)


# =====================================================================
# Serialization Helpers
# =====================================================================

def _serialize_for_json(
    value: Any,
) -> Any:
    """
    Convert project objects into JSON-safe values.
    """

    if hasattr(
        value,
        "__dataclass_fields__",
    ):

        from dataclasses import asdict

        return asdict(value)

    if isinstance(
        value,
        Path,
    ):
        return str(value)

    return value


# =====================================================================
# Workbook Results -> Records
# =====================================================================

def workbook_results_to_records(
    workbooks: list[WorkbookScanResult],
) -> list[dict[str, Any]]:
    """
    Flatten workbook/sheet results into report records.

    One record = one sheet.
    """

    records: list[
        dict[str, Any]
    ] = []

    for workbook in workbooks:

        for sheet in workbook.sheets:

            record = {
                "file_name": sheet.file_name,
                "relative_path": sheet.relative_path,
                "sheet_name": sheet.sheet_name,
                "sheet_index": sheet.sheet_index,
                "read_status": sheet.read_status,
                "preview_rows_requested": (
                    sheet.preview_rows_requested
                ),
                "preview_rows_read": (
                    sheet.preview_rows_read
                ),
                "preview_columns": (
                    sheet.preview_columns
                ),
                "full_rows_read": (
                    sheet.full_rows_read
                ),
                "full_columns_read": (
                    sheet.full_columns_read
                ),
                "is_empty": sheet.is_empty,
                "first_non_empty_row": (
                    sheet.first_non_empty_row
                ),
                "last_non_empty_row": (
                    sheet.last_non_empty_row
                ),
                "first_non_empty_column": (
                    sheet.first_non_empty_column
                ),
                "last_non_empty_column": (
                    sheet.last_non_empty_column
                ),
                "completely_empty_rows_in_preview": (
                    json.dumps(
                        sheet.completely_empty_rows_in_preview,
                        ensure_ascii=False,
                    )
                ),
                "completely_empty_columns_in_preview": (
                    json.dumps(
                        sheet.completely_empty_columns_in_preview,
                        ensure_ascii=False,
                    )
                ),
                "possible_header_rows": (
                    json.dumps(
                        sheet.possible_header_rows,
                        ensure_ascii=False,
                    )
                ),
                "possible_header_rows_excel": (
                    json.dumps(
                        sheet.possible_header_rows_excel,
                        ensure_ascii=False,
                    )
                ),
                "detected_header_values": (
                    json.dumps(
                        sheet.detected_header_values,
                        ensure_ascii=False,
                    )
                ),
                "header_occurrences": (
                    json.dumps(
                        sheet.header_occurrences,
                        ensure_ascii=False,
                    )
                ),
                "expected_headers_count": (
                    sheet.expected_headers_count
                ),
                "detected_headers_count": (
                    sheet.detected_headers_count
                ),
                "header_match_count": (
                    sheet.header_match_count
                ),
                "header_mismatch_count": (
                    sheet.header_mismatch_count
                ),
                "header_not_found_count": (
                    sheet.header_not_found_count
                ),
                "header_diagnostics": (
                    json.dumps(
                        sheet.header_diagnostics,
                        ensure_ascii=False,
                    )
                ),
                "duplicate_values_in_first_rows": (
                    json.dumps(
                        sheet.duplicate_values_in_first_rows,
                        ensure_ascii=False,
                    )
                ),
                "possible_total_rows": (
                    json.dumps(
                        sheet.possible_total_rows,
                        ensure_ascii=False,
                    )
                ),
                "possible_total_rows_excel": (
                    json.dumps(
                        sheet.possible_total_rows_excel,
                        ensure_ascii=False,
                    )
                ),
                "possible_note_rows": (
                    json.dumps(
                        sheet.possible_note_rows,
                        ensure_ascii=False,
                    )
                ),
                "possible_note_rows_excel": (
                    json.dumps(
                        sheet.possible_note_rows_excel,
                        ensure_ascii=False,
                    )
                ),
                "brand_column_candidates": (
                    json.dumps(
                        sheet.brand_column_candidates,
                        ensure_ascii=False,
                    )
                ),
                "preview": (
                    json.dumps(
                        sheet.preview,
                        ensure_ascii=False,
                    )
                ),
                "error": sheet.error,
            }

            records.append(record)

    records.sort(
        key=lambda record: (
            str(
                record[
                    "relative_path"
                ]
            ).casefold(),
            int(
                record[
                    "sheet_index"
                ]
            ),
        )
    )

    return records


# =====================================================================
# Structural Scan CSV
# =====================================================================

def save_structural_scan_csv(
    workbooks: list[WorkbookScanResult],
    output_path: Path | str,
) -> None:
    """
    Save flattened Structural Scan results to CSV.
    """

    output_path = Path(
        output_path
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    records = workbook_results_to_records(
        workbooks
    )

    dataframe = pd.DataFrame(
        records
    )

    dataframe.to_csv(
        output_path,
        index=False,
        encoding="utf-8-sig",
    )


# =====================================================================
# Structural Scan JSON
# =====================================================================

def save_structural_scan_json(
    workbooks: list[WorkbookScanResult],
    output_path: Path | str,
) -> None:
    """
    Save complete hierarchical workbook/sheet evidence to JSON.
    """

    output_path = Path(
        output_path
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    payload = [
        _serialize_for_json(workbook)
        for workbook in workbooks
    ]

    with output_path.open(
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            payload,
            file,
            ensure_ascii=False,
            indent=2,
            default=str,
        )


# =====================================================================
# Brand Reference
# =====================================================================

def save_brand_names_json(
    brand_reference: dict[str, Any]
    | BrandReference,
    output_path: Path | str,
) -> None:
    """
    Save the brand reference artifact.
    """

    output_path = Path(
        output_path
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if isinstance(
        brand_reference,
        BrandReference,
    ):

        payload = _serialize_for_json(
            brand_reference
        )

    else:

        payload = brand_reference

    with output_path.open(
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            payload,
            file,
            ensure_ascii=False,
            indent=2,
            default=str,
        )


# =====================================================================
# Console Display
# =====================================================================

def display_structural_summary(
    workbooks: list[WorkbookScanResult],
) -> None:
    """
    Display a compact Structural Scan summary.
    """

    records = workbook_results_to_records(
        workbooks
    )

    if not records:

        print(
            "No Structural Scan results."
        )

        return

    dataframe = pd.DataFrame(
        records
    )

    columns = [
        "file_name",
        "sheet_name",
        "sheet_index",
        "read_status",
        "full_rows_read",
        "full_columns_read",
        "possible_header_rows_excel",
        "expected_headers_count",
        "detected_headers_count",
        "header_match_count",
        "header_mismatch_count",
        "header_not_found_count",
        "possible_total_rows_excel",
        "possible_note_rows_excel",
        "brand_column_candidates",
        "error",
    ]

    available = [
        column
        for column in columns
        if column in dataframe.columns
    ]

    print(
        "\n=== STRUCTURAL SCAN SUMMARY ===\n"
    )

    print(
        dataframe[
            available
        ].to_string(
            index=False
        )
    )


def display_brand_summary(
    brand_reference: dict[str, Any],
) -> None:
    """
    Display brand reference summary.
    """

    print(
        "\n=== BRAND REFERENCE SUMMARY ===\n"
    )

    print(
        "Files scanned: "
        f"{brand_reference['files_scanned']}"
    )

    print(
        "Files with brand column: "
        f"{brand_reference['files_with_brand_column']}"
    )

    print(
        "Sheets with brand column: "
        f"{brand_reference['sheets_with_brand_column']}"
    )

    print(
        "Unique brands: "
        f"{brand_reference['unique_brand_count']}"
    )

    print(
        "\n=== UNIQUE BRAND NAMES ===\n"
    )

    for brand in brand_reference[
        "brands"
    ]:

        print(
            f"- {brand}"
        )


# =====================================================================
# Output Summary
# =====================================================================

def display_output_paths(
    structural_csv_path: Path | str,
    structural_json_path: Path | str,
    brand_json_path: Path | str,
) -> None:
    """
    Display generated artifact paths.
    """

    print(
        "\n=== OUTPUTS ===\n"
    )

    print(
        f"Structural CSV : "
        f"{structural_csv_path}"
    )

    print(
        f"Structural JSON: "
        f"{structural_json_path}"
    )

    print(
        f"Brand JSON     : "
        f"{brand_json_path}"
    )