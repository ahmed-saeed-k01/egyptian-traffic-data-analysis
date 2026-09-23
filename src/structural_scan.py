"""
src/structural_scan.py

Traffic_Data - Structural Scan Orchestrator
===========================================

Purpose
-------
Coordinate the Structural Scan stage.

Pipeline
--------
    discovery
        ↓
    mapping reference
        ↓
    Excel structural scan
        ↓
    brand aggregation
        ↓
    reports
        ↓
    structural_scan.csv
    structural_scan.json
    brand_names.json

This module is intentionally small.

It does NOT contain the detailed detection logic.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Sequence

from src.config import (
    BRAND_NAMES_JSON_PATH,
    MAPPING_PATH,
    PREVIEW_ROWS,
    RAW_DIR,
    STRUCTURAL_SCAN_CSV_PATH,
    STRUCTURAL_SCAN_JSON_PATH,
)

from src.discovery import discover_raw_files

from src.excel_structural import (
    aggregate_brands_from_workbooks,
    load_mapping_reference,
    scan_workbook,
)

from src.structural_reports import (
    display_brand_summary,
    display_output_paths,
    display_structural_summary,
    save_brand_names_json,
    save_structural_scan_csv,
    save_structural_scan_json,
)

from src.structural_models import (
    WorkbookScanResult,
)


logger = logging.getLogger(
    "traffic_data.structural_scan"
)


# =====================================================================
# Logging
# =====================================================================

def configure_logging(
    level: int = logging.INFO,
) -> None:
    """
    Configure application logging.
    """

    logging.basicConfig(
        level=level,
        format=(
            "%(asctime)s | %(levelname)s | "
            "%(name)s | %(message)s"
        ),
    )


# =====================================================================
# Scan All Workbooks
# =====================================================================

def scan_all_workbooks(
    files: Sequence[Path],
    raw_dir: Path | str,
    expected_headers: Sequence[str],
    n_preview_rows: int = PREVIEW_ROWS,
) -> tuple[
    list[WorkbookScanResult],
    dict[str, dict],
]:
    """
    Scan all discovered Excel workbooks.
    """

    raw_dir = Path(
        raw_dir
    )

    workbook_results: list[
        WorkbookScanResult
    ] = []

    workbook_data: dict[
        str,
        dict,
    ] = {}

    for path in files:

        path = Path(path)

        if not path.exists():

            logger.error(
                "File does not exist: %s",
                path,
            )

            continue

        try:

            result, sheet_data = (
                scan_workbook(
                    path=path,
                    raw_dir=raw_dir,
                    expected_headers=expected_headers,
                    n_preview_rows=n_preview_rows,
                )
            )

            workbook_results.append(
                result
            )

            workbook_data[
                path.name
            ] = sheet_data

        except Exception as exc:

            logger.exception(
                "Unexpected workbook scan error: %s",
                path,
            )

    workbook_results.sort(
        key=lambda result: (
            result.relative_path.casefold()
        )
    )

    return (
        workbook_results,
        workbook_data,
    )


# =====================================================================
# Main Structural Scan
# =====================================================================

def run_structural_scan(
    files: Sequence[Path],
    raw_dir: Path | str,
    mapping_path: Path | str,
    structural_csv_path: Path | str,
    structural_json_path: Path | str,
    brand_json_path: Path | str,
    n_preview_rows: int = PREVIEW_ROWS,
) -> tuple[
    list[WorkbookScanResult],
    dict[str, dict],
    dict,
]:
    """
    Execute the complete Structural Scan stage.

    Steps
    -----

    1. Load mapping reference.
    2. Scan every workbook.
    3. Detect structural evidence.
    4. Aggregate brands.
    5. Save CSV.
    6. Save JSON.
    7. Save brand reference.
    """

    logger.info(
        "Starting Structural Scan."
    )

    # ---------------------------------------------------------------
    # Mapping
    # ---------------------------------------------------------------

    mapping_reference = (
        load_mapping_reference(
            mapping_path
        )
    )

    logger.info(
        "Mapping status: %s",
        mapping_reference.status,
    )

    if (
        mapping_reference.status
        != "SUCCESS"
    ):

        logger.warning(
            "Mapping validation is unavailable. "
            "Structural Scan will continue using "
            "generic structural detection."
        )

    expected_headers = (
        mapping_reference.expected_headers
    )

    logger.info(
        "Expected headers loaded: %d",
        len(expected_headers),
    )

    # ---------------------------------------------------------------
    # Workbook Scan
    # ---------------------------------------------------------------

    (
        workbook_results,
        workbook_data,
    ) = scan_all_workbooks(
        files=files,
        raw_dir=raw_dir,
        expected_headers=expected_headers,
        n_preview_rows=n_preview_rows,
    )

    # ---------------------------------------------------------------
    # Brand Aggregation
    # ---------------------------------------------------------------

    brand_reference = (
        aggregate_brands_from_workbooks(
            files=files,
            workbook_data=workbook_data,
            expected_headers=expected_headers,
        )
    )

    # ---------------------------------------------------------------
    # Reports
    # ---------------------------------------------------------------

    save_structural_scan_csv(
        workbooks=workbook_results,
        output_path=structural_csv_path,
    )

    save_structural_scan_json(
        workbooks=workbook_results,
        output_path=structural_json_path,
    )

    save_brand_names_json(
        brand_reference=brand_reference,
        output_path=brand_json_path,
    )

    logger.info(
        "Structural Scan completed."
    )

    return (
        workbook_results,
        workbook_data,
        brand_reference,
    )


# =====================================================================
# Standalone Execution
# =====================================================================

def main() -> int:
    """
    Standalone Structural Scan entry point.
    """

    configure_logging()

    print(
        "\n"
        "============================================================\n"
        " TRAFFIC_DATA - STRUCTURAL SCAN\n"
        "============================================================\n"
    )

    # ---------------------------------------------------------------
    # Discovery
    # ---------------------------------------------------------------

    files = discover_raw_files(
        raw_dir=RAW_DIR,
        recursive=True,
    )

    print(
        f"Discovered Excel files: {len(files)}"
    )

    if not files:

        print(
            "\nNo supported Excel files were discovered."
        )

        return 1

    # ---------------------------------------------------------------
    # Structural Scan
    # ---------------------------------------------------------------

    (
        workbook_results,
        workbook_data,
        brand_reference,
    ) = run_structural_scan(
        files=files,
        raw_dir=RAW_DIR,
        mapping_path=MAPPING_PATH,
        structural_csv_path=(
            STRUCTURAL_SCAN_CSV_PATH
        ),
        structural_json_path=(
            STRUCTURAL_SCAN_JSON_PATH
        ),
        brand_json_path=(
            BRAND_NAMES_JSON_PATH
        ),
        n_preview_rows=PREVIEW_ROWS,
    )

    # ---------------------------------------------------------------
    # Console Summary
    # ---------------------------------------------------------------

    display_structural_summary(
        workbook_results
    )

    display_brand_summary(
        brand_reference
    )

    display_output_paths(
        structural_csv_path=(
            STRUCTURAL_SCAN_CSV_PATH
        ),
        structural_json_path=(
            STRUCTURAL_SCAN_JSON_PATH
        ),
        brand_json_path=(
            BRAND_NAMES_JSON_PATH
        ),
    )

    print(
        "\nStructural Scan completed successfully."
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )