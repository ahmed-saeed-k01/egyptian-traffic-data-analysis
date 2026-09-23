from pathlib import Path
from typing import Any

import pandas as pd


def scan_excel_file(
    path: Path,
    n_preview_rows: int = 15,
) -> dict[str, Any]:
    """
    Perform a raw structural scan of an Excel workbook.

    No header is assumed.
    """
    result = {
        "file": path.stem,
        "filename": path.name,
        "status": "SUCCESS",
        "error": None,
        "sheet_names": [],
        "sheets": {},
    }

    try:
        xls = pd.ExcelFile(path)

        result["sheet_names"] = xls.sheet_names

        for sheet_name in xls.sheet_names:
            preview = pd.read_excel(
                path,
                sheet_name=sheet_name,
                header=None,
                nrows=n_preview_rows,
            )

            result["sheets"][sheet_name] = {
                "preview": preview,
                "rows_scanned": len(preview),
                "columns_scanned": preview.shape[1],
            }

    except Exception as exc:
        result["status"] = "FAILED"
        result["error"] = f"{type(exc).__name__}: {exc}"

    return result


def run_structural_scan(
    files: list[Path],
    n_preview_rows: int = 15,
) -> list[dict[str, Any]]:
    """Run structural scanning for all raw files."""
    results = []

    for path in files:
        result = scan_excel_file(
            path,
            n_preview_rows=n_preview_rows,
        )
        results.append(result)

    return results


def display_structural_scan(results: list[dict[str, Any]]) -> None:
    """
    Print structural scan results for manual inspection.
    """
    for result in results:
        print("=" * 80)
        print(
            f"FILE: {result['filename']} | "
            f"STATUS: {result['status']}"
        )

        if result["status"] == "FAILED":
            print(f"ERROR: {result['error']}")
            continue

        print(f"SHEETS: {result['sheet_names']}")

        for sheet_name, sheet_info in result["sheets"].items():
            print("-" * 80)
            print(f"SHEET: {sheet_name}")
            print(sheet_info["preview"].to_string(index=False))
            print()