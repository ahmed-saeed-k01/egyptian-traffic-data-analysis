import pandas as pd


def profile_dataframe(
    df: pd.DataFrame,
    file_name: str,
    sheet_name: str,
) -> dict:
    """Profile one DataFrame."""
    unnamed_columns = [
        str(column)
        for column in df.columns
        if str(column).startswith("Unnamed")
    ]

    fully_null_columns = [
        str(column)
        for column in df.columns
        if df[column].isna().all()
    ]

    dtype_map = {
        str(column): str(dtype)
        for column, dtype in df.dtypes.items()
    }

    return {
        "file": file_name,
        "sheet": sheet_name,
        "rows": len(df),
        "columns": df.shape[1],
        "column_names": " | ".join(
            str(column) for column in df.columns
        ),
        "dtypes": str(dtype_map),
        "unnamed_columns_count": len(unnamed_columns),
        "unnamed_columns": " | ".join(unnamed_columns),
        "fully_null_columns_count": len(
            fully_null_columns
        ),
        "fully_null_columns": " | ".join(
            fully_null_columns
        ),
        "missing_total": int(
            df.isna().sum().sum()
        ),
        "duplicate_rows": int(
            df.duplicated().sum()
        ),
        "memory_mb": round(
            df.memory_usage(deep=True).sum()
            / (1024 ** 2),
            4,
        ),
        "is_empty": bool(df.empty),
    }


def build_inventory(
    all_raw: dict[str, dict[str, pd.DataFrame]],
) -> pd.DataFrame:
    """Build inventory report for all loaded datasets."""
    records = []

    for file_name, sheets in all_raw.items():

        if not sheets:
            records.append(
                {
                    "file": file_name,
                    "sheet": None,
                    "rows": 0,
                    "columns": 0,
                    "column_names": "",
                    "dtypes": "",
                    "unnamed_columns_count": None,
                    "unnamed_columns": "",
                    "fully_null_columns_count": None,
                    "fully_null_columns": "",
                    "missing_total": None,
                    "duplicate_rows": None,
                    "memory_mb": None,
                    "is_empty": True,
                    "ingestion_status": "FAILED",
                }
            )
            continue

        for sheet_name, df in sheets.items():
            record = profile_dataframe(
                df,
                file_name=file_name,
                sheet_name=sheet_name,
            )

            record["ingestion_status"] = "SUCCESS"
            records.append(record)

    return pd.DataFrame(records)


def save_inventory(
    inventory: pd.DataFrame,
    output_path,
) -> None:
    """Save inventory report."""
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    inventory.to_csv(
        output_path,
        index=False,
        encoding="utf-8-sig",
    )