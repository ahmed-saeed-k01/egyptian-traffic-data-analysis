import hashlib

import pandas as pd


def normalize_column_name(column) -> str:
    """Create a stable representation of a column name."""
    return " ".join(
        str(column).strip().lower().split()
    )


def get_schema_columns(
    df: pd.DataFrame,
) -> list[str]:
    """Return normalized column names in their existing order."""
    return [
        normalize_column_name(column)
        for column in df.columns
    ]


def calculate_schema_signature(
    df: pd.DataFrame,
) -> str:
    """
    Generate a stable SHA-256 schema signature
    based on ordered column names.
    """
    columns = get_schema_columns(df)

    schema_text = "||".join(columns)

    return hashlib.sha256(
        schema_text.encode("utf-8")
    ).hexdigest()


def profile_schema(
    df: pd.DataFrame,
) -> dict:
    """Return basic schema information."""
    columns = get_schema_columns(df)

    return {
        "column_count": len(columns),
        "columns": columns,
        "schema_signature": calculate_schema_signature(df),
    }