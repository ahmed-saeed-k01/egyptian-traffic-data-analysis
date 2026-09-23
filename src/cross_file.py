import pandas as pd


def compare_schema_signatures(
    schema_records: list[dict],
) -> pd.DataFrame:
    """
    Compare schema signatures across datasets.
    """
    if not schema_records:
        return pd.DataFrame(
            columns=[
                "file",
                "sheet",
                "schema_signature",
            ]
        )

    return pd.DataFrame(schema_records)


def detect_schema_drift(
    schema_records: list[dict],
) -> pd.DataFrame:
    """
    Identify datasets whose schema signature differs
    from the most common signature.
    """
    df = compare_schema_signatures(
        schema_records
    )

    if df.empty:
        return df

    signature_counts = (
        df["schema_signature"]
        .value_counts()
    )

    reference_signature = (
        signature_counts.index[0]
    )

    result = df.copy()

    result["reference_signature"] = (
        reference_signature
    )

    result["schema_status"] = result[
        "schema_signature"
    ].apply(
        lambda value:
        "MATCH"
        if value == reference_signature
        else "DRIFT"
    )

    return result


def compare_columns(
    left: list[str],
    right: list[str],
) -> dict:
    """Compare two schemas by column names."""
    left_set = set(left)
    right_set = set(right)

    return {
        "only_in_left": sorted(
            left_set - right_set
        ),
        "only_in_right": sorted(
            right_set - left_set
        ),
        "common": sorted(
            left_set & right_set
        ),
    }