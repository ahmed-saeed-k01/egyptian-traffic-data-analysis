from itertools import combinations

import pandas as pd


def test_key_uniqueness(
    df: pd.DataFrame,
    columns: list[str] | tuple[str, ...],
) -> float | None:
    """
    Calculate uniqueness ratio for a candidate key.

    1.0 means every row has a unique combination.
    """
    columns = list(columns)

    if len(df) == 0:
        return None

    if not set(columns).issubset(df.columns):
        return None

    unique_count = (
        df.drop_duplicates(
            subset=columns
        ).shape[0]
    )

    return round(
        unique_count / len(df),
        4,
    )


def find_candidate_keys(
    df: pd.DataFrame,
    candidate_columns: list[str],
    max_combo_size: int = 3,
) -> pd.DataFrame:
    """
    Test candidate key combinations up to max_combo_size.

    This function does not declare the Grain automatically.
    It only provides statistical evidence.
    """
    if max_combo_size < 1:
        raise ValueError(
            "max_combo_size must be at least 1"
        )

    candidate_columns = [
        column
        for column in candidate_columns
        if column in df.columns
    ]

    results = []

    max_size = min(
        max_combo_size,
        len(candidate_columns),
    )

    for size in range(1, max_size + 1):

        for combo in combinations(
            candidate_columns,
            size,
        ):
            ratio = test_key_uniqueness(
                df,
                combo,
            )

            if ratio is not None:
                results.append(
                    {
                        "columns": " + ".join(combo),
                        "column_count": size,
                        "uniqueness_ratio": ratio,
                    }
                )

    if not results:
        return pd.DataFrame(
            columns=[
                "columns",
                "column_count",
                "uniqueness_ratio",
            ]
        )

    return (
        pd.DataFrame(results)
        .sort_values(
            by=[
                "uniqueness_ratio",
                "column_count",
            ],
            ascending=[
                False,
                True,
            ],
        )
        .reset_index(drop=True)
    )