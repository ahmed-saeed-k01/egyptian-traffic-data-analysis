"""
src/structural_detectors.py

Traffic_Data - Structural Detectors
===================================

Purpose
-------
Generic structural detection logic used by the Structural Scan stage.

This module detects structural evidence such as:

    - empty rows / columns
    - non-empty boundaries
    - possible header rows
    - expected header occurrences
    - header mismatches
    - repeated values
    - possible total rows
    - possible note/source/comment rows
    - possible brand columns
    - brand values

Design principles
-----------------
1. Detection only.
2. No final analytical schema decisions.
3. No source-file modification.
4. No report writing.
5. No Excel file opening.
6. No PDF file opening.
7. Deterministic output.
8. Preserve discovered positions.
9. Compare headers by content, not by expected position.
10. Treat all classifications as evidence, not truth.

Input
-----
The detector functions operate on already-loaded tabular data,
normally represented as pandas.DataFrame.

Output
------
Most functions return lightweight Python structures
(dict/list/tuple) that can be consumed by:

    structural_scan.py
    structural_models.py
    structural_reports.py

This module deliberately does not depend on the file-reading layer.
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter
from typing import Any, Iterable, Sequence

import pandas as pd


# ============================================================================
# Constants
# ============================================================================

FIRST_ROWS_FOR_REPEAT_CHECK = 10

HEADER_MIN_NON_EMPTY_CELLS = 2

HEADER_MIN_TEXT_RATIO = 0.60

HEADER_MISMATCH_MIN_SIMILARITY = 0.60


# ---------------------------------------------------------------------------
# Total / subtotal keywords
# ---------------------------------------------------------------------------

TOTAL_KEYWORDS = (
    "total",
    "grand total",
    "subtotal",
    "overall total",
    "sum",
    "الإجمالي",
    "اجمالي",
    "الإجمالى",
    "اجمالى",
    "المجموع",
    "المجموع الكلي",
    "المجموع الكلى",
    "الإجمالي الكلي",
    "الإجمالى الكلى",
)


# ---------------------------------------------------------------------------
# Note / source / comment keywords
# ---------------------------------------------------------------------------

NOTE_KEYWORDS = (
    "note",
    "notes",
    "source",
    "remark",
    "remarks",
    "comment",
    "comments",
    "reference",
    "references",
    "ملاحظة",
    "ملاحظات",
    "المصدر",
    "مصدر",
    "ملحوظة",
    "ملحوظات",
    "تعليق",
    "تعليقات",
    "مرجع",
    "مراجع",
)


# ---------------------------------------------------------------------------
# Brand / make header aliases
# ---------------------------------------------------------------------------

BRAND_HEADER_ALIASES = (
    "الماركة",
    "ماركة",
    "الماركات",
    "ماركات",
    "العلامة التجارية",
    "العلامة",
    "نوع الماركة",
    "اسم الماركة",
    "اسم العلامة",
    "manufacturer",
    "manufacturers",
    "brand",
    "brands",
    "make",
    "vehicle make",
    "vehicle brand",
)


# ============================================================================
# Text Normalization
# ============================================================================

def normalize_cell_value(value: Any) -> str:
    """
    Convert a cell value into a readable string.

    Important
    ---------
    This function does not modify the original DataFrame value.

    It creates a comparison/display representation only.
    """

    if value is None:
        return ""

    try:
        missing = pd.isna(value)

        if isinstance(missing, bool) and missing:
            return ""

    except (TypeError, ValueError):
        pass

    text = str(value)

    # Normalize non-breaking spaces.
    text = text.replace("\u00a0", " ")

    # Normalize Unicode representation without altering semantic content.
    text = unicodedata.normalize(
        "NFC",
        text,
    )

    return text.strip()


def normalize_for_comparison(value: Any) -> str:
    """
    Normalize a value for structural comparison only.

    The original source value remains untouched.

    Operations:
        - Unicode normalization
        - remove invisible directional/BOM characters
        - collapse repeated whitespace
        - case-folding
        - trim
    """

    text = normalize_cell_value(value)

    if not text:
        return ""

    text = text.replace("\u200f", "")
    text = text.replace("\u200e", "")
    text = text.replace("\ufeff", "")

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.casefold().strip()


# ============================================================================
# Excel Position Helpers
# ============================================================================

def excel_column_letter(
    column_index: int,
) -> str:
    """
    Convert a zero-based column index to an Excel column letter.

    Examples
    --------
    0  -> A
    1  -> B
    25 -> Z
    26 -> AA
    27 -> AB
    """

    if column_index < 0:
        raise ValueError(
            "column_index must be >= 0."
        )

    number = column_index + 1

    result = ""

    while number > 0:

        number, remainder = divmod(
            number - 1,
            26,
        )

        result = (
            chr(65 + remainder)
            + result
        )

    return result


# ============================================================================
# Basic Cell Classification
# ============================================================================

def is_empty_value(
    value: Any,
) -> bool:
    """
    Return True when a value should be treated as structurally empty.
    """

    return normalize_cell_value(value) == ""


def is_text_like(
    value: Any,
) -> bool:
    """
    Determine whether a value behaves like textual content.

    Numeric values are not considered text-like even though they can
    technically be converted to strings.
    """

    if value is None:
        return False

    try:
        missing = pd.isna(value)

        if isinstance(missing, bool) and missing:
            return False

    except (TypeError, ValueError):
        pass

    if isinstance(value, bool):
        return False

    if isinstance(
        value,
        (int, float),
    ):
        return False

    text = normalize_cell_value(value)

    return bool(text)


# ============================================================================
# Empty Row / Column Detection
# ============================================================================

def row_is_empty(
    row: pd.Series,
) -> bool:
    """
    Return True when every cell in the row is structurally empty.
    """

    return all(
        is_empty_value(value)
        for value in row.tolist()
    )


def column_is_empty(
    column: pd.Series,
) -> bool:
    """
    Return True when every cell in the column is structurally empty.
    """

    return all(
        is_empty_value(value)
        for value in column.tolist()
    )


def find_empty_rows(
    dataframe: pd.DataFrame,
) -> list[int]:
    """
    Return zero-based indexes of completely empty rows.

    The returned order follows the DataFrame row order.
    """

    if dataframe.empty:
        return []

    return [
        int(index)
        for index, row in dataframe.iterrows()
        if row_is_empty(row)
    ]


def find_empty_columns(
    dataframe: pd.DataFrame,
) -> list[int]:
    """
    Return zero-based positional indexes of completely empty columns.

    Positional indexes are used intentionally because DataFrame column
    labels are not necessarily Excel column names.
    """

    if dataframe.empty:
        return []

    result: list[int] = []

    for position in range(
        len(dataframe.columns)
    ):

        column = dataframe.iloc[
            :,
            position,
        ]

        if column_is_empty(column):
            result.append(position)

    return result


# ============================================================================
# Boundary Detection
# ============================================================================

def find_non_empty_boundaries(
    dataframe: pd.DataFrame,
) -> tuple[
    int | None,
    int | None,
    int | None,
    int | None,
]:
    """
    Find the structural non-empty boundaries of a DataFrame.

    Returns
    -------
    (
        first_non_empty_row,
        last_non_empty_row,
        first_non_empty_column,
        last_non_empty_column,
    )

    All indexes are zero-based positional indexes.

    Returns all None values when the DataFrame contains no non-empty cells.
    """

    if dataframe.empty:
        return (
            None,
            None,
            None,
            None,
        )

    non_empty_rows: list[int] = []

    non_empty_columns: list[int] = []

    for row_position in range(
        len(dataframe)
    ):

        row = dataframe.iloc[
            row_position
        ]

        if not row_is_empty(row):
            non_empty_rows.append(
                row_position
            )

    for column_position in range(
        len(dataframe.columns)
    ):

        column = dataframe.iloc[
            :,
            column_position,
        ]

        if not column_is_empty(column):
            non_empty_columns.append(
                column_position
            )

    if not non_empty_rows:
        return (
            None,
            None,
            None,
            None,
        )

    return (
        min(non_empty_rows),
        max(non_empty_rows),
        (
            min(non_empty_columns)
            if non_empty_columns
            else None
        ),
        (
            max(non_empty_columns)
            if non_empty_columns
            else None
        ),
    )


# ============================================================================
# Row Statistics
# ============================================================================

def calculate_row_statistics(
    row: pd.Series,
) -> dict[str, Any]:
    """
    Calculate structural statistics for one row.

    This is evidence only.
    """

    values = row.tolist()

    normalized_values = [
        normalize_cell_value(value)
        for value in values
    ]

    non_empty_values = [
        value
        for value in normalized_values
        if value
    ]

    non_empty_count = len(
        non_empty_values
    )

    text_count = sum(
        1
        for value in values
        if is_text_like(value)
    )

    numeric_count = sum(
        1
        for value in values
        if (
            isinstance(
                value,
                (int, float),
            )
            and not isinstance(
                value,
                bool,
            )
            and not pd.isna(value)
        )
    )

    text_ratio = (
        text_count / non_empty_count
        if non_empty_count
        else 0.0
    )

    numeric_ratio = (
        numeric_count / non_empty_count
        if non_empty_count
        else 0.0
    )

    return {
        "non_empty_count": non_empty_count,
        "text_count": text_count,
        "numeric_count": numeric_count,
        "text_ratio": round(
            text_ratio,
            4,
        ),
        "numeric_ratio": round(
            numeric_ratio,
            4,
        ),
        "is_completely_empty": (
            non_empty_count == 0
        ),
    }


# ============================================================================
# Possible Header Detection
# ============================================================================

def possible_header_rows(
    dataframe: pd.DataFrame,
) -> list[int]:
    """
    Detect rows that structurally resemble header rows.

    Criteria:
        - at least HEADER_MIN_NON_EMPTY_CELLS cells
        - at least HEADER_MIN_TEXT_RATIO of non-empty cells are text-like

    This does NOT choose a final header.

    Every qualifying row is returned.
    """

    candidates: list[int] = []

    if dataframe.empty:
        return candidates

    for row_index in range(
        len(dataframe)
    ):

        row = dataframe.iloc[
            row_index
        ]

        statistics = calculate_row_statistics(
            row
        )

        if (
            statistics["non_empty_count"]
            < HEADER_MIN_NON_EMPTY_CELLS
        ):
            continue

        if (
            statistics["text_ratio"]
            >= HEADER_MIN_TEXT_RATIO
        ):
            candidates.append(
                row_index
            )

    return candidates


def build_header_row_evidence(
    dataframe: pd.DataFrame,
    candidate_rows: Sequence[int],
    expected_headers: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """
    Build detailed evidence for possible header rows.
    """

    expected_normalized = {
        normalize_for_comparison(
            header
        )
        for header in (
            expected_headers or ()
        )
    }

    expected_normalized.discard("")

    evidence: list[
        dict[str, Any]
    ] = []

    for row_index in candidate_rows:

        if not (
            0 <= row_index < len(dataframe)
        ):
            continue

        row = dataframe.iloc[
            row_index
        ]

        statistics = calculate_row_statistics(
            row
        )

        expected_hits = 0

        detected_header_count = 0

        for value in row.tolist():

            normalized = (
                normalize_for_comparison(
                    value
                )
            )

            if not normalized:
                continue

            detected_header_count += 1

            if normalized in expected_normalized:
                expected_hits += 1

        confidence = statistics[
            "text_ratio"
        ]

        if expected_normalized:
            confidence = min(
                1.0,
                confidence
                + min(
                    expected_hits * 0.10,
                    0.30,
                ),
            )

        evidence.append(
            {
                "row_index": row_index,
                "row_index_excel": (
                    row_index + 1
                ),
                "non_empty_count": statistics[
                    "non_empty_count"
                ],
                "text_count": statistics[
                    "text_count"
                ],
                "numeric_count": statistics[
                    "numeric_count"
                ],
                "text_ratio": statistics[
                    "text_ratio"
                ],
                "expected_header_hits": (
                    expected_hits
                ),
                "detected_header_count": (
                    detected_header_count
                ),
                "confidence": round(
                    confidence,
                    4,
                ),
                "reason": (
                    "text-dominant row"
                    if expected_hits == 0
                    else "text-dominant row with expected header matches"
                ),
            }
        )

    return evidence


# ============================================================================
# Header Cell Detection
# ============================================================================

def detect_header_cells(
    dataframe: pd.DataFrame,
    candidate_rows: Sequence[int],
) -> list[dict[str, Any]]:
    """
    Extract all non-empty cells from candidate header rows.

    Repeated header rows are preserved.
    """

    result: list[
        dict[str, Any]
    ] = []

    seen: set[
        tuple[int, int]
    ] = set()

    for row_index in candidate_rows:

        if not (
            0 <= row_index < len(dataframe)
        ):
            continue

        row = dataframe.iloc[
            row_index
        ]

        for column_index, value in enumerate(
            row.tolist()
        ):

            text = normalize_cell_value(
                value
            )

            if not text:
                continue

            key = (
                row_index,
                column_index,
            )

            if key in seen:
                continue

            seen.add(key)

            result.append(
                {
                    "value": text,
                    "row": int(row_index),
                    "excel_row": int(row_index) + 1,
                    "column": int(column_index),
                    "excel_column": (
                        excel_column_letter(
                            int(column_index)
                        )
                    ),
                }
            )

    return result


def detect_header_rows_from_mapping(
    dataframe: pd.DataFrame,
    expected_headers: Sequence[str],
) -> list[int]:
    """
    Detect every row containing one or more expected headers.

    This function deliberately does not assume:
        - a fixed header row
        - a fixed column
        - a fixed position

    Repeated headers are retained.
    """

    expected_normalized = {
        normalize_for_comparison(
            header
        )
        for header in expected_headers
    }

    expected_normalized.discard("")

    if not expected_normalized:
        return []

    detected_rows: list[int] = []

    for row_index in range(
        len(dataframe)
    ):

        row = dataframe.iloc[
            row_index
        ]

        for value in row.tolist():

            normalized = (
                normalize_for_comparison(
                    value
                )
            )

            if (
                normalized
                and normalized
                in expected_normalized
            ):
                detected_rows.append(
                    row_index
                )
                break

    return detected_rows


def extract_header_values(
    dataframe: pd.DataFrame,
    header_rows: Sequence[int],
) -> dict[str, list[str]]:
    """
    Preserve all visible values from detected header rows.

    Values are not converted into final field names.
    """

    result: dict[
        str,
        list[str],
    ] = {}

    for row_index in header_rows:

        if not (
            0 <= row_index < len(dataframe)
        ):
            continue

        values = [
            normalize_cell_value(value)
            for value in dataframe.iloc[
                row_index
            ].tolist()
        ]

        result[
            f"row_{row_index}"
        ] = values

    return result


# ============================================================================
# Header Similarity
# ============================================================================

def normalized_similarity(
    left: str,
    right: str,
) -> float:
    """
    Calculate a deterministic lightweight similarity score.

    Strategy
    --------
    1. exact normalized equality -> 1.0
    2. substring relationship -> length ratio
    3. otherwise -> character-set Jaccard similarity

    This score is diagnostic only.

    It must not be interpreted as proof of semantic equivalence.
    """

    left_normalized = (
        normalize_for_comparison(left)
    )

    right_normalized = (
        normalize_for_comparison(right)
    )

    if (
        not left_normalized
        or not right_normalized
    ):
        return 0.0

    if (
        left_normalized
        == right_normalized
    ):
        return 1.0

    if (
        left_normalized
        in right_normalized
        or right_normalized
        in left_normalized
    ):

        shorter = min(
            len(left_normalized),
            len(right_normalized),
        )

        longer = max(
            len(left_normalized),
            len(right_normalized),
        )

        return round(
            shorter / longer,
            4,
        )

    left_chars = set(
        left_normalized
    )

    right_chars = set(
        right_normalized
    )

    union = (
        left_chars
        | right_chars
    )

    if not union:
        return 0.0

    intersection = (
        left_chars
        & right_chars
    )

    return round(
        len(intersection)
        / len(union),
        4,
    )


# ============================================================================
# Exact Header Matching
# ============================================================================

def find_exact_header_matches(
    dataframe: pd.DataFrame,
    expected_headers: Sequence[str],
) -> list[dict[str, Any]]:
    """
    Find every exact occurrence of every expected header.

    Matching is content-based.

    The actual discovered row and column are returned.
    """

    occurrences: list[
        dict[str, Any]
    ] = []

    expected_by_normalized: dict[
        str,
        str,
    ] = {}

    for header in expected_headers:

        normalized = (
            normalize_for_comparison(
                header
            )
        )

        if normalized:
            expected_by_normalized[
                normalized
            ] = str(header)

    if not expected_by_normalized:
        return occurrences

    for row_index in range(
        len(dataframe)
    ):

        row = dataframe.iloc[
            row_index
        ]

        for column_index, value in enumerate(
            row.tolist()
        ):

            detected = normalize_cell_value(
                value
            )

            if not detected:
                continue

            normalized = (
                normalize_for_comparison(
                    detected
                )
            )

            expected = (
                expected_by_normalized.get(
                    normalized
                )
            )

            if expected is None:
                continue

            occurrences.append(
                {
                    "row_index": row_index,
                    "excel_row": (
                        row_index + 1
                    ),
                    "column_index": column_index,
                    "excel_column": (
                        excel_column_letter(
                            column_index
                        )
                    ),
                    "detected_value": detected,
                    "expected_header": expected,
                    "comparison_status": "MATCH",
                    "similarity": 1.0,
                }
            )

    return occurrences


# ============================================================================
# Header Diagnostics
# ============================================================================

def build_header_diagnostics(
    dataframe: pd.DataFrame,
    expected_headers: Sequence[str],
    candidate_rows: Sequence[int],
) -> list[dict[str, Any]]:
    """
    Compare expected headers with observed candidate-header cells.

    Status values:

        MATCH
        MISMATCH
        NOT_FOUND

    Important
    ---------
    One expected header can have multiple MATCH occurrences because
    repeated headers are valid structural evidence.
    """

    diagnostics: list[
        dict[str, Any]
    ] = []

    if not expected_headers:
        return diagnostics

    candidate_cells = detect_header_cells(
        dataframe,
        candidate_rows,
    )

    for expected_header in expected_headers:

        exact_matches = [
            cell
            for cell in candidate_cells
            if (
                normalize_for_comparison(
                    cell["value"]
                )
                == normalize_for_comparison(
                    expected_header
                )
            )
        ]

        if exact_matches:

            for match in exact_matches:

                diagnostics.append(
                    {
                        "detected_header": match[
                            "value"
                        ],
                        "expected_header": (
                            expected_header
                        ),
                        "comparison_status": "MATCH",
                        "detected_row": match[
                            "row"
                        ],
                        "detected_row_excel": (
                            match["excel_row"]
                        ),
                        "detected_column": match[
                            "column"
                        ],
                        "detected_column_excel": (
                            match[
                                "excel_column"
                            ]
                        ),
                        "similarity": 1.0,
                    }
                )

            continue

        best_candidate: dict[
            str,
            Any,
        ] | None = None

        best_similarity = 0.0

        for cell in candidate_cells:

            similarity = normalized_similarity(
                cell["value"],
                expected_header,
            )

            if (
                similarity
                > best_similarity
            ):
                best_similarity = similarity
                best_candidate = cell

        if (
            best_candidate is not None
            and best_similarity
            >= HEADER_MISMATCH_MIN_SIMILARITY
        ):

            diagnostics.append(
                {
                    "detected_header": (
                        best_candidate[
                            "value"
                        ]
                    ),
                    "expected_header": (
                        expected_header
                    ),
                    "comparison_status": "MISMATCH",
                    "detected_row": (
                        best_candidate[
                            "row"
                        ]
                    ),
                    "detected_row_excel": (
                        best_candidate[
                            "excel_row"
                        ]
                    ),
                    "detected_column": (
                        best_candidate[
                            "column"
                        ]
                    ),
                    "detected_column_excel": (
                        best_candidate[
                            "excel_column"
                        ]
                    ),
                    "similarity": (
                        best_similarity
                    ),
                }
            )

        else:

            diagnostics.append(
                {
                    "detected_header": "",
                    "expected_header": (
                        expected_header
                    ),
                    "comparison_status": (
                        "NOT_FOUND"
                    ),
                    "detected_row": None,
                    "detected_row_excel": None,
                    "detected_column": None,
                    "detected_column_excel": None,
                    "similarity": 0.0,
                }
            )

    return diagnostics


def summarize_header_diagnostics(
    diagnostics: Sequence[
        dict[str, Any]
    ],
) -> dict[str, int]:
    """
    Summarize MATCH / MISMATCH / NOT_FOUND diagnostics.

    Repeated MATCH occurrences are counted as occurrences.

    NOT_FOUND is counted per diagnostic entry.
    """

    summary = {
        "expected_headers_count": 0,
        "detected_headers_count": 0,
        "header_match_count": 0,
        "header_mismatch_count": 0,
        "header_not_found_count": 0,
    }

    expected_seen: set[str] = set()

    for diagnostic in diagnostics:

        expected = normalize_for_comparison(
            diagnostic.get(
                "expected_header",
                "",
            )
        )

        if expected:
            expected_seen.add(expected)

        status = str(
            diagnostic.get(
                "comparison_status",
                "",
            )
        ).upper()

        if status == "MATCH":

            summary[
                "header_match_count"
            ] += 1

            summary[
                "detected_headers_count"
            ] += 1

        elif status == "MISMATCH":

            summary[
                "header_mismatch_count"
            ] += 1

            summary[
                "detected_headers_count"
            ] += 1

        elif status == "NOT_FOUND":

            summary[
                "header_not_found_count"
            ] += 1

    summary[
        "expected_headers_count"
    ] = len(expected_seen)

    return summary


# ============================================================================
# Repeated Value Detection
# ============================================================================

def detect_repeated_values_in_first_rows(
    preview: pd.DataFrame,
    n_rows: int = FIRST_ROWS_FOR_REPEAT_CHECK,
) -> dict[str, list[str]]:
    """
    Detect repeated non-empty values within the first N rows.

    The result is grouped by positional column index.

    This is structural evidence.

    It does not mean repeated values are errors.
    """

    result: dict[
        str,
        list[str],
    ] = {}

    if preview.empty:
        return result

    row_limit = min(
        n_rows,
        len(preview),
    )

    first_rows = preview.iloc[
        :row_limit
    ]

    for column_position in range(
        len(first_rows.columns)
    ):

        column = first_rows.iloc[
            :,
            column_position,
        ]

        normalized_values: list[
            str
        ] = []

        originals: dict[
            str,
            str,
        ] = {}

        for value in column.tolist():

            text = normalize_cell_value(
                value
            )

            if not text:
                continue

            normalized = (
                normalize_for_comparison(
                    text
                )
            )

            if not normalized:
                continue

            normalized_values.append(
                normalized
            )

            originals.setdefault(
                normalized,
                text,
            )

        counts = Counter(
            normalized_values
        )

        duplicates = [
            originals[key]
            for key, count in counts.items()
            if count > 1
        ]

        if duplicates:

            duplicates.sort(
                key=lambda value: (
                    normalize_for_comparison(
                        value
                    ),
                    value,
                )
            )

            result[
                str(column_position)
            ] = duplicates

    return result


def detect_repeated_value_evidence(
    dataframe: pd.DataFrame,
    n_rows: int = FIRST_ROWS_FOR_REPEAT_CHECK,
) -> list[dict[str, Any]]:
    """
    Produce detailed repeated-value evidence.

    Unlike detect_repeated_values_in_first_rows(),
    this function also records row positions.
    """

    result: list[
        dict[str, Any]
    ] = []

    if dataframe.empty:
        return result

    row_limit = min(
        n_rows,
        len(dataframe),
    )

    preview = dataframe.iloc[
        :row_limit
    ]

    for column_position in range(
        len(preview.columns)
    ):

        value_locations: dict[
            str,
            list[int],
        ] = {}

        original_values: dict[
            str,
            str,
        ] = {}

        for row_position in range(
            len(preview)
        ):

            value = preview.iloc[
                row_position,
                column_position,
            ]

            text = normalize_cell_value(
                value
            )

            if not text:
                continue

            normalized = (
                normalize_for_comparison(
                    text
                )
            )

            if not normalized:
                continue

            value_locations.setdefault(
                normalized,
                [],
            ).append(
                row_position
            )

            original_values.setdefault(
                normalized,
                text,
            )

        for normalized, rows in (
            value_locations.items()
        ):

            if len(rows) <= 1:
                continue

            result.append(
                {
                    "value": original_values[
                        normalized
                    ],
                    "occurrence_count": len(
                        rows
                    ),
                    "rows": rows,
                    "columns": [
                        column_position
                    ],
                    "reason": (
                        "repeated value within preview area"
                    ),
                }
            )

    return result


# ============================================================================
# Keyword Detection
# ============================================================================

def row_contains_keyword(
    row: pd.Series,
    keywords: Iterable[str],
) -> bool:
    """
    Return True when any non-empty cell contains a supplied keyword.

    Matching is normalized and substring-based.

    This intentionally favors recall over aggressive classification.
    """

    normalized_keywords = tuple(
        normalize_for_comparison(
            keyword
        )
        for keyword in keywords
        if normalize_for_comparison(
            keyword
        )
    )

    if not normalized_keywords:
        return False

    for value in row.tolist():

        text = normalize_for_comparison(
            value
        )

        if not text:
            continue

        if any(
            keyword in text
            for keyword in normalized_keywords
        ):
            return True

    return False


def find_keyword_matches_in_row(
    row: pd.Series,
    keywords: Iterable[str],
) -> list[str]:
    """
    Return the normalized keyword terms matched by a row.
    """

    normalized_keywords = [
        normalize_for_comparison(
            keyword
        )
        for keyword in keywords
        if normalize_for_comparison(
            keyword
        )
    ]

    matched: list[str] = []

    for value in row.tolist():

        text = normalize_for_comparison(
            value
        )

        if not text:
            continue

        for keyword in normalized_keywords:

            if keyword in text:
                matched.append(keyword)

    return sorted(
        set(matched)
    )


# ============================================================================
# Total Row Detection
# ============================================================================

def detect_possible_total_rows(
    dataframe: pd.DataFrame,
) -> list[int]:
    """
    Detect possible total/subtotal rows.

    This is evidence only.

    No row is deleted or transformed.
    """

    if dataframe.empty:
        return []

    return [
        row_index
        for row_index in range(
            len(dataframe)
        )
        if row_contains_keyword(
            dataframe.iloc[
                row_index
            ],
            TOTAL_KEYWORDS,
        )
    ]


def build_total_row_evidence(
    dataframe: pd.DataFrame,
) -> list[dict[str, Any]]:
    """
    Build detailed total-row evidence.
    """

    result: list[
        dict[str, Any]
    ] = []

    for row_index in detect_possible_total_rows(
        dataframe
    ):

        row = dataframe.iloc[
            row_index
        ]

        matched_terms = (
            find_keyword_matches_in_row(
                row,
                TOTAL_KEYWORDS,
            )
        )

        numeric_count = 0

        for value in row.tolist():

            if (
                isinstance(
                    value,
                    (int, float),
                )
                and not isinstance(
                    value,
                    bool,
                )
                and not pd.isna(value)
            ):
                numeric_count += 1

        confidence = min(
            1.0,
            0.50
            + min(
                len(matched_terms)
                * 0.10,
                0.30,
            )
            + (
                0.20
                if numeric_count > 0
                else 0.0
            ),
        )

        result.append(
            {
                "row_index": row_index,
                "row_index_excel": (
                    row_index + 1
                ),
                "matched_terms": tuple(
                    matched_terms
                ),
                "numeric_value_count": (
                    numeric_count
                ),
                "confidence": round(
                    confidence,
                    4,
                ),
                "reason": (
                    "total keyword detected"
                ),
            }
        )

    return result


# ============================================================================
# Note Row Detection
# ============================================================================

def detect_possible_note_rows(
    dataframe: pd.DataFrame,
) -> list[int]:
    """
    Detect possible note/source/comment rows.
    """

    if dataframe.empty:
        return []

    return [
        row_index
        for row_index in range(
            len(dataframe)
        )
        if row_contains_keyword(
            dataframe.iloc[
                row_index
            ],
            NOTE_KEYWORDS,
        )
    ]


def build_note_row_evidence(
    dataframe: pd.DataFrame,
) -> list[dict[str, Any]]:
    """
    Build detailed note/source/comment evidence.
    """

    result: list[
        dict[str, Any]
    ] = []

    for row_index in detect_possible_note_rows(
        dataframe
    ):

        row = dataframe.iloc[
            row_index
        ]

        matched_terms = (
            find_keyword_matches_in_row(
                row,
                NOTE_KEYWORDS,
            )
        )

        row_text = " | ".join(
            normalize_cell_value(value)
            for value in row.tolist()
            if normalize_cell_value(value)
        )

        confidence = min(
            1.0,
            0.60
            + min(
                len(matched_terms)
                * 0.10,
                0.30,
            ),
        )

        result.append(
            {
                "row_index": row_index,
                "row_index_excel": (
                    row_index + 1
                ),
                "text": row_text or None,
                "confidence": round(
                    confidence,
                    4,
                ),
                "reason": (
                    "note/source/comment keyword detected"
                ),
            }
        )

    return result


# ============================================================================
# Brand Header Detection
# ============================================================================

def _normalized_brand_aliases() -> set[str]:
    """
    Return normalized Brand/Make aliases.
    """

    return {
        normalize_for_comparison(
            alias
        )
        for alias in BRAND_HEADER_ALIASES
        if normalize_for_comparison(
            alias
        )
    }


def find_brand_column_candidates(
    dataframe: pd.DataFrame,
    header_rows: Sequence[int],
) -> list[dict[str, Any]]:
    """
    Detect columns whose observed header matches a known Brand/Make alias.

    Multiple header occurrences are retained as evidence.

    No final schema decision is made.
    """

    aliases = _normalized_brand_aliases()

    candidates: list[
        dict[str, Any]
    ] = []

    seen: set[
        tuple[int, int, str]
    ] = set()

    for row_index in header_rows:

        if not (
            0 <= row_index < len(dataframe)
        ):
            continue

        row = dataframe.iloc[
            row_index
        ]

        for column_index, value in enumerate(
            row.tolist()
        ):

            text = normalize_cell_value(
                value
            )

            if not text:
                continue

            normalized = (
                normalize_for_comparison(
                    text
                )
            )

            if normalized not in aliases:
                continue

            key = (
                row_index,
                column_index,
                normalized,
            )

            if key in seen:
                continue

            seen.add(key)

            candidates.append(
                {
                    "header_value": text,
                    "header_row": row_index,
                    "header_row_excel": (
                        row_index + 1
                    ),
                    "column_index": column_index,
                    "column_excel": (
                        excel_column_letter(
                            column_index
                        )
                    ),
                    "source": "known_alias",
                    "expected_header": None,
                    "similarity": 1.0,
                    "confidence": 1.0,
                }
            )

    return candidates


# ============================================================================
# Brand Header Similarity Detection
# ============================================================================

def find_fuzzy_brand_column_candidates(
    dataframe: pd.DataFrame,
    header_rows: Sequence[int],
    minimum_similarity: float = 0.70,
) -> list[dict[str, Any]]:
    """
    Detect possible Brand/Make columns using similarity.

    Exact aliases should be preferred.

    Fuzzy candidates are diagnostic evidence only and must be manually
    reviewed before being used as a reference mapping.
    """

    if minimum_similarity < 0.0:
        raise ValueError(
            "minimum_similarity must be >= 0."
        )

    if minimum_similarity > 1.0:
        raise ValueError(
            "minimum_similarity must be <= 1."
        )

    aliases = tuple(
        _normalized_brand_aliases()
    )

    candidates: list[
        dict[str, Any]
    ] = []

    seen: set[
        tuple[int, int]
    ] = set()

    for row_index in header_rows:

        if not (
            0 <= row_index < len(dataframe)
        ):
            continue

        row = dataframe.iloc[
            row_index
        ]

        for column_index, value in enumerate(
            row.tolist()
        ):

            text = normalize_cell_value(
                value
            )

            if not text:
                continue

            best_alias = ""

            best_similarity = 0.0

            for alias in aliases:

                similarity = normalized_similarity(
                    text,
                    alias,
                )

                if similarity > best_similarity:
                    best_similarity = similarity
                    best_alias = alias

            if (
                best_similarity
                < minimum_similarity
            ):
                continue

            key = (
                row_index,
                column_index,
            )

            if key in seen:
                continue

            seen.add(key)

            candidates.append(
                {
                    "header_value": text,
                    "header_row": row_index,
                    "header_row_excel": (
                        row_index + 1
                    ),
                    "column_index": column_index,
                    "column_excel": (
                        excel_column_letter(
                            column_index
                        )
                    ),
                    "source": "fuzzy_alias",
                    "expected_header": best_alias,
                    "similarity": round(
                        best_similarity,
                        4,
                    ),
                    "confidence": round(
                        best_similarity,
                        4,
                    ),
                }
            )

    return candidates


# ============================================================================
# Brand Value Extraction
# ============================================================================

def clean_brand_for_reference(
    value: Any,
) -> str:
    """
    Prepare a brand value for reference-level deduplication.

    This does NOT modify the source DataFrame.
    """

    text = normalize_cell_value(
        value
    )

    if not text:
        return ""

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def extract_brands_from_column(
    dataframe: pd.DataFrame,
    candidate: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Extract observed brand values below a detected Brand header.

    Extraction rules
    ----------------
    - starts after the detected header row
    - preserves row/column position
    - skips empty cells
    - skips possible total rows
    - skips possible note rows
    - stops when another Brand header is encountered

    The function does not modify the DataFrame.
    """

    try:
        column_index = int(
            candidate["column_index"]
        )

        header_row = int(
            candidate["header_row"]
        )

    except (
        KeyError,
        TypeError,
        ValueError,
    ) as exc:

        raise ValueError(
            "Invalid brand-column candidate."
        ) from exc

    if not (
        0 <= column_index
        < len(dataframe.columns)
    ):
        return []

    if not (
        0 <= header_row
        < len(dataframe)
    ):
        return []

    brand_aliases = _normalized_brand_aliases()

    results: list[
        dict[str, Any]
    ] = []

    for row_index in range(
        header_row + 1,
        len(dataframe),
    ):

        value = dataframe.iloc[
            row_index,
            column_index,
        ]

        text = clean_brand_for_reference(
            value
        )

        normalized = (
            normalize_for_comparison(
                text
            )
        )

        # A repeated Brand header indicates the beginning
        # of another structural block.
        if normalized in brand_aliases:
            break

        if not text:
            continue

        row = dataframe.iloc[
            row_index
        ]

        if row_contains_keyword(
            row,
            TOTAL_KEYWORDS,
        ):
            continue

        if row_contains_keyword(
            row,
            NOTE_KEYWORDS,
        ):
            continue

        results.append(
            {
                "brand": text,
                "row_index": row_index,
                "excel_row": (
                    row_index + 1
                ),
                "column_index": column_index,
                "column_excel": (
                    excel_column_letter(
                        column_index
                    )
                ),
            }
        )

    return results


def collect_brand_names_from_sheet(
    dataframe: pd.DataFrame,
    brand_candidates: Sequence[
        dict[str, Any]
    ],
) -> tuple[
    list[str],
    dict[str, list[dict[str, Any]]],
]:
    """
    Collect unique brand names from one sheet.

    Returns
    -------
    brands:
        Deterministically sorted unique brand names.

    sources:
        Mapping from normalized brand name to observed source locations.

    Deduplication is comparison-level only.
    Original observed spelling is retained as the reference value.
    """

    brands_by_normalized: dict[
        str,
        str,
    ] = {}

    sources: dict[
        str,
        list[dict[str, Any]],
    ] = {}

    for candidate in brand_candidates:

        extracted = (
            extract_brands_from_column(
                dataframe,
                candidate,
            )
        )

        for item in extracted:

            brand = item["brand"]

            normalized = (
                normalize_for_comparison(
                    brand
                )
            )

            if not normalized:
                continue

            brands_by_normalized.setdefault(
                normalized,
                brand,
            )

            source = {
                "header_row": candidate[
                    "header_row"
                ],
                "header_row_excel": candidate[
                    "header_row_excel"
                ],
                "column_index": candidate[
                    "column_index"
                ],
                "column_excel": candidate[
                    "column_excel"
                ],
                "brand_row": item[
                    "row_index"
                ],
                "brand_row_excel": item[
                    "excel_row"
                ],
            }

            sources.setdefault(
                normalized,
                [],
            )

            if source not in sources[
                normalized
            ]:
                sources[
                    normalized
                ].append(source)

    sorted_keys = sorted(
        brands_by_normalized.keys()
    )

    brands = [
        brands_by_normalized[key]
        for key in sorted_keys
    ]

    return (
        brands,
        sources,
    )


# ============================================================================
# Preview Helpers
# ============================================================================

def build_preview(
    dataframe: pd.DataFrame,
    max_rows: int,
) -> list[list[Any]]:
    """
    Return a deterministic list-of-lists preview.

    No DataFrame mutation occurs.
    """

    if max_rows < 0:
        raise ValueError(
            "max_rows must be >= 0."
        )

    if dataframe.empty or max_rows == 0:
        return []

    preview = dataframe.iloc[
        :max_rows
    ]

    result: list[
        list[Any]
    ] = []

    for row in preview.itertuples(
        index=False,
        name=None,
    ):

        result.append(
            [
                value
                for value in row
            ]
        )

    return result


# ============================================================================
# Structural Summary
# ============================================================================

def summarize_dataframe_structure(
    dataframe: pd.DataFrame,
) -> dict[str, Any]:
    """
    Produce a compact structural summary.

    This function combines low-level detectors but does not make
    final schema decisions.
    """

    (
        first_non_empty_row,
        last_non_empty_row,
        first_non_empty_column,
        last_non_empty_column,
    ) = find_non_empty_boundaries(
        dataframe
    )

    empty_rows = find_empty_rows(
        dataframe
    )

    empty_columns = find_empty_columns(
        dataframe
    )

    possible_headers = possible_header_rows(
        dataframe
    )

    possible_totals = (
        detect_possible_total_rows(
            dataframe
        )
    )

    possible_notes = (
        detect_possible_note_rows(
            dataframe
        )
    )

    return {
        "rows": int(len(dataframe)),
        "columns": int(len(dataframe.columns)),
        "is_empty": bool(
            first_non_empty_row is None
        ),
        "first_non_empty_row": (
            first_non_empty_row
        ),
        "last_non_empty_row": (
            last_non_empty_row
        ),
        "first_non_empty_column": (
            first_non_empty_column
        ),
        "last_non_empty_column": (
            last_non_empty_column
        ),
        "empty_row_count": len(
            empty_rows
        ),
        "empty_column_count": len(
            empty_columns
        ),
        "empty_rows": empty_rows,
        "empty_columns": empty_columns,
        "possible_header_rows": (
            possible_headers
        ),
        "possible_total_rows": (
            possible_totals
        ),
        "possible_note_rows": (
            possible_notes
        ),
    }
