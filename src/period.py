from pathlib import Path
import re
from typing import Optional


MONTH_PATTERNS = {
    "01": r"\b(jan|january|يناير)\b",
    "02": r"\b(feb|february|فبراير)\b",
    "03": r"\b(mar|march|مارس)\b",
    "04": r"\b(apr|april|أبريل|ابريل)\b",
    "05": r"\b(may|مايو)\b",
    "06": r"\b(jun|june|يونيو)\b",
    "07": r"\b(jul|july|يوليو)\b",
    "08": r"\b(aug|august|أغسطس|اغسطس)\b",
    "09": r"\b(sep|sept|september|سبتمبر)\b",
    "10": r"\b(oct|october|أكتوبر|اكتوبر)\b",
    "11": r"\b(nov|november|نوفمبر)\b",
    "12": r"\b(dec|december|ديسمبر)\b",
}


def detect_period_from_text(text: str) -> Optional[str]:
    """
    Detect YYYY-MM from explicit month/year text.

    Returns None if no sufficiently clear period is found.
    """
    if not text:
        return None

    text_lower = text.lower()

    year_match = re.search(
        r"\b(20\d{2})\b",
        text_lower,
    )

    if not year_match:
        return None

    year = year_match.group(1)

    for month_number, pattern in MONTH_PATTERNS.items():
        if re.search(pattern, text_lower):
            return f"{year}-{month_number}"

    return None


def detect_period_from_filename(
    path: Path,
) -> Optional[str]:
    """
    Detect period from filename only when an explicit
    month/year representation exists.
    """
    return detect_period_from_text(path.stem)