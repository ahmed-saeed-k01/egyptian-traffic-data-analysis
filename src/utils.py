from pathlib import Path


def ensure_parent_directory(path: Path) -> None:
    """Ensure the parent directory of a file exists."""
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )


def safe_text(value) -> str:
    """Convert a value to clean text safely."""
    if value is None:
        return ""

    return str(value).strip()