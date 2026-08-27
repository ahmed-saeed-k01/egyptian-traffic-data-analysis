from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

RAW_DIR = ROOT / "data" / "raw"

STANDARDIZED_DIR = (
    ROOT / "data" / "processed" / "standardized"
)

VALIDATED_DIR = (
    ROOT / "data" / "processed" / "validated"
)

MODELED_DIR = (
    ROOT / "data" / "processed" / "modeled"
)

REPORTS_DIR = ROOT / "reports"

RAW_FILES = list(RAW_DIR.glob("*.xlsx"))