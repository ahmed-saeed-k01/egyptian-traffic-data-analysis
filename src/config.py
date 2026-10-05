from pathlib import Path


# ============================================================
# Project Root
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]


# ============================================================
# Main Data Directories
# ============================================================

DATA_DIR = PROJECT_ROOT / "data"

RAW_DIR = DATA_DIR / "raw"
PROCESSED_DIR = DATA_DIR / "processed"
REFERENCE_DIR = DATA_DIR / "reference"


# ============================================================
# Processed Data Directories
# ============================================================

MODELED_DIR = PROCESSED_DIR / "modeled"
STANDARDIZED_DIR = PROCESSED_DIR / "standardized"
VALIDATED_DIR = PROCESSED_DIR / "validated"


# ============================================================
# Reports
# ============================================================

REPORTS_DIR = PROJECT_ROOT / "reports"

INVENTORY_REPORTS_DIR = REPORTS_DIR / "inventory"
QUALITY_REPORTS_DIR = REPORTS_DIR / "quality"
ANALYSIS_REPORTS_DIR = REPORTS_DIR / "analysis"


# ============================================================
# Outputs
# ============================================================

OUTPUTS_DIR = PROJECT_ROOT / "outputs"

CHARTS_DIR = OUTPUTS_DIR / "charts"
DASHBOARD_DIR = OUTPUTS_DIR / "dashboard"
TABLES_DIR = OUTPUTS_DIR / "tables"


# ============================================================
# Notebooks
# ============================================================

NOTEBOOKS_DIR = PROJECT_ROOT / "notebooks"


# ============================================================
# Reference Files
# ============================================================

# Existing reference supplied to the project.
# It is read-only during Discovery / Structural Scan.
MAPPING_PATH = REFERENCE_DIR / "mapping.xlsx"

# Reference artifact created by Structural Scan.
# It is NOT an input source file and must not be required
# to exist before the first Structural Scan run.
BRAND_NAMES_JSON_PATH = REFERENCE_DIR / "brand_names.json"


# ============================================================
# Inventory Outputs
# ============================================================

# Discovery output. The discovery.py module uses this path by default.
FILE_MANIFEST_PATH = INVENTORY_REPORTS_DIR / "file_manifest.csv"

# Structural Scan outputs
STRUCTURAL_SCAN_CSV_PATH = INVENTORY_REPORTS_DIR / "structural_scan.csv"
STRUCTURAL_SCAN_JSON_PATH = INVENTORY_REPORTS_DIR / "structural_scan.json"
STRUCTURAL_NOTES_PATH = INVENTORY_REPORTS_DIR / "structural_notes.md"

# General Inventory outputs
INVENTORY_REPORT_PATH = INVENTORY_REPORTS_DIR / "inventory_report.csv"

GRAIN_DEFINITIONS_PATH = INVENTORY_REPORTS_DIR / "grain_definitions.md"

INGESTION_LOG_PATH = INVENTORY_REPORTS_DIR / "ingestion.log"


# ============================================================
# Pipeline Versions
# ============================================================

# Required by discovery.py for manifest traceability and validation.
MANIFEST_SCHEMA_VERSION = "1.0"
PIPELINE_VERSION = "1.0.0"


# ============================================================
# Supported Raw File Types
# ============================================================

# Excel files that the project may inspect.
# These extensions are also recognized by discovery.py.
EXCEL_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".xlsm",
    ".xlsb",
    ".ods",
}

# Original source PDFs are allowed to remain in data/raw/.
# Structural Scan can inspect PDF pages without converting them
# into Excel first.
PDF_EXTENSIONS = {
    ".pdf",
}

# Master set used by Discovery and downstream structural stages.
# Keep this as the single project-level allow-list.
SUPPORTED_EXTENSIONS = frozenset(
    EXCEL_EXTENSIONS
    | PDF_EXTENSIONS
)

# Reference files used by downstream mapping, standardization,
# and validation stages. These are intentionally separate from
# the raw-source extension allow-list above.
REFERENCE_EXCEL_EXTENSIONS = {
    '.xlsx',
}

REFERENCE_JSON_EXTENSIONS = {
    '.json',
}

REFERENCE_SUPPORTED_EXTENSIONS = frozenset(
    REFERENCE_EXCEL_EXTENSIONS
    | REFERENCE_JSON_EXTENSIONS
)

REFERENCE_MANIFEST_PATH = (
    INVENTORY_REPORTS_DIR / 'reference_manifest.csv'
)

REFERENCE_MANIFEST_SCHEMA_VERSION = '1.0'


# ============================================================
# Structural Scan Settings
# ============================================================

# Number of initial rows/pages of content used as structural evidence.
# This does NOT define a final header.
PREVIEW_ROWS = 15

# Maximum number of PDF pages Structural Scan will inspect from
# one source file.
MAX_PDF_PAGES = 200

# Maximum text extracted from one PDF page for structural inspection.
MAX_PDF_TEXT_CHARS_PER_PAGE = 20_000


# ============================================================
# General Project Settings
# ============================================================

DEFAULT_ENCODING = "utf-8"


# ============================================================
# Directory Creation
# ============================================================

def ensure_project_directories() -> None:
    """Create the static project directories if they do not exist."""

    directories = [
        DATA_DIR,
        RAW_DIR,
        PROCESSED_DIR,
        MODELED_DIR,
        STANDARDIZED_DIR,
        VALIDATED_DIR,
        REFERENCE_DIR,
        REPORTS_DIR,
        INVENTORY_REPORTS_DIR,
        QUALITY_REPORTS_DIR,
        ANALYSIS_REPORTS_DIR,
        OUTPUTS_DIR,
        CHARTS_DIR,
        DASHBOARD_DIR,
        TABLES_DIR,
        NOTEBOOKS_DIR,
    ]

    for directory in directories:
        directory.mkdir(
            parents=True,
            exist_ok=True,
        )


if __name__ == "__main__":
    ensure_project_directories()
    print(f"Project directories are ready under: {PROJECT_ROOT}")
