
"""
src/structural_models.py

Traffic_Data - Structural Models
================================

Purpose
-------
Define the typed data models used by the Structural Scan stage.

This module is intentionally model-only.

It does NOT:
    - read Excel files
    - read PDF files
    - detect headers
    - detect totals
    - detect notes
    - detect brands
    - clean data
    - normalize data
    - save reports
    - modify raw source files

It represents structural evidence discovered by other modules.

Architecture
------------
config.py
    ↓
discovery.py
    ↓
structural_scan.py
    ├── structural_models.py
    ├── structural_detectors.py
    ├── excel_structural.py
    ├── pdf_structural.py
    └── structural_reports.py

Design principles
-----------------
1. Models contain data, not discovery logic.
2. Raw files are represented by immutable metadata.
3. Structural evidence is separated from analytical decisions.
4. Excel and PDF evidence can coexist in the same scan result.
5. Coordinates remain explicit.
6. JSON serialization is straightforward.
7. Mutable collections use default_factory.
8. Optional evidence is represented explicitly with None/empty values.
9. Models remain independent of pandas/openpyxl/PyMuPDF.
10. No model silently changes or normalizes source data.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


# ============================================================================
# Type Aliases
# ============================================================================

JsonValue = (
    str
    | int
    | float
    | bool
    | None
    | list[Any]
    | dict[str, Any]
)


# ============================================================================
# Generic Structural Location
# ============================================================================

@dataclass(frozen=True)
class CellLocation:
    """
    Exact logical position of a cell in a tabular source.

    row_index / column_index
        Zero-based internal indexes.

    excel_row / excel_column
        Human-readable Excel coordinates.

    For PDF structural work, these Excel-specific fields may remain None.
    """

    row_index: int | None = None
    column_index: int | None = None

    excel_row: int | None = None
    excel_column: str | None = None


# ============================================================================
# Header Models
# ============================================================================

@dataclass(frozen=True)
class HeaderOccurrence:
    """
    One detected occurrence of a header-like value.

    This represents observed evidence only.

    It does not decide whether the value is the final schema header.
    """

    detected_value: str

    row_index: int
    excel_row: int | None

    column_index: int
    excel_column: str | None

    expected_header: str | None = None

    comparison_status: str = "UNMATCHED"

    similarity: float = 0.0

    @property
    def location(self) -> CellLocation:
        """Return the occurrence location as a typed object."""

        return CellLocation(
            row_index=self.row_index,
            column_index=self.column_index,
            excel_row=self.excel_row,
            excel_column=self.excel_column,
        )


@dataclass(frozen=True)
class HeaderDiagnostic:
    """
    Diagnostic comparison between expected and detected headers.

    This is deliberately diagnostic.

    A mismatch is reported as evidence and is NOT silently corrected.
    """

    detected_header: str | None

    expected_header: str

    comparison_status: str

    detected_row: int | None = None
    detected_row_excel: int | None = None

    detected_column: int | None = None
    detected_column_excel: str | None = None

    similarity: float = 0.0


@dataclass(frozen=True)
class HeaderRowEvidence:
    """
    Evidence describing a possible header row.
    """

    row_index: int

    row_index_excel: int | None

    non_empty_count: int

    text_count: int

    numeric_count: int

    text_ratio: float

    expected_header_hits: int = 0

    detected_header_count: int = 0

    confidence: float = 0.0

    reason: str | None = None


# ============================================================================
# Mapping Reference
# ============================================================================

@dataclass(frozen=True)
class MappingReference:
    """
    Reference information loaded from mapping.xlsx.

    mapping.xlsx is reference-only.

    It may describe expected:
        - file identity
        - sheet identity
        - header names

    It must NOT force:
        - detected positions
        - final schema
        - page numbers
        - table blocks
        - cleaning decisions
    """

    status: str

    source_path: str | None = None

    expected_headers: tuple[str, ...] = ()

    header_column: str | None = None

    available_columns: tuple[str, ...] = ()

    error: str | None = None


# ============================================================================
# Brand Models
# ============================================================================

@dataclass(frozen=True)
class BrandOccurrence:
    """
    One observed source occurrence of a vehicle brand/make.

    The value is preserved as detected.
    """

    brand: str

    file_name: str

    relative_path: str

    file_type: str

    sheet_name: str | None

    page_number: int | None

    header_row: int | None

    header_row_excel: int | None

    column_index: int | None

    column_excel: str | None


@dataclass
class BrandReference:
    """
    Aggregated brand reference artifact.

    This is a reference artifact, not an analytical dataset.
    """

    artifact: str = "brand_names_reference"

    purpose: str = (
        "Unique vehicle brand/make names detected "
        "across the raw source files."
    )

    reference_role: str = (
        "Reference artifact only. Not an analytical dataset."
    )

    files_scanned: int = 0

    files_with_brand_column: int = 0

    sheets_with_brand_column: int = 0

    pages_with_brand_column: int = 0

    unique_brand_count: int = 0

    brands: list[str] = field(
        default_factory=list
    )

    sources: list[BrandOccurrence] = field(
        default_factory=list
    )


# ============================================================================
# Row Evidence
# ============================================================================

@dataclass(frozen=True)
class RowEvidence:
    """
    Structural evidence for one row.

    This does not classify the row as a final analytical record.
    """

    row_index: int

    row_index_excel: int | None

    non_empty_count: int

    text_count: int

    numeric_count: int

    text_ratio: float

    numeric_ratio: float

    is_completely_empty: bool = False

    possible_total: bool = False

    possible_note: bool = False

    reason: str | None = None


# ============================================================================
# Total Row Evidence
# ============================================================================

@dataclass(frozen=True)
class TotalRowEvidence:
    """
    Evidence indicating that a row may contain totals/subtotals.
    """

    row_index: int

    row_index_excel: int | None

    matched_terms: tuple[str, ...] = ()

    numeric_value_count: int = 0

    confidence: float = 0.0

    reason: str | None = None


# ============================================================================
# Note Row Evidence
# ============================================================================

@dataclass(frozen=True)
class NoteRowEvidence:
    """
    Evidence indicating that a row may be a note/comment/source row.
    """

    row_index: int

    row_index_excel: int | None

    text: str | None = None

    confidence: float = 0.0

    reason: str | None = None


# ============================================================================
# Duplicate / Repeated Values
# ============================================================================

@dataclass(frozen=True)
class RepeatedValueEvidence:
    """
    Repeated value observed in an early structural area.
    """

    value: str

    occurrence_count: int

    rows: tuple[int, ...] = ()

    columns: tuple[int, ...] = ()

    reason: str | None = None


# ============================================================================
# Brand Column Candidate
# ============================================================================

@dataclass(frozen=True)
class BrandColumnCandidate:
    """
    Candidate column that may contain vehicle brand/make values.
    """

    column_index: int

    column_excel: str | None

    detected_header: str | None

    confidence: float

    non_empty_value_count: int = 0

    unique_value_count: int = 0

    sample_values: tuple[str, ...] = ()

    reason: str | None = None


# ============================================================================
# Excel Structural Evidence
# ============================================================================

@dataclass(frozen=True)
class ExcelStructuralEvidence:
    """
    Excel-specific structural evidence.

    This model intentionally does not depend on openpyxl/pandas.
    """

    workbook_file_name: str

    sheet_name: str

    sheet_index: int

    read_status: str

    preview_rows_requested: int

    preview_rows_read: int

    preview_columns: int

    full_rows_read: int

    full_columns_read: int

    is_empty: bool

    first_non_empty_row: int | None = None

    last_non_empty_row: int | None = None

    first_non_empty_column: int | None = None

    last_non_empty_column: int | None = None

    completely_empty_rows_in_preview: list[int] = field(
        default_factory=list
    )

    completely_empty_columns_in_preview: list[int] = field(
        default_factory=list
    )

    possible_header_rows: list[int] = field(
        default_factory=list
    )

    possible_header_rows_excel: list[int] = field(
        default_factory=list
    )

    header_occurrences: list[HeaderOccurrence] = field(
        default_factory=list
    )

    header_diagnostics: list[HeaderDiagnostic] = field(
        default_factory=list
    )

    header_row_evidence: list[HeaderRowEvidence] = field(
        default_factory=list
    )

    expected_headers_count: int = 0

    detected_headers_count: int = 0

    header_match_count: int = 0

    header_mismatch_count: int = 0

    header_not_found_count: int = 0

    duplicate_values_in_first_rows: dict[
        str,
        list[str],
    ] = field(
        default_factory=dict
    )

    repeated_value_evidence: list[
        RepeatedValueEvidence
    ] = field(
        default_factory=list
    )

    possible_total_rows: list[int] = field(
        default_factory=list
    )

    possible_total_rows_excel: list[int] = field(
        default_factory=list
    )

    total_row_evidence: list[
        TotalRowEvidence
    ] = field(
        default_factory=list
    )

    possible_note_rows: list[int] = field(
        default_factory=list
    )

    possible_note_rows_excel: list[int] = field(
        default_factory=list
    )

    note_row_evidence: list[
        NoteRowEvidence
    ] = field(
        default_factory=list
    )

    brand_column_candidates: list[
        BrandColumnCandidate
    ] = field(
        default_factory=list
    )

    preview: list[list[JsonValue]] = field(
        default_factory=list
    )

    error: str | None = None


# ============================================================================
# PDF Structural Evidence
# ============================================================================

@dataclass(frozen=True)
class PDFWordEvidence:
    """
    One extracted PDF word with preserved coordinates.

    Coordinates are kept exactly as supplied by the PDF extraction layer.
    """

    text: str

    x0: float

    y0: float

    x1: float

    y1: float

    block_number: int | None = None

    line_number: int | None = None

    word_number: int | None = None


@dataclass(frozen=True)
class PDFHeaderOccurrence:
    """
    One detected PDF header occurrence.

    Coordinates allow later reconstruction of columns/regions.
    """

    detected_value: str

    page_number: int

    x0: float

    y0: float

    x1: float

    y1: float

    expected_header: str | None = None

    comparison_status: str = "UNMATCHED"

    similarity: float = 0.0


@dataclass(frozen=True)
class PDFRowEvidence:
    """
    Structural row reconstructed from PDF word coordinates.
    """

    page_number: int

    row_index: int

    y0: float

    y1: float

    words: tuple[PDFWordEvidence, ...] = ()

    text: str = ""

    possible_header: bool = False

    possible_total: bool = False

    possible_note: bool = False

    confidence: float = 0.0


@dataclass(frozen=True)
class PDFColumnEvidence:
    """
    Structural column/region detected from PDF coordinates.
    """

    column_index: int

    x0: float

    x1: float

    detected_header: str | None = None

    confidence: float = 0.0


@dataclass(frozen=True)
class PDFPageStructuralEvidence:
    """
    Structural evidence for one PDF page.
    """

    page_number: int

    read_status: str

    text_layer_available: bool

    extracted_text_char_count: int

    words: tuple[PDFWordEvidence, ...] = ()

    rows: tuple[PDFRowEvidence, ...] = ()

    columns: tuple[PDFColumnEvidence, ...] = ()

    header_occurrences: tuple[
        PDFHeaderOccurrence, ...
    ] = ()

    error: str | None = None


@dataclass(frozen=True)
class PDFStructuralEvidence:
    """
    Structural evidence for a complete PDF document.

    The page count is detected from the actual document.
    No fixed page count is assumed by the model.
    """

    file_name: str

    relative_path: str

    read_status: str

    page_count: int

    pages_scanned: int

    pages: tuple[
        PDFPageStructuralEvidence, ...
    ] = ()

    text_layer_pages: int = 0

    pages_without_text_layer: int = 0

    error: str | None = None


# ============================================================================
# Cross-Page Semantic Aggregation
# ============================================================================

@dataclass(frozen=True)
class CrossPageColumnOccurrence:
    """
    One occurrence of a logical column/header on a PDF page.
    """

    logical_header: str

    page_number: int

    detected_header: str

    column_index: int | None

    x0: float | None

    x1: float | None

    values: tuple[float, ...] = ()

    raw_values: tuple[str, ...] = ()


@dataclass(frozen=True)
class CrossPageAggregate:
    """
    Aggregated values for the same logical column across PDF pages.

    Example
    -------
    page 1 -> Xt1
    page 2 -> Xt2
    total   -> XT

    The aggregation layer can represent:
        XT = Xt1 + Xt2
    without modifying the source PDF.
    """

    logical_header: str

    occurrences: tuple[
        CrossPageColumnOccurrence, ...
    ] = ()

    page_count: int = 0

    numeric_total: float | None = None

    raw_total: str | None = None

    aggregation_method: str = "SUM"

    confidence: float = 0.0

    error: str | None = None


# ============================================================================
# Sheet Scan Result
# ============================================================================

@dataclass
class SheetScanResult:
    """
    Structural result for one workbook sheet.

    This is evidence, not a final analytical schema.
    """

    file_name: str

    relative_path: str

    sheet_name: str

    sheet_index: int

    read_status: str

    preview_rows_requested: int

    preview_rows_read: int

    preview_columns: int

    full_rows_read: int

    full_columns_read: int

    is_empty: bool

    first_non_empty_row: int | None

    last_non_empty_row: int | None

    first_non_empty_column: int | None

    last_non_empty_column: int | None

    completely_empty_rows_in_preview: list[int] = field(
        default_factory=list
    )

    completely_empty_columns_in_preview: list[int] = field(
        default_factory=list
    )

    possible_header_rows: list[int] = field(
        default_factory=list
    )

    possible_header_rows_excel: list[int] = field(
        default_factory=list
    )

    detected_header_values: dict[
        str,
        list[str],
    ] = field(
        default_factory=dict
    )

    header_occurrences: list[
        HeaderOccurrence
    ] = field(
        default_factory=list
    )

    expected_headers_count: int = 0

    detected_headers_count: int = 0

    header_match_count: int = 0

    header_mismatch_count: int = 0

    header_not_found_count: int = 0

    header_diagnostics: list[
        HeaderDiagnostic
    ] = field(
        default_factory=list
    )

    duplicate_values_in_first_rows: dict[
        str,
        list[str],
    ] = field(
        default_factory=dict
    )

    possible_total_rows: list[int] = field(
        default_factory=list
    )

    possible_total_rows_excel: list[int] = field(
        default_factory=list
    )

    possible_note_rows: list[int] = field(
        default_factory=list
    )

    possible_note_rows_excel: list[int] = field(
        default_factory=list
    )

    brand_column_candidates: list[
        BrandColumnCandidate
    ] = field(
        default_factory=list
    )

    preview: list[list[JsonValue]] = field(
        default_factory=list
    )

    error: str | None = None

    def to_excel_evidence(self) -> ExcelStructuralEvidence:
        """
        Convert the legacy/general sheet result into the
        strongly typed Excel evidence model.

        This keeps backward compatibility while allowing
        newer Structural Scan components to use typed evidence.
        """

        return ExcelStructuralEvidence(
            workbook_file_name=self.file_name,
            sheet_name=self.sheet_name,
            sheet_index=self.sheet_index,
            read_status=self.read_status,
            preview_rows_requested=(
                self.preview_rows_requested
            ),
            preview_rows_read=self.preview_rows_read,
            preview_columns=self.preview_columns,
            full_rows_read=self.full_rows_read,
            full_columns_read=self.full_columns_read,
            is_empty=self.is_empty,
            first_non_empty_row=self.first_non_empty_row,
            last_non_empty_row=self.last_non_empty_row,
            first_non_empty_column=(
                self.first_non_empty_column
            ),
            last_non_empty_column=(
                self.last_non_empty_column
            ),
            completely_empty_rows_in_preview=(
                list(
                    self.completely_empty_rows_in_preview
                )
            ),
            completely_empty_columns_in_preview=(
                list(
                    self.completely_empty_columns_in_preview
                )
            ),
            possible_header_rows=(
                list(self.possible_header_rows)
            ),
            possible_header_rows_excel=(
                list(self.possible_header_rows_excel)
            ),
            header_occurrences=(
                list(self.header_occurrences)
            ),
            header_diagnostics=(
                list(self.header_diagnostics)
            ),
            expected_headers_count=(
                self.expected_headers_count
            ),
            detected_headers_count=(
                self.detected_headers_count
            ),
            header_match_count=(
                self.header_match_count
            ),
            header_mismatch_count=(
                self.header_mismatch_count
            ),
            header_not_found_count=(
                self.header_not_found_count
            ),
            duplicate_values_in_first_rows=(
                dict(
                    self.duplicate_values_in_first_rows
                )
            ),
            possible_total_rows=(
                list(self.possible_total_rows)
            ),
            possible_total_rows_excel=(
                list(self.possible_total_rows_excel)
            ),
            possible_note_rows=(
                list(self.possible_note_rows)
            ),
            possible_note_rows_excel=(
                list(self.possible_note_rows_excel)
            ),
            brand_column_candidates=(
                list(self.brand_column_candidates)
            ),
            preview=list(self.preview),
            error=self.error,
        )


# ============================================================================
# Workbook Result
# ============================================================================

@dataclass
class WorkbookScanResult:
    """
    Structural evidence for one complete Excel workbook.
    """

    file_name: str

    relative_path: str

    read_status: str

    sheet_count: int

    sheets: list[
        SheetScanResult
    ] = field(
        default_factory=list
    )

    error: str | None = None


# ============================================================================
# Complete Source File Result
# ============================================================================

@dataclass
class SourceScanResult:
    """
    Unified structural result for one discovered source file.

    file_type determines which structural evidence is populated.

    Excel:
        workbook is populated.

    PDF:
        pdf is populated.
    """

    file_name: str

    relative_path: str

    file_type: str

    read_status: str

    workbook: WorkbookScanResult | None = None

    pdf: PDFStructuralEvidence | None = None

    error: str | None = None


# ============================================================================
# Complete Structural Scan Result
# ============================================================================

@dataclass
class StructuralScanResult:
    """
    Complete Structural Scan output.

    Contains:
        - mapping reference
        - source-file structural evidence
        - brand reference
        - cross-page PDF aggregates
    """

    mapping_reference: MappingReference

    sources: list[
        SourceScanResult
    ] = field(
        default_factory=list
    )

    # ------------------------------------------------------------------------
    # Backward-compatible Excel collection
    # ------------------------------------------------------------------------

    workbooks: list[
        WorkbookScanResult
    ] = field(
        default_factory=list
    )

    # ------------------------------------------------------------------------
    # Reference artifact
    # ------------------------------------------------------------------------

    brand_reference: BrandReference | None = None

    # ------------------------------------------------------------------------
    # PDF semantic aggregation
    # ------------------------------------------------------------------------

    cross_page_aggregates: list[
        CrossPageAggregate
    ] = field(
        default_factory=list
    )

    # ------------------------------------------------------------------------
    # Run-level information
    # ------------------------------------------------------------------------

    scan_status: str = "NOT_STARTED"

    files_requested: int = 0

    files_scanned: int = 0

    files_failed: int = 0

    excel_files_scanned: int = 0

    pdf_files_scanned: int = 0

    error: str | None = None


# ============================================================================
# Serialization Helpers
# ============================================================================

def model_to_dict(
    model: Any,
) -> dict[str, Any]:
    """
    Convert a dataclass model into a plain dictionary.

    This is intentionally dependency-free.

    The resulting dictionary can be passed to:
        - json.dump()
        - report generators
        - debugging utilities
    """

    if not hasattr(
        model,
        "__dataclass_fields__",
    ):
        raise TypeError(
            "model_to_dict() expects a dataclass instance."
        )

    return asdict(model)


def models_to_dicts(
    models: list[Any],
) -> list[dict[str, Any]]:
    """
    Convert a list of dataclass models into dictionaries.
    """

    return [
        model_to_dict(model)
        for model in models
    ]
