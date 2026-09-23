"""
File Discovery
==============

Traffic_Data - File Discovery Stage
===================================

Purpose
-------
Discover eligible raw source files and create a reproducible,
auditable, file-level manifest.

Pipeline position
-----------------
config.py
    ↓
discovery.py
    ↓
file_manifest.csv
    ↓
structural_scan.py
    ↓
ingestion.py

Core principle
--------------
Discovery is FILE-LEVEL ONLY.

It does not interpret PDF or Excel contents.

It records:
    - filesystem eligibility
    - file type
    - relative path
    - source root
    - size
    - modification time
    - discovery timestamp
    - deterministic file identity
    - optional SHA-256 content identity
    - discovery run identity
    - manifest schema version
    - pipeline version

It does NOT:
    - open Excel workbooks
    - open PDF documents
    - inspect worksheets
    - inspect PDF pages
    - detect headers
    - extract tables
    - infer schemas
    - infer periods
    - clean data
    - normalize data
    - convert PDF to Excel
    - modify raw files

All raw source files remain read-only and authoritative.
"""

from __future__ import annotations

import hashlib
import logging
import os
import uuid

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd


try:
    from src.config import (
        FILE_MANIFEST_PATH,
        RAW_DIR,
        SUPPORTED_EXTENSIONS,
        MANIFEST_SCHEMA_VERSION,
        PIPELINE_VERSION,
    )
except ModuleNotFoundError:
    from config import (  # type: ignore
        FILE_MANIFEST_PATH,
        RAW_DIR,
        SUPPORTED_EXTENSIONS,
        MANIFEST_SCHEMA_VERSION,
        PIPELINE_VERSION,
    )


# ============================================================================
# Constants
# ============================================================================

LOGGER_NAME = "traffic_data.discovery"

DEFAULT_HASH_CHUNK_SIZE = 1024 * 1024  # 1 MiB

EXCLUDED_TEMPORARY_PREFIXES = (
    "~$",
)

EXCLUDED_HIDDEN_PREFIXES = (
    ".",
)

DISCOVERY_STATUS = "DISCOVERED"

METADATA_ERROR_STATUS = "METADATA_ERROR"

EXCLUDED_TEMPORARY_STATUS = "EXCLUDED_TEMPORARY"

EXCLUDED_HIDDEN_STATUS = "EXCLUDED_HIDDEN"

EXCLUDED_UNSUPPORTED_STATUS = "EXCLUDED_UNSUPPORTED"

EXCLUDED_REPARSE_STATUS = "EXCLUDED_REPARSE"

EXCLUDED_SYMLINK_STATUS = "EXCLUDED_SYMLINK"


# ============================================================================
# Logging
# ============================================================================

logger = logging.getLogger(LOGGER_NAME)


def configure_logging(
    level: int = logging.INFO,
) -> None:
    """
    Configure basic logging.

    basicConfig() does not override an existing application
    logging configuration.
    """

    logging.basicConfig(
        level=level,
        format=(
            "%(asctime)s | %(levelname)s | "
            "%(name)s | %(message)s"
        ),
    )


# ============================================================================
# Run Identity
# ============================================================================

def generate_discovery_run_id() -> str:
    """
    Generate a unique identifier for one Discovery execution.

    Format:
        YYYYMMDDTHHMMSSZ_<8-char UUID>

    Example:
        20260909T172130Z_a81f42c7
    """

    timestamp = datetime.now(
        timezone.utc
    ).strftime(
        "%Y%m%dT%H%M%SZ"
    )

    unique_suffix = uuid.uuid4().hex[:8]

    return (
        f"{timestamp}_{unique_suffix}"
    )


def get_utc_timestamp() -> str:
    """
    Return the current UTC timestamp as ISO-8601.
    """

    return datetime.now(
        timezone.utc
    ).isoformat()


# ============================================================================
# Data Models
# ============================================================================

@dataclass(frozen=True)
class DiscoveredFile:
    """
    Immutable file-level representation.
    """

    path: Path
    relative_path: str
    file_name: str

    extension: str
    file_type: str

    parent_directory: str
    depth: int

    size_bytes: int
    modified_time: str

    file_id: str
    content_id: str | None = None


@dataclass(frozen=True)
class DiscoveryCandidate:
    """
    Immutable filesystem candidate.

    Excluded candidates remain observable.
    """

    path: Path
    relative_path: str
    file_name: str

    extension: str
    file_type: str

    parent_directory: str
    depth: int

    status: str
    reason: str | None = None


# ============================================================================
# Extension Utilities
# ============================================================================

def normalize_extensions(
    extensions: Iterable[str],
) -> frozenset[str]:
    """
    Normalize extensions into lowercase dotted form.

    Examples
    --------
    "xlsx"  -> ".xlsx"
    ".XLSX" -> ".xlsx"
    "PDF"   -> ".pdf"
    """

    normalized: set[str] = set()

    for extension in extensions:

        value = str(
            extension
        ).strip().lower()

        if not value:
            continue

        if not value.startswith("."):
            value = f".{value}"

        normalized.add(value)

    return frozenset(
        normalized
    )


# ============================================================================
# File Type Classification
# ============================================================================

def classify_file_type(
    extension: str,
    supported_extensions: Iterable[str] = SUPPORTED_EXTENSIONS,
) -> str:
    """
    Classify a file using the centralized extension policy.

    IMPORTANT
    ---------
    No independent extension list is maintained here.

    The supported extension policy comes from config.py.

    Returns
    -------
    str
        "excel"
        "pdf"
        "other"
    """

    normalized_extension = (
        str(extension)
        .strip()
        .lower()
    )

    supported = normalize_extensions(
        supported_extensions
    )

    if normalized_extension not in supported:
        return "other"

    if normalized_extension == ".pdf":
        return "pdf"

    excel_extensions = {
        extension
        for extension in supported
        if extension in {
            ".xlsx",
            ".xls",
            ".xlsm",
            ".xlsb",
            ".ods",
        }
    }

    if normalized_extension in excel_extensions:
        return "excel"

    return "other"


# ============================================================================
# Path Utilities
# ============================================================================

def get_relative_path(
    path: Path,
    raw_dir: Path,
) -> str:
    """
    Return deterministic POSIX-style relative path.

    Raises
    ------
    ValueError
        If path is outside raw_dir.
    """

    try:

        relative_path = path.relative_to(
            raw_dir
        )

    except ValueError as exc:

        raise ValueError(
            f"Path is outside raw directory: {path}"
        ) from exc

    return relative_path.as_posix()


def get_parent_directory(
    relative_path: str,
) -> str:
    """
    Return parent directory from relative path.
    """

    parent = Path(
        relative_path
    ).parent

    if str(parent) == ".":
        return ""

    return parent.as_posix()


def get_path_depth(
    relative_path: str,
) -> int:
    """
    Return directory depth of a relative file path.
    """

    parent_directory = (
        get_parent_directory(
            relative_path
        )
    )

    if not parent_directory:
        return 0

    return len(
        Path(
            parent_directory
        ).parts
    )


# ============================================================================
# Filesystem Safety
# ============================================================================

def is_symlink(
    path: Path,
) -> bool:
    """
    Return True if path is a symbolic link.
    """

    try:
        return path.is_symlink()

    except OSError:
        return False


def is_reparse_point(
    entry: os.DirEntry,
) -> bool:
    """
    Detect Windows reparse points where possible.

    Reparse points are not followed by Discovery.

    On non-Windows systems this returns False unless the
    filesystem exposes the relevant attribute.
    """

    try:

        stat_result = entry.stat(
            follow_symlinks=False
        )

        file_attributes = getattr(
            stat_result,
            "st_file_attributes",
            0,
        )

        # Windows FILE_ATTRIBUTE_REPARSE_POINT
        return bool(
            file_attributes & 0x0400
        )

    except OSError:

        return False


# ============================================================================
# File Classification
# ============================================================================

def classify_file(
    path: Path,
    supported_extensions: Iterable[str] = SUPPORTED_EXTENSIONS,
) -> tuple[str, str | None]:
    """
    Classify a filesystem file according to Discovery rules.
    """

    if is_symlink(path):

        return (
            EXCLUDED_SYMLINK_STATUS,
            "Symbolic link is not followed.",
        )

    if path.name.startswith(
        EXCLUDED_TEMPORARY_PREFIXES
    ):

        return (
            EXCLUDED_TEMPORARY_STATUS,
            "Temporary or lock file.",
        )

    if path.name.startswith(
        EXCLUDED_HIDDEN_PREFIXES
    ):

        return (
            EXCLUDED_HIDDEN_STATUS,
            "Hidden file.",
        )

    normalized_extensions = (
        normalize_extensions(
            supported_extensions
        )
    )

    extension = path.suffix.lower()

    if extension not in normalized_extensions:

        return (
            EXCLUDED_UNSUPPORTED_STATUS,
            (
                "Unsupported extension: "
                f"{extension or '<none>'}"
            ),
        )

    return (
        DISCOVERY_STATUS,
        None,
    )


def is_supported_raw_file(
    path: Path,
    supported_extensions: Iterable[str] = SUPPORTED_EXTENSIONS,
) -> bool:
    """
    Return True if path is an eligible raw file.
    """

    if not path.is_file():
        return False

    status, _ = classify_file(
        path=path,
        supported_extensions=supported_extensions,
    )

    return status == DISCOVERY_STATUS


# ============================================================================
# Candidate Discovery
# ============================================================================

def discover_candidates(
    raw_dir: Path | str,
    recursive: bool = True,
    supported_extensions: Iterable[str] = SUPPORTED_EXTENSIONS,
) -> list[DiscoveryCandidate]:
    """
    Discover filesystem candidates.

    Uses os.scandir() instead of Path.glob() so filesystem
    traversal errors can be isolated at directory level.

    Reparse points and symbolic links are never followed.
    """

    raw_dir = Path(
        raw_dir
    ).resolve()

    if not raw_dir.exists():

        raise FileNotFoundError(
            f"Raw data directory does not exist: {raw_dir}"
        )

    if not raw_dir.is_dir():

        raise NotADirectoryError(
            f"Raw data path is not a directory: {raw_dir}"
        )

    candidates: list[
        DiscoveryCandidate
    ] = []

    def scan_directory(
        directory: Path,
    ) -> None:

        try:

            entries = list(
                os.scandir(directory)
            )

        except (
            PermissionError,
            OSError,
        ) as exc:

            logger.error(
                "Unable to scan directory: %s | %s",
                directory,
                exc,
            )

            return

        entries.sort(
            key=lambda entry: (
                entry.name.casefold()
            )
        )

        for entry in entries:

            entry_path = Path(
                entry.path
            )

            try:

                relative_path = (
                    get_relative_path(
                        path=entry_path,
                        raw_dir=raw_dir,
                    )
                )

            except ValueError as exc:

                logger.error(
                    "Skipping path outside raw root: %s | %s",
                    entry_path,
                    exc,
                )

                continue

            # --------------------------------------------------------
            # Reparse points
            # --------------------------------------------------------

            try:

                reparse = is_reparse_point(
                    entry
                )

            except OSError:

                reparse = False

            if reparse:

                logger.warning(
                    "Skipping reparse point: %s",
                    entry_path,
                )

                continue

            # --------------------------------------------------------
            # Symbolic links
            # --------------------------------------------------------

            try:

                if entry.is_symlink():

                    logger.warning(
                        "Skipping symbolic link: %s",
                        entry_path,
                    )

                    continue

            except OSError as exc:

                logger.warning(
                    "Unable to inspect symbolic-link state: %s | %s",
                    entry_path,
                    exc,
                )

                continue

            # --------------------------------------------------------
            # Directory traversal
            # --------------------------------------------------------

            try:

                if entry.is_dir(
                    follow_symlinks=False
                ):

                    if recursive:
                        scan_directory(
                            entry_path
                        )

                    continue

            except OSError as exc:

                logger.error(
                    "Unable to inspect directory entry: %s | %s",
                    entry_path,
                    exc,
                )

                continue

            # --------------------------------------------------------
            # File candidate
            # --------------------------------------------------------

            try:

                if not entry.is_file(
                    follow_symlinks=False
                ):
                    continue

            except OSError as exc:

                logger.error(
                    "Unable to inspect file entry: %s | %s",
                    entry_path,
                    exc,
                )

                continue

            extension = (
                entry_path.suffix.lower()
            )

            file_type = classify_file_type(
                extension=extension,
                supported_extensions=(
                    supported_extensions
                ),
            )

            status, reason = classify_file(
                path=entry_path,
                supported_extensions=(
                    supported_extensions
                ),
            )

            candidates.append(
                DiscoveryCandidate(
                    path=entry_path,
                    relative_path=relative_path,
                    file_name=entry_path.name,
                    extension=extension,
                    file_type=file_type,
                    parent_directory=(
                        get_parent_directory(
                            relative_path
                        )
                    ),
                    depth=get_path_depth(
                        relative_path
                    ),
                    status=status,
                    reason=reason,
                )
            )

    scan_directory(
        raw_dir
    )

    # ------------------------------------------------------------------------
    # Deterministic ordering
    # ------------------------------------------------------------------------

    candidates.sort(
        key=lambda candidate: (
            candidate.relative_path.casefold()
        )
    )

    # ------------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------------

    discovered_count = sum(
        candidate.status
        == DISCOVERY_STATUS
        for candidate in candidates
    )

    excluded_count = (
        len(candidates)
        - discovered_count
    )

    pdf_count = sum(
        candidate.status
        == DISCOVERY_STATUS
        and candidate.file_type == "pdf"
        for candidate in candidates
    )

    excel_count = sum(
        candidate.status
        == DISCOVERY_STATUS
        and candidate.file_type == "excel"
        for candidate in candidates
    )

    logger.info(
        (
            "Filesystem scan completed: "
            "%d candidate file(s), "
            "%d discovered, "
            "%d excluded, "
            "%d PDF, "
            "%d Excel."
        ),
        len(candidates),
        discovered_count,
        excluded_count,
        pdf_count,
        excel_count,
    )

    return candidates


def discover_raw_files(
    raw_dir: Path | str,
    recursive: bool = True,
    supported_extensions: Iterable[str] = SUPPORTED_EXTENSIONS,
) -> list[Path]:
    """
    Return only accepted raw files.
    """

    candidates = discover_candidates(
        raw_dir=raw_dir,
        recursive=recursive,
        supported_extensions=supported_extensions,
    )

    files = [
        candidate.path
        for candidate in candidates
        if candidate.status
        == DISCOVERY_STATUS
    ]

    logger.info(
        "Discovered %d supported raw file(s).",
        len(files),
    )

    return files


# ============================================================================
# Hashing
# ============================================================================

def calculate_sha256(
    path: Path,
    chunk_size: int = DEFAULT_HASH_CHUNK_SIZE,
) -> str:
    """
    Calculate SHA-256 without loading the complete file into memory.
    """

    if chunk_size <= 0:

        raise ValueError(
            (
                "chunk_size must be positive, "
                f"got {chunk_size}"
            )
        )

    sha256 = hashlib.sha256()

    with path.open("rb") as file:

        while True:

            chunk = file.read(
                chunk_size
            )

            if not chunk:
                break

            sha256.update(
                chunk
            )

    return sha256.hexdigest()


def calculate_file_id(
    relative_path: str,
) -> str:
    """
    Calculate deterministic file identity from relative path.

    file_id identifies a file within the raw directory structure.

    It does NOT identify content.
    """

    normalized_path = (
        relative_path
        .replace("\\", "/")
        .casefold()
    )

    return hashlib.sha256(
        normalized_path.encode(
            "utf-8"
        )
    ).hexdigest()


# ============================================================================
# Metadata
# ============================================================================

def get_file_metadata(
    path: Path,
    raw_dir: Path,
    calculate_hash: bool = False,
) -> DiscoveredFile:
    """
    Extract file-level metadata.

    No PDF or Excel content is interpreted.
    """

    path = path.resolve()
    raw_dir = raw_dir.resolve()

    relative_path = get_relative_path(
        path=path,
        raw_dir=raw_dir,
    )

    stat = path.stat()

    modified_time = (
        datetime.fromtimestamp(
            stat.st_mtime,
            tz=timezone.utc,
        ).isoformat()
    )

    extension = path.suffix.lower()

    file_type = classify_file_type(
        extension=extension,
        supported_extensions=(
            SUPPORTED_EXTENSIONS
        ),
    )

    content_id = (
        calculate_sha256(path)
        if calculate_hash
        else None
    )

    return DiscoveredFile(
        path=path,
        relative_path=relative_path,
        file_name=path.name,
        extension=extension,
        file_type=file_type,
        parent_directory=(
            get_parent_directory(
                relative_path
            )
        ),
        depth=get_path_depth(
            relative_path
        ),
        size_bytes=stat.st_size,
        modified_time=modified_time,
        file_id=calculate_file_id(
            relative_path
        ),
        content_id=content_id,
    )


# ============================================================================
# Manifest Schema
# ============================================================================

MANIFEST_COLUMNS = [
    # Provenance
    "discovery_run_id",
    "discovered_at",
    "source_root",
    "manifest_schema_version",
    "pipeline_version",

    # File identity
    "file_id",
    "content_id",

    # Path
    "file_name",
    "relative_path",
    "parent_directory",
    "depth",

    # File classification
    "extension",
    "file_type",

    # Filesystem metadata
    "size_bytes",
    "modified_time",

    # Discovery result
    "discovery_status",
    "error",
]


# ============================================================================
# Manifest Construction
# ============================================================================

def build_manifest(
    files: Sequence[Path],
    raw_dir: Path | str,
    calculate_hash: bool = False,
    discovery_run_id: str | None = None,
    discovered_at: str | None = None,
) -> pd.DataFrame:
    """
    Build an auditable file-level manifest.

    Provenance fields are identical for all rows belonging
    to the same Discovery execution.
    """

    raw_dir = Path(
        raw_dir
    ).resolve()

    if discovery_run_id is None:
        discovery_run_id = (
            generate_discovery_run_id()
        )

    if discovered_at is None:
        discovered_at = (
            get_utc_timestamp()
        )

    source_root = raw_dir.as_posix()

    records: list[
        dict[str, object]
    ] = []

    for path in files:

        path = Path(path)

        try:

            metadata = get_file_metadata(
                path=path,
                raw_dir=raw_dir,
                calculate_hash=calculate_hash,
            )

            records.append(
                {
                    # Provenance
                    "discovery_run_id": (
                        discovery_run_id
                    ),
                    "discovered_at": (
                        discovered_at
                    ),
                    "source_root": (
                        source_root
                    ),
                    "manifest_schema_version": (
                        MANIFEST_SCHEMA_VERSION
                    ),
                    "pipeline_version": (
                        PIPELINE_VERSION
                    ),

                    # Identity
                    "file_id": (
                        metadata.file_id
                    ),
                    "content_id": (
                        metadata.content_id
                    ),

                    # Path
                    "file_name": (
                        metadata.file_name
                    ),
                    "relative_path": (
                        metadata.relative_path
                    ),
                    "parent_directory": (
                        metadata.parent_directory
                    ),
                    "depth": (
                        metadata.depth
                    ),

                    # Classification
                    "extension": (
                        metadata.extension
                    ),
                    "file_type": (
                        metadata.file_type
                    ),

                    # Filesystem
                    "size_bytes": (
                        metadata.size_bytes
                    ),
                    "modified_time": (
                        metadata.modified_time
                    ),

                    # Result
                    "discovery_status": (
                        DISCOVERY_STATUS
                    ),
                    "error": None,
                }
            )

        except (
            OSError,
            ValueError,
        ) as exc:

            logger.exception(
                (
                    "Failed to extract metadata "
                    "for file: %s"
                ),
                path,
            )

            try:

                relative_path = (
                    get_relative_path(
                        path=path.resolve(),
                        raw_dir=raw_dir,
                    )
                )

            except ValueError:

                relative_path = str(path)

            extension = (
                path.suffix.lower()
            )

            file_type = classify_file_type(
                extension=extension,
                supported_extensions=(
                    SUPPORTED_EXTENSIONS
                ),
            )

            records.append(
                {
                    # Provenance
                    "discovery_run_id": (
                        discovery_run_id
                    ),
                    "discovered_at": (
                        discovered_at
                    ),
                    "source_root": (
                        source_root
                    ),
                    "manifest_schema_version": (
                        MANIFEST_SCHEMA_VERSION
                    ),
                    "pipeline_version": (
                        PIPELINE_VERSION
                    ),

                    # Identity
                    "file_id": (
                        calculate_file_id(
                            relative_path
                        )
                        if relative_path
                        else None
                    ),
                    "content_id": None,

                    # Path
                    "file_name": path.name,
                    "relative_path": (
                        relative_path
                    ),

                    "parent_directory": (
                        get_parent_directory(
                            relative_path
                        )
                        if relative_path
                        else None
                    ),

                    "depth": (
                        get_path_depth(
                            relative_path
                        )
                        if relative_path
                        else None
                    ),

                    # Classification
                    "extension": extension,
                    "file_type": file_type,

                    # Filesystem
                    "size_bytes": None,
                    "modified_time": None,

                    # Result
                    "discovery_status": (
                        METADATA_ERROR_STATUS
                    ),
                    "error": str(exc),
                }
            )

    manifest = pd.DataFrame(
        records,
        columns=MANIFEST_COLUMNS,
    )

    if not manifest.empty:

        manifest = (
            manifest
            .sort_values(
                by="relative_path",
                key=lambda column: (
                    column.astype(str)
                    .str.casefold()
                ),
                na_position="last",
            )
            .reset_index(drop=True)
        )

    return manifest


# ============================================================================
# Manifest Validation
# ============================================================================

def validate_manifest(
    manifest: pd.DataFrame,
    raw_dir: Path | str,
    expected_file_count: int | None = None,
) -> dict[str, object]:
    """
    Validate the generated manifest.

    Validation remains file-level.

    Duplicate content IDs are findings, not automatic failures.
    """

    raw_dir = Path(
        raw_dir
    ).resolve()

    required_columns = set(
        MANIFEST_COLUMNS
    )

    missing_columns = sorted(
        required_columns
        - set(manifest.columns)
    )

    if missing_columns:

        return {
            "valid": False,
            "manifest_row_count": len(
                manifest
            ),
            "missing_columns": (
                missing_columns
            ),
            "missing_files": [],
            "outside_raw_dir": [],
            "unsupported_files": [],
            "metadata_error_count": 0,
            "duplicate_relative_paths": [],
            "duplicate_file_ids": [],
            "duplicate_content_ids": [],
            "pdf_count": 0,
            "excel_count": 0,
            "expected_file_count": (
                expected_file_count
            ),
            "expected_count_match": None,
        }

    missing_files: list[str] = []
    outside_raw_dir: list[str] = []
    unsupported_files: list[str] = []

    duplicate_relative_paths: list[str] = []
    duplicate_file_ids: list[str] = []
    duplicate_content_ids: list[str] = []

    metadata_error_count = int(
        (
            manifest["discovery_status"]
            == METADATA_ERROR_STATUS
        ).sum()
    )

    # ------------------------------------------------------------------------
    # Validate rows
    # ------------------------------------------------------------------------

    for _, row in manifest.iterrows():

        relative_path = row[
            "relative_path"
        ]

        if not isinstance(
            relative_path,
            str,
        ):
            continue

        path = (
            raw_dir
            / Path(relative_path)
        )

        if not path.exists():

            missing_files.append(
                relative_path
            )

            continue

        try:

            path.resolve().relative_to(
                raw_dir
            )

        except ValueError:

            outside_raw_dir.append(
                relative_path
            )

        extension = str(
            row["extension"]
        ).lower()

        if extension not in (
            normalize_extensions(
                SUPPORTED_EXTENSIONS
            )
        ):

            unsupported_files.append(
                relative_path
            )

    # ------------------------------------------------------------------------
    # Duplicate relative paths
    # ------------------------------------------------------------------------

    relative_path_counts = (
        manifest[
            "relative_path"
        ]
        .dropna()
        .astype(str)
        .value_counts()
    )

    duplicate_relative_paths = sorted(
        relative_path_counts[
            relative_path_counts > 1
        ].index.tolist()
    )

    # ------------------------------------------------------------------------
    # Duplicate file IDs
    # ------------------------------------------------------------------------

    file_id_counts = (
        manifest[
            "file_id"
        ]
        .dropna()
        .astype(str)
        .value_counts()
    )

    duplicate_file_ids = sorted(
        file_id_counts[
            file_id_counts > 1
        ].index.tolist()
    )

    # ------------------------------------------------------------------------
    # Duplicate content IDs
    # ------------------------------------------------------------------------

    content_id_series = (
        manifest[
            "content_id"
        ]
        .dropna()
        .astype(str)
    )

    content_id_counts = (
        content_id_series.value_counts()
    )

    duplicate_content_ids = sorted(
        content_id_counts[
            content_id_counts > 1
        ].index.tolist()
    )

    # ------------------------------------------------------------------------
    # File type counts
    # ------------------------------------------------------------------------

    pdf_count = int(
        (
            manifest[
                "file_type"
            ]
            == "pdf"
        ).sum()
    )

    excel_count = int(
        (
            manifest[
                "file_type"
            ]
            == "excel"
        ).sum()
    )

    # ------------------------------------------------------------------------
    # Expected count
    # ------------------------------------------------------------------------

    if expected_file_count is None:

        expected_count_match = None

    else:

        if expected_file_count < 0:

            raise ValueError(
                "expected_file_count cannot be negative."
            )

        expected_count_match = (
            len(manifest)
            == expected_file_count
        )

    # ------------------------------------------------------------------------
    # Validation result
    # ------------------------------------------------------------------------

    valid = not (
        missing_files
        or outside_raw_dir
        or unsupported_files
        or metadata_error_count > 0
        or duplicate_relative_paths
        or duplicate_file_ids
        or (
            expected_count_match is False
        )
    )

    result: dict[str, object] = {
        "valid": valid,
        "manifest_row_count": len(
            manifest
        ),
        "missing_columns": [],
        "missing_files": (
            missing_files
        ),
        "outside_raw_dir": (
            outside_raw_dir
        ),
        "unsupported_files": (
            unsupported_files
        ),
        "metadata_error_count": (
            metadata_error_count
        ),
        "duplicate_relative_paths": (
            duplicate_relative_paths
        ),
        "duplicate_file_ids": (
            duplicate_file_ids
        ),
        "duplicate_content_ids": (
            duplicate_content_ids
        ),
        "pdf_count": pdf_count,
        "excel_count": excel_count,
        "expected_file_count": (
            expected_file_count
        ),
        "expected_count_match": (
            expected_count_match
        ),
    }

    if valid:

        logger.info(
            (
                "Manifest validation passed: "
                "%d file(s) | PDF=%d | Excel=%d"
            ),
            len(manifest),
            pdf_count,
            excel_count,
        )

    else:

        logger.warning(
            "Manifest validation found file-level issues."
        )

    if duplicate_content_ids:

        logger.warning(
            (
                "Detected %d duplicate "
                "content hash group(s)."
            ),
            len(
                duplicate_content_ids
            ),
        )

    return result


# ============================================================================
# Discovery Findings
# ============================================================================

def summarize_candidates(
    candidates: Sequence[
        DiscoveryCandidate
    ],
) -> dict[str, object]:
    """
    Summarize accepted and excluded candidates.
    """

    summary: dict[str, int] = {
        "total_candidates": len(
            candidates
        ),
        "discovered": 0,
        "discovered_pdf": 0,
        "discovered_excel": 0,
        "excluded_temporary": 0,
        "excluded_hidden": 0,
        "excluded_unsupported": 0,
        "excluded_symlink": 0,
        "excluded_reparse": 0,
    }

    for candidate in candidates:

        if candidate.status == (
            DISCOVERY_STATUS
        ):

            summary[
                "discovered"
            ] += 1

            if candidate.file_type == "pdf":

                summary[
                    "discovered_pdf"
                ] += 1

            elif candidate.file_type == "excel":

                summary[
                    "discovered_excel"
                ] += 1

        elif candidate.status == (
            EXCLUDED_TEMPORARY_STATUS
        ):

            summary[
                "excluded_temporary"
            ] += 1

        elif candidate.status == (
            EXCLUDED_HIDDEN_STATUS
        ):

            summary[
                "excluded_hidden"
            ] += 1

        elif candidate.status == (
            EXCLUDED_UNSUPPORTED_STATUS
        ):

            summary[
                "excluded_unsupported"
            ] += 1

        elif candidate.status == (
            EXCLUDED_SYMLINK_STATUS
        ):

            summary[
                "excluded_symlink"
            ] += 1

        elif candidate.status == (
            EXCLUDED_REPARSE_STATUS
        ):

            summary[
                "excluded_reparse"
            ] += 1

    return summary


# ============================================================================
# Manifest Persistence
# ============================================================================

def save_manifest(
    manifest: pd.DataFrame,
    output_path: Path | str = FILE_MANIFEST_PATH,
) -> pd.DataFrame:
    """
    Save manifest to CSV.

    UTF-8 with BOM is retained for Windows/Excel compatibility.
    """

    output_path = Path(
        output_path
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    manifest.to_csv(
        output_path,
        index=False,
        encoding="utf-8-sig",
    )

    logger.info(
        "File manifest saved: %s",
        output_path,
    )

    return manifest


# ============================================================================
# High-Level Discovery Pipeline
# ============================================================================

def run_discovery(
    raw_dir: Path | str = RAW_DIR,
    manifest_path: Path | str = FILE_MANIFEST_PATH,
    recursive: bool = True,
    calculate_hash: bool = True,
    expected_file_count: int | None = None,
) -> pd.DataFrame:
    """
    Execute the complete Discovery stage.

    Pipeline
    --------
    1. Validate raw directory.
    2. Generate run identity.
    3. Generate discovery timestamp.
    4. Discover filesystem candidates.
    5. Classify accepted/excluded candidates.
    6. Build file-level metadata.
    7. Build provenance-aware manifest.
    8. Validate manifest.
    9. Save manifest.
    10. Return manifest.

    Discovery does not open PDF or Excel documents.
    """

    raw_dir = Path(
        raw_dir
    ).resolve()

    manifest_path = Path(
        manifest_path
    )

    discovery_run_id = (
        generate_discovery_run_id()
    )

    discovered_at = (
        get_utc_timestamp()
    )

    logger.info(
        "=" * 72
    )

    logger.info(
        "Starting file discovery stage"
    )

    logger.info(
        "Discovery run ID: %s",
        discovery_run_id,
    )

    logger.info(
        "Discovery timestamp UTC: %s",
        discovered_at,
    )

    logger.info(
        "Raw directory: %s",
        raw_dir,
    )

    logger.info(
        "Manifest schema version: %s",
        MANIFEST_SCHEMA_VERSION,
    )

    logger.info(
        "Pipeline version: %s",
        PIPELINE_VERSION,
    )

    logger.info(
        "Recursive search: %s",
        recursive,
    )

    logger.info(
        "SHA-256 enabled: %s",
        calculate_hash,
    )

    logger.info(
        "Supported extensions: %s",
        sorted(
            normalize_extensions(
                SUPPORTED_EXTENSIONS
            )
        ),
    )

    # ------------------------------------------------------------------------
    # Discover filesystem candidates
    # ------------------------------------------------------------------------

    candidates = discover_candidates(
        raw_dir=raw_dir,
        recursive=recursive,
        supported_extensions=(
            SUPPORTED_EXTENSIONS
        ),
    )

    # ------------------------------------------------------------------------
    # Candidate summary
    # ------------------------------------------------------------------------

    candidate_summary = (
        summarize_candidates(
            candidates
        )
    )

    logger.info(
        "Discovery summary: %s",
        candidate_summary,
    )

    # ------------------------------------------------------------------------
    # Accepted files
    # ------------------------------------------------------------------------

    discovered_files = [
        candidate.path
        for candidate in candidates
        if candidate.status
        == DISCOVERY_STATUS
    ]

    # ------------------------------------------------------------------------
    # Build manifest
    # ------------------------------------------------------------------------

    manifest = build_manifest(
        files=discovered_files,
        raw_dir=raw_dir,
        calculate_hash=calculate_hash,
        discovery_run_id=discovery_run_id,
        discovered_at=discovered_at,
    )

    # ------------------------------------------------------------------------
    # Validate
    # ------------------------------------------------------------------------

    validation = validate_manifest(
        manifest=manifest,
        raw_dir=raw_dir,
        expected_file_count=(
            expected_file_count
        ),
    )

    logger.info(
        "Manifest validation: %s",
        validation,
    )

    if not bool(
        validation["valid"]
    ):

        logger.warning(
            (
                "Discovery validation found "
                "blocking file-level issues."
            )
        )

    # ------------------------------------------------------------------------
    # Persist
    # ------------------------------------------------------------------------

    save_manifest(
        manifest=manifest,
        output_path=manifest_path,
    )

    logger.info(
        (
            "Discovery stage completed. "
            "Run=%s | Discovered files=%d"
        ),
        discovery_run_id,
        len(manifest),
    )

    logger.info(
        "=" * 72
    )

    return manifest


# ============================================================================
# Display Helpers
# ============================================================================

def display_discovery_summary(
    manifest: pd.DataFrame,
) -> None:
    """
    Display a concise human-readable summary.
    """

    if manifest.empty:

        print(
            "\nNo supported raw files were discovered."
        )

        return

    pdf_count = int(
        (
            manifest[
                "file_type"
            ]
            == "pdf"
        ).sum()
    )

    excel_count = int(
        (
            manifest[
                "file_type"
            ]
            == "excel"
        ).sum()
    )

    print(
        "\n=== Discovery Summary ==="
    )

    if "discovery_run_id" in manifest.columns:

        print(
            "Run ID      : "
            f"{manifest['discovery_run_id'].iloc[0]}"
        )

    if "discovered_at" in manifest.columns:

        print(
            "Discovered  : "
            f"{manifest['discovered_at'].iloc[0]}"
        )

    if "source_root" in manifest.columns:

        print(
            "Source root : "
            f"{manifest['source_root'].iloc[0]}"
        )

    print(
        f"Total files : {len(manifest)}"
    )

    print(
        f"PDF files   : {pdf_count}"
    )

    print(
        f"Excel files : {excel_count}"
    )

    print(
        "\n=== Discovered Files ==="
    )

    display_columns = [
        "file_name",
        "relative_path",
        "file_type",
        "extension",
        "size_bytes",
        "modified_time",
        "discovery_status",
    ]

    available_columns = [
        column
        for column in display_columns
        if column in manifest.columns
    ]

    print(
        manifest[
            available_columns
        ].to_string(
            index=False
        )
    )


# ============================================================================
# Command-Line Execution
# ============================================================================

def main() -> int:
    """
    Run Discovery as standalone module/script.
    """

    configure_logging()

    print(
        "\n"
        "======================================================================\n"
        " TRAFFIC_DATA - FILE DISCOVERY\n"
        " PDF + EXCEL SOURCE DISCOVERY\n"
        "======================================================================\n"
    )

    manifest = run_discovery(
        raw_dir=RAW_DIR,
        manifest_path=FILE_MANIFEST_PATH,
        recursive=True,
        calculate_hash=True,
        expected_file_count=None,
    )

    display_discovery_summary(
        manifest
    )

    print(
        "\n=== Output ==="
    )

    print(
        f"Manifest: {FILE_MANIFEST_PATH}"
    )

    print(
        "\nDiscovery stage completed successfully."
    )

    return 0


# ============================================================================
# Entry Point
# ============================================================================

if __name__ == "__main__":
    raise SystemExit(
        main()
    )