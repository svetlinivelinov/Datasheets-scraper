import argparse
import hashlib
import os
import re
import sqlite3

import fitz  # PyMuPDF


DEFAULT_LOCAL_LIBRARY_ROOT = r"C:\Users\H588238\OneDrive - Honeywell\TSI Team Site (GPS Sofia) - General\3 DATASHEETS\Datasheets Library"
DEFAULT_FOLDER_FILTER_PATTERN = r"datasheets?"  # Matches 'datasheet' or 'datasheets'


def configure_mupdf_logging() -> None:
    # Some PDFs emit MuPDF warnings/errors (for example bad embedded color profiles)
    # that are non-fatal for this indexing workflow.
    tools = getattr(fitz, "TOOLS", None)
    if not tools:
        return

    for method_name in ("mupdf_display_errors", "mupdf_display_warnings"):
        method = getattr(tools, method_name, None)
        if callable(method):
            try:
                method(False)
            except Exception:
                pass


def local_app_data_file(file_name: str) -> str:
    base = os.getenv("LOCALAPPDATA")
    if not base:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)
    target_dir = os.path.join(base, "DatasheetsScraper")
    os.makedirs(target_dir, exist_ok=True)
    return os.path.join(target_dir, file_name)


def parse_args():
    parser = argparse.ArgumentParser(description="Build or update local PDF index database.")
    parser.add_argument(
        "--database-file",
        default=os.getenv("PDF_INDEX_DB", local_app_data_file("pdf_text_map.db")),
        help="Output SQLite DB path (local per user).",
    )
    parser.add_argument(
        "--log-file",
        default=local_app_data_file("pdf_index_log.txt"),
        help="Log file path.",
    )
    parser.add_argument(
        "--root-folder",
        action="append",
        help="Local root folder to scan. Can be passed multiple times.",
    )
    parser.add_argument(
        "--folder-filter-pattern",
        default=DEFAULT_FOLDER_FILTER_PATTERN,
        help="Regex used on folder names to include PDFs.",
    )
    parser.add_argument(
        "--prune-missing",
        action="store_true",
        help="Remove DB rows for files that no longer exist under scanned root folders.",
    )
    return parser.parse_args()


def compute_file_hash(file_path: str) -> str:
    hasher = hashlib.sha256()
    with open(file_path, "rb") as f:
        while chunk := f.read(8192):
            hasher.update(chunk)
    return hasher.hexdigest()


def normalize_text(value: str) -> str:
    return (
        (value or "")
        .lower()
        .replace(".", "")
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
        .replace("/", "")
        .replace("\\", "")
    )


def ensure_schema(cursor: sqlite3.Cursor) -> None:
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS pdf_index (
            file_path TEXT PRIMARY KEY,
            file_size INTEGER,
            file_hash TEXT,
            file_name_norm TEXT,
            first_page_text TEXT,
            first_page_text_norm TEXT,
            last_modified REAL
        )
        """
    )

    columns = {row[1] for row in cursor.execute("PRAGMA table_info(pdf_index)").fetchall()}
    if "file_name_norm" not in columns:
        cursor.execute("ALTER TABLE pdf_index ADD COLUMN file_name_norm TEXT")
    if "first_page_text_norm" not in columns:
        cursor.execute("ALTER TABLE pdf_index ADD COLUMN first_page_text_norm TEXT")

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_pdf_index_file_size ON pdf_index(file_size)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_pdf_index_file_hash ON pdf_index(file_hash)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_pdf_index_file_name_norm ON pdf_index(file_name_norm)")


def collect_pdf_files(root_folders: list[str], folder_filter_pattern: str) -> list[str]:
    pdf_files = []
    for root_folder in root_folders:
        for root, _, files in os.walk(root_folder):
            if any(re.search(folder_filter_pattern, part, re.IGNORECASE) for part in root.split(os.sep)):
                for file_name in files:
                    if file_name.lower().endswith(".pdf"):
                        pdf_files.append(os.path.join(root, file_name))
    return pdf_files


def main():
    args = parse_args()
    configure_mupdf_logging()

    root_folders = args.root_folder or [os.getenv("PDF_LOCAL_LIBRARY_ROOT", DEFAULT_LOCAL_LIBRARY_ROOT)]
    database_file = os.path.abspath(args.database_file)
    log_file = os.path.abspath(args.log_file)
    folder_filter_pattern = args.folder_filter_pattern
    prune_missing = args.prune_missing

    for root_folder in root_folders:
        if root_folder.lower().startswith(("http://", "https://")):
            raise ValueError(f"ROOT_FOLDERS must contain local paths, not web URLs: {root_folder}")
        if not os.path.isdir(root_folder):
            raise FileNotFoundError(f"Root folder not found: {root_folder}")

    db_dir = os.path.dirname(database_file)
    if db_dir:
        os.makedirs(db_dir, exist_ok=True)
    log_dir = os.path.dirname(log_file)
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)

    conn = sqlite3.connect(database_file)
    cursor = conn.cursor()
    ensure_schema(cursor)
    conn.commit()

    updated = 0
    skipped = 0
    pruned = 0

    with open(log_file, "w", encoding="utf-8") as log:
        log.write("PDF Index Log\n")
        log.write("=====================\n")
        log.write(f"Database file: {database_file}\n")
        log.write(f"Root folders: {root_folders}\n")

        pdf_files = collect_pdf_files(root_folders, folder_filter_pattern)
        scanned_file_set = set(pdf_files)
        log.write(f'Found {len(pdf_files)} PDF files in folders matching "{folder_filter_pattern}"\n\n')

        existing_rows = cursor.execute(
            "SELECT file_path, file_size, file_hash, last_modified FROM pdf_index"
        ).fetchall()

        size_to_entries: dict[int, list[tuple[str, str, float]]] = {}
        hash_to_best_mtime: dict[str, float] = {}

        for existing_path, existing_size, existing_hash, existing_mtime in existing_rows:
            if existing_size is None or not existing_hash:
                continue
            mtime_value = float(existing_mtime or 0)
            entry = (existing_path, existing_hash, mtime_value)
            size_to_entries.setdefault(existing_size, []).append(entry)
            if existing_hash not in hash_to_best_mtime or mtime_value > hash_to_best_mtime[existing_hash]:
                hash_to_best_mtime[existing_hash] = mtime_value

        if prune_missing:
            stale_paths = []
            for existing_path, _, _, _ in existing_rows:
                if not existing_path:
                    continue
                if existing_path not in scanned_file_set:
                    stale_paths.append(existing_path)

            if stale_paths:
                cursor.executemany(
                    "DELETE FROM pdf_index WHERE file_path = ?",
                    [(p,) for p in stale_paths],
                )
                pruned = len(stale_paths)
                log.write(f"Pruned {pruned} stale DB rows for missing files.\n")
            else:
                log.write("Prune enabled: no stale DB rows found.\n")

        rows_to_upsert = []
        paths_to_delete = set()

        for file_path in pdf_files:
            try:
                file_size = os.path.getsize(file_path)
                last_modified = os.path.getmtime(file_path)
                file_hash = compute_file_hash(file_path)

                should_skip_current = False
                potential_duplicates = size_to_entries.get(file_size, [])

                for existing_path, existing_hash, existing_mtime in potential_duplicates:
                    if existing_hash != file_hash:
                        continue

                    best_known_mtime = hash_to_best_mtime.get(existing_hash, existing_mtime)
                    if last_modified > best_known_mtime:
                        paths_to_delete.add(existing_path)
                        break

                    skipped += 1
                    log.write(f"Skipped (duplicate): {file_path}\n")
                    should_skip_current = True
                    break

                if should_skip_current:
                    continue

                try:
                    with fitz.open(file_path) as doc:
                        first_page_text = doc[0].get_text("text") if len(doc) > 0 else ""
                except Exception as ex:
                    # Keep indexing by filename even if first-page text extraction fails.
                    first_page_text = ""
                    log.write(f"Warning: text extraction failed for {file_path}: {ex}\n")

                rows_to_upsert.append(
                    (
                        file_path,
                        file_size,
                        file_hash,
                        normalize_text(os.path.basename(file_path)),
                        first_page_text,
                        normalize_text(first_page_text),
                        last_modified,
                    )
                )

                current_entry = (file_path, file_hash, last_modified)
                best_mtime = hash_to_best_mtime.get(file_hash)
                if best_mtime is None or last_modified > best_mtime:
                    hash_to_best_mtime[file_hash] = last_modified
                size_to_entries.setdefault(file_size, []).append(current_entry)

                updated += 1
                log.write(f"Updated: {file_path}\n")
            except Exception as ex:
                log.write(f"Error processing {file_path}: {ex}\n")

    if paths_to_delete:
        cursor.executemany("DELETE FROM pdf_index WHERE file_path = ?", [(p,) for p in paths_to_delete])

    if rows_to_upsert:
        cursor.executemany(
            """
            INSERT INTO pdf_index (
                file_path,
                file_size,
                file_hash,
                file_name_norm,
                first_page_text,
                first_page_text_norm,
                last_modified
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(file_path) DO UPDATE SET
                file_size=excluded.file_size,
                file_hash=excluded.file_hash,
                file_name_norm=excluded.file_name_norm,
                first_page_text=excluded.first_page_text,
                first_page_text_norm=excluded.first_page_text_norm,
                last_modified=excluded.last_modified
            """,
            rows_to_upsert,
        )

    conn.commit()
    conn.close()

    with open(log_file, "a", encoding="utf-8") as log:
        log.write(
            f"\nIndexing complete. {updated} files updated, {skipped} files skipped due to duplication, {pruned} stale rows pruned.\n"
        )

    print(f"Database: {database_file}")
    print(f"Log: {log_file}")
    print(f"Indexing complete. Updated={updated}, Skipped duplicates={skipped}, Pruned={pruned}")


if __name__ == "__main__":
    main()
