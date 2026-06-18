import os
import argparse
import sqlite3
import time
import pandas as pd


DEFAULT_LOCAL_LIBRARY_ROOT = r"C:\Users\H588238\OneDrive - Honeywell\TSI Team Site (GPS Sofia) - General\3 DATASHEETS\Datasheets Library"


def normalize_text(value: str) -> str:
    return (
        (value or "")
        .lower()
        .replace('.', '')
        .replace(' ', '')
        .replace('-', '')
        .replace('_', '')
        .replace('/', '')
        .replace('\\', '')
    )


def load_search_rows(conn: sqlite3.Connection):
    # Supports both legacy DBs (without normalized columns) and newer schema.
    columns = {
        row[1]
        for row in conn.execute("PRAGMA table_info(pdf_index)").fetchall()
    }

    if "file_name_norm" in columns and "first_page_text_norm" in columns:
        rows = conn.execute(
            "SELECT file_path, file_name_norm, first_page_text_norm, first_page_text FROM pdf_index"
        ).fetchall()
        optimized_rows = []
        for file_path, file_name_norm, first_page_text_norm, first_page_text in rows:
            fp = str(file_path or "")
            # Some DBs contain these columns but values are empty for many rows.
            file_name_norm_final = str(file_name_norm or "") or normalize_text(os.path.basename(fp))
            first_page_text_norm_final = str(first_page_text_norm or "") or normalize_text(str(first_page_text or ""))
            optimized_rows.append((fp, file_name_norm_final, first_page_text_norm_final))
        return optimized_rows

    rows = conn.execute("SELECT file_path, first_page_text FROM pdf_index").fetchall()
    optimized_rows = []
    for file_path, first_page_text in rows:
        fp = str(file_path or "")
        optimized_rows.append(
            (
                fp,
                normalize_text(os.path.basename(fp)),
                normalize_text(str(first_page_text or "")),
            )
        )
    return optimized_rows


def prepare_search_index(rows):
    file_name_rows = [(file_path, file_name_norm) for file_path, file_name_norm, _ in rows]
    first_page_rows = [(file_path, first_page_text_norm) for file_path, _, first_page_text_norm in rows]
    return file_name_rows, first_page_rows


def find_best_match(keyword_norm: str, file_name_rows, first_page_rows):
    if not keyword_norm:
        return None

    for file_path, file_name_norm in file_name_rows:
        if keyword_norm in file_name_norm:
            return file_path

    for file_path, first_page_text_norm in first_page_rows:
        if keyword_norm in first_page_text_norm:
            return file_path

    return None


def resolve_existing_file(candidates):
    existing = [p for p in candidates if os.path.isfile(p)]
    if not existing:
        return candidates[0]
    return max(existing, key=os.path.getmtime)


def to_local_path_if_possible(path_value: str, local_library_root: str) -> str:
    path_text = str(path_value or "")
    if not path_text or path_text == "Not Found":
        return path_text

    try:
        relative_path = os.path.relpath(path_text, local_library_root)
    except ValueError:
        return path_text

    if relative_path.startswith(".."):
        return path_text

    return path_text


def parse_args():
    parser = argparse.ArgumentParser(description="Find best PDF matches and write mapped links to Excel.")
    parser.add_argument("--excel-file", help="Path to PART_MODEL_LIST file (.xlsm/.xlsx).")
    parser.add_argument("--database-file", help="Path to sqlite database file.")
    parser.add_argument("--output-excel", help="Output Excel file path.")
    parser.add_argument(
        "--local-library-root",
        default=os.getenv("PDF_LOCAL_LIBRARY_ROOT", DEFAULT_LOCAL_LIBRARY_ROOT),
        help="Local OneDrive-synced Datasheets Library root.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    # File paths (resolved from script location for portability)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(script_dir, "..", ".."))
    current_dir = os.getcwd()

    excel_candidates = [
        os.path.join(script_dir, "PART_MODEL_LIST.xlsm"),
        os.path.join(script_dir, "PART_MODEL_LIST.xlsx"),
        os.path.join(current_dir, "PART_MODEL_LIST.xlsm"),
        os.path.join(current_dir, "PART_MODEL_LIST.xlsx"),
        os.path.join(project_root, "PART_MODEL_LIST.xlsm"),
        os.path.join(project_root, "PART_MODEL_LIST.xlsx"),
    ]
    excel_file = args.excel_file or resolve_existing_file(excel_candidates)

    db_candidates = [
        os.path.join(args.local_library_root, "_index", "pdf_text_map.db"),
        os.path.join(script_dir, "pdf_text_map.db"),
        os.path.join(script_dir, "pdf_text_map_All.db"),
        os.path.join(current_dir, "pdf_text_map.db"),
        os.path.join(current_dir, "pdf_text_map_All.db"),
        os.path.join(project_root, "@Test", "pdf_text_map_All.db"),
        os.path.join(project_root, "pdf_text_map.db"),
        os.path.join(project_root, "pdf_text_map_All.db"),
    ]
    database_file = args.database_file or resolve_existing_file(db_candidates)

    output_excel = args.output_excel or os.path.join(
        os.path.dirname(os.path.abspath(excel_file)),
        "PART_MODEL_with_results_PDF.xlsx",
    )

    # Validate paths
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"Excel file not found: {excel_file}")
    if not os.path.isfile(database_file):
        raise FileNotFoundError(f"Database file not found: {database_file}")

    # Load Excel file
    df = pd.read_excel(excel_file, engine='openpyxl')
    if df.shape[1] < 2:
        raise ValueError("Expected at least two columns (A and B) in the Excel file")

    # Ensure column C exists
    if df.shape[1] == 2:
        df.insert(2, 'Result', '')  # Insert column C at index 2

    # Extract keywords from column B
    keywords = df.iloc[:, 1].fillna('').astype(str).tolist()
    normalized_keywords = [normalize_text(kw) for kw in keywords]

    # Connect to the database
    conn = sqlite3.connect(database_file)
    search_rows = load_search_rows(conn)
    file_name_rows, first_page_rows = prepare_search_index(search_rows)

    # Search logic
    start_time = time.time()
    results = []
    match_cache = {}

    for keyword_norm in normalized_keywords:
        if not keyword_norm:
            results.append("Not Found")
            continue

        if keyword_norm not in match_cache:
            found_path = find_best_match(keyword_norm, file_name_rows, first_page_rows)
            match_cache[keyword_norm] = found_path if found_path else "Not Found"

        results.append(
            to_local_path_if_possible(
                match_cache[keyword_norm],
                args.local_library_root,
            )
        )

    # Close DB connection
    conn.close()

    # Write results to column C
    df.iloc[:, 2] = results  # Column C is index 2

    # Save updated Excel file
    df.to_excel(output_excel, index=False)
    print(f"Input Excel: {excel_file}")
    print(f"Database: {database_file}")
    print(f"Results saved to {output_excel}")
    print(f"Search completed in {time.time() - start_time:.2f} seconds")


if __name__ == "__main__":
    main()
