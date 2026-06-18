import os
import sys
import threading
import subprocess
import importlib
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BUILD_MAP_SCRIPT = os.path.join(SCRIPT_DIR, "01_build_pdf_map.py")
SEARCH_SCRIPT = os.path.join(SCRIPT_DIR, "02_file_search_PDF.py")
DOWNLOAD_SCRIPT = os.path.join(SCRIPT_DIR, "03_download_sharepoint_links.ps1")

DEFAULT_LOCAL_ROOT = r"C:\Users\H588238\OneDrive - Honeywell\TSI Team Site (GPS Sofia) - General\3 DATASHEETS\Datasheets Library"


def local_app_data_file(file_name: str) -> str:
    base = os.getenv("LOCALAPPDATA")
    if not base:
        return os.path.join(SCRIPT_DIR, file_name)
    target_dir = os.path.join(base, "DatasheetsScraper")
    os.makedirs(target_dir, exist_ok=True)
    return os.path.join(target_dir, file_name)


class PipelineUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Datasheets Pipeline UI")
        self.geometry("980x680")

        self.local_root_var = tk.StringVar(value=DEFAULT_LOCAL_ROOT)
        self.part_model_var = tk.StringVar(value="")
        self.db_file_var = tk.StringVar(value=local_app_data_file("pdf_text_map.db"))
        self.last_output_excel = ""

        self._build_ui()

    def _build_ui(self):
        row = 0

        tk.Label(self, text="Local synced Datasheets folder:").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        tk.Entry(self, textvariable=self.local_root_var, width=110).grid(row=row, column=1, sticky="we", padx=8, pady=6)
        tk.Button(self, text="Browse", command=self.pick_local_root).grid(row=row, column=2, padx=8, pady=6)

        row += 1
        tk.Label(self, text="PART_MODEL_LIST file:").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        tk.Entry(self, textvariable=self.part_model_var, width=110).grid(row=row, column=1, sticky="we", padx=8, pady=6)
        tk.Button(self, text="Browse", command=self.pick_part_model).grid(row=row, column=2, padx=8, pady=6)

        row += 1
        tk.Label(self, text="Database file (optional):").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        tk.Entry(self, textvariable=self.db_file_var, width=110).grid(row=row, column=1, sticky="we", padx=8, pady=6)
        tk.Button(self, text="Browse", command=self.pick_db_file).grid(row=row, column=2, padx=8, pady=6)

        row += 1
        btn_frame = tk.Frame(self)
        btn_frame.grid(row=row, column=0, columnspan=3, sticky="w", padx=8, pady=10)
        tk.Button(btn_frame, text="Step 1: Run 01 Build Map", command=self.run_step1).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Step 2: Run 02 Search", command=self.run_step2).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Step 3: Run 03 Download", command=self.run_step3).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Run All", command=self.run_all).pack(side="left", padx=6)

        row += 1
        self.log = ScrolledText(self, height=28)
        self.log.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(row, weight=1)

    def pick_local_root(self):
        path = filedialog.askdirectory(title="Select local synced Datasheets folder")
        if path:
            self.local_root_var.set(path)

    def pick_part_model(self):
        path = filedialog.askopenfilename(
            title="Select PART_MODEL_LIST file",
            filetypes=[("Excel", "*.xlsm *.xlsx")],
        )
        if path:
            self.part_model_var.set(path)

    def pick_db_file(self):
        path = filedialog.askopenfilename(
            title="Select DB file",
            filetypes=[("SQLite DB", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")],
        )
        if path:
            self.db_file_var.set(path)

    def append_log(self, text: str):
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def validate_common_inputs(self) -> bool:
        part_model = self.part_model_var.get().strip()
        local_root = self.local_root_var.get().strip()
        if not part_model or not os.path.isfile(part_model):
            messagebox.showerror("Missing file", "Select a valid PART_MODEL_LIST file.")
            return False
        if not local_root or not os.path.isdir(local_root):
            messagebox.showerror("Missing folder", "Select a valid local synced Datasheets folder.")
            return False
        return True

    def run_command(self, cmd, cwd=None):
        self.append_log("Command: " + " ".join(cmd))
        proc = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
        if proc.stdout:
            self.append_log(proc.stdout.rstrip())
        if proc.stderr:
            self.append_log(proc.stderr.rstrip())
        if proc.returncode != 0:
            raise RuntimeError(f"Command failed with exit code {proc.returncode}")

    def ensure_pipeline_dependencies(self):
        required = {
            "pandas": "pandas",
            "openpyxl": "openpyxl",
            "fitz": "pymupdf",
        }

        missing = []
        for module_name, package_name in required.items():
            try:
                importlib.import_module(module_name)
            except Exception:
                missing.append(package_name)

        if not missing:
            return

        unique_missing = sorted(set(missing))
        self.append_log("Installing missing packages: " + ", ".join(unique_missing))
        self.run_command([sys.executable, "-m", "pip", "install", *unique_missing], cwd=SCRIPT_DIR)

    def output_excel_path(self) -> str:
        base_dir = os.path.dirname(os.path.abspath(self.part_model_var.get().strip()))
        return os.path.join(base_dir, "PART_MODEL_with_results_PDF.xlsx")

    def _is_file_writable_or_free(self, path: str) -> bool:
        if not os.path.exists(path):
            return True
        try:
            with open(path, "a+b"):
                return True
        except OSError:
            return False

    def _search_output_excel_path(self) -> str:
        preferred = self.output_excel_path()
        if self._is_file_writable_or_free(preferred):
            return preferred

        base_dir = os.path.dirname(preferred)
        base_name = os.path.splitext(os.path.basename(preferred))[0]
        ext = os.path.splitext(preferred)[1] or ".xlsx"
        timestamp = threading.current_thread().name  # deterministic non-empty fallback
        # Use monotonic-ish timestamp from wall clock for readability.
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(base_dir, f"{base_name}_{timestamp}{ext}")

    def run_step1(self):
        self._run_async(self._step1)

    def run_step2(self):
        self._run_async(self._step2)

    def run_step3(self):
        self._run_async(self._step3)

    def run_all(self):
        self._run_async(self._run_all)

    def _run_async(self, target):
        t = threading.Thread(target=target, daemon=True)
        t.start()

    def _step1(self):
        try:
            if not self.validate_common_inputs():
                return False

            self.ensure_pipeline_dependencies()

            _, db_file = self._build_db()
            self.append_log(f"Step 1 completed. DB build/update completed: {db_file}")
            return True
        except Exception as exc:
            self.append_log(f"ERROR (Step 1): {exc}")
            return False

    def _build_db(self):
        local_root = self.local_root_var.get().strip()
        db_file = self.db_file_var.get().strip() or local_app_data_file("pdf_text_map.db")

        # Always build/update each user's local DB before searching.
        build_cmd = [
            sys.executable,
            BUILD_MAP_SCRIPT,
            "--database-file",
            db_file,
            "--root-folder",
            local_root,
            "--prune-missing",
        ]
        self.run_command(build_cmd, cwd=SCRIPT_DIR)
        return local_root, db_file

    def _step2(self):
        try:
            if not self.validate_common_inputs():
                return False

            self.ensure_pipeline_dependencies()

            local_root = self.local_root_var.get().strip()
            part_model = self.part_model_var.get().strip()
            db_file = self.db_file_var.get().strip() or local_app_data_file("pdf_text_map.db")
            output_excel = self._search_output_excel_path()

            if not os.path.isfile(db_file):
                raise RuntimeError(
                    f"Step 1 DB file not found: {db_file}. Run Step 1 first."
                )

            cmd = [
                sys.executable,
                SEARCH_SCRIPT,
                "--excel-file", part_model,
                "--output-excel", output_excel,
                "--local-library-root", local_root,
                "--database-file", db_file,
            ]

            self.run_command(cmd, cwd=SCRIPT_DIR)
            self.last_output_excel = output_excel
            self.append_log(f"Excel with links created: {output_excel}")
            self.append_log("Step 2 completed.")
            return True
        except Exception as exc:
            self.append_log(f"ERROR (Step 2): {exc}")
            return False

    def _step3(self):
        try:
            if not self.validate_common_inputs():
                return False

            local_root = self.local_root_var.get().strip()
            output_excel = self.last_output_excel or self.output_excel_path()

            if not os.path.isfile(output_excel):
                raise RuntimeError(f"Step 2 output not found: {output_excel}")

            out_dir = os.path.dirname(os.path.abspath(self.part_model_var.get().strip()))

            cmd = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                DOWNLOAD_SCRIPT,
                "-ExcelFile",
                output_excel,
                "-OutDir",
                out_dir,
                "-DatasheetsLibraryRoot",
                local_root,
            ]

            self.run_command(cmd, cwd=SCRIPT_DIR)
            self.append_log("Step 3 completed.")
            return True
        except Exception as exc:
            self.append_log(f"ERROR (Step 3): {exc}")
            return False

    def _run_all(self):
        if not self._step1():
            self.append_log("Run All stopped at Step 1 due to error.")
            return
        if not self._step2():
            self.append_log("Run All stopped at Step 2 due to error.")
            return
        self._step3()


if __name__ == "__main__":
    app = PipelineUI()
    app.mainloop()
