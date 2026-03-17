#!/usr/bin/env python3
"""
Hardball Dynasty Draft Optimizer — GUI launcher.

Three actions: Fetch data (scrape draft pool to Excel), Sort master list (reapply formula and sort), Push to Hardball Dynasty (push order to the site). Log output appears in the window.
"""
import io
import logging
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

from app_dir import get_app_dir

# Default paths (same as main.py)
DEFAULT_TEMPLATE = get_app_dir() / "Season x amateur draft-template.xlsx"
OUTPUTS_DIR = get_app_dir() / "outputs"


def _latest_output() -> Path | None:
    if not OUTPUTS_DIR.is_dir():
        return None
    files = sorted(OUTPUTS_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0] if files else None


class QueueHandler(logging.Handler):
    """Send log records to a queue for the GUI thread to display."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        try:
            self.log_queue.put(self.format(record))
        except Exception:
            pass


def run_fetch(
    excel_path: Path,
    output_dir: Path,
    top_n: int,
    headless: bool,
    user_data_dir: str | None,
    log_queue: queue.Queue | None = None,
) -> tuple[bool, str]:
    """Run fetch in this thread. Returns (success, message)."""
    try:
        from credentials import get_headless
        from excel_draft import validate_template
        from web_draft import run_sync_from_web_to_excel

        headless = headless or get_headless()
        validation = validate_template(excel_path)
        if validation:
            return False, "Template validation failed:\n  " + "\n  ".join(validation)

        if log_queue is not None:
            old_stdout = sys.stdout
            sys.stdout = io.StringIO()
        try:
            run_sync_from_web_to_excel(
                str(excel_path),
                headless=headless,
                user_data_dir=user_data_dir,
                top_n=top_n,
                output_dir=str(output_dir),
            )
            if log_queue is not None and sys.stdout.getvalue():
                for line in sys.stdout.getvalue().strip().splitlines():
                    log_queue.put(line)
        finally:
            if log_queue is not None:
                sys.stdout = old_stdout
        return True, "Fetch completed successfully."
    except Exception as e:
        logging.exception("Fetch failed")
        return False, str(e)


def run_apply_order_sort_only(excel_path: Path) -> tuple[bool, str]:
    """Reapply formula and sort Master List. Returns (success, message)."""
    try:
        from excel_draft import reapply_formula_and_sort_master_list
        ok = reapply_formula_and_sort_master_list(excel_path)
        return ok, "Master list updated (sorted by adjusted score)." if ok else "Could not sort Master List (Excel COM required on Windows)."
    except Exception as e:
        logging.exception("Apply order sort failed")
        return False, str(e)


def run_apply_order_push(
    excel_path: Path,
    headless: bool,
    user_data_dir: str | None,
    log_queue: queue.Queue | None = None,
) -> tuple[bool, str]:
    """Push Excel order to Hardball Dynasty web. Returns (success, message)."""
    try:
        from credentials import get_headless
        from web_draft import run_apply_excel_order_to_web

        headless = headless or get_headless()
        if log_queue is not None:
            old_stdout = sys.stdout
            sys.stdout = io.StringIO()
        try:
            run_apply_excel_order_to_web(
                str(excel_path),
                headless=headless,
                user_data_dir=user_data_dir,
            )
            if log_queue is not None and sys.stdout.getvalue():
                for line in sys.stdout.getvalue().strip().splitlines():
                    log_queue.put(line)
        finally:
            if log_queue is not None:
                sys.stdout = old_stdout
        return True, "Order applied and saved to Hardball Dynasty."
    except Exception as e:
        logging.exception("Apply order push failed")
        return False, str(e)


class DraftOptimizerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Hardball Dynasty Draft Optimizer")
        self.root.minsize(500, 400)
        self.root.geometry("620x480")

        self.log_queue: queue.Queue = queue.Queue()
        self.running = False

        self._build_ui()
        self._setup_logging()
        self._poll_log_queue()
        self._update_file_labels()

    def _build_ui(self):
        main = tk.Frame(self.root, padx=12, pady=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Buttons row
        btn_frame = tk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(0, 8))

        self.btn_fetch = tk.Button(
            btn_frame,
            text="Fetch data",
            command=self._on_fetch,
            width=16,
            font=("Segoe UI", 10),
        )
        self.btn_fetch.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_sort = tk.Button(
            btn_frame,
            text="Sort master list",
            command=self._on_sort_master_list,
            width=16,
            font=("Segoe UI", 10),
        )
        self.btn_sort.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_push = tk.Button(
            btn_frame,
            text="Push to Hardball Dynasty",
            command=self._on_push,
            width=22,
            font=("Segoe UI", 10),
        )
        self.btn_push.pack(side=tk.LEFT, padx=(0, 8))

        # File labels
        file_frame = tk.Frame(main)
        file_frame.pack(fill=tk.X, pady=(0, 4))

        tk.Label(file_frame, text="Template:", font=("Segoe UI", 9), width=10, anchor="w").pack(side=tk.LEFT)
        self.lbl_template = tk.Label(file_frame, text="", font=("Segoe UI", 9), anchor="w", fg="gray")
        self.lbl_template.pack(side=tk.LEFT, fill=tk.X, expand=True)

        file_frame2 = tk.Frame(main)
        file_frame2.pack(fill=tk.X, pady=(0, 8))

        tk.Label(file_frame2, text="Output file:", font=("Segoe UI", 9), width=10, anchor="w").pack(side=tk.LEFT)
        self.lbl_output_file = tk.Label(file_frame2, text="", font=("Segoe UI", 9), anchor="w", fg="gray")
        self.lbl_output_file.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Log area
        tk.Label(main, text="Log", font=("Segoe UI", 9)).pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(
            main,
            height=16,
            font=("Consolas", 9),
            state=tk.DISABLED,
            wrap=tk.WORD,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

    def _setup_logging(self):
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        handler = QueueHandler(self.log_queue)
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        root_logger.addHandler(handler)

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
        except queue.Empty:
            pass
        self.root.after(200, self._poll_log_queue)

    def _append_log(self, msg: str):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg.strip() + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _update_file_labels(self):
        if DEFAULT_TEMPLATE.exists():
            self.lbl_template.config(text=DEFAULT_TEMPLATE.name, fg="black")
        else:
            self.lbl_template.config(text="(none — choose before Fetch)", fg="gray")

        latest = _latest_output()
        if latest:
            self.lbl_output_file.config(text=latest.name, fg="black")
        else:
            self.lbl_output_file.config(text="(none — run Fetch first)", fg="gray")

    def _set_buttons_enabled(self, enabled: bool):
        self.btn_fetch.config(state=tk.NORMAL if enabled else tk.DISABLED)
        self.btn_sort.config(state=tk.NORMAL if enabled else tk.DISABLED)
        self.btn_push.config(state=tk.NORMAL if enabled else tk.DISABLED)
        self.running = not enabled

    def _get_output_excel_path(self) -> Path | None:
        """Return latest output file or prompt user. Returns None if cancelled."""
        excel_path = _latest_output()
        if not excel_path or not excel_path.exists():
            path = filedialog.askopenfilename(
                title="Select Excel file (draft output)",
                filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
                initialdir=OUTPUTS_DIR if OUTPUTS_DIR.is_dir() else get_app_dir(),
            )
            if not path or not Path(path).exists():
                return None
            return Path(path)
        return Path(excel_path)

    def _on_fetch(self):
        excel_path = DEFAULT_TEMPLATE
        if not excel_path.exists():
            excel_path = Path(
                filedialog.askopenfilename(
                    title="Select Excel template",
                    filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
                    initialdir=get_app_dir(),
                )
            )
            if not excel_path or not Path(excel_path).exists():
                return
        output_dir = OUTPUTS_DIR
        output_dir.mkdir(parents=True, exist_ok=True)

        self._append_log("Starting Fetch... (browser will open; log in if prompted)")
        self._set_buttons_enabled(False)

        def work():
            success, msg = run_fetch(
                Path(excel_path),
                output_dir,
                top_n=500,
                headless=False,
                user_data_dir=None,
                log_queue=self.log_queue,
            )
            self.root.after(0, lambda: self._task_done("Fetch", success, msg))

        threading.Thread(target=work, daemon=True).start()

    def _on_sort_master_list(self):
        excel_path = self._get_output_excel_path()
        if excel_path is None:
            return
        self._append_log("Sort master list: Reapplying formula and sorting...")
        self._set_buttons_enabled(False)

        def work():
            success, msg = run_apply_order_sort_only(excel_path)
            self.root.after(0, lambda: self._task_done("Sort master list", success, msg))

        threading.Thread(target=work, daemon=True).start()

    def _on_push(self):
        excel_path = self._get_output_excel_path()
        if excel_path is None:
            return
        self._append_log("Push to Hardball Dynasty... (browser will open)")
        self._set_buttons_enabled(False)

        def work():
            success, msg = run_apply_order_push(
                excel_path, headless=False, user_data_dir=None, log_queue=self.log_queue
            )
            self.root.after(0, lambda: self._task_done("Push to Hardball Dynasty", success, msg))

        threading.Thread(target=work, daemon=True).start()

    def _task_done(self, name: str, success: bool, message: str):
        self._set_buttons_enabled(True)
        self._update_file_labels()
        self._append_log(message)
        if success:
            messagebox.showinfo(name, message)
        else:
            messagebox.showerror(name, message)

    def run(self):
        self.root.mainloop()


def main():
    app = DraftOptimizerApp()
    app.run()


if __name__ == "__main__":
    main()
