"""
Aplicació GUI per convertir fitxers DOCX a PDF i desar-los en un ZIP
amb contrasenya.

Millores:
- Execució paral·lela configurable (1-16)
- Reintents agressius + mètode alternatiu via win32com (si disponible)
- Informe final detallat (resum, reintents resolts, errors pendents)
- Correccions d'estil PEP8, docstrings i tipat bàsic
Requisits: pip install docx2pdf pyzipper psutil
"""

from __future__ import annotations

import gc
import os
import sys
import queue
import shutil
import tempfile
import threading
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Tuple

try:
    import psutil  # type: ignore
except Exception:
    psutil = None

try:
    import pyzipper  # type: ignore
except Exception:
    pyzipper = None

try:

    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

from conversion_report import ConversionReport
from report_dialog import ReportDialog
from docx_converter import DocxConverter

# Constants
CONVERSION_TIMEOUT = 600
MAX_RETRIES = 5
MEMORY_THRESHOLD_MB = 500
DEFAULT_PARALLEL = 8
MAX_PARALLEL_ALLOWED = 16
STATUS_COLORS = {
    "info": "blue",
    "success": "green",
    "error": "red",
    "warning": "orange",
}


# Redirigir stderr per evitar soroll en executables empaquetats
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    try:
        sys.stderr = open(os.devnull, "w")
    except Exception:
        pass


class DocxToPdfZipApp:
    """Aplicació principal amb UI i lògica de conversió + ZIP."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Convertidor DOCX a PDF amb ZIP (Paral·lel)")
        self.root.geometry("760x640")
        self.root.resizable(True, True)

        # Estat
        self.docx_files: List[str] = []
        self.is_processing = False
        self.cancel_flag = False
        self.message_queue: queue.Queue[Tuple[str, Dict | int]] = queue.Queue()
        self.conversion_thread: Optional[threading.Thread] = None
        self.completed_conversions = 0
        self.conversion_report = ConversionReport()

        # UI refs
        self.files_listbox: tk.Listbox
        self.files_label: ttk.Label
        self.password_entry: ttk.Entry
        self.password_confirm_entry: ttk.Entry
        self.dest_entry: ttk.Entry
        self.parallel_spinbox: ttk.Spinbox
        self.progress: ttk.Progressbar
        self.status_label: ttk.Label
        self.active_label: ttk.Label
        self.use_password_var: tk.BooleanVar

        self.check_dependencies()
        self.setup_ui()
        self.process_queue()

    def check_dependencies(self) -> None:
        missing = []
        if not DocxConverter.is_docx2pdf_available():
            missing.append("docx2pdf")
        if not DocxConverter.is_pyzipper_available():
            missing.append("pyzipper")
        if missing:
            messagebox.showerror(
                "Dependències",
                "Falten dependències: "
                + ", ".join(missing)
                + "\n\nInstal·la-les amb: pip install "
                + " ".join(missing),
            )
            self.root.destroy()

    def setup_ui(self) -> None:
        main = ttk.Frame(self.root, padding="10")
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        self._setup_title(main)
        self._setup_file_selection(main)
        self._setup_file_list(main)
        self._setup_config(main)
        self._setup_destination(main)
        self._setup_actions(main)
        self._setup_progress_and_status(main)

    def _setup_title(self, main: ttk.Frame) -> None:
        ttk.Label(
            main,
            text="Convertidor DOCX → PDF → ZIP (Paral·lel)",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, columnspan=3, pady=8)

    def _setup_file_selection(self, main: ttk.Frame) -> None:
        ttk.Button(
            main,
            text="📁 Seleccionar fitxers DOCX",
            command=self.select_files,
            width=30,
        ).grid(row=1, column=0, pady=8, padx=5, sticky="w")
        self.files_label = ttk.Label(main, text="Cap fitxer seleccionat")
        self.files_label.grid(row=1, column=1, columnspan=2, pady=8, sticky="w")

    def _setup_file_list(self, main: ttk.Frame) -> None:
        list_frame = ttk.LabelFrame(main, text="Fitxers seleccionats", padding="6")
        list_frame.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=8)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self.files_listbox = tk.Listbox(
            list_frame, yscrollcommand=scroll.set, height=10, width=90
        )
        self.files_listbox.grid(row=0, column=0, sticky="nsew")
        scroll.config(command=self.files_listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")

    def _setup_config(self, main: ttk.Frame) -> None:
        cfg = ttk.LabelFrame(main, text="Configuració", padding="8")
        cfg.grid(row=3, column=0, columnspan=3, sticky="ew", pady=8)
        cfg.columnconfigure(1, weight=1)

        ttk.Label(cfg, text="Contrasenya:").grid(
            row=0, column=0, padx=5, pady=2, sticky="w"
        )
        self.password_entry = ttk.Entry(cfg, show="*", width=25)
        self.password_entry.grid(row=0, column=1, padx=5, pady=2, sticky="w")

        ttk.Label(cfg, text="Confirmar:").grid(
            row=1, column=0, padx=5, pady=2, sticky="w"
        )
        self.password_confirm_entry = ttk.Entry(cfg, show="*", width=25)
        self.password_confirm_entry.grid(row=1, column=1, padx=5, pady=2, sticky="w")

        self.use_password_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            cfg,
            text="Utilitzar contrasenya",
            variable=self.use_password_var,
            command=self.toggle_password_fields,
        ).grid(row=2, column=0, columnspan=2, pady=2, sticky="w")

        ttk.Label(cfg, text="Conversions paral·leles:").grid(
            row=3, column=0, padx=5, pady=2, sticky="w"
        )
        self.parallel_spinbox = ttk.Spinbox(
            cfg, from_=1, to=MAX_PARALLEL_ALLOWED, width=23
        )
        self.parallel_spinbox.set(DEFAULT_PARALLEL)
        self.parallel_spinbox.grid(row=3, column=1, padx=5, pady=2, sticky="w")

    def toggle_password_fields(self) -> None:
        if self.use_password_var.get():
            self.password_entry.grid()
            self.password_confirm_entry.grid()
        else:
            self.password_entry.grid_remove()
            self.password_confirm_entry.grid_remove()

    def _setup_destination(self, main: ttk.Frame) -> None:
        dest = ttk.Frame(main)
        dest.grid(row=4, column=0, columnspan=3, sticky="ew", pady=6)
        dest.columnconfigure(1, weight=1)

        ttk.Label(dest, text="Desar ZIP a:").grid(row=0, column=0, padx=5, sticky="w")
        self.dest_entry = ttk.Entry(dest, width=60)
        self.dest_entry.grid(row=0, column=1, padx=5, sticky="ew")
        self.dest_entry.insert(
            0, str(Path.home() / f"documents_{time.strftime('%Y%m%d_%H%M%S')}.zip")
        )

        ttk.Button(dest, text="📂 Triar", command=self.select_destination).grid(
            row=0, column=2, padx=5
        )

    def _setup_actions(self, main: ttk.Frame) -> None:
        actions = ttk.Frame(main)
        actions.grid(row=5, column=0, columnspan=3, pady=10)

        self.convert_btn = ttk.Button(
            actions,
            text="🚀 Convertir i Crear ZIP",
            command=self.start_conversion,
            width=30,
        )
        self.convert_btn.grid(row=0, column=0, padx=5)

        self.cancel_btn = ttk.Button(
            actions,
            text="❌ Cancel·lar",
            command=self.cancel_conversion,
            width=15,
            state="disabled",
        )
        self.cancel_btn.grid(row=0, column=1, padx=5)

        ttk.Button(actions, text="🗑️ Netejar", command=self.clear_all, width=15).grid(
            row=0, column=2, padx=5
        )

    def _setup_progress_and_status(self, main: ttk.Frame) -> None:
        self.progress = ttk.Progressbar(main, mode="determinate", maximum=100)
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=6)

        self.status_label = ttk.Label(
            main, text="Llest per començar", foreground="blue"
        )
        self.status_label.grid(row=7, column=0, columnspan=3, pady=4)

        self.active_label = ttk.Label(
            main, text="Conversions completades: 0 / 0", foreground="gray"
        )
        self.active_label.grid(row=8, column=0, columnspan=3, pady=2)

    def select_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Selecciona fitxers DOCX",
            filetypes=[("Word Documents", "*.docx"), ("Tots els fitxers", "*.*")],
        )
        if not files:
            return
        self.docx_files = list(files)
        self.files_listbox.delete(0, tk.END)
        for path in self.docx_files:
            self.files_listbox.insert(tk.END, Path(path).name)
        self.files_label.config(text=f"{len(self.docx_files)} fitxer(s) seleccionat(s)")
        self.update_status("Fitxers carregats correctament", "success")

    def select_destination(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Desa el fitxer ZIP",
            defaultextension=".zip",
            filetypes=[("ZIP", "*.zip"), ("Tots", "*.*")],
            initialfile=f"documents_{time.strftime('%Y%m%d_%H%M%S')}.zip",
        )
        if file_path:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, file_path)

    def clear_all(self) -> None:
        if self.is_processing:
            messagebox.showwarning(
                "Processant", "No es pot netejar durant la conversió."
            )
            return
        self.docx_files.clear()
        self.files_listbox.delete(0, tk.END)
        self.files_label.config(text="Cap fitxer seleccionat")
        self.password_entry.delete(0, tk.END)
        self.password_confirm_entry.delete(0, tk.END)
        self.progress["value"] = 0
        self.update_status("Llest per començar", "info")
        self.active_label.config(text="Conversions completades: 0 / 0")

    def update_status(self, message: str, status_type: str = "info") -> None:
        self.status_label.config(
            text=message, foreground=STATUS_COLORS.get(status_type, "blue")
        )

    def update_progress(self, value: int) -> None:
        self.progress["value"] = value
        self.root.update_idletasks()

    def update_active_conversions(self, completed: int, total: int) -> None:
        self.active_label.config(text=f"Conversions completades: {completed} / {total}")

    def process_queue(self) -> None:
        try:
            while True:
                msg_type, data = self.message_queue.get_nowait()
                if msg_type == "status":
                    self.update_status(
                        data.get("message", ""), data.get("type", "info")
                    )
                elif msg_type == "progress":
                    self.update_progress(data)
                elif msg_type == "active":
                    self.update_active_conversions(data["completed"], data["total"])
                elif msg_type == "finished":
                    self.conversion_finished(data)
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)

    def start_conversion(self) -> None:
        if not self.docx_files:
            messagebox.showwarning("Atenció", "Selecciona almenys un fitxer DOCX.")
            return

        pwd = ""
        if self.use_password_var.get():
            pwd = self.password_entry.get()
            pwd2 = self.password_confirm_entry.get()
            if not pwd:
                messagebox.showwarning("Atenció", "Introdueix una contrasenya.")
                return
            if pwd != pwd2:
                messagebox.showerror("Error", "Les contrasenyes no coincideixen.")
                return

        zip_path = self.dest_entry.get().strip()
        if not zip_path:
            messagebox.showwarning("Atenció", "Selecciona un destí per al ZIP.")
            return

        try:
            max_workers = int(self.parallel_spinbox.get())
            if not (1 <= max_workers <= MAX_PARALLEL_ALLOWED):
                raise ValueError
        except ValueError:
            messagebox.showwarning("Atenció", "Número de fils entre 1 i 16.")
            return

        # Reinicia informe i estat
        self.conversion_report = ConversionReport()
        self.conversion_report.total_files = len(self.docx_files)
        self.conversion_report.start_time = time.time()

        self.is_processing = True
        self.cancel_flag = False
        self.completed_conversions = 0
        self.convert_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.progress["value"] = 0

        self.conversion_thread = threading.Thread(
            target=self.convert_and_zip_thread,
            args=(self.docx_files.copy(), pwd, zip_path, max_workers),
            daemon=True,
        )
        self.conversion_thread.start()

    def cancel_conversion(self) -> None:
        if messagebox.askyesno("Confirmar", "Vols cancel·lar la conversió?"):
            self.cancel_flag = True
            self.message_queue.put(
                ("status", {"message": "Cancel·lant...", "type": "warning"})
            )

    def convert_single_file(
        self, docx_file: str, temp_dir: str, file_num: int, total_files: int
    ) -> Tuple[Optional[Path], int, Optional[str]]:
        converter = DocxConverter(
            max_retries=MAX_RETRIES, timeout=CONVERSION_TIMEOUT, has_win32=HAS_WIN32
        )
        return converter.convert_single_file(docx_file, temp_dir, file_num, total_files)

    def convert_and_zip_thread(
        self, docx_files: List[str], password: str, zip_path: str, max_workers: int
    ) -> None:
        temp_dir = None
        try:
            temp_dir = tempfile.mkdtemp()
            total = len(docx_files)

            self.message_queue.put(
                (
                    "status",
                    {
                        "message": f"Convertint {total} fitxer(s) amb {max_workers} fils...",
                        "type": "info",
                    },
                )
            )

            pdf_files: List[Path] = []

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_file = {
                    executor.submit(self.convert_single_file, f, temp_dir, i, total): f
                    for i, f in enumerate(docx_files, 1)
                }

                for fut in as_completed(future_to_file):
                    if self.cancel_flag:
                        for other in future_to_file:
                            other.cancel()
                        raise InterruptedError("Conversió cancel·lada per l'usuari")

                    src = future_to_file[fut]
                    name = Path(src).name

                    try:
                        pdf_path, attempts, error = fut.result()
                        if pdf_path:
                            pdf_files.append(pdf_path)
                            self.conversion_report.add_success(name)
                            if attempts > 1:
                                self.conversion_report.add_retry(name, attempts)
                            self.message_queue.put(
                                (
                                    "status",
                                    {
                                        "message": f"✓ Convertit: {name}"
                                        + (
                                            f" ({attempts} intents)"
                                            if attempts > 1
                                            else ""
                                        ),
                                        "type": "success",
                                    },
                                )
                            )
                        else:
                            self.conversion_report.add_failure(
                                name, error or "Error desconegut"
                            )
                            self.message_queue.put(
                                (
                                    "status",
                                    {
                                        "message": f"✗ Error final: {name}",
                                        "type": "error",
                                    },
                                )
                            )
                    except Exception as exc:
                        self.conversion_report.add_failure(name, str(exc))
                        self.message_queue.put(
                            (
                                "status",
                                {
                                    "message": f"✗ Excepció: {name} - {str(exc)}",
                                    "type": "error",
                                },
                            )
                        )

                    self.completed_conversions += 1
                    progress = (self.completed_conversions / total) * 80
                    self.message_queue.put(("progress", int(progress)))
                    self.message_queue.put(
                        (
                            "active",
                            {"completed": self.completed_conversions, "total": total},
                        )
                    )

                    if (
                        self.completed_conversions % 3 == 0
                        or self._check_memory_usage()
                    ):
                        gc.collect()

            if not pdf_files:
                raise RuntimeError("No s'ha pogut convertir cap fitxer.")

            if self.cancel_flag:
                raise InterruptedError("Conversió cancel·lada per l'usuari")

            self.message_queue.put(
                (
                    "status",
                    {
                        "message": f"Creant ZIP amb {len(pdf_files)} PDF(s)...",
                        "type": "info",
                    },
                )
            )
            self.message_queue.put(("progress", 85))
            gc.collect()

            with pyzipper.AESZipFile(
                zip_path,
                "w",
                compression=pyzipper.ZIP_DEFLATED,
                encryption=pyzipper.WZ_AES if password else None,
            ) as zipf:
                if password:
                    zipf.setpassword(password.encode("utf-8"))
                for idx, pdf in enumerate(pdf_files, 1):
                    if self.cancel_flag:
                        raise InterruptedError("Conversió cancel·lada per l'usuari")
                    zipf.write(pdf, pdf.name)
                    progress = 85 + (idx / len(pdf_files)) * 15
                    self.message_queue.put(("progress", int(progress)))

            self.message_queue.put(("progress", 100))
            self.conversion_report.end_time = time.time()
            self.message_queue.put(
                ("finished", {"success": True, "zip_path": zip_path})
            )

        except InterruptedError as exc:
            self.conversion_report.end_time = time.time()
            self.message_queue.put(("finished", {"success": False, "error": str(exc)}))
        except Exception as exc:
            self.conversion_report.end_time = time.time()
            self.message_queue.put(
                (
                    "finished",
                    {"success": False, "error": f"Error durant la conversió: {exc}"},
                )
            )
        finally:
            if temp_dir and Path(temp_dir).exists():
                shutil.rmtree(temp_dir)
            gc.collect()

    def _check_memory_usage(self) -> bool:
        if psutil is None:
            return False
        try:
            proc = psutil.Process(os.getpid())
            mb = proc.memory_info().rss / 1024 / 1024
            return mb > MEMORY_THRESHOLD_MB
        except Exception:
            return False

    def conversion_finished(self, data: Dict[str, object]) -> None:
        self.is_processing = False
        self.convert_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")

        if data.get("success"):
            summary = self.conversion_report.get_summary()
            zip_path = str(data.get("zip_path", ""))

            if summary["failed"] == 0:
                self.update_status(
                    f"✅ Tots els {int(summary['total'])} fitxers convertits!",
                    "success",
                )
            else:
                self.update_status(
                    f"⚠️ {int(summary['success'])}/{int(summary['total'])} convertits",
                    "warning",
                )

            report_text = self.conversion_report.generate_detailed_report()
            ReportDialog(self.root, report_text, zip_path)
        else:
            error_msg = str(data.get("error", "Error desconegut"))
            self.update_status(f"Error: {error_msg}", "error")
            if "cancel·lada" not in error_msg.lower():
                messagebox.showerror(
                    "Error", f"Error durant la conversió:\n{error_msg}"
                )


def main() -> None:
    root = tk.Tk()
    _app = DocxToPdfZipApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
