import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


class ReportDialog(tk.Toplevel):
    """Diàleg per mostrar i desar l'informe final."""

    def __init__(self, parent: tk.Tk, report_text: str, zip_path: str) -> None:
        super().__init__(parent)
        self.title("Informe de Conversió")
        self.geometry("800x600")
        self.resizable(True, True)

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.text_area = scrolledtext.ScrolledText(
            text_frame, wrap=tk.WORD, width=90, height=30, font=("Courier New", 9)
        )
        self.text_area.pack(fill=tk.BOTH, expand=True)
        self.text_area.insert("1.0", report_text)
        self.text_area.config(state="disabled")

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(
            button_frame,
            text="📂 Obrir ubicació del ZIP",
            command=lambda: self.open_zip_location(zip_path),
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="💾 Desar informe",
            command=lambda: self.save_report(report_text),
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="✓ Tancar", command=self.destroy).pack(
            side=tk.RIGHT, padx=5
        )

        self.transient(parent)
        self.grab_set()

    def open_zip_location(self, zip_path: str) -> None:
        """Obre el directori on s'ha desat el ZIP."""
        try:
            folder = os.path.dirname(zip_path) or "."
            os.startfile(folder)  # Windows
        except Exception as exc:
            messagebox.showerror("Error", f"No es pot obrir la carpeta: {exc}")

    def save_report(self, report_text: str) -> None:
        """Desa l'informe en un fitxer .txt."""
        file_path = filedialog.asksaveasfilename(
            title="Desar informe",
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("Tots", "*.*")],
            initialfile=f"informe_conversio_{time.strftime('%Y%m%d_%H%M%S')}.txt",
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as fh:
                fh.write(report_text)
            messagebox.showinfo("Èxit", "Informe desat correctament.")
        except Exception as exc:
            messagebox.showerror("Error", f"No s'ha pogut desar l'informe: {exc}")
