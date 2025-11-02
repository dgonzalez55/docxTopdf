"""
Gestió de la conversió de fitxers DOCX a PDF.

Millores:
- Execució paral·lela configurable (1-16)
- Reintents agressius + mètode alternatiu via win32com (si disponible)
Requisits: pip install docx2pdf
"""

from __future__ import annotations

import gc
import importlib.util
import os
import time
from concurrent.futures import (
    ThreadPoolExecutor,
    TimeoutError as FuturesTimeoutError,
)
from pathlib import Path
from typing import Optional, Tuple

# Dependències opcionals
try:
    from docx2pdf import convert  # type: ignore
except Exception:
    convert = None

try:
    import psutil  # type: ignore
except Exception:
    psutil = None

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False


class DocxConverter:
    """Gestió de la conversió de fitxers DOCX a PDF."""

    MEMORY_THRESHOLD_MB = 500

    def __init__(self, max_retries: int, timeout: int, has_win32: bool) -> None:
        self.max_retries = max_retries
        self.timeout = timeout
        self.has_win32 = has_win32

    def convert_with_win32(self, docx_path: Path, pdf_path: Path) -> bool:
        """Mètode alternatiu via COM (Word)."""
        if not self.has_win32:
            return False
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path))
            # 17 = wdFormatPDF
            doc.SaveAs(str(pdf_path), FileFormat=17)
            doc.Close()
            word.Quit()
            pythoncom.CoUninitialize()
            return pdf_path.exists() and pdf_path.stat().st_size > 0
        except Exception:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            return False

    def convert_single_file(
        self, docx_file: str, temp_dir: str, file_num: int, total_files: int
    ) -> Tuple[Optional[Path], int, Optional[str]]:
        """Converteix un DOCX amb reintents i timeout."""
        docx_path = Path(docx_file)
        pdf_path = Path(temp_dir) / (docx_path.stem + ".pdf")
        last_error: Optional[str] = None

        # Comprovar si el fitxer DOCX existeix
        if not docx_path.exists() or not docx_path.is_file():
            print(f"❌ El fitxer {docx_path} no existeix o no és vàlid.")
            return None, 0, f"El fitxer {docx_path} no existeix o no és vàlid."

        # Comprovar si el directori temporal existeix
        if not Path(temp_dir).exists():
            print(f"❌ El directori temporal {temp_dir} no existeix.")
            return None, 0, f"El directori temporal {temp_dir} no existeix."

        for attempt in range(1, self.max_retries + 1):
            # Netejar residus d'intents previs
            if pdf_path.exists():
                try:
                    pdf_path.unlink()
                except Exception as e:
                    print(f"⚠️ No s'ha pogut eliminar un PDF residual: {e}")

            try:
                # Mètode principal: docx2pdf amb timeout
                if convert is None:
                    raise RuntimeError("El mòdul docx2pdf no està disponible.")

                print(
                    f"🔄 Intentant convertir {docx_path.name} amb docx2pdf "
                    f"(Intent {attempt})..."
                )
                with ThreadPoolExecutor(max_workers=1) as ex:
                    fut = ex.submit(convert, str(docx_path), str(pdf_path))
                    fut.result(timeout=self.timeout)

                if pdf_path.exists() and pdf_path.stat().st_size > 0:
                    print(f"✅ Conversió completada: {docx_path.name}")
                    return pdf_path, attempt, None

                raise ValueError("PDF buit o inexistent després de docx2pdf")

            except (FuturesTimeoutError, ValueError) as exc:
                last_error = str(exc)
                print(f"⚠️ Error durant la conversió amb docx2pdf: {last_error}")

                # Provar mètode alternatiu si disponible
                if self.has_win32 and attempt < self.max_retries:
                    print(
                        f"🔄 Intentant convertir {docx_path.name} amb win32com "
                        f"(Intent {attempt})..."
                    )
                    if self.convert_with_win32(docx_path, pdf_path):
                        print(f"✅ Conversió completada amb win32com: {docx_path.name}")
                        return pdf_path, attempt, None

            except Exception as exc:  # Altres errors
                last_error = str(exc)
                print(f"⚠️ Error inesperat durant la conversió: {last_error}")

            # Informar i esperar abans del següent intent
            print(f"⚠️ {docx_path.name}: intent {attempt}/{self.max_retries} fallit.")
            if attempt < self.max_retries:
                time.sleep(min(attempt * 3, 15))
                gc.collect()

        print(
            f"❌ Conversió fallida per {docx_path.name} després de "
            f"{self.max_retries} intents."
        )
        return (
            None,
            self.max_retries,
            f"Fallo després de {self.max_retries} intents: {last_error}",
        )

    def _check_memory_usage(self) -> bool:
        if psutil is None:
            return False
        try:
            proc = psutil.Process(os.getpid())
            mb = proc.memory_info().rss / 1024 / 1024
            return mb > self.MEMORY_THRESHOLD_MB
        except Exception:
            return False

    @staticmethod
    def is_docx2pdf_available() -> bool:
        return importlib.util.find_spec("docx2pdf") is not None

    @staticmethod
    def is_pyzipper_available() -> bool:
        return importlib.util.find_spec("pyzipper") is not None
