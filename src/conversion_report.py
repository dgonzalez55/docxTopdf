from typing import Dict, List, Optional, Tuple


class ConversionReport:
    """Gestió d'estadístiques i informe final de conversió."""

    def __init__(self) -> None:
        self.total_files: int = 0
        self.successful_conversions: List[str] = []
        self.failed_conversions: List[Tuple[str, str]] = []
        self.retried_files: List[Tuple[str, int]] = []
        self.start_time: Optional[float] = None
        self.end_time: Optional[float] = None

    def add_success(self, filename: str) -> None:
        self.successful_conversions.append(filename)

    def add_failure(self, filename: str, error: str) -> None:
        self.failed_conversions.append((filename, error))

    def add_retry(self, filename: str, attempts: int) -> None:
        self.retried_files.append((filename, attempts))

    def get_summary(self) -> Dict[str, float]:
        duration = 0.0
        if self.start_time and self.end_time:
            duration = max(0.0, self.end_time - self.start_time)
        return {
            "total": float(self.total_files),
            "success": float(len(self.successful_conversions)),
            "failed": float(len(self.failed_conversions)),
            "retried": float(len(self.retried_files)),
            "duration": duration,
        }

    def generate_detailed_report(self) -> str:
        summary = self.get_summary()
        total = int(summary["total"]) or 1  # evitar divisió per zero
        success = int(summary["success"])
        failed = int(summary["failed"])
        retried = int(summary["retried"])
        duration = summary["duration"]
        duration_min = int(duration // 60)
        duration_sec = int(duration % 60)

        success_pct = (success / total) * 100
        failed_pct = (failed / total) * 100

        lines: List[str] = []
        sep = "=" * 72
        sub = "-" * 72

        lines.append(sep)
        lines.append("INFORME DE CONVERSIÓ DOCX → PDF")
        lines.append(sep)
        lines.append("")
        lines.append("📊 RESUM GENERAL")
        lines.append(sub)
        lines.append(f"{'Total de fitxers processats:':<35} {total}")
        lines.append(f"{'✓ Conversions exitoses:':<35}{success} ({success_pct:.1f}%)")
        lines.append(f"{'✗ Conversions fallides:':<35}{failed} ({failed_pct:.1f}%)")
        lines.append(f"{'🔄 Fitxers amb reintents:':<35}{retried}")
        lines.append(f"{'⏱️  Temps total:':<35}{duration_min}m {duration_sec}s")
        lines.append("")

        if self.retried_files:
            lines.append("🔄 FITXERS RESOLTS DESPRÉS DE REINTENTS")
            lines.append(sub)
            for filename, attempts in self.retried_files:
                lines.append(f"  • {filename}")
                lines.append(f"    └─ Resolt després de {attempts} intent(s)")
            lines.append("")

        if success and not self.retried_files:
            lines.append("✓ FITXERS CONVERTITS SENSE REINTENTS")
            lines.append(sub)
            for filename in self.successful_conversions[:10]:
                lines.append(f"  ✓ {filename}")
            if len(self.successful_conversions) > 10:
                rest = len(self.successful_conversions) - 10
                lines.append(f"  ... i {rest} més")
            lines.append("")

        if self.failed_conversions:
            lines.append("✗ FITXERS AMB ERRORS NO RESOLTS")
            lines.append(sub)
            for filename, error in self.failed_conversions:
                lines.append(f"  ✗ {filename}")
                lines.append(f"    └─ Error: {error}")
                lines.append("    └─ Suggeriments:")
                lines.append("       • Verifica si el DOCX és corrupte")
                lines.append("       • Comprova que MS Word estigui instal·lat")
                lines.append("       • Obre i desa de nou el DOCX")
            lines.append("")

        lines.append(sep)
        if failed == 0:
            lines.append("✅ CONVERSIÓ COMPLETADA AMB ÈXIT TOTAL")
        elif success > 0:
            lines.append("⚠️  CONVERSIÓ COMPLETADA AMB ALGUNS ERRORS")
        else:
            lines.append("❌ NO S'HA POGUT CONVERTIR CAP FITXER")
        lines.append(sep)

        return "\n".join(lines)
