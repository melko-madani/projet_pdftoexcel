"""Détection des tableaux dans les fichiers PDF."""

import logging
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber

logger = logging.getLogger(__name__)


@dataclass
class TableInfo:
    """Informations sur un tableau détecté dans une page PDF."""
    page_num: int
    rows: list[list[str]]
    headers: list[str]
    col_count: int

    @property
    def row_count(self) -> int:
        return len(self.rows)


@dataclass
class ScanResult:
    """Résultat complet du scan d'un PDF."""
    pdf_path: Path
    total_pages: int
    tables: list[TableInfo] = field(default_factory=list)
    pages_without_tables: list[int] = field(default_factory=list)


def scan_pdf(pdf_path: Path, min_cols: int = 8, min_rows: int = 1) -> ScanResult:
    """Scanne un PDF et détecte tous les tableaux.

    Args:
        pdf_path: Chemin vers le fichier PDF.
        min_cols: Nombre minimum de colonnes pour considérer un tableau valide.
        min_rows: Nombre minimum de lignes de données (hors header) pour un tableau valide.

    Returns:
        ScanResult avec tous les tableaux détectés.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"Fichier PDF introuvable : {pdf_path}")
    if not pdf_path.suffix.lower() == ".pdf":
        raise ValueError(f"Le fichier n'est pas un PDF : {pdf_path}")

    result = ScanResult(pdf_path=pdf_path, total_pages=0)

    logger.info("Ouverture de %s...", pdf_path.name)

    with pdfplumber.open(pdf_path) as pdf:
        result.total_pages = len(pdf.pages)
        logger.info("PDF : %d page(s) detectee(s)", result.total_pages)

        for page in pdf.pages:
            page_num = page.page_number
            tables = page.extract_tables()

            if not tables:
                result.pages_without_tables.append(page_num)
                logger.info("Page %d : aucun tableau", page_num)
                continue

            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    logger.debug(
                        "Page %d, tableau %d : ignoré (moins de 2 lignes)",
                        page_num, table_idx + 1,
                    )
                    continue

                # Première ligne = headers
                raw_headers = table[0]
                headers = [
                    str(h).strip().replace("\n", " ") if h else ""
                    for h in raw_headers
                ]
                rows = table[1:]
                col_count = len(headers)

                # Filtrage : faux tableaux (en-tête courrier)
                first_cells = " ".join(
                    str(c).strip() for c in raw_headers[:3] if c
                ).upper()
                if any(kw in first_cells for kw in ["MELKO", "ENERGIE", "POUR LE COMPTE"]):
                    logger.debug(
                        "Page %d, tableau %d : ignoré (faux tableau en-tête courrier)",
                        page_num, table_idx + 1,
                    )
                    continue

                # Filtrage : min colonnes
                if col_count < min_cols:
                    logger.debug(
                        "Page %d, tableau %d : ignoré (%d colonnes < %d min)",
                        page_num, table_idx + 1, col_count, min_cols,
                    )
                    continue

                # Filtrage : min lignes de données
                if len(rows) < min_rows:
                    logger.debug(
                        "Page %d, tableau %d : ignoré (%d lignes < %d min)",
                        page_num, table_idx + 1, len(rows), min_rows,
                    )
                    continue

                table_info = TableInfo(
                    page_num=page_num,
                    rows=rows,
                    headers=headers,
                    col_count=col_count,
                )
                result.tables.append(table_info)

                logger.info(
                    "Page %d : tableau %d — %d lignes, %d colonnes",
                    page_num, table_idx + 1, len(rows), col_count,
                )

    return result


def format_scan_report(result: ScanResult) -> str:
    """Génère un rapport textuel du scan."""
    lines = []
    lines.append(f"{'='*60}")
    lines.append(f"Rapport de scan : {result.pdf_path.name}")
    lines.append(f"{'='*60}")
    lines.append(f"Pages totales : {result.total_pages}")
    lines.append(f"Tableaux trouvés : {len(result.tables)}")
    lines.append("")

    if result.tables:
        lines.append("Tableaux détectés :")
        lines.append(f"{'-'*60}")
        for t in result.tables:
            header_preview = " | ".join(t.headers[:3])
            if len(t.headers) > 3:
                header_preview += " | ..."
            lines.append(
                f"  Page {t.page_num} : {t.row_count} lignes, "
                f"{t.col_count} colonnes"
            )
            lines.append(f"    Headers : [{header_preview}]")
        lines.append("")

    if result.pages_without_tables:
        pages_str = ", ".join(str(p) for p in result.pages_without_tables)
        lines.append(f"Pages sans tableau : {pages_str}")
    else:
        lines.append("Toutes les pages contiennent au moins un tableau.")

    lines.append(f"{'='*60}")
    return "\n".join(lines)
