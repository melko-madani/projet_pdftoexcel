"""Extraction ciblée de colonnes spécifiques depuis les tableaux PDF.

Extrait les colonnes : Référence de l'avis, Adresse, Montant de dégrèvement
et les exporte dans un fichier JSON (03.json).
"""

import json
import logging
import re
import tempfile
from pathlib import Path

from .parser import normalize_header, process_tables
from .scanner import scan_pdf

logger = logging.getLogger(__name__)

# Colonnes cibles à extraire
TARGET_COLUMNS = [
    "Référence de l'avis",
    "Adresse",
    "Montant de dégrèvement",
]

# Variantes possibles pour matcher les headers du PDF
_COLUMN_ALIASES = {
    "référence de l'avis": "Référence de l'avis",
    "reference de l'avis": "Référence de l'avis",
    "référence de l avis": "Référence de l'avis",
    "reference de l avis": "Référence de l'avis",
    "référence avis": "Référence de l'avis",
    "reference avis": "Référence de l'avis",
    "réf. avis": "Référence de l'avis",
    "ref. avis": "Référence de l'avis",
    "réf avis": "Référence de l'avis",
    "ref avis": "Référence de l'avis",
    "n° avis": "Référence de l'avis",
    "adresse": "Adresse",
    "adresse du bien": "Adresse",
    "adresse du local": "Adresse",
    "adresse immeuble": "Adresse",
    "montant de dégrèvement": "Montant de dégrèvement",
    "montant de degrevement": "Montant de dégrèvement",
    "montant dégrèvement": "Montant de dégrèvement",
    "montant degrevement": "Montant de dégrèvement",
    "dégrèvement": "Montant de dégrèvement",
    "degrevement": "Montant de dégrèvement",
    "montant du dégrèvement": "Montant de dégrèvement",
    "montant du degrevement": "Montant de dégrèvement",
}


def _match_column(header: str) -> str | None:
    """Tente de matcher un header PDF avec une colonne cible.

    Returns:
        Le nom canonique de la colonne cible, ou None si pas de match.
    """
    normalized = normalize_header(header)

    # Match exact via alias
    if normalized in _COLUMN_ALIASES:
        return _COLUMN_ALIASES[normalized]

    # Match partiel : vérifier si le header contient un alias
    for alias, canonical in _COLUMN_ALIASES.items():
        if alias in normalized or normalized in alias:
            return canonical

    return None


def _find_column_indices(headers: list[str]) -> dict[str, int]:
    """Trouve les indices des colonnes cibles dans les headers du tableau.

    Returns:
        Dict {nom_canonique: index} pour les colonnes trouvées.
    """
    found = {}
    for idx, header in enumerate(headers):
        match = _match_column(header)
        if match and match not in found:
            found[match] = idx
    return found


def extract_target_columns_from_pdf(pdf_bytes: bytes, filename: str = "") -> list[dict]:
    """Extrait les 3 colonnes cibles depuis un PDF.

    Args:
        pdf_bytes: Contenu du fichier PDF.
        filename: Nom du fichier (pour le logging).

    Returns:
        Liste de dicts avec les colonnes trouvées pour chaque ligne.
    """
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    try:
        tmp.write(pdf_bytes)
        tmp.close()
        tmp_path = Path(tmp.name)

        scan_result = scan_pdf(tmp_path, min_cols=3)

        if not scan_result.tables:
            logger.info("Aucun tableau trouvé dans %s", filename or "le PDF")
            return []

        datasets = process_tables(scan_result.tables)

        all_rows = []
        for ds in datasets:
            col_map = _find_column_indices(ds.headers)

            if not col_map:
                logger.debug(
                    "Dataset '%s' : aucune colonne cible trouvée (headers: %s)",
                    ds.name,
                    ds.headers,
                )
                continue

            logger.info(
                "Dataset '%s' : colonnes trouvées = %s",
                ds.name,
                list(col_map.keys()),
            )

            for row in ds.data_rows:
                entry = {}
                for col_name, col_idx in col_map.items():
                    value = row[col_idx] if col_idx < len(row) else None
                    entry[col_name] = value if value is not None else ""
                # Ne garder que les lignes qui ont au moins une valeur non vide
                if any(str(v).strip() for v in entry.values()):
                    all_rows.append(entry)

        return all_rows

    except Exception as e:
        logger.error("Erreur extraction colonnes depuis %s : %s", filename, e)
        return []
    finally:
        Path(tmp.name).unlink(missing_ok=True)


def extract_and_save_json(
    pdf_files: list[tuple[str, bytes]],
    output_path: str | Path = "03.json",
) -> dict:
    """Extrait les colonnes cibles de plusieurs PDFs et sauvegarde en JSON.

    Args:
        pdf_files: Liste de tuples (filename, pdf_bytes).
        output_path: Chemin du fichier JSON de sortie.

    Returns:
        Le dict complet exporté en JSON.
    """
    output_path = Path(output_path)
    result = {
        "colonnes_extraites": TARGET_COLUMNS,
        "nombre_fichiers": len(pdf_files),
        "fichiers": [],
    }

    total_rows = 0

    for filename, pdf_bytes in pdf_files:
        logger.info("Extraction de %s...", filename)
        rows = extract_target_columns_from_pdf(pdf_bytes, filename)

        fichier_entry = {
            "nom_fichier": filename,
            "nombre_lignes": len(rows),
            "donnees": rows,
        }
        result["fichiers"].append(fichier_entry)
        total_rows += len(rows)

    result["nombre_total_lignes"] = total_rows

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    logger.info(
        "Export JSON terminé : %d lignes dans %s", total_rows, output_path
    )

    return result
