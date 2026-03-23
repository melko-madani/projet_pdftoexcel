"""Extraction, nettoyage et consolidation des données tabulaires."""

import logging
from dataclasses import dataclass, field

from .scanner import TableInfo
from .utils import (
    clean_cell,
    detect_column_type,
    is_empty_row,
    is_total_row,
    parse_date,
    parse_euro,
    parse_mois_vacance,
)

logger = logging.getLogger(__name__)


@dataclass
class Dataset:
    """Un jeu de données prêt à être exporté en Excel."""
    name: str
    headers: list[str]
    data_rows: list[list]
    total_rows: list[list] = field(default_factory=list)
    column_types: list[str] = field(default_factory=list)
    source_pages: list[int] = field(default_factory=list)


def normalize_header(header: str) -> str:
    """Normalise un header individuel pour la comparaison."""
    import re
    text = clean_cell(header).lower().strip()
    # Supprimer retours à la ligne, espaces multiples
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_headers(headers: list[str]) -> list[str]:
    """Normalise les headers pour la comparaison."""
    return [normalize_header(h) for h in headers]


def headers_match(headers_a: list[str], headers_b: list[str]) -> bool:
    """Vérifie si deux listes de headers représentent la même structure.

    Comparaison tolérante : compare les 5 premiers headers,
    considère un match si 80%+ des headers comparés sont identiques.
    """
    norm_a = normalize_headers(headers_a)
    norm_b = normalize_headers(headers_b)

    # Comparer sur les 5 premiers headers
    compare_count = min(5, len(norm_a), len(norm_b))
    if compare_count == 0:
        return False

    matches = sum(
        1 for i in range(compare_count) if norm_a[i] == norm_b[i]
    )
    match_ratio = matches / compare_count
    return match_ratio >= 0.8


def is_header_row(row: list[str], reference_headers: list[str]) -> bool:
    """Vérifie si une ligne est une répétition du header.

    Utilise la même logique tolérante que headers_match.
    """
    cleaned = [clean_cell(cell) for cell in row]
    return headers_match(cleaned, reference_headers)


def group_tables(tables: list[TableInfo]) -> list[list[TableInfo]]:
    """Regroupe les tableaux consécutifs ayant la même structure.

    Deux tableaux sont dans le même groupe s'ils ont les mêmes headers
    (même nombre et mêmes noms de colonnes, insensible à la casse).
    """
    if not tables:
        return []

    groups: list[list[TableInfo]] = []
    current_group: list[TableInfo] = [tables[0]]

    for table in tables[1:]:
        if headers_match(current_group[0].headers, table.headers):
            current_group.append(table)
        else:
            groups.append(current_group)
            current_group = [table]

    groups.append(current_group)
    return groups


def clean_value(value: str, col_type: str):
    """Nettoie et convertit une valeur selon le type de colonne."""
    cleaned = clean_cell(value)
    if not cleaned:
        return None

    if col_type == "euro":
        result = parse_euro(cleaned)
        return result if result is not None else cleaned
    elif col_type == "date":
        result = parse_date(cleaned)
        return result if result is not None else cleaned
    elif col_type == "mois_vacance":
        result = parse_mois_vacance(cleaned)
        return result if result is not None else cleaned
    else:
        return cleaned


def _is_annexe_fiscale(headers: list[str]) -> bool:
    """Détecte si un tableau est une annexe fiscale (contient programme + référence + avis)."""
    normalized = [normalize_header(h) for h in headers]
    joined = " ".join(normalized)
    has_programme = "programme" in joined
    has_reference_avis = ("référence" in joined or "reference" in joined) and "avis" in joined
    return has_programme and has_reference_avis


def process_tables(tables: list[TableInfo]) -> list[Dataset]:
    """Traite tous les tableaux détectés et retourne des Datasets prêts à l'export.

    - Regroupe les tableaux consécutifs de même structure
    - Déduplique les headers répétés
    - Nettoie et convertit les données
    - Sépare les lignes de total
    """
    if not tables:
        logger.warning("Aucun tableau à traiter.")
        return []

    groups = group_tables(tables)
    datasets = []

    for group_idx, group in enumerate(groups):
        reference_headers = group[0].headers
        headers = [clean_cell(h) for h in reference_headers]
        source_pages = sorted(set(t.page_num for t in group))

        # Nom du dataset
        is_annexe = _is_annexe_fiscale(headers)

        if is_annexe:
            # Compter combien de groupes sont des annexes pour numéroter si besoin
            annexe_count = sum(
                1 for g in groups[:group_idx + 1]
                if _is_annexe_fiscale([clean_cell(h) for h in g[0].headers])
            )
            total_annexes = sum(
                1 for g in groups
                if _is_annexe_fiscale([clean_cell(h) for h in g[0].headers])
            )
            if total_annexes == 1:
                name = "Annexe"
            else:
                name = f"Annexe {annexe_count}"
        elif len(groups) == 1:
            name = "Données consolidées"
        else:
            pages_str = ", ".join(str(p) for p in source_pages)
            name = f"Tableau_Pages_{pages_str}"

        # Collecter toutes les lignes brutes (en dédupliquant les headers)
        all_raw_rows = []
        for table in group:
            for row in table.rows:
                # Ignorer les lignes vides
                if is_empty_row(row):
                    continue
                # Dédupliquer les headers répétés
                if is_header_row(row, reference_headers):
                    logger.debug(
                        "Header dupliqué supprimé (page %d)", table.page_num
                    )
                    continue
                # Normaliser la longueur de la ligne
                normalized = list(row)
                while len(normalized) < len(headers):
                    normalized.append("")
                if len(normalized) > len(headers):
                    normalized = normalized[:len(headers)]
                all_raw_rows.append(normalized)

        # Détecter les types de colonnes à partir des valeurs
        column_types = []
        for col_idx, header in enumerate(headers):
            sample_values = [
                clean_cell(row[col_idx])
                for row in all_raw_rows[:20]
                if col_idx < len(row) and not is_total_row(row)
            ]
            col_type = detect_column_type(header, sample_values)
            column_types.append(col_type)

        logger.info(
            "Dataset '%s' : %d colonnes, %d lignes brutes",
            name, len(headers), len(all_raw_rows),
        )

        # Séparer les lignes de données et les lignes de total
        data_rows = []
        total_rows = []

        for row in all_raw_rows:
            cleaned_row = [
                clean_value(clean_cell(row[i]), column_types[i])
                for i in range(len(headers))
            ]
            if is_total_row(row):
                total_rows.append(cleaned_row)
            else:
                data_rows.append(cleaned_row)

        logger.info(
            "  → %d lignes de données, %d lignes de total",
            len(data_rows), len(total_rows),
        )

        datasets.append(Dataset(
            name=name,
            headers=headers,
            data_rows=data_rows,
            total_rows=total_rows,
            column_types=column_types,
            source_pages=source_pages,
        ))

    return datasets
