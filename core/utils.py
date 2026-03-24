"""Fonctions utilitaires pour le parsing des données fiscales."""

import re
from datetime import datetime
from typing import Any

# Mots-clés identifiant les lignes de total/sous-total
TOTAL_KEYWORDS = [
    "total",
    "sous-total",
    "sous total",
    "montant total",
    "total général",
    "total general",
]


def clean_cell(value: Any) -> str:
    """Nettoie une cellule : strip, supprime les retours à la ligne internes."""
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s*\n\s*", " ", text)
    return text


def parse_euro(value: str) -> float | None:
    """Convertit une valeur monétaire française en float.

    Exemples : '542,78 €' → 542.78, '1 234,56' → 1234.56
    """
    if not value or not value.strip():
        return None
    text = value.strip()
    # Supprimer le symbole € et les espaces autour
    text = text.replace("€", "").strip()
    # Supprimer les espaces insécables et espaces utilisés comme séparateur de milliers
    text = re.sub(r"[\s\u00a0\u202f]", "", text)
    # Remplacer la virgule décimale par un point
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value: str) -> datetime | None:
    """Convertit une date au format dd/mm/yyyy en datetime."""
    if not value or not value.strip():
        return None
    text = value.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_mois_vacance(value: str) -> int | None:
    """Convertit '12 douzièmes' ou '6 douzièmes' en entier."""
    if not value or not value.strip():
        return None
    match = re.search(r"(\d+)\s*douzi[eè]mes?", value.strip(), re.IGNORECASE)
    if match:
        return int(match.group(1))
    # Essayer juste un nombre seul
    match = re.match(r"^\s*(\d+)\s*$", value.strip())
    if match:
        return int(match.group(1))
    return None


def is_total_row(row: list[str]) -> bool:
    """Detecte si une ligne est une ligne de total/sous-total.

    Verifie uniquement la premiere cellule non vide pour eviter les faux positifs
    (ex: 'SDIF DE LA SOMME' dans une cellule de donnees).
    """
    for cell in row:
        cleaned = clean_cell(cell).lower()
        if cleaned:
            return any(keyword in cleaned for keyword in TOTAL_KEYWORDS)
    return False


def is_empty_row(row: list[str]) -> bool:
    """Détecte si une ligne est entièrement vide."""
    return all(not clean_cell(cell) for cell in row)


def detect_column_type(header: str, sample_values: list[str]) -> str:
    """Détecte le type d'une colonne à partir du header et des valeurs.

    Retourne : 'euro', 'date', 'mois_vacance', 'text'
    """
    header_lower = header.lower()

    # Détection par header
    euro_keywords = [
        "montant", "cotisation", "frais", "part", "teom",
        "intercommunalité", "intercommunalite", "dégrèvement",
        "degrevement", "sous-total", "total", "(€)", "€",
    ]
    date_keywords = ["date"]
    mois_keywords = ["mois", "douzième", "douzieme", "vacance"]

    if any(kw in header_lower for kw in euro_keywords):
        return "euro"
    if any(kw in header_lower for kw in date_keywords):
        return "date"
    if any(kw in header_lower for kw in mois_keywords):
        # Vérifier avec les valeurs pour distinguer mois_vacance vs date
        for val in sample_values:
            if val and "douzi" in val.lower():
                return "mois_vacance"
        # Si le header contient "date", c'est une date
        if "date" in header_lower:
            return "date"
        return "mois_vacance"

    # Détection par valeurs
    for val in sample_values:
        if val and parse_euro(val) is not None and ("," in val or "€" in val):
            return "euro"
        if val and parse_date(val) is not None:
            return "date"

    return "text"
