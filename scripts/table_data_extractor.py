"""Extraction de donnees specifiques depuis les colonnes des tableaux annexes."""

import re
import unicodedata
from dataclasses import dataclass, field


def _normalize_header(header: str) -> str:
    """Normalise un header : strip, sans retour ligne, lowercase, sans accents."""
    text = header.replace("\n", " ").strip().lower()
    nfkd = unicodedata.normalize("NFD", text)
    return "".join(c for c in nfkd if unicodedata.category(c) != "Mn")


def find_column_index(headers: list[str], keywords: list[str]) -> int | None:
    """Trouve l'index d'une colonne par mots-cles dans le header.

    Tous les keywords doivent etre presents dans le header normalise.
    Retourne None si pas trouve.
    """
    for i, h in enumerate(headers):
        h_norm = _normalize_header(h)
        if all(kw in h_norm for kw in keywords):
            return i
    return None


def find_column_index_exact(headers: list[str], keyword: str) -> int | None:
    """Trouve l'index d'une colonne par match exact du keyword."""
    for i, h in enumerate(headers):
        h_norm = _normalize_header(h)
        if h_norm == keyword:
            return i
    return None


def is_valid_value(value, min_length: int = 3) -> bool:
    """Verifie qu'une valeur n'est pas un parasite."""
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    if len(text) < min_length:
        return False
    if "%" in text:
        return False
    # Doit contenir au moins un chiffre ou une lettre
    if not re.search(r"[a-zA-Z0-9]", text):
        return False
    return True


def is_valid_address(value) -> bool:
    """Verifie qu'une valeur est une adresse valide."""
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    if "%" in text:
        return False
    # Doit contenir au moins une lettre
    if not re.search(r"[a-zA-Z]", text):
        return False
    return True


def is_valid_programme(value) -> bool:
    """Verifie qu'une valeur est un N de programme valide (3-5 chiffres)."""
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    if "%" in text or "Taux" in text or "\n" in text:
        return False
    # Doit etre un nombre de 3 a 5 chiffres
    return bool(re.match(r"^\d{3,5}$", text))


def parse_montant_cell(value) -> float:
    """Convertit une cellule montant en float.

    Gere : '791 euro', '791,00 euro', '791', '1 002', 791.0, vide → 0.0
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    # Supprimer euro/EUR et symbole
    text = re.sub(r"[€]", "", text)
    text = re.sub(r"\s*euros?\s*", "", text, flags=re.IGNORECASE)
    # Supprimer espaces (separateurs milliers)
    text = re.sub(r"\s+", "", text)
    # Virgule → point
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


@dataclass
class TableExtractedData:
    """Donnees extraites des colonnes des tableaux annexes."""
    references_avis: str = ""  # "ref1, ref2"
    adresses: str = ""  # "addr1 / addr2 / ..."
    montant_degrevement_total: float = 0.0
    numero_programme: str = ""  # "1146" ou "1101, 1146"
    commune: str = ""  # "AMIENS"
    adresses_liste: list = field(default_factory=list)


def extract_from_datasets(datasets: list) -> TableExtractedData:
    """Extrait les donnees cibles depuis les datasets en memoire.

    Args:
        datasets: Liste d'objets avec .headers (list[str]) et .data_rows (list[list])

    Returns:
        TableExtractedData avec les donnees extraites.
    """
    refs_set: list[str] = []  # ordered unique
    addrs_set: list[str] = []
    progs_set: list[str] = []
    communes_set: list[str] = []
    montant_sum = 0.0

    refs_seen = set()
    addrs_seen = set()
    progs_seen = set()
    communes_seen = set()

    for ds in datasets:
        headers = ds.headers

        # Trouver les colonnes cibles
        ref_col = find_column_index(headers, ["reference", "avis"])
        addr_col = find_column_index_exact(headers, "adresse")
        montant_col = find_column_index(headers, ["montant", "degrevement"])
        prog_col = find_column_index(headers, ["programme"])
        commune_col = find_column_index_exact(headers, "commune")

        for row in ds.data_rows:
            # References avis
            if ref_col is not None and ref_col < len(row):
                val = str(row[ref_col]).strip() if row[ref_col] is not None else ""
                if is_valid_value(val, min_length=5) and val not in refs_seen:
                    refs_seen.add(val)
                    refs_set.append(val)

            # Adresses
            if addr_col is not None and addr_col < len(row):
                val = str(row[addr_col]).strip() if row[addr_col] is not None else ""
                if is_valid_address(val) and val not in addrs_seen:
                    addrs_seen.add(val)
                    addrs_set.append(val)

            # Montant degrevement (somme)
            if montant_col is not None and montant_col < len(row):
                montant_sum += parse_montant_cell(row[montant_col])

            # Programme
            if prog_col is not None and prog_col < len(row):
                val = str(row[prog_col]).strip() if row[prog_col] is not None else ""
                if is_valid_programme(val) and val not in progs_seen:
                    progs_seen.add(val)
                    progs_set.append(val)

            # Commune
            if commune_col is not None and commune_col < len(row):
                val = str(row[commune_col]).strip() if row[commune_col] is not None else ""
                if is_valid_value(val) and val not in communes_seen:
                    communes_seen.add(val)
                    communes_set.append(val)

    return TableExtractedData(
        references_avis=", ".join(refs_set),
        adresses=";".join(addrs_set),
        montant_degrevement_total=montant_sum,
        numero_programme=", ".join(progs_set),
        commune=", ".join(communes_set) if len(communes_set) > 1 else (communes_set[0] if communes_set else ""),
        adresses_liste=addrs_set,
    )
