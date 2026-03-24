"""Extraction de donnees specifiques depuis les colonnes des tableaux annexes."""

import logging
import re
import unicodedata
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)


def _strip_accents(text: str) -> str:
    """Supprime les accents d'un texte."""
    return "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")


def _normalize_header(header: str) -> str:
    """Normalise un header : strip, sans retour ligne, lowercase, sans accents."""
    text = header.replace("\n", " ").strip().lower()
    nfkd = unicodedata.normalize("NFD", text)
    return "".join(c for c in nfkd if unicodedata.category(c) != "Mn")


def _normalize_for_search(text: str) -> str:
    """Normalise un texte pour la recherche : sans accents, sans espaces, sans apostrophes, lowercase."""
    text = text.replace("\n", "").lower()
    text = _strip_accents(text)
    text = text.replace(" ", "").replace("'", "").replace("\u2019", "")
    return text


def find_column_index(headers: list[str], keywords: list[str]) -> int | None:
    """Trouve l'index d'une colonne par mots-cles dans le header.

    Tous les keywords doivent etre presents dans le header normalise.
    Normalise en supprimant accents, espaces et apostrophes pour tolerer
    les variantes introduites par pdfplumber.
    Retourne None si pas trouve.
    """
    for i, h in enumerate(headers):
        if not h:
            continue
        h_clean = _normalize_for_search(h)
        if all(_normalize_for_search(kw) in h_clean for kw in keywords):
            return i
    return None


def find_column_index_exact(headers: list[str], keyword: str) -> int | None:
    """Trouve l'index d'une colonne par match exact du keyword (sans accents)."""
    kw_norm = _strip_accents(keyword.lower().strip())
    for i, h in enumerate(headers):
        h_norm = _normalize_header(h)
        if h_norm == kw_norm:
            return i
    return None


# Mots-cles a chercher dans les headers pour chaque champ metadonnee
COLUMN_SEARCH = {
    "ref_avis":       ["davis"],                     # "N° d'avis"
    "programme":      ["programme"],                 # "N° Programme"
    "commune":        ["commune"],                   # "Commune"
    "adresse":        ["adresse", "travaux"],         # "Adresse des travaux"
    "montant_deg":    ["degrevement"],               # "Montant degrevement demande"
    "montant_ht":     ["montant", "ht", "facture"],  # "Montant HT facture"
    "taux_tva":       ["taux", "tva"],               # "Taux de TVA facture"
    "montant_ttc":    ["montant", "ttc", "facture"], # "Montant TTC facture"
    "subvention":     ["subvention"],                # "Montant subventions encaisses"
    "nature_travaux": ["nature", "travaux"],         # "NATURE DES TRAVAUX"
    "installateur":   ["installateur"],              # "Installateur"
    "n_operation":    ["operation"],                  # "N°OPERATION"
}


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
    # Gerer les floats type 1146.0 -> "1146"
    if isinstance(value, float) and value == int(value):
        value = int(value)
    text = str(value).strip()
    if not text:
        return False
    if "%" in text or "Taux" in text or "\n" in text:
        return False
    # Doit etre un nombre de 3 a 5 chiffres
    return bool(re.match(r"^\d{3,5}$", text))


def parse_montant_cell(value) -> float:
    """Convertit une cellule montant en float.

    Gere : '791 euro', '791,00 euro', '791', '1 002', 791.0, vide -> 0.0
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
    # Virgule -> point
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _extract_cell_str(row, col_idx) -> str:
    """Extrait une valeur de cellule en string, gere float->int."""
    if col_idx is None or col_idx >= len(row):
        return ""
    val = row[col_idx]
    if val is None:
        return ""
    if isinstance(val, float) and val == int(val):
        val = int(val)
    return str(val).strip()


def _parse_taux_tva(value) -> str:
    """Extrait le taux de TVA d'une cellule. Ex: '20', '20%', 0.2 -> '20%'."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        # 0.2 -> 20%, 20.0 -> 20%
        v = float(value)
        if v < 1:
            v = v * 100
        return f"{v:g}%"
    text = str(value).strip()
    if not text:
        return ""
    text = text.replace("%", "").replace(",", ".").strip()
    try:
        v = float(text)
        if v < 1:
            v = v * 100
        return f"{v:g}%"
    except ValueError:
        return ""


@dataclass
class TableExtractedData:
    """Donnees extraites des colonnes des tableaux annexes."""
    references_avis: str = ""       # valeurs uniques de "N° d'avis" jointes par ";"
    numero_programme: str = ""      # valeurs uniques de "N° Programme" jointes par ","
    commune: str = ""               # valeur unique de "Commune"
    adresse: str = ""               # valeurs uniques de "Adresse des travaux" jointes par ";"
    montant_degrevement: float = 0.0  # somme de "Montant degrevement demande"
    montant_ht_total: str = ""      # somme de "Montant HT facture"
    taux_tva: str = ""              # valeurs uniques de "Taux de TVA facture" jointes par ";"
    montant_ttc_total: str = ""     # somme de "Montant TTC facture"
    montant_subvention: str = ""    # somme de "Montant subventions encaisses"
    nature_travaux: str = ""        # valeurs uniques de "NATURE DES TRAVAUX" jointes par ";"
    nom_entreprise: str = ""        # valeurs uniques de "Installateur" jointes par ";"
    n_operation: str = ""           # valeurs uniques de "N°OPERATION" jointes par ","


def extract_from_datasets(datasets: list) -> TableExtractedData:
    """Parcourt les datasets (feuilles Annexe uniquement) et extrait les valeurs.

    Args:
        datasets: Liste d'objets avec .name, .headers (list[str]) et .data_rows (list[list])

    Returns:
        TableExtractedData avec les donnees extraites.
    """
    all_ref_avis: set[str] = set()
    all_programmes: set[str] = set()
    all_communes: set[str] = set()
    all_adresses: set[str] = set()
    total_degrevement = 0.0
    total_ht = 0.0
    total_ttc = 0.0
    total_subvention = 0.0
    all_natures: set[str] = set()
    all_installateurs: set[str] = set()
    all_operations: set[str] = set()
    all_taux_tva: set[str] = set()

    for ds in datasets:
        # Ne traiter que les feuilles Annexe
        if not ds.name or "annexe" not in ds.name.lower():
            continue

        headers = ds.headers

        # Trouver l'index de chaque colonne
        idx_ref = find_column_index(headers, COLUMN_SEARCH["ref_avis"])
        idx_prog = find_column_index(headers, COLUMN_SEARCH["programme"])
        idx_commune = find_column_index(headers, COLUMN_SEARCH["commune"])
        idx_addr = find_column_index(headers, COLUMN_SEARCH["adresse"])
        idx_deg = find_column_index(headers, COLUMN_SEARCH["montant_deg"])
        idx_ht = find_column_index(headers, COLUMN_SEARCH["montant_ht"])
        idx_tva = find_column_index(headers, COLUMN_SEARCH["taux_tva"])
        idx_ttc = find_column_index(headers, COLUMN_SEARCH["montant_ttc"])
        idx_sub = find_column_index(headers, COLUMN_SEARCH["subvention"])
        idx_nature = find_column_index(headers, COLUMN_SEARCH["nature_travaux"])
        idx_install = find_column_index(headers, COLUMN_SEARCH["installateur"])
        idx_oper = find_column_index(headers, COLUMN_SEARCH["n_operation"])

        logger.info(
            "Annexe '%s': ref=%s, prog=%s, commune=%s, addr=%s, deg=%s, ht=%s, oper=%s",
            ds.name, idx_ref, idx_prog, idx_commune, idx_addr, idx_deg, idx_ht, idx_oper,
        )

        for row in ds.data_rows:
            # References avis
            val = _extract_cell_str(row, idx_ref)
            if val and is_valid_value(val, min_length=5):
                all_ref_avis.add(val)

            # Programme
            val = _extract_cell_str(row, idx_prog)
            if val and is_valid_programme(val):
                all_programmes.add(val)

            # Commune
            val = _extract_cell_str(row, idx_commune)
            if val and is_valid_value(val):
                all_communes.add(val)

            # Adresse
            val = _extract_cell_str(row, idx_addr)
            if val and is_valid_address(val):
                all_adresses.add(val)

            # Montant degrevement (somme)
            if idx_deg is not None and idx_deg < len(row):
                total_degrevement += parse_montant_cell(row[idx_deg])

            # Montant HT (somme)
            if idx_ht is not None and idx_ht < len(row):
                total_ht += parse_montant_cell(row[idx_ht])

            # Montant TTC (somme)
            if idx_ttc is not None and idx_ttc < len(row):
                total_ttc += parse_montant_cell(row[idx_ttc])

            # Montant subvention (somme)
            if idx_sub is not None and idx_sub < len(row):
                total_subvention += parse_montant_cell(row[idx_sub])

            # Taux TVA (valeurs uniques)
            if idx_tva is not None and idx_tva < len(row) and not all_taux_tva:
                parsed = _parse_taux_tva(row[idx_tva])
                if parsed:
                    all_taux_tva.add(parsed)

            # Nature travaux (valeurs uniques)
            val = _extract_cell_str(row, idx_nature)
            if val:
                all_natures.add(val)

            # Installateur (valeurs uniques)
            val = _extract_cell_str(row, idx_install)
            if val:
                all_installateurs.add(val)

            # N° Operation (valeurs uniques)
            val = _extract_cell_str(row, idx_oper)
            if val:
                all_operations.add(val)

    return TableExtractedData(
        references_avis=";".join(sorted(all_ref_avis)),
        numero_programme=", ".join(sorted(all_programmes)),
        commune=", ".join(sorted(all_communes)) if len(all_communes) > 1 else (next(iter(all_communes)) if all_communes else ""),
        adresse=";".join(sorted(all_adresses)),
        montant_degrevement=total_degrevement,
        montant_ht_total=f"{total_ht:.2f}" if total_ht > 0 else "",
        taux_tva=";".join(sorted(all_taux_tva)),
        montant_ttc_total=f"{total_ttc:.2f}" if total_ttc > 0 else "",
        montant_subvention=f"{total_subvention:.2f}" if total_subvention > 0 else "",
        nature_travaux=";".join(sorted(all_natures)),
        nom_entreprise=";".join(sorted(all_installateurs)),
        n_operation=", ".join(sorted(all_operations)),
    )
