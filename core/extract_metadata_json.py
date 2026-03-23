"""Extraction ciblée de métadonnées pour 04.json.

Extrait : Sous-catégorie, Montant HT, Nom de l'entreprise,
Taux de TVA, Nature des travaux, Montant TTC, Montant de la subvention.
Sous-catégorie déterminée via l'arbre Dégrèvement pour Travaux.
"""

import json
import logging
import re
import tempfile
from pathlib import Path

import pdfplumber

logger = logging.getLogger(__name__)

# --- Arbre de sous-catégories ---
# Type: Dégrèvement pour Travaux
#   Catégorie: Accessibilité PMR
#     - Aménagement parties communes
#     - Aménagement parties privatives
#     - Ascenseur
#     - Cheminements parties communes
#     - Élargissement/Aménagement parking
#     - Global
#   Catégorie: Économie d'énergie
#     - Isolation
#     - Chauffage/Refroidissement
#     - Éclairage
#     - Eau chaude
#     - Global

_SOUS_CAT_PMR = {
    "parties communes": "Amenagement parties communes",
    "amenagement parties communes": "Amenagement parties communes",
    "aménagement parties communes": "Amenagement parties communes",
    "parties privatives": "Amenagement parties privatives",
    "amenagement parties privatives": "Amenagement parties privatives",
    "aménagement parties privatives": "Amenagement parties privatives",
    "ascenseur": "Ascenseur",
    "cheminement": "Cheminements parties communes",
    "cheminements": "Cheminements parties communes",
    "elargissement": "Elargissement/Amenagement parking",
    "élargissement": "Elargissement/Amenagement parking",
    "parking": "Elargissement/Amenagement parking",
}

_SOUS_CAT_ENERGIE = {
    "isolation": "Isolation",
    "chauffage": "Chauffage/Refroidissement",
    "refroidissement": "Chauffage/Refroidissement",
    "climatisation": "Chauffage/Refroidissement",
    "eclairage": "Eclairage",
    "éclairage": "Eclairage",
    "eau chaude": "Eau chaude",
    "ecs": "Eau chaude",
}


def _detect_categorie_travaux(text: str) -> str:
    """Détecte la catégorie de travaux : Accessibilité PMR ou Économie d'énergie."""
    lower = text.lower()
    pmr_keywords = ["pmr", "accessibilit", "handicap", "mobilité réduite", "mobilite reduite"]
    energie_keywords = ["énergie", "energie", "thermique", "énergétique", "energetique"]

    for kw in pmr_keywords:
        if kw in lower:
            return "Accessibilite PMR"
    for kw in energie_keywords:
        if kw in lower:
            return "Economie d'energie"
    return ""


def _detect_sous_categorie(text: str, categorie: str) -> str:
    """Détecte la sous-catégorie selon la catégorie et le texte du PDF."""
    lower = text.lower()

    if categorie == "Accessibilite PMR":
        for keyword, sous_cat in _SOUS_CAT_PMR.items():
            if keyword in lower:
                return sous_cat
        return "Global"

    if categorie == "Economie d'energie":
        for keyword, sous_cat in _SOUS_CAT_ENERGIE.items():
            if keyword in lower:
                return sous_cat
        return "Global"

    return ""


def _extract_all_text(pdf_bytes: bytes) -> str:
    """Extrait le texte de toutes les pages d'un PDF."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    try:
        tmp.write(pdf_bytes)
        tmp.close()
        with pdfplumber.open(tmp.name) as pdf:
            pages = []
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    pages.append(t)
            return "\n".join(pages)
    except Exception as e:
        logger.warning("Impossible de lire le PDF : %s", e)
        return ""
    finally:
        Path(tmp.name).unlink(missing_ok=True)


def _parse_montant(raw: str) -> float:
    """Convertit '55 245,50' ou '106 281' en float."""
    if not raw or not raw.strip():
        return 0.0
    text = re.sub(r"[€\s\u00a0]+", "", raw.strip())
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _search(pattern: re.Pattern, text: str, group: int = 1) -> str:
    """Recherche un pattern et retourne le groupe ou ''."""
    match = pattern.search(text)
    return match.group(group).strip() if match else ""


# Patterns pour extraction
_PAT_ENTREPRISE = [
    re.compile(r"mandataire\s+de\s+(?:la\s+)?(.+?)(?:\n|,|\.)", re.IGNORECASE),
    re.compile(r"pour\s+le\s+compte\s+de\s+(?:la\s+)?(.+?)(?:\n|,|\.)", re.IGNORECASE),
    re.compile(r"entreprise\s*:\s*(.+?)(?:\n|,|\.)", re.IGNORECASE),
    re.compile(r"réalis[ée]s?\s+par\s+(?:la\s+société\s+|l'entreprise\s+)?(.+?)(?:\n|,|\.)", re.IGNORECASE),
]

_PAT_TVA = re.compile(r"(?:taux\s+de\s+)?TVA\s*[:\s]*(\d+[\.,]?\d*)\s*%", re.IGNORECASE)
_PAT_TVA_ALT = re.compile(r"(\d+[\.,]?\d*)\s*%\s*(?:de\s+)?TVA", re.IGNORECASE)

_PAT_MONTANT_HT = [
    re.compile(r"montant\s+HT\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)?", re.IGNORECASE),
    re.compile(r"hors\s+taxes?\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)?", re.IGNORECASE),
    re.compile(r"HT\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)", re.IGNORECASE),
    re.compile(r"montant\s+[ée]gal\s+[àa]\s*([\d\s.,]+)\s*(?:euros?|€)", re.IGNORECASE),
]

_PAT_MONTANT_TTC = [
    re.compile(r"montant\s+TTC\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)?", re.IGNORECASE),
    re.compile(r"toutes\s+taxes\s+comprises?\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)?", re.IGNORECASE),
    re.compile(r"TTC\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)", re.IGNORECASE),
    re.compile(r"montant\s+total\s+de\s*([\d\s.,]+)\s*(?:euros?|€)", re.IGNORECASE),
]

_PAT_SUBVENTION = [
    re.compile(r"subvention\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)?", re.IGNORECASE),
    re.compile(r"montant\s+de\s+la\s+subvention\s*[:\s]*([\d\s.,]+)\s*(?:€|euros?)?", re.IGNORECASE),
]

_PAT_NATURE_TRAVAUX = [
    re.compile(r"nature\s+des\s+travaux\s*[:\s]*(.+?)(?:\n|$)", re.IGNORECASE),
    re.compile(r"travaux\s+de\s+(.+?)(?:\n|,\s*(?:pour|dans|sur))", re.IGNORECASE),
    re.compile(r"Motif\s+de\s+la\s+vacance\s*:\s*(.+?)(?=Pi[èe]ces|$)", re.DOTALL | re.IGNORECASE),
]

_PAT_OBJET = re.compile(
    r"Objet\s*:\s*(.+?)(?=Adresses\s+concern[ée]es)", re.DOTALL | re.IGNORECASE
)


def extract_04_from_pdf(pdf_bytes: bytes, filename: str = "") -> dict:
    """Extrait les 7 champs cibles depuis un PDF courrier.

    Returns:
        Dict avec les 7 champs.
    """
    text = _extract_all_text(pdf_bytes)
    if not text:
        return {
            "Sous-categorie": "",
            "Montant HT": 0.0,
            "Nom de l'entreprise": "",
            "Taux de TVA": "0%",
            "Nature des travaux": "",
            "Montant TTC": 0.0,
            "Montant de la subvention": 0.0,
        }

    # Objet complet pour détection catégorie/sous-catégorie
    objet = _search(_PAT_OBJET, text)
    full_context = objet + " " + text

    # Catégorie → Sous-catégorie
    categorie = _detect_categorie_travaux(full_context)
    sous_categorie = _detect_sous_categorie(full_context, categorie)

    # Entreprise
    nom_entreprise = ""
    for pat in _PAT_ENTREPRISE:
        nom_entreprise = _search(pat, text)
        if nom_entreprise:
            break

    # TVA
    taux_tva = _search(_PAT_TVA, text)
    if not taux_tva:
        taux_tva = _search(_PAT_TVA_ALT, text)
    taux_tva = f"{taux_tva}%" if taux_tva else "0%"

    # Montant HT
    montant_ht = 0.0
    for pat in _PAT_MONTANT_HT:
        raw = _search(pat, text)
        if raw:
            montant_ht = _parse_montant(raw)
            if montant_ht > 0:
                break

    # Montant TTC
    montant_ttc = 0.0
    for pat in _PAT_MONTANT_TTC:
        raw = _search(pat, text)
        if raw:
            montant_ttc = _parse_montant(raw)
            if montant_ttc > 0:
                break

    # Subvention
    montant_subvention = 0.0
    for pat in _PAT_SUBVENTION:
        raw = _search(pat, text)
        if raw:
            montant_subvention = _parse_montant(raw)
            if montant_subvention > 0:
                break

    # Nature des travaux
    nature_travaux = ""
    for pat in _PAT_NATURE_TRAVAUX:
        nature_travaux = _search(pat, text)
        if nature_travaux:
            # Nettoyer
            nature_travaux = re.sub(r"\s+", " ", nature_travaux).strip()
            break

    return {
        "Sous-categorie": sous_categorie,
        "Montant HT": montant_ht,
        "Nom de l'entreprise": nom_entreprise,
        "Taux de TVA": taux_tva,
        "Nature des travaux": nature_travaux,
        "Montant TTC": montant_ttc,
        "Montant de la subvention": montant_subvention,
    }


def extract_and_save_04_json(
    pdf_files: list[tuple[str, bytes]],
    output_path: str | Path = "04.json",
) -> dict:
    """Extrait les métadonnées de plusieurs PDFs et sauvegarde en 04.json.

    Args:
        pdf_files: Liste de tuples (filename, pdf_bytes).
        output_path: Chemin du fichier JSON de sortie.

    Returns:
        Le dict complet exporté.
    """
    output_path = Path(output_path)
    result = {
        "colonnes_extraites": [
            "Sous-categorie",
            "Montant HT",
            "Nom de l'entreprise",
            "Taux de TVA",
            "Nature des travaux",
            "Montant TTC",
            "Montant de la subvention",
        ],
        "arbre_sous_categories": {
            "Degrevement pour Travaux": {
                "Accessibilite PMR": [
                    "Amenagement parties communes",
                    "Amenagement parties privatives",
                    "Ascenseur",
                    "Cheminements parties communes",
                    "Elargissement/Amenagement parking",
                    "Global",
                ],
                "Economie d'energie": [
                    "Isolation",
                    "Chauffage/Refroidissement",
                    "Eclairage",
                    "Eau chaude",
                    "Global",
                ],
            }
        },
        "nombre_fichiers": len(pdf_files),
        "fichiers": [],
    }

    for filename, pdf_bytes in pdf_files:
        logger.info("Extraction 04 de %s...", filename)
        data = extract_04_from_pdf(pdf_bytes, filename)
        result["fichiers"].append({
            "nom_fichier": filename,
            "donnees": data,
        })

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    logger.info("Export 04.json terminé : %s", output_path)
    return result
