"""Extraction des métadonnées des fichiers PDF de dossiers fiscaux TFPB.

Ce module sert aussi de wrapper pour les nouveaux modules scripts/ :
- scripts.raw_extractor : RawMetadata, extraction brute
- scripts.metadata_transformer : ComputedMetadata, transformation metier
"""

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber

# Re-exports depuis scripts/ pour compatibilite
from scripts.raw_extractor import (  # noqa: F401
    RawMetadata,
    build_raw_metadata,
    extract_raw_from_ar,
    extract_raw_from_courrier,
    extract_raw_from_depot,
)
from scripts.metadata_transformer import (  # noqa: F401
    ComputedMetadata,
    compute_metadata,
    computed_metadata_to_rows,
)

logger = logging.getLogger(__name__)


@dataclass
class DossierMetadata:
    """Métadonnées agrégées d'un dossier de demande."""
    numero_demande: str = ""
    objet: str = ""
    libelle: str = ""  # Adresses concernées
    motif_vacance: str = ""
    date_courrier: str = ""
    responsable: str = ""
    numero_lr_depot: str = ""  # N° LR de la preuve de dépôt
    numero_lr_ar: str = ""  # N° LR de l'accusé de réception
    date_presentation_ar: str = ""
    date_distribution_ar: str = ""
    type_fichiers: dict = field(default_factory=dict)
    # {"courrier": "336-xxx.pdf", "depot": "334-xxx.pdf", "ar": "333-xxx.pdf"}


def _clean_text(text: str) -> str:
    """Nettoie le texte extrait : supprime retours à la ligne, espaces multiples."""
    text = text.replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def _extract_first_page_text(pdf_path: Path) -> str:
    """Extrait le texte de la première page d'un PDF."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if pdf.pages:
                text = pdf.pages[0].extract_text() or ""
                return text
    except Exception as e:
        logger.warning("Impossible de lire %s : %s", pdf_path.name, e)
    return ""


def detect_pdf_type(filename: str, text: str = "") -> str:
    """Détecte le type de fichier PDF.

    Args:
        filename: Nom du fichier PDF.
        text: Texte de la première page (optionnel).

    Returns:
        "courrier", "depot", "ar" ou "inconnu".
    """
    name_lower = filename.lower()

    # Détection par nom de fichier
    if "courrier" in name_lower:
        return "courrier"
    if "preuve" in name_lower or "dépôt" in name_lower or "depot" in name_lower:
        return "depot"
    if "ar_n" in name_lower or "ar " in name_lower or name_lower.startswith("ar_"):
        return "ar"

    # Fallback : détection par contenu texte
    if text:
        text_upper = text.upper()
        if "OBJET :" in text_upper and "AFFAIRE SUIVIE PAR" in text_upper:
            return "courrier"
        if "PREUVE DE DÉPÔT" in text_upper or "PREUVE DE DEPOT" in text_upper or "DATE DE DÉPÔT" in text_upper or "DATE DE DEPOT" in text_upper:
            return "depot"
        if "AVIS DE RÉCEPTION" in text_upper or "AVIS DE RECEPTION" in text_upper or "PRÉSENTÉE" in text_upper or "PRESENTEE" in text_upper or "AVISÉE LE" in text_upper or "AVISEE LE" in text_upper:
            return "ar"

    return "inconnu"


def _extract_numero_demande(text: str, filename: str) -> str:
    """Extrait le N° de demande du texte ou du nom de fichier."""
    # Tentative depuis le contenu
    match = re.search(r"demande\s*N°\s*(\d+)", text, re.IGNORECASE)
    if match:
        return match.group(1)

    # Fallback : préfixe du nom de fichier
    match = re.match(r"^(\d+)", filename)
    if match:
        return match.group(1)

    return ""


def _extract_objet(text: str) -> str:
    """Extrait l'objet du courrier."""
    match = re.search(r"Objet\s*:\s*(.+?)(?=Adresses\s+concern)", text, re.DOTALL | re.IGNORECASE)
    if match:
        return _clean_text(match.group(1))
    return ""


def _extract_libelle(text: str) -> str:
    """Extrait les adresses concernées (libellé de la demande)."""
    match = re.search(r"Adresses\s+concern[ée]es\s*:\s*(.+?)(?=Motif)", text, re.DOTALL | re.IGNORECASE)
    if match:
        return _clean_text(match.group(1))
    return ""


def _extract_motif(text: str) -> str:
    """Extrait le motif de la vacance."""
    match = re.search(r"Motif\s+de\s+la\s+vacance\s*:\s*(.+?)(?=Pi[èe]ces|$)", text, re.DOTALL | re.IGNORECASE)
    if match:
        return _clean_text(match.group(1))
    return ""


def _extract_date_courrier(text: str) -> str:
    """Extrait la date du courrier."""
    # Format : "Fait à ..., le ... 15 janvier 2025"
    match = re.search(r"(?:Fait\s+[àa]\s+.+?,\s*le\s+|le\s+lundi\s+|le\s+)(\d{1,2}\s+\w+\s+\d{4})", text, re.IGNORECASE)
    if match:
        return _clean_text(match.group(1))
    # Format dd/mm/yyyy
    match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
    if match:
        return match.group(1)
    return ""


def _extract_responsable(text: str) -> str:
    """Extrait le responsable (Affaire suivie par)."""
    match = re.search(r"Affaire\s+suivie\s+par\s*:\s*(.+?)(?=[,(]|\n|$)", text, re.IGNORECASE)
    if match:
        return _clean_text(match.group(1))
    return ""


def _extract_numero_lr(text: str) -> str:
    """Extrait le N° de lettre recommandée (15 chiffres)."""
    match = re.search(r"(\d{15})", text)
    if match:
        return match.group(1)
    return ""


def _extract_date_presentation_ar(text: str) -> str:
    """Extrait la date de présentation de l'AR."""
    match = re.search(r"(?:Pr[ée]sent[ée]e|avis[ée]e)\s+(?:le\s+)?(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)
    if match:
        return match.group(1)
    return ""


def _extract_date_distribution_ar(text: str) -> str:
    """Extrait la date de distribution de l'AR."""
    match = re.search(r"Distribu[ée]e\s+le\s+(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)
    if match:
        return match.group(1)
    return ""


def extract_metadata(pdf_path: Path) -> tuple[str, dict]:
    """Extrait les métadonnées d'un seul PDF.

    Args:
        pdf_path: Chemin vers le fichier PDF.

    Returns:
        Tuple (type_fichier, dict_metadata).
    """
    pdf_path = Path(pdf_path)
    text = _extract_first_page_text(pdf_path)
    pdf_type = detect_pdf_type(pdf_path.name, text)

    metadata = {"type": pdf_type, "filename": pdf_path.name}

    if pdf_type == "courrier":
        metadata["numero_demande"] = _extract_numero_demande(text, pdf_path.name)
        metadata["objet"] = _extract_objet(text)
        metadata["libelle"] = _extract_libelle(text)
        metadata["motif_vacance"] = _extract_motif(text)
        metadata["date_courrier"] = _extract_date_courrier(text)
        metadata["responsable"] = _extract_responsable(text)
    elif pdf_type == "depot":
        metadata["numero_lr"] = _extract_numero_lr(text)
    elif pdf_type == "ar":
        metadata["numero_lr"] = _extract_numero_lr(text)
        metadata["date_presentation"] = _extract_date_presentation_ar(text)
        metadata["date_distribution"] = _extract_date_distribution_ar(text)

    logger.info("PDF '%s' : type=%s", pdf_path.name, pdf_type)
    return pdf_type, metadata


def process_dossier(dir_path: Path) -> DossierMetadata:
    """Traite un dossier complet et agrège les métadonnées.

    Un dossier = une demande. Tous les PDFs sont regroupés.

    Args:
        dir_path: Chemin vers le dossier contenant les PDFs.

    Returns:
        DossierMetadata agrégé.
    """
    dir_path = Path(dir_path)
    pdf_files = sorted(dir_path.glob("*.pdf"))

    if not pdf_files:
        logger.warning("Aucun PDF trouvé dans %s", dir_path)
        return DossierMetadata()

    dossier = DossierMetadata()

    for pdf_file in pdf_files:
        pdf_type, meta = extract_metadata(pdf_file)

        dossier.type_fichiers[pdf_type] = pdf_file.name

        if pdf_type == "courrier":
            dossier.numero_demande = meta.get("numero_demande", "")
            dossier.objet = meta.get("objet", "")
            dossier.libelle = meta.get("libelle", "")
            dossier.motif_vacance = meta.get("motif_vacance", "")
            dossier.date_courrier = meta.get("date_courrier", "")
            dossier.responsable = meta.get("responsable", "")
        elif pdf_type == "depot":
            dossier.numero_lr_depot = meta.get("numero_lr", "")
        elif pdf_type == "ar":
            dossier.numero_lr_ar = meta.get("numero_lr", "")
            dossier.date_presentation_ar = meta.get("date_presentation", "")
            dossier.date_distribution_ar = meta.get("date_distribution", "")

    # Fallback N° demande depuis n'importe quel nom de fichier
    if not dossier.numero_demande:
        for pdf_file in pdf_files:
            match = re.match(r"^(\d+)", pdf_file.name)
            if match:
                dossier.numero_demande = match.group(1)
                break

    logger.info(
        "Dossier '%s' : demande N°%s, %d fichier(s) détecté(s)",
        dir_path.name, dossier.numero_demande, len(dossier.type_fichiers),
    )

    return dossier


def format_metadata_report(dossier: DossierMetadata) -> str:
    """Génère un rapport textuel des métadonnées."""
    lines = []
    lines.append(f"{'='*60}")
    lines.append(f"Métadonnées du dossier — Demande N°{dossier.numero_demande}")
    lines.append(f"{'='*60}")

    fields = [
        ("N° Demande", dossier.numero_demande),
        ("Objet", dossier.objet),
        ("Libellé (Adresses)", dossier.libelle),
        ("Motif vacance", dossier.motif_vacance),
        ("Date courrier", dossier.date_courrier),
        ("Responsable", dossier.responsable),
        ("N° LR Dépôt", dossier.numero_lr_depot),
        ("N° LR AR", dossier.numero_lr_ar),
        ("Date présentation AR", dossier.date_presentation_ar),
        ("Date distribution AR", dossier.date_distribution_ar),
    ]

    for label, value in fields:
        display = value if value else "(non trouvé)"
        lines.append(f"  {label:25s} : {display}")

    if dossier.type_fichiers:
        lines.append(f"\n{'Fichiers détectés':25s} :")
        for ftype, fname in dossier.type_fichiers.items():
            lines.append(f"  {ftype:25s} : {fname}")

    lines.append(f"{'='*60}")
    return "\n".join(lines)


def metadata_to_rows(dossier: DossierMetadata) -> list[tuple[str, str]]:
    """Convertit les métadonnées en liste de tuples (clé, valeur) pour Excel."""
    return [
        ("N° Demande", dossier.numero_demande),
        ("Objet", dossier.objet),
        ("Libellé (Adresses concernées)", dossier.libelle),
        ("Motif de la vacance", dossier.motif_vacance),
        ("Date du courrier", dossier.date_courrier),
        ("Responsable", dossier.responsable),
        ("N° LR Preuve de dépôt", dossier.numero_lr_depot),
        ("N° LR Accusé de réception", dossier.numero_lr_ar),
        ("Date de présentation AR", dossier.date_presentation_ar),
        ("Date de distribution AR", dossier.date_distribution_ar),
        ("Fichier courrier", dossier.type_fichiers.get("courrier", "")),
        ("Fichier preuve de dépôt", dossier.type_fichiers.get("depot", "")),
        ("Fichier AR", dossier.type_fichiers.get("ar", "")),
    ]
