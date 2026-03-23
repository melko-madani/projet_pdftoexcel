"""Etape 2 : Transformation des donnees brutes en donnees metier."""

import logging
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

from .raw_extractor import RawMetadata, _extract_all_text

logger = logging.getLogger(__name__)

MOIS_FR = {
    "janvier": "01", "février": "02", "fevrier": "02", "mars": "03",
    "avril": "04", "mai": "05", "juin": "06", "juillet": "07",
    "août": "08", "aout": "08", "septembre": "09", "octobre": "10",
    "novembre": "11", "décembre": "12", "decembre": "12",
}


@dataclass
class ComputedMetadata:
    """Donnees metier calculees a partir des donnees brutes."""
    # Identifiant
    libelle_demande: str = ""
    responsable: str = ""

    # Classification
    type_demande: str = ""
    categorie: str = ""
    sous_categorie: str = ""

    # Montants
    montant_ht: float = 0.0
    taux_tva: str = "0%"
    montant_ttc: float = 0.0
    montant_demande: float = 0.0

    # Entreprise
    nom_entreprise: str = ""
    nature_travaux: str = ""
    nature_depenses: str = "Degrevement taxe fonciere"

    # References
    ref_avis: str = ""
    adresse: str = ""
    numero_programme: str = ""
    nombre_logements: int = 0
    numero_operation: str = ""

    # Envoi
    date_limite_envoi: str = ""
    type_envoi: str = ""
    numero_recommande: str = ""

    # Interlocuteur
    nom_interlocuteur: str = ""
    prenom_interlocuteur: str = ""
    mail_interlocuteur: str = ""
    tel_interlocuteur: str = ""

    # Divers
    commentaire: str = ""
    lien_escale: str = ""
    montant_subvention: float = 0.0


def parse_montant(raw: str) -> float:
    """Convertit '55 245' ou '106 281,50' en float."""
    if not raw or not raw.strip():
        return 0.0
    text = raw.strip()
    # Supprimer espaces (separateurs de milliers)
    text = re.sub(r"\s+", "", text)
    # Remplacer virgule par point
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def format_date_fr(raw: str) -> str:
    """Convertit '29 decembre 2025' en '29/12/2025'. Passe-plat si deja au format dd/mm/yyyy."""
    if not raw or not raw.strip():
        return ""
    text = raw.strip()

    # Deja au format dd/mm/yyyy ?
    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
        return text

    # Format "29 decembre 2025"
    match = re.match(r"(\d{1,2})\s+(\w+)\s+(\d{4})", text)
    if match:
        day = match.group(1).zfill(2)
        month_str = match.group(2).lower()
        year = match.group(3)
        month = MOIS_FR.get(month_str, "")
        if month:
            return f"{day}/{month}/{year}"

    return text


def deduce_type(objet: str) -> str:
    """Deduit le type depuis l'objet."""
    lower = objet.lower()
    if "degrevement" in lower or "dégrèvement" in lower:
        return "Degrevement"
    if "exoneration" in lower or "exonération" in lower:
        return "Exoneration"
    return "Autre"


def deduce_categorie(objet: str) -> str:
    """Deduit la categorie depuis l'objet."""
    lower = objet.lower()
    if "tfpb" in lower or "taxe fonciere" in lower or "taxe foncière" in lower:
        return "TFPB"
    if "tfpnb" in lower:
        return "TFPNB"
    return "Autre"


def deduce_sous_categorie(motif: str, objet: str) -> str:
    """Deduit la sous-categorie depuis le motif et l'objet."""
    text = (motif + " " + objet).lower()
    if "anru" in text or "renouvellement urbain" in text:
        return "Vacance ANRU"
    if "demolir" in text or "demolition" in text or "démolir" in text or "démolition" in text:
        return "Vacance demolition"
    if "securite" in text or "sécurité" in text or "mitoyen" in text:
        return "Vacance securite"
    if "vacance" in text or "vacant" in text:
        return "Vacance"
    return "Autre"


def extract_entreprise(courrier_bytes: bytes | None) -> str:
    """Extrait le nom de l'entreprise/bailleur depuis le texte du courrier."""
    if not courrier_bytes:
        return ""
    text = _extract_all_text(courrier_bytes)
    if not text:
        return ""

    # Chercher "mandataire de la ..."
    match = re.search(
        r"mandataire\s+de\s+(?:la\s+)?(.+?)(?:\n|,|\.)",
        text, re.IGNORECASE,
    )
    if match:
        return match.group(1).strip()

    # Chercher "pour le compte de ..."
    match = re.search(
        r"pour\s+le\s+compte\s+de\s+(?:la\s+)?(.+?)(?:\n|,|\.)",
        text, re.IGNORECASE,
    )
    if match:
        return match.group(1).strip()

    return ""


def extract_numero_programme_from_tables(datasets) -> str:
    """Extrait le N de programme depuis la premiere ligne du tableau annexe."""
    if not datasets:
        return ""

    for ds in datasets:
        # Chercher la colonne programme
        headers_lower = [h.lower().replace("\n", " ") for h in ds.headers]
        prog_idx = None
        for i, h in enumerate(headers_lower):
            if "programme" in h:
                prog_idx = i
                break

        if prog_idx is not None and ds.data_rows:
            # Premiere valeur non-vide
            for row in ds.data_rows:
                if prog_idx < len(row) and row[prog_idx]:
                    val = str(row[prog_idx]).strip()
                    if val:
                        return val

    return ""


def compute_metadata(
    raw: RawMetadata,
    datasets=None,
    courrier_bytes: bytes | None = None,
) -> ComputedMetadata:
    """Transforme RawMetadata en ComputedMetadata."""
    c = ComputedMetadata()

    # Identifiant
    addr_short = raw.adresses[:80] if raw.adresses else ""
    annee = raw.annee_fiscale or ""
    if addr_short and annee:
        c.libelle_demande = f"{addr_short} - TFPB {annee}"
    elif addr_short:
        c.libelle_demande = addr_short
    else:
        c.libelle_demande = f"Demande N\u00b0{raw.numero_demande}"

    c.responsable = raw.responsable

    # Classification
    c.type_demande = deduce_type(raw.objet_complet)
    c.categorie = deduce_categorie(raw.objet_complet)
    c.sous_categorie = deduce_sous_categorie(raw.motif_vacance, raw.objet_complet)

    # Montants
    c.montant_ht = parse_montant(raw.montant_cotisations_sans_frais)
    c.taux_tva = "0%"
    c.montant_ttc = parse_montant(raw.montant_total_imposition)
    c.montant_demande = parse_montant(raw.montant_degrevement)

    # Entreprise
    c.nom_entreprise = extract_entreprise(courrier_bytes)
    c.nature_travaux = raw.motif_vacance
    c.nature_depenses = "Degrevement taxe fonciere"

    # References
    c.ref_avis = raw.ref_avis_imposition.strip()
    c.adresse = raw.adresses
    c.numero_programme = extract_numero_programme_from_tables(datasets)
    try:
        c.nombre_logements = int(raw.nombre_logements) if raw.nombre_logements else 0
    except ValueError:
        c.nombre_logements = 0
    c.numero_operation = ""

    # Envoi
    c.date_limite_envoi = format_date_fr(raw.date_limite_envoi)
    c.type_envoi = "Recommande avec AR" if raw.numero_lr_ar else "Recommande"
    c.numero_recommande = raw.numero_lr_depot

    # Interlocuteur
    dest = raw.nom_destinataire.strip()
    parts = dest.split() if dest else []
    if len(parts) >= 2:
        c.nom_interlocuteur = parts[-1]
        c.prenom_interlocuteur = " ".join(parts[:-1])
    elif len(parts) == 1:
        c.nom_interlocuteur = parts[0]
    c.mail_interlocuteur = ""
    c.tel_interlocuteur = ""

    # Divers
    c.commentaire = ""
    c.lien_escale = ""
    c.montant_subvention = 0.0

    return c


def computed_metadata_to_rows(c: ComputedMetadata) -> list[tuple[str, object]]:
    """Convertit ComputedMetadata en liste ordonnee de (label, valeur) pour Excel.

    Les valeurs float sont laissees telles quelles pour le formatage euro.
    """
    return [
        ("Libelle de la Demande", c.libelle_demande),
        ("Responsable", c.responsable),
        ("Type", c.type_demande),
        ("Categorie", c.categorie),
        ("Sous-categorie", c.sous_categorie),
        ("Montant HT", c.montant_ht),
        ("Nom de l'entreprise", c.nom_entreprise),
        ("Taux de TVA", c.taux_tva),
        ("Nature des travaux", c.nature_travaux),
        ("Montant TTC", c.montant_ttc),
        ("Montant de la subvention", c.montant_subvention),
        ("Reference(s) Avis", c.ref_avis),
        ("Adresse", c.adresse),
        ("Montant de la demande", c.montant_demande),
        ("Date limite d'envoi", c.date_limite_envoi),
        ("Type d'envoi", c.type_envoi),
        ("Numero de recommande", c.numero_recommande),
        ("Commentaire", c.commentaire),
        ("Lien escale", c.lien_escale),
        ("Nom interlocuteur", c.nom_interlocuteur),
        ("Prenom interlocuteur", c.prenom_interlocuteur),
        ("Mail interlocuteur", c.mail_interlocuteur),
        ("Tel interlocuteur", c.tel_interlocuteur),
        ("N° Programme", c.numero_programme),
        ("Nombre de logements", c.nombre_logements),
        ("Nature de depenses", c.nature_depenses),
        ("N° d'operation", c.numero_operation),
    ]
