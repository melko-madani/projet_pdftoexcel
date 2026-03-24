"""Arborescence des types/categories/sous-categories de demandes fiscales.

Classification basee sur des mots-cles dans le texte du courrier.
Recherche insensible a la casse ET aux accents.
"""

import unicodedata


def _normalize(text: str) -> str:
    """Normalise un texte : minuscule, sans accents."""
    text = text.lower()
    # Decomposer les caracteres accentues puis supprimer les diacritiques
    nfkd = unicodedata.normalize("NFD", text)
    return "".join(c for c in nfkd if unicodedata.category(c) != "Mn")


def _contains(text_norm: str, *keywords: str) -> bool:
    """Verifie si le texte normalise contient un des mots-cles."""
    return any(kw in text_norm for kw in keywords)


CLASSIFICATION = {
    "Dégrèvement pour Travaux": {
        "Accessibilité PMR": [
            "Aménagement parties communes",
            "Aménagement parties privatives",
            "Ascenceur",
            "Cheminements parties communes",
            "Élargissement/Aménagement parking",
            "Global",
            "Autre",
        ],
        "Economie d'énergie": [
            "Isolation",
            "Chauffage/Refroidissement",
            "Eclairage",
            "Eau chaude",
            "Global",
            "Autre",
        ],
        "Autre": [],
    },
    "Vacance": {
        "Travaux": [],
        "Démolition": [],
        "Locative": [],
        "Autre": [],
    },
    "Régularisation Abattement/Exonération": {
        "ElementsDeConfort": [],
        "Coefficients": [],
        "TypeDeBien": [],
        "HorsPatrimoine": [],
        "FinDeGestion": [],
        "Autre": [],
        "VideOrdures": [],
        "Categorie": [],
        "Adresse": [],
        "ThLogementsVacants": [],
    },
    "Autre Régularisation": {},
}


# --- Type ---

def deduce_type(objet: str, full_text: str = "") -> str:
    """Deduit le type depuis l'objet du courrier.

    Regles :
    - "vacance" → Vacance
    - "travaux" sans "vacance" → Dégrèvement pour Travaux
    - "regularisation" + ("abattement" ou "exoneration") → Régularisation Abattement/Exonération
    - "regularisation" seul → Autre Régularisation
    - sinon → ""
    """
    obj = _normalize(objet)
    txt = _normalize(full_text) if full_text else obj

    if _contains(obj, "vacance"):
        return "Vacance"

    if _contains(obj, "travaux") and not _contains(obj, "vacance"):
        return "Dégrèvement pour Travaux"

    if _contains(obj, "regularisation"):
        if _contains(obj, "abattement", "exoneration"):
            return "Régularisation Abattement/Exonération"
        return "Autre Régularisation"

    # Fallback sur le texte complet
    if _contains(txt, "vacance"):
        return "Vacance"

    if _contains(txt, "degrevement") and _contains(txt, "travaux"):
        return "Dégrèvement pour Travaux"

    return ""


# --- Categorie ---

def deduce_categorie(type_demande: str, motif: str, objet: str, full_text: str = "") -> str:
    """Deduit la categorie selon le type.

    Pour Vacance : Démolition / Travaux / Locative / Autre
    Pour Dégrèvement pour Travaux : Accessibilité PMR / Economie d'énergie / Autre
    Pour Régularisation Abattement/Exonération : sous-types specifiques
    """
    motif_n = _normalize(motif)
    objet_n = _normalize(objet)
    text_n = _normalize(full_text) if full_text else motif_n + " " + objet_n
    combined = motif_n + " " + objet_n + " " + text_n

    if type_demande == "Vacance":
        if _contains(combined, "demolir", "demolition", "demolis"):
            return "Démolition"
        if _contains(combined, "travaux"):
            return "Travaux"
        if _contains(combined, "locati", "location"):
            return "Locative"
        # Defaut pour vacance (ANRU = demolition par defaut)
        return "Démolition"

    if type_demande == "Dégrèvement pour Travaux":
        if _contains(combined, "pmr", "accessibilite", "handicap"):
            return "Accessibilité PMR"
        if _contains(combined, "energie", "isolation", "chauffage", "thermique"):
            return "Economie d'énergie"
        return "Autre"

    if type_demande == "Régularisation Abattement/Exonération":
        if _contains(combined, "confort"):
            return "ElementsDeConfort"
        if _contains(combined, "coefficient"):
            return "Coefficients"
        if _contains(combined, "type de bien"):
            return "TypeDeBien"
        if _contains(combined, "hors patrimoine"):
            return "HorsPatrimoine"
        if _contains(combined, "fin de gestion"):
            return "FinDeGestion"
        if _contains(combined, "vide ordures", "vide-ordures"):
            return "VideOrdures"
        if _contains(combined, "categorie"):
            return "Categorie"
        if _contains(combined, "adresse"):
            return "Adresse"
        if _contains(combined, "logements vacants"):
            return "ThLogementsVacants"
        return "Autre"

    return ""


# --- Sous-categorie ---

def _match_sous_cat_pmr(text_n: str) -> str:
    """Cherche une sous-categorie PMR dans un texte normalise."""
    if _contains(text_n, "parties communes"):
        return "Aménagement parties communes"
    if _contains(text_n, "parties privatives"):
        return "Aménagement parties privatives"
    if _contains(text_n, "ascenseur"):
        return "Ascenceur"
    if _contains(text_n, "cheminement"):
        return "Cheminements parties communes"
    if _contains(text_n, "parking", "elargissement"):
        return "Élargissement/Aménagement parking"
    if _contains(text_n, "global"):
        return "Global"
    return ""


def _match_sous_cat_energie(text_n: str) -> str:
    """Cherche une sous-categorie Economie d'energie dans un texte normalise."""
    if _contains(text_n, "isolation"):
        return "Isolation"
    if _contains(text_n, "chauffage", "refroidissement"):
        return "Chauffage/Refroidissement"
    if _contains(text_n, "eclairage"):
        return "Eclairage"
    if _contains(text_n, "eau chaude"):
        return "Eau chaude"
    if _contains(text_n, "global"):
        return "Global"
    return ""


def deduce_sous_categorie(type_demande: str, categorie: str, objet: str = "", nature_travaux: str = "") -> str:
    """Deduit la sous-categorie selon le type et la categorie.

    Cherche dans la nature des travaux du tableau (plus fiable car specifique),
    puis dans l'objet du courrier.
    Ne cherche PAS dans le texte complet pour eviter les faux positifs.
    Si rien ne matche, retourne "" (vide).
    """
    nature_n = _normalize(nature_travaux) if nature_travaux else ""
    objet_n = _normalize(objet)

    if categorie == "Accessibilité PMR":
        if nature_n:
            result = _match_sous_cat_pmr(nature_n)
            if result:
                return result
        return _match_sous_cat_pmr(objet_n)

    if categorie == "Economie d'énergie":
        if nature_n:
            result = _match_sous_cat_energie(nature_n)
            if result:
                return result
        return _match_sous_cat_energie(objet_n)

    return ""


# --- Libelle ---

def build_libelle(annee: str, categorie: str, numero_programme: str, commune: str) -> str:
    """Construit le libelle de la demande.

    Format : 'Degrevement TFPB {annee} - {categorie} {N} programme(s) immobilier(s) {COMMUNE}'
    """
    base = "Degrevement TFPB"
    if annee:
        base += f" {annee}"

    if not categorie:
        return base

    # Compter les programmes uniques
    programmes = [p.strip() for p in numero_programme.split(",") if p.strip()] if numero_programme else []
    nb_prog = len(programmes)

    if nb_prog == 0:
        prog_text = ""
    elif nb_prog == 1:
        prog_text = "1 programme immobilier"
    else:
        prog_text = f"{nb_prog} programmes immobiliers"

    parts = [base, "-", categorie]
    if prog_text:
        parts.append(prog_text)
    if commune:
        parts.append(commune)

    return " ".join(parts)
