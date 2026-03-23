"""Toutes les regex centralisees pour l'extraction de metadonnees fiscales."""

import re

# --- Patterns courrier ---

PATTERN_NUM_DEMANDE = re.compile(r"demande\s*N[°o]\s*(\d+)", re.IGNORECASE)
PATTERN_ANNEE_FISCALE = re.compile(r"TFPB\s+(20\d{2})")
PATTERN_OBJET = re.compile(
    r"Objet\s*:\s*(.+?)(?=Adresses\s+concern[ée]es)", re.DOTALL | re.IGNORECASE
)
PATTERN_ADRESSES = re.compile(
    r"Adresses\s+concern[ée]es\s*:\s*(.+?)(?=Motif)", re.DOTALL | re.IGNORECASE
)
PATTERN_MOTIF = re.compile(
    r"Motif\s+de\s+la\s+vacance\s*:\s*(.+?)(?=Pi[èe]ces|$)", re.DOTALL | re.IGNORECASE
)
PATTERN_RESPONSABLE = re.compile(
    r"Affaire\s+suivie\s+par\s*:\s*([^,(]+)", re.IGNORECASE
)
PATTERN_EMAIL = re.compile(r"[\w.-]+@[\w.-]+\.\w+")
PATTERN_DATE_COURRIER = re.compile(
    r"Fait\s+[àa]\s+.+?,\s*le\s+\w+\s+(\d{1,2}\s+\w+\s+\d{4})", re.IGNORECASE
)
PATTERN_DATE_COURRIER_ALT = re.compile(
    r"le\s+lundi\s+(\d{1,2}\s+\w+\s+\d{4})", re.IGNORECASE
)
PATTERN_DATE_LIMITE = re.compile(
    r"au\s+plus\s+tard\s+le\s+(\d{1,2}\s+\w+\s+\d{4})", re.IGNORECASE
)
PATTERN_REF_AVIS = re.compile(
    r"r[ée]f[ée]rence\s*:\s*([\d\s]+\d)", re.IGNORECASE
)
PATTERN_MONTANT_TOTAL = re.compile(
    r"montant\s+total\s+de\s*([\d\s.,]+)\s*(?:euros?|€)", re.IGNORECASE
)
PATTERN_MONTANT_DEGREVEMENT = re.compile(
    r"d[ée]gr[èe]vement\s+de\s*([\d\s.,]+)\s*(?:euros?|€)", re.IGNORECASE
)
PATTERN_MONTANT_COTISATIONS = re.compile(
    r"montant\s+[ée]gal\s+[àa]\s*([\d\s.,]+)\s*(?:euros?|€)", re.IGNORECASE
)
PATTERN_FRAIS_GESTION = re.compile(
    r"frais\s+de\s+gestion[^€]*?([\d\s.,]+)\s*(?:euros?|€)", re.IGNORECASE
)
PATTERN_NB_LOGEMENTS = re.compile(r"(\d+)\s*logements", re.IGNORECASE)
PATTERN_ARTICLES_CGI = re.compile(
    r"articles?\s+([\d\w-]+)\s+du\s+CGI", re.IGNORECASE
)
PATTERN_CODE_POSTAL = re.compile(r"(\d{5})\s+([A-Z]{3,})")
PATTERN_DESTINATAIRE = re.compile(
    r"[AÀ]\s+l['\u2019]attention\s+de\s+(?:Monsieur|Madame|M\.|Mme)?\s*(\w+(?:\s+\w+)?)",
    re.IGNORECASE,
)
PATTERN_ENTREPRISE_MANDATAIRE = re.compile(
    r"mandataire\s+de\s+(?:la\s+)?(.+?)(?:\n|,|\.)", re.IGNORECASE
)
PATTERN_ENTREPRISE_COMPTE = re.compile(
    r"pour\s+le\s+compte\s+de\s+(?:la\s+)?(.+?)(?:\n|,|\.)", re.IGNORECASE
)

# --- Patterns AR et depot ---

PATTERN_NUM_LR = re.compile(r"(\d{15})")
PATTERN_DATE_DDMMYYYY = re.compile(r"(\d{2}/\d{2}/\d{4})")
PATTERN_PRESENTATION_AR = re.compile(
    r"(?:Pr[ée]sent[ée]e|avis[ée]e)\s+(?:le\s+)?(\d{1,2}/\d{1,2}/\d{4})",
    re.IGNORECASE,
)
PATTERN_DISTRIBUTION_AR = re.compile(
    r"Distribu[ée]e\s+le\s+(\d{1,2}/\d{1,2}/\d{4})", re.IGNORECASE
)
PATTERN_RECEPTIONNAIRE = re.compile(
    r"(?:Nom\s+du\s+destinataire|mandataire)\s*:\s*(.+?)(?:\n|$)", re.IGNORECASE
)
