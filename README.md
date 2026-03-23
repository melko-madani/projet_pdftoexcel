# PDF Table Extractor

Extracteur de tableaux PDF vers Excel, conçu pour les dossiers de dégrèvement fiscal (TFPB).

## Installation

```bash
# Créer un environnement virtuel
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# Installer les dépendances
pip install -r requirements.txt
```

## Utilisation

### Analyser un PDF (rapport de détection)

```bash
python extractor.py scan mon_fichier.pdf
```

Affiche un rapport avec le nombre de tableaux détectés par page, le nombre de lignes/colonnes, et un aperçu des 3 premiers headers.

### Extraire les tableaux en Excel

```bash
python extractor.py extract mon_fichier.pdf
python extractor.py extract mon_fichier.pdf --output resultat.xlsx
```

### Traitement batch d'un dossier

```bash
python extractor.py batch ./dossier_pdfs/
python extractor.py batch ./dossier_pdfs/ --output-dir ./resultats/
```

### Interface web (Streamlit)

```bash
streamlit run app.py
```

Ouvre le navigateur sur `localhost:8501` avec une interface drag & drop pour uploader des PDFs, visualiser les tableaux détectés et télécharger les fichiers Excel.

### Options CLI

- `--output` / `-o` : chemin du fichier Excel de sortie
- `--output-dir` / `-o` : dossier de sortie (mode batch)
- `--min-cols` : nombre minimum de colonnes pour un tableau valide (défaut: 8)
- `--verbose` / `-v` : logs détaillés

## Fonctionnalités

- Détection automatique des tableaux dans les PDF multi-pages
- Consolidation des tableaux de même structure sur des pages consécutives
- Dédupplication des headers répétés en haut de chaque page
- Nettoyage des données : valeurs monétaires (`542,78 €` → `542.78`), dates, mois de vacance
- Séparation des lignes de total/sous-total
- Formatage Excel professionnel : headers colorés, lignes alternées, formats monétaires/dates, filtres auto

## Stack technique

- Python 3.11+
- `pdfplumber` — lecture PDF et détection de tableaux
- `openpyxl` — création Excel avec formatage
- `typer` — interface CLI
- `streamlit` — interface web
