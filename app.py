"""Interface Streamlit pour PDF Table Extractor.

Refonte UX basee sur les 10 heuristiques de Nielsen.
"""

import io
import json
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from core.extract_columns import extract_target_columns_from_pdf
from core.extract_metadata_json import extract_04_from_pdf
from core.pipeline import (
    DemandeDossier,
    DemandeResult,
    build_output_zip,
    extract_zip_to_dossiers,
    process_demande,
    process_zip,
)

DEFAULT_MIN_COLS = 8


# --- Helpers ---

def _readable_demande_name(result: DemandeResult) -> str:
    """Genere un nom lisible pour une demande."""
    name = f"Demande N\u00b0{result.prefix}"
    if result.metadata and result.metadata.libelle:
        short = result.metadata.libelle[:50]
        if len(result.metadata.libelle) > 50:
            short += "..."
        name += f" \u2014 {short}"
    return name


def _check_zip_structure(zip_data: bytes) -> tuple[bool, int]:
    """Verifie la structure du ZIP. Retourne (has_standard_structure, pdf_count)."""
    has_mails = False
    has_proof = False
    pdf_count = 0

    try:
        with zipfile.ZipFile(io.BytesIO(zip_data)) as zf:
            for entry in zf.namelist():
                if entry.endswith("/") or "__MACOSX" in entry:
                    continue
                if entry.lower().endswith(".pdf"):
                    pdf_count += 1
                parts = Path(entry).parts
                if len(parts) >= 2:
                    if parts[0].lower() == "mails" or (len(parts) >= 3 and parts[1].lower() == "mails"):
                        has_mails = True
                    if parts[0].lower() == "proof" or (len(parts) >= 3 and parts[1].lower() == "proof"):
                        has_proof = True
    except zipfile.BadZipFile:
        return False, 0

    return (has_mails and has_proof), pdf_count


def _file_size_mb(uploaded_file) -> float:
    """Retourne la taille du fichier en MB."""
    return len(uploaded_file.getvalue()) / (1024 * 1024)


# --- Main ---

def main():
    st.set_page_config(
        page_title="PDF Table Extractor",
        page_icon="\U0001f4c4",
    )

    # Init session state
    if "results" not in st.session_state:
        st.session_state.results = None
    if "output_data" not in st.session_state:
        st.session_state.output_data = None
    if "processing_done" not in st.session_state:
        st.session_state.processing_done = False
    if "errors" not in st.session_state:
        st.session_state.errors = []
    if "json_03" not in st.session_state:
        st.session_state.json_03 = None
    if "json_04" not in st.session_state:
        st.session_state.json_04 = None

    st.title("\U0001f4c4 PDF Table Extractor")

    # --- Sidebar ---
    with st.sidebar:
        st.header("\U0001f4c4 PDF Table Extractor")
        st.divider()

        # Bouton principal change selon l'etape
        if st.session_state.processing_done:
            if st.button(
                "\U0001f504 Recommencer",
                type="primary",
                use_container_width=True,
                help="Reinitialiser et deposer de nouveaux fichiers",
            ):
                st.session_state.results = None
                st.session_state.output_data = None
                st.session_state.processing_done = False
                st.session_state.errors = []
                st.session_state.json_03 = None
                st.session_state.json_04 = None
                st.rerun()

        with st.expander("Options avancees"):
            min_cols = st.slider(
                "Colonnes minimum par tableau",
                min_value=3,
                max_value=25,
                value=DEFAULT_MIN_COLS,
                help="Les vrais tableaux fiscaux ont 20+ colonnes. "
                     "Augmentez pour filtrer les faux tableaux.",
            )

    # --- Etape 3 : Resultats (si traitement termine) ---
    if st.session_state.processing_done:
        _render_results()
        return

    # --- Etape 1 : Upload ---

    # Aide
    with st.expander("Comment utiliser cet outil ?"):
        st.markdown(
            "**Structure attendue du fichier ZIP :**\n"
            "- `mails/` \u2014 contient les courriers PDF (ex: `336-Courrier_TFPB_2024.pdf`)\n"
            "- `proof/` \u2014 contient les AR et preuves de depot "
            "(ex: `336-AR_n_xxx.pdf`, `336-Preuve_de_Depot.pdf`)\n\n"
            "**Convention de nommage :** le prefixe numerique (avant le `-`) "
            "relie les fichiers d'une meme demande.\n\n"
            "**Fichiers acceptes :** `.zip` (lot de demandes) ou `.pdf` (fichiers individuels)\n\n"
            "**Resultat :** un dossier ZIP structure par demande, contenant les PDF originaux, "
            "les tableaux extraits en Excel, et les informations du dossier."
        )

    uploaded_files = st.file_uploader(
        "Deposez vos courriers ici",
        type=["pdf", "zip"],
        accept_multiple_files=True,
        help="Glissez un fichier ZIP (lot de demandes) ou des fichiers PDF individuels.",
    )

    if not uploaded_files:
        return

    # Resume des fichiers deposes
    zip_files = [f for f in uploaded_files if f.name.lower().endswith(".zip")]
    pdf_files = [f for f in uploaded_files if f.name.lower().endswith(".pdf")]

    if zip_files:
        st.caption(f"\U0001f4e6 {len(zip_files)} archive(s) ZIP")
    if pdf_files:
        st.caption(f"\U0001f4c4 {pdf_files_label(pdf_files)}")

    # Validations (heuristique 5 : prevention des erreurs)
    warnings = []

    if zip_files:
        zip_data = zip_files[0].getvalue()
        has_structure, pdf_count = _check_zip_structure(zip_data)

        if pdf_count == 0:
            st.error(
                "Le fichier ZIP ne contient aucun PDF. "
                "Verifiez que vos courriers sont dans le dossier `mails/` "
                "et vos justificatifs dans `proof/`."
            )
            return

        if not has_structure:
            warnings.append(
                "Structure atypique detectee \u2014 les fichiers seront analyses individuellement. "
                "Structure attendue : `mails/` et `proof/`."
            )

    for f in uploaded_files:
        size = _file_size_mb(f)
        if size > 50:
            warnings.append(f"`{f.name}` fait {size:.0f} MB \u2014 le traitement peut prendre du temps.")

    for w in warnings:
        st.warning(w)

    # Bouton lancer (dans la page, pas le sidebar, pour le rendre visible)
    st.divider()

    if st.button(
        "Lancer l'extraction",
        type="primary",
        use_container_width=True,
        help="Analyser les fichiers deposes et generer le dossier structure",
    ):
        min_cols_val = min_cols if "min_cols" in dir() else DEFAULT_MIN_COLS
        _run_extraction(uploaded_files, min_cols_val)


def pdf_files_label(pdf_files: list) -> str:
    """Label pour les fichiers PDF uploades."""
    if len(pdf_files) == 1:
        return f"1 courrier : {pdf_files[0].name}"
    return f"{len(pdf_files)} courriers deposes"


def _run_extraction(uploaded_files: list, min_cols: int):
    """Execute l'extraction et stocke les resultats en session_state."""
    zip_files = [f for f in uploaded_files if f.name.lower().endswith(".zip")]
    pdf_files = [f for f in uploaded_files if f.name.lower().endswith(".pdf")]

    errors = []
    results = []

    if zip_files:
        # Mode ZIP
        zip_data = zip_files[0].getvalue()
        status = st.empty()
        progress = st.progress(0, text="Analyse du lot de demandes...")

        def on_progress(current, total, prefix):
            if total > 0 and prefix:
                pct = current / total
                progress.progress(pct, text=f"Traitement de la demande {prefix} ({current + 1}/{total})...")

        results = process_zip(zip_data, min_cols=min_cols, on_progress=on_progress)

        progress.progress(1.0, text="Extraction terminee.")

        # Collecter les erreurs
        for r in results:
            if r.error:
                errors.append(
                    f"Le fichier du dossier {r.prefix} n'a pas pu etre lu. "
                    f"Il sera ignore. Verifiez le fichier et re-deposez-le."
                )
            if not r.datasets and r.source_pdfs.get("courrier"):
                errors.append(
                    f"Le courrier N\u00b0{r.prefix} ne contient aucun tableau fiscal. "
                    f"Seul le fichier d'informations sera genere."
                )

    else:
        # Mode PDF
        progress = st.progress(0, text="Analyse des courriers...")
        total = len(pdf_files)

        for idx, f in enumerate(pdf_files):
            progress.progress(
                idx / total,
                text=f"Traitement du fichier {idx + 1}/{total} : {f.name}...",
            )

            prefix = Path(f.name).stem.split("-")[0]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            try:
                tmp.write(f.getbuffer())
                tmp.close()

                dossier = DemandeDossier(
                    prefix=prefix,
                    courrier_pdf=Path(tmp.name).read_bytes(),
                    courrier_filename=f.name,
                )
                result = process_demande(dossier, min_cols=min_cols)
                results.append(result)

                if result.error:
                    errors.append(
                        f"Le fichier `{f.name}` n'a pas pu etre lu. "
                        f"Verifiez le fichier et re-deposez-le."
                    )
                if not result.datasets:
                    errors.append(
                        f"Le courrier N\u00b0{prefix} ne contient aucun tableau fiscal. "
                        f"Seul le fichier d'informations sera genere."
                    )

            except Exception as e:
                errors.append(f"Erreur sur `{f.name}` : {e}. Les autres fichiers sont traites normalement.")
            finally:
                Path(tmp.name).unlink(missing_ok=True)

        progress.progress(1.0, text="Extraction terminee.")

    # --- Extraction ciblée : 03.json ---
    json_data = {
        "colonnes_extraites": [
            "Référence de l'avis",
            "Adresse",
            "Montant de dégrèvement",
        ],
        "nombre_fichiers": len(pdf_files) if not zip_files else 0,
        "fichiers": [],
        "nombre_total_lignes": 0,
    }

    if not zip_files:
        for f in pdf_files:
            f.seek(0)
            rows = extract_target_columns_from_pdf(f.getvalue(), f.name)
            json_data["fichiers"].append({
                "nom_fichier": f.name,
                "nombre_lignes": len(rows),
                "donnees": rows,
            })
            json_data["nombre_total_lignes"] += len(rows)

    json_bytes = json.dumps(json_data, ensure_ascii=False, indent=2).encode("utf-8")
    st.session_state.json_03 = json_bytes

    # --- Extraction ciblée : 04.json ---
    json_04_data = {
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
        "nombre_fichiers": len(pdf_files) if not zip_files else 0,
        "fichiers": [],
    }

    if not zip_files:
        for f in pdf_files:
            f.seek(0)
            data_04 = extract_04_from_pdf(f.getvalue(), f.name)
            json_04_data["fichiers"].append({
                "nom_fichier": f.name,
                "donnees": data_04,
            })

    json_04_bytes = json.dumps(json_04_data, ensure_ascii=False, indent=2).encode("utf-8")
    st.session_state.json_04 = json_04_bytes

    # Stocker les resultats
    if results:
        st.session_state.results = results
        st.session_state.output_data = build_output_zip(results)
        st.session_state.errors = errors
    else:
        st.session_state.results = []
        st.session_state.errors = errors

    st.session_state.processing_done = True
    st.rerun()


def _render_results():
    """Affiche les resultats apres traitement."""
    results = st.session_state.results
    output_data = st.session_state.output_data
    errors = st.session_state.errors

    if not results:
        st.error("Aucune demande trouvee. Verifiez le contenu de votre fichier et re-deposez-le.")
        return

    # Succes
    total_rows = sum(r.row_count for r in results)
    total_tables = sum(r.table_count for r in results)

    st.success(
        f"{len(results)} demande(s) traitee(s) \u2014 "
        f"{total_tables} annexe(s) fiscale(s) extraite(s), "
        f"{total_rows} lignes au total."
    )

    # Erreurs / avertissements
    for err in errors:
        st.warning(err)

    st.divider()

    # Liste des demandes
    st.markdown("**Dossiers generes :**")
    for r in results:
        name = _readable_demande_name(r)
        if r.datasets:
            st.markdown(f"- \U0001f4c2 {name}")
        elif r.error:
            st.markdown(f"- \u274c {name}")
        else:
            st.markdown(f"- \U0001f4c2 {name} *(pas de tableau)*")

    st.divider()

    # Apercu optionnel
    with st.expander("Voir le detail par demande"):
        for r in results:
            st.markdown(f"**Demande N\u00b0{r.prefix}**")

            cols = st.columns(4)
            cols[0].caption(
                f"Courrier : {'oui' if r.source_pdfs.get('courrier') else 'non fourni'}"
            )
            cols[1].caption(
                f"AR : {'oui' if r.source_pdfs.get('ar') else 'non fourni'}"
            )
            cols[2].caption(
                f"Depot : {'oui' if r.source_pdfs.get('depot') else 'non fourni'}"
            )
            cols[3].caption(f"Lignes : {r.row_count}")

            if r.datasets:
                for ds in r.datasets:
                    preview_rows = ds.data_rows[:5]
                    if preview_rows:
                        preview_data = []
                        for row in preview_rows:
                            row_dict = {}
                            for i, header in enumerate(ds.headers):
                                row_dict[header] = row[i] if i < len(row) else None
                            preview_data.append(row_dict)
                        st.dataframe(preview_data, use_container_width=True)
                        if len(ds.data_rows) > 5:
                            st.caption(f"Apercu : 5 lignes sur {len(ds.data_rows)}")

            st.divider()

    # Telechargement
    if output_data:
        st.download_button(
            label="Telecharger le dossier structure (.zip)",
            data=output_data,
            file_name="resultats_extraction.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
            help="Archive ZIP contenant un dossier par demande avec les PDF originaux et les Excel generes.",
        )

    # Telechargement 03.json
    json_03 = st.session_state.get("json_03")
    if json_03:
        st.download_button(
            label="Telecharger 03.json (Ref. avis, Adresse, Montant)",
            data=json_03,
            file_name="03.json",
            mime="application/json",
            use_container_width=True,
        )

    # Telechargement 04.json
    json_04 = st.session_state.get("json_04")
    if json_04:
        st.download_button(
            label="Telecharger 04.json (Sous-cat, Montants, Entreprise, TVA)",
            data=json_04,
            file_name="04.json",
            mime="application/json",
            use_container_width=True,
        )

    # Footer
    st.divider()
    st.caption(
        "[PDF Table Extractor](https://github.com/melko-madani/projet_pdftoexcel) "
        "\u2014 Extraction de tableaux fiscaux"
    )


if __name__ == "__main__":
    main()
