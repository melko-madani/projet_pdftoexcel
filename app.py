"""Interface Streamlit pour l'extraction de tableaux PDF vers Excel."""

import io
import tempfile
from pathlib import Path

import streamlit as st

from core.excel_writer import write_excel
from core.parser import process_tables
from core.scanner import format_scan_report, scan_pdf

# =====================
# Configuration
# =====================
st.set_page_config(
    page_title="Extraction Courrier PDF",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Extracteur PDF → Excel")
st.markdown(
    "Uploadez vos courriers PDF (dossiers fiscaux TFPB) pour extraire "
    "les tableaux dans un fichier Excel formaté."
)

# =====================
# Paramètres
# =====================
with st.sidebar:
    st.header("⚙️ Paramètres")
    min_cols = st.number_input(
        "Nombre minimum de colonnes",
        min_value=1,
        max_value=20,
        value=8,
        help="Nombre minimum de colonnes pour qu'un tableau soit considéré valide.",
    )
    show_report = st.checkbox("Afficher le rapport de scan", value=True)

# =====================
# Upload
# =====================
uploaded_files = st.file_uploader(
    "Choisir un ou plusieurs fichiers PDF",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} fichier(s) sélectionné(s)")

    if st.button("🚀 Extraire les données", type="primary"):
        progress = st.progress(0)
        status = st.empty()

        all_results = []  # (filename, scan_result, datasets, excel_bytes)

        for i, pdf_file in enumerate(uploaded_files):
            status.text(f"Traitement de : {pdf_file.name}...")

            try:
                # Sauvegarder temporairement le fichier uploadé
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=".pdf"
                ) as tmp:
                    tmp.write(pdf_file.read())
                    tmp_path = Path(tmp.name)

                # Scan avec le core du coworker
                scan_result = scan_pdf(tmp_path, min_cols=min_cols)

                if not scan_result.tables:
                    st.warning(f"**{pdf_file.name}** : aucun tableau détecté.")
                    all_results.append((pdf_file.name, scan_result, [], None))
                    continue

                # Traitement avec le core du coworker
                datasets = process_tables(scan_result.tables)

                if not datasets:
                    st.warning(
                        f"**{pdf_file.name}** : aucune donnée exploitable après nettoyage."
                    )
                    all_results.append((pdf_file.name, scan_result, [], None))
                    continue

                # Export Excel en mémoire
                output_path = tmp_path.with_suffix(".xlsx")
                write_excel(datasets, output_path)

                excel_bytes = output_path.read_bytes()
                all_results.append(
                    (pdf_file.name, scan_result, datasets, excel_bytes)
                )

                # Nettoyage
                tmp_path.unlink(missing_ok=True)
                output_path.unlink(missing_ok=True)

            except Exception as e:
                st.error(f"Erreur sur **{pdf_file.name}** : {e}")

            progress.progress((i + 1) / len(uploaded_files))

        # =====================
        # Résultats
        # =====================
        status.success(
            f"Extraction terminée ! "
            f"{sum(1 for _, _, ds, _ in all_results if ds)} / "
            f"{len(uploaded_files)} PDF traité(s) avec succès."
        )

        for filename, scan_result, datasets, excel_bytes in all_results:
            st.divider()
            st.subheader(f"📁 {filename}")

            # Rapport de scan
            if show_report:
                with st.expander("Rapport de scan", expanded=False):
                    report = format_scan_report(scan_result)
                    st.code(report, language=None)

            if not datasets:
                continue

            # Aperçu des données par dataset
            for ds in datasets:
                with st.expander(
                    f"📊 {ds.name} — {len(ds.data_rows)} lignes, "
                    f"{len(ds.total_rows)} total(s)",
                    expanded=True,
                ):
                    # Convertir en tableau pour affichage
                    import pandas as pd

                    if ds.data_rows:
                        df = pd.DataFrame(ds.data_rows, columns=ds.headers)
                        st.dataframe(df, use_container_width=True)

                    if ds.total_rows:
                        st.caption("Lignes de total :")
                        df_totals = pd.DataFrame(
                            ds.total_rows, columns=ds.headers
                        )
                        st.dataframe(df_totals, use_container_width=True)

            # Bouton téléchargement
            if excel_bytes:
                excel_name = Path(filename).stem + ".xlsx"
                st.download_button(
                    label=f"📥 Télécharger {excel_name}",
                    data=excel_bytes,
                    file_name=excel_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        # Si plusieurs PDFs, proposer un téléchargement groupé (ZIP)
        if len([r for r in all_results if r[3]]) > 1:
            import zipfile

            st.divider()
            st.subheader("📦 Téléchargement groupé")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename, _, datasets, excel_bytes in all_results:
                    if excel_bytes:
                        excel_name = Path(filename).stem + ".xlsx"
                        zf.writestr(excel_name, excel_bytes)
            zip_buffer.seek(0)

            st.download_button(
                label="📥 Télécharger tous les fichiers Excel (ZIP)",
                data=zip_buffer,
                file_name="extraction_courriers.zip",
                mime="application/zip",
                type="primary",
            )
