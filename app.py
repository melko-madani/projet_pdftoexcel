"""Interface Streamlit pour l'extraction de tableaux PDF vers Excel."""

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from core.excel_writer import write_excel
from core.parser import process_tables
from core.scanner import scan_pdf

# =====================
# Configuration
# =====================
LOGO_PATH = Path(__file__).parent / "image.jpg"

st.set_page_config(
    page_title="Extraction Courrier PDF",
    page_icon="image.jpg",
)

# En-tête avec logo
col_logo, col_title = st.columns([1, 4])
with col_logo:
    st.image(str(LOGO_PATH), width=120)
with col_title:
    st.title("Extracteur PDF → Excel")
    st.caption("Outil interne - Extraction des courriers fiscaux")

st.divider()

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

        all_excel = []  # (filename, excel_bytes)

        for i, pdf_file in enumerate(uploaded_files):
            status.text(f"Traitement de : {pdf_file.name}...")

            try:
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=".pdf"
                ) as tmp:
                    tmp.write(pdf_file.read())
                    tmp_path = Path(tmp.name)

                scan_result = scan_pdf(tmp_path, min_cols=8)

                if not scan_result.tables:
                    st.warning(f"**{pdf_file.name}** : aucun tableau détecté.")
                    continue

                datasets = process_tables(scan_result.tables)

                if not datasets:
                    st.warning(f"**{pdf_file.name}** : aucune donnée exploitable.")
                    continue

                output_path = tmp_path.with_suffix(".xlsx")
                write_excel(datasets, output_path)

                excel_bytes = output_path.read_bytes()
                all_excel.append((pdf_file.name, excel_bytes))

                tmp_path.unlink(missing_ok=True)
                output_path.unlink(missing_ok=True)

            except Exception as e:
                st.error(f"Erreur sur **{pdf_file.name}** : {e}")

            progress.progress((i + 1) / len(uploaded_files))

        # =====================
        # Résultats
        # =====================
        status.success(
            f"Extraction terminée ! {len(all_excel)} / "
            f"{len(uploaded_files)} PDF traité(s) avec succès."
        )

        # Téléchargements
        for filename, excel_bytes in all_excel:
            excel_name = Path(filename).stem + ".xlsx"
            st.download_button(
                label=f"📥 Télécharger {excel_name}",
                data=excel_bytes,
                file_name=excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if len(all_excel) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename, excel_bytes in all_excel:
                    excel_name = Path(filename).stem + ".xlsx"
                    zf.writestr(excel_name, excel_bytes)
            zip_buffer.seek(0)

            st.download_button(
                label="📥 Télécharger tout (ZIP)",
                data=zip_buffer,
                file_name="extraction_courriers.zip",
                mime="application/zip",
                type="primary",
            )
