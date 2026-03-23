"""Interface Streamlit pour PDF Table Extractor."""

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from core.pipeline import (
    DemandeDossier,
    build_output_zip,
    process_demande,
    process_zip,
)

DEFAULT_MIN_COLS = 8

# =====================
# Configuration
# =====================
LOGO_PATH = Path(__file__).parent / "image.jpg"

st.set_page_config(
    page_title="Extraction Courrier PDF",
    page_icon="📄",
)

# En-tête avec logo
col_logo, col_title = st.columns([1, 4])
with col_logo:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=120)
with col_title:
    st.title("Extracteur PDF → Excel")
    st.caption("Outil interne - Extraction des courriers fiscaux")

st.divider()

# =====================
# Upload
# =====================
uploaded_files = st.file_uploader(
    "Choisir un ou plusieurs fichiers PDF / ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True,
)

if uploaded_files:
    zip_files = [f for f in uploaded_files if f.name.lower().endswith(".zip")]
    pdf_files = [f for f in uploaded_files if f.name.lower().endswith(".pdf")]

    parts = []
    if zip_files:
        parts.append(f"{len(zip_files)} archive(s) ZIP")
    if pdf_files:
        parts.append(f"{len(pdf_files)} courrier(s) PDF")
    st.info(f"{len(uploaded_files)} fichier(s) sélectionné(s) — {', '.join(parts)}")

    if st.button("🚀 Extraire les données", type="primary", use_container_width=True):
        errors = []
        results = []

        if zip_files:
            # Mode ZIP
            zip_data = zip_files[0].getvalue()
            progress = st.progress(0, text="Analyse du lot de demandes...")

            def on_progress(current, total, prefix):
                if total > 0 and prefix:
                    pct = current / total
                    progress.progress(
                        pct,
                        text=f"Traitement de la demande {prefix} ({current + 1}/{total})...",
                    )

            results = process_zip(
                zip_data, min_cols=DEFAULT_MIN_COLS, on_progress=on_progress
            )
            progress.progress(1.0, text="Extraction terminée.")

            for r in results:
                if r.error:
                    errors.append(f"Dossier {r.prefix} : fichier illisible, ignoré.")
                if not r.datasets and r.source_pdfs.get("courrier"):
                    errors.append(
                        f"Courrier N°{r.prefix} : aucun tableau fiscal détecté."
                    )
        else:
            # Mode PDF
            progress = st.progress(0, text="Analyse des courriers...")
            total = len(pdf_files)

            for idx, f in enumerate(pdf_files):
                progress.progress(
                    idx / total,
                    text=f"Traitement {idx + 1}/{total} : {f.name}...",
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
                    result = process_demande(dossier, min_cols=DEFAULT_MIN_COLS)
                    results.append(result)

                    if result.error:
                        errors.append(f"`{f.name}` : fichier illisible.")
                    if not result.datasets:
                        errors.append(
                            f"`{f.name}` : aucun tableau fiscal détecté."
                        )
                except Exception as e:
                    errors.append(f"Erreur sur `{f.name}` : {e}")
                finally:
                    Path(tmp.name).unlink(missing_ok=True)

            progress.progress(1.0, text="Extraction terminée.")

        # =====================
        # Résultats
        # =====================
        if results:
            total_rows = sum(r.row_count for r in results)
            total_tables = sum(r.table_count for r in results)

            st.success(
                f"{len(results)} demande(s) traitée(s) — "
                f"{total_tables} annexe(s) extraite(s), "
                f"{total_rows} lignes au total."
            )

            for err in errors:
                st.warning(err)

            output_data = build_output_zip(results)

            if output_data:
                st.download_button(
                    label="📥 Télécharger le dossier structuré (.zip)",
                    data=output_data,
                    file_name="resultats_extraction.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                )
        else:
            st.error("Aucune demande trouvée. Vérifiez vos fichiers.")
            for err in errors:
                st.warning(err)
