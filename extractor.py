"""CLI pour l'extraction de tableaux PDF vers Excel."""

import logging
import sys
from pathlib import Path

import typer

from core.excel_writer import write_excel
from core.metadata import (
    extract_metadata,
    format_metadata_report,
    metadata_to_rows,
    process_dossier,
)
from core.parser import process_tables
from core.pipeline import build_output_zip, process_zip
from core.scanner import format_scan_report, scan_pdf

app = typer.Typer(
    name="pdf-table-extractor",
    help="Extracteur de tableaux PDF vers Excel pour dossiers fiscaux TFPB.",
    add_completion=False,
)


def setup_logging(verbose: bool = False) -> None:
    """Configure le logging."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stderr,
    )


@app.command()
def scan(
    pdf_path: Path = typer.Argument(..., help="Chemin vers le fichier PDF à analyser."),
    min_cols: int = typer.Option(8, "--min-cols", help="Nombre minimum de colonnes pour un tableau valide."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Logs détaillés."),
) -> None:
    """Analyse un PDF et affiche un rapport de détection des tableaux."""
    setup_logging(verbose)

    if not pdf_path.exists():
        typer.echo(f"Erreur : fichier introuvable — {pdf_path}", err=True)
        raise typer.Exit(1)

    result = scan_pdf(pdf_path, min_cols=min_cols)
    report = format_scan_report(result)
    typer.echo(report)


@app.command()
def extract(
    pdf_path: Path = typer.Argument(..., help="Chemin vers le fichier PDF."),
    output: Path = typer.Option(
        None, "--output", "-o",
        help="Chemin du fichier Excel de sortie. Par défaut : même nom que le PDF avec .xlsx.",
    ),
    min_cols: int = typer.Option(8, "--min-cols", help="Nombre minimum de colonnes pour un tableau valide."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Logs détaillés."),
) -> None:
    """Extrait les tableaux d'un PDF et les exporte en fichier Excel formaté."""
    setup_logging(verbose)
    logger = logging.getLogger(__name__)

    if not pdf_path.exists():
        typer.echo(f"Erreur : fichier introuvable — {pdf_path}", err=True)
        raise typer.Exit(1)

    # Chemin de sortie par défaut
    if output is None:
        output = pdf_path.with_suffix(".xlsx")

    # Scan
    typer.echo(f"Scan de {pdf_path.name}...")
    result = scan_pdf(pdf_path, min_cols=min_cols)

    if not result.tables:
        typer.echo("Aucun tableau détecté dans ce PDF.")
        raise typer.Exit(0)

    typer.echo(f"{len(result.tables)} tableau(x) détecté(s).")

    # Traitement
    typer.echo("Traitement et nettoyage des données...")
    datasets = process_tables(result.tables)

    if not datasets:
        typer.echo("Aucune donnée exploitable après nettoyage.")
        raise typer.Exit(0)

    # Export Excel
    typer.echo(f"Export vers {output}...")
    write_excel(datasets, output)

    # Résumé
    total_rows = sum(len(ds.data_rows) + len(ds.total_rows) for ds in datasets)
    typer.echo(f"\nTerminé ! {total_rows} lignes exportées dans {len(datasets)} feuille(s).")
    for ds in datasets:
        if ds.name.startswith("Annexe"):
            pages_str = ", ".join(str(p) for p in ds.source_pages)
            typer.echo(f'INFO: La feuille "{ds.name}" contient les tableaux consolidés des pages {pages_str}')
    typer.echo(f"Fichier : {output.resolve()}")


@app.command()
def batch(
    input_dir: Path = typer.Argument(..., help="Dossier contenant les fichiers PDF."),
    output_dir: Path = typer.Option(
        None, "--output-dir", "-o",
        help="Dossier de sortie pour les fichiers Excel. Par défaut : même dossier que les PDFs.",
    ),
    min_cols: int = typer.Option(8, "--min-cols", help="Nombre minimum de colonnes pour un tableau valide."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Logs détaillés."),
) -> None:
    """Traite tous les fichiers PDF d'un dossier et exporte les tableaux en Excel."""
    setup_logging(verbose)

    if not input_dir.exists() or not input_dir.is_dir():
        typer.echo(f"Erreur : dossier introuvable — {input_dir}", err=True)
        raise typer.Exit(1)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        typer.echo(f"Aucun fichier PDF trouvé dans {input_dir}")
        raise typer.Exit(0)

    if output_dir is None:
        output_dir = input_dir

    output_dir.mkdir(parents=True, exist_ok=True)

    typer.echo(f"Traitement batch : {len(pdf_files)} fichier(s) PDF trouvé(s)")
    typer.echo(f"{'='*60}")

    success_count = 0
    error_count = 0
    skip_count = 0

    for pdf_file in pdf_files:
        typer.echo(f"\n--- {pdf_file.name} ---")
        output_path = output_dir / pdf_file.with_suffix(".xlsx").name

        try:
            result = scan_pdf(pdf_file, min_cols=min_cols)

            if not result.tables:
                typer.echo("  Aucun tableau détecté, fichier ignoré.")
                skip_count += 1
                continue

            datasets = process_tables(result.tables)
            if not datasets:
                typer.echo("  Aucune donnée exploitable, fichier ignoré.")
                skip_count += 1
                continue

            write_excel(datasets, output_path)
            total_rows = sum(len(ds.data_rows) + len(ds.total_rows) for ds in datasets)
            typer.echo(
                f"  OK — {total_rows} lignes dans {len(datasets)} feuille(s) → {output_path.name}"
            )
            success_count += 1

        except Exception as e:
            typer.echo(f"  ERREUR : {e}", err=True)
            error_count += 1

    typer.echo(f"\n{'='*60}")
    typer.echo(
        f"Résumé : {success_count} réussi(s), {skip_count} ignoré(s), {error_count} erreur(s)"
    )


@app.command()
def metadata(
    pdf_path: Path = typer.Argument(..., help="Chemin vers le fichier PDF."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Logs détaillés."),
) -> None:
    """Extrait et affiche les métadonnées d'un fichier PDF."""
    setup_logging(verbose)

    if not pdf_path.exists():
        typer.echo(f"Erreur : fichier introuvable — {pdf_path}", err=True)
        raise typer.Exit(1)

    pdf_type, meta = extract_metadata(pdf_path)
    typer.echo(f"Type détecté : {pdf_type}")
    typer.echo(f"Fichier : {pdf_path.name}")
    typer.echo("")
    for key, value in meta.items():
        if key in ("type", "filename"):
            continue
        display = value if value else "(non trouvé)"
        typer.echo(f"  {key:25s} : {display}")


@app.command()
def preprocess(
    input_dir: Path = typer.Argument(..., help="Dossier contenant les fichiers PDF du dossier de demande."),
    output: Path = typer.Option(
        None, "--output", "-o",
        help="Chemin du fichier Excel de sortie.",
    ),
    min_cols: int = typer.Option(8, "--min-cols", help="Nombre minimum de colonnes pour un tableau valide."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Logs détaillés."),
) -> None:
    """Pré-processing d'un dossier : détecte les types, extrait métadonnées et tableaux."""
    setup_logging(verbose)

    if not input_dir.exists() or not input_dir.is_dir():
        typer.echo(f"Erreur : dossier introuvable — {input_dir}", err=True)
        raise typer.Exit(1)

    # Extraction des métadonnées
    typer.echo(f"Analyse du dossier {input_dir.name}...")
    dossier = process_dossier(input_dir)
    report = format_metadata_report(dossier)
    typer.echo(report)

    # Extraction des tableaux depuis tous les PDF
    pdf_files = sorted(input_dir.glob("*.pdf"))
    all_tables = []

    for pdf_file in pdf_files:
        result = scan_pdf(pdf_file, min_cols=min_cols)
        if result.tables:
            typer.echo(f"  {pdf_file.name} : {len(result.tables)} tableau(x)")
            all_tables.extend(result.tables)

    if not all_tables:
        typer.echo("\nAucun tableau détecté dans les PDF du dossier.")
        raise typer.Exit(0)

    # Traitement et export
    datasets = process_tables(all_tables)
    if not datasets:
        typer.echo("\nAucune donnée exploitable après nettoyage.")
        raise typer.Exit(0)

    if output is None:
        demande_id = dossier.numero_demande or "dossier"
        output = input_dir / f"demande_{demande_id}_extraction.xlsx"

    meta_rows = metadata_to_rows(dossier)
    write_excel(datasets, output, metadata_rows=meta_rows)

    total_rows = sum(len(ds.data_rows) + len(ds.total_rows) for ds in datasets)
    typer.echo(f"\nTerminé ! {total_rows} lignes + métadonnées exportées.")
    for ds in datasets:
        if ds.name.startswith("Annexe"):
            pages_str = ", ".join(str(p) for p in ds.source_pages)
            typer.echo(f'INFO: La feuille "{ds.name}" contient les tableaux consolidés des pages {pages_str}')
    typer.echo(f"Fichier : {output.resolve()}")


@app.command(name="process-zip")
def process_zip_cmd(
    zip_path: Path = typer.Argument(..., help="Chemin vers le fichier ZIP contenant les dossiers de demandes."),
    output: Path = typer.Option(
        None, "--output", "-o",
        help="Chemin du fichier ZIP de sortie. Par défaut : resultats_extraction.zip",
    ),
    min_cols: int = typer.Option(8, "--min-cols", help="Nombre minimum de colonnes pour un tableau valide."),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Logs détaillés."),
) -> None:
    """Traite un ZIP de dossiers de demandes et génère un ZIP structuré."""
    setup_logging(verbose)

    if not zip_path.exists():
        typer.echo(f"Erreur : fichier introuvable — {zip_path}", err=True)
        raise typer.Exit(1)

    if output is None:
        output = zip_path.parent / "resultats_extraction.zip"

    typer.echo(f"Traitement de {zip_path.name}...")

    zip_data = zip_path.read_bytes()

    def on_progress(current, total, prefix):
        if prefix:
            typer.echo(f"  [{current + 1}/{total}] Demande {prefix}...")
        else:
            typer.echo("Traitement terminé.")

    results = process_zip(zip_data, min_cols=min_cols, on_progress=on_progress)

    if not results:
        typer.echo("Aucune demande trouvée dans le ZIP.")
        raise typer.Exit(0)

    # Rapport
    typer.echo(f"\n{'='*60}")
    typer.echo(f"{'Demandes traitées':30s} : {len(results)}")
    for r in results:
        status = "OK" if not r.error else f"ERREUR: {r.error}"
        courrier = "oui" if r.source_pdfs.get('courrier') else "non"
        ar = "oui" if r.source_pdfs.get('ar') else "non"
        depot = "oui" if r.source_pdfs.get('depot') else "non"
        typer.echo(
            f"  Demande {r.prefix} (N°{r.numero_demande}) : "
            f"courrier={courrier}, AR={ar}, dépôt={depot}, "
            f"{r.row_count} lignes — {status}"
        )
        for ds in r.datasets:
            if ds.name.startswith("Annexe"):
                pages_str = ", ".join(str(p) for p in ds.source_pages)
                typer.echo(f'    → Feuille "{ds.name}" : pages {pages_str}')
    typer.echo(f"{'='*60}")

    # Génération du ZIP
    typer.echo(f"\nGénération de {output}...")
    output_data = build_output_zip(results)
    output.write_bytes(output_data)
    typer.echo(f"Fichier créé : {output.resolve()} ({len(output_data) // 1024} Ko)")


if __name__ == "__main__":
    app()
