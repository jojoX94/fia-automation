import streamlit as st
import sys
import os
import time
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv

# Add parent directory to Python path
sys.path.append(str(Path(__file__).parent.parent))

from core import automate, matching
from core.excel_processor import ExcelProcessor
from core.lgestat import LGEstatAutomation, PersonData

def process_excel_file(uploaded_file, output_dir=None):
    """Process uploaded Excel file and return the processor instance."""
    # Save uploaded file temporarily
    temp_path = Path("temp_upload.xlsx")
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getvalue())

    try:
        # Process the file
        processor = ExcelProcessor(temp_path, output_dir=output_dir)
        summary = processor.process()
        return processor, summary
    finally:
        # Clean up temp file
        if temp_path.exists():
            os.unlink(temp_path)

def prepare_verification_data(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare DataFrame for group verification."""
    # Map columns if they come from cleaned_rows.csv
    if "code_perso" in df.columns and "numero_personne" not in df.columns:
        df = df.rename(columns={"code_perso": "numero_personne"})

    if "groupe_attendu" not in df.columns and "numero_groupe" in df.columns:
        df = df.rename(columns={"numero_groupe": "groupe_attendu"})

    # Ensure required columns exist
    required_cols = ["numero_personne", "groupe_attendu"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Colonnes manquantes: {', '.join(missing_cols)}")

    # Clean data
    df["numero_personne"] = df["numero_personne"].astype(str).str.strip()
    df["groupe_attendu"] = df["groupe_attendu"].astype(str).str.strip().str.upper()

    # Remove empty rows
    df = df.loc[df["numero_personne"].str.len() > 0].copy()

    return df

def verify_groups_tab():
    """Interface for group verification functionality."""
    st.header("2. Vérification des Groupes")

    # Load environment variables
    load_dotenv()
    credentials = {
        "client_id": os.getenv("LGESTAT_CLIENT_ID"),
        "email": os.getenv("LGESTAT_EMAIL"),
        "password": os.getenv("LGESTAT_PASSWORD")
    }

    if not all(credentials.values()):
        st.error("Configuration LGEstat manquante. Vérifiez le fichier .env")
        return

    # Data source selection
    data_source = st.radio(
        "Source des données",
        ["Utiliser les données nettoyées", "Charger un nouveau fichier"],
        index=0
    )

    # Get verification data
    verification_df = None
    if data_source == "Utiliser les données nettoyées":
        try:
            # Look for most recent cleaned data
            runs = sorted(Path(".").glob("run_*"))
            if not runs:
                st.warning("Aucune donnée nettoyée trouvée. Traitez d'abord un fichier Excel ou chargez un fichier directement.")
                return

            latest_run = max(runs)
            cleaned_file = latest_run / "cleaned_rows.csv"
            if cleaned_file.exists():
                verification_df = pd.read_csv(cleaned_file)
                st.success(f"Données chargées depuis: {cleaned_file}")
            else:
                st.warning("Fichier cleaned_rows.csv non trouvé dans le dernier traitement.")
                return
        except Exception as e:
            st.error(f"Erreur lors du chargement des données nettoyées: {str(e)}")
            return
    else:
        # File upload
        uploaded_file = st.file_uploader(
            "Fichier de vérification (.xlsx, .csv, .numbers)",
            type=["xlsx", "csv", "numbers"]
        )

    if data_source == "Charger un nouveau fichier" and not uploaded_file:
        return

    try:
        # Prepare verification data
        if data_source == "Charger un nouveau fichier":
            # Save and read uploaded file
            temp_path = Path("temp_verify.xlsx")
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getvalue())

            # Read the file based on its type
            if uploaded_file.name.endswith('.csv'):
                verification_df = pd.read_csv(temp_path)
            else:  # xlsx or numbers
                verification_df = pd.read_excel(temp_path)

        # Prepare data for verification
        if verification_df is not None:
            try:
                verification_df = prepare_verification_data(verification_df)
            except ValueError as ve:
                st.error(f"Erreur dans le format des données: {str(ve)}")
                return

            # Show preview of data to verify
            st.subheader("Aperçu des données à vérifier")
            st.dataframe(verification_df.head())
            temp_path = None
        if data_source == "Charger un nouveau fichier":
            # Handle uploaded file
            temp_path = Path("temp_verify.xlsx")
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getvalue())

            # Read the file based on its type
            if uploaded_file.name.endswith('.csv'):
                verification_df = pd.read_csv(temp_path)
            else:  # xlsx or numbers
                verification_df = pd.read_excel(temp_path)

        # Prepare and validate data
        if verification_df is not None:
            try:
                verification_df = prepare_verification_data(verification_df)
            except ValueError as ve:
                st.error(f"Erreur dans le format des données: {str(ve)}")
                return

            # Show preview of data to verify
            st.subheader("Aperçu des données à vérifier")
            st.dataframe(verification_df.head())

            # Initialize automation when ready
            if st.button("Lancer la vérification"):
                lgestat = None
                try:
                    lgestat = LGEstatAutomation(
                        client_id=credentials["client_id"],
                        email=credentials["email"],
                        password=credentials["password"],
                        headless=False  # Temporarily set to False to see the browser
                    )

                    # Process verifications
                    with st.spinner("Vérification en cours..."):
                        results = []
                        progress_bar = st.progress(0)
                        total_rows = len(verification_df)

                        # Login first
                        if not lgestat.login():
                            st.error("Échec de connexion à LGEstat")
                            return

                        # Process each person
                        for idx, row in verification_df.iterrows():
                            person = PersonData(
                                numero=row["numero_personne"],
                                groupe_attendu=row["groupe_attendu"],
                                nom=row.get("nom_prenom", "")
                            )
                            result = lgestat.verify_person(person)
                            results.append(result)

                            # Update progress
                            progress = (idx + 1) / total_rows
                            progress_bar.progress(progress)

                        # Create results DataFrame
                        results_df = pd.DataFrame(results)

                        # Display results
                        st.success("✅ Vérification terminée!")

                        # Statistics
                        total = len(results_df)
                        verified = results_df["est_dans_groupe"].sum()
                        failed = total - verified

                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total", total)
                        with col2:
                            st.metric("Vérifiés", int(verified))
                        with col3:
                            st.metric("Échoués", int(failed))

                        # Detailed results
                        st.subheader("Résultats détaillés")
                        st.dataframe(results_df)

                        # Export option
                        if st.button("Exporter les résultats"):
                            timestamp = time.strftime("%Y%m%d_%H%M%S")
                            export_path = f"verification_resultats_{timestamp}.csv"
                            results_df.to_csv(export_path, index=False)
                            st.success(f"Résultats exportés: {export_path}")

                finally:
                    if lgestat:
                        lgestat.close()

    except Exception as e:
        st.error(f"Erreur lors de la vérification: {str(e)}")

    finally:
        # Cleanup temporary file
        if temp_path and temp_path.exists():
            os.unlink(temp_path)

def main():
    st.title("FIA Automation Tools")

    # Create tabs
    tab1, tab2 = st.tabs(["Traitement Excel", "Vérification Groupes"])

    with tab1:
        # Excel file processing section
        st.header("1. Traitement du fichier Excel")
        uploaded_file = st.file_uploader("Choisir un fichier Excel", type=["xlsx"], key="excel_upload")

        if uploaded_file:
            try:
                processor, summary = process_excel_file(uploaded_file)

                # Display summary
                st.success("✅ Fichier traité avec succès!")

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Lignes traitées", summary["rows_cleaned"])
                with col2:
                    st.metric("Lignes avec erreurs", summary["rows_errors"])

                st.subheader("Détails")
                st.json(summary)

                # Show preview of cleaned data
                if Path(summary["outdir"]).joinpath("cleaned_rows.csv").exists():
                    cleaned_data = processor.get_processed_data()
                    st.subheader("Aperçu des données nettoyées")
                    st.dataframe(cleaned_data.head())

            except Exception as e:
                st.error(f"Erreur lors du traitement: {str(e)}")

    with tab2:
        verify_groups_tab()

if __name__ == "__main__":
    main()