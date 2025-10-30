import streamlit as st
import sys
from pathlib import Path
import os

# Add parent directory to Python path
sys.path.append(str(Path(__file__).parent.parent))

from core import automate, matching
from core.excel_processor import ExcelProcessor

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

def main():
    st.title("FIA Automation Tools")

    # Excel file upload section
    st.header("1. Traitement du fichier Excel")
    uploaded_file = st.file_uploader("Choisir un fichier Excel", type=["xlsx"])

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

                # Option to proceed with automation
                if st.button("Continuer avec l'automatisation"):
                    # Here you can integrate with your automation logic
                    st.info("Fonctionnalité d'automatisation en développement...")

        except Exception as e:
            st.error(f"Erreur lors du traitement: {str(e)}")

if __name__ == "__main__":
    main()