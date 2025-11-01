#!/usr/bin/env python3
"""
verify_groups.py — Lecture d'un fichier (.numbers/.xlsx/.csv) contenant
(numero_personne, groupe_attendu, etc.) et vérification dans Lgestat via Selenium
que la personne appartient bien au groupe indiqué.

Usage:
  python verify_groups.py --input chemin/vers/fichier.numbers --headless 1
Prérequis:
  pip install pandas openpyxl selenium python-dotenv loguru numbers-parser
"""

import os
import sys
import argparse
from pathlib import Path

import pandas as pd
from loguru import logger
from dotenv import load_dotenv

# Add parent directory to Python path for imports
sys.path.append(str(Path(__file__).parent.parent))

from core.lgestat import LGEstatAutomation
from core.utils import read_table, normalize

# Load environment variables
load_dotenv()

# Required environment variables
LGESTAT_CLIENT_ID = os.getenv("LGESTAT_CLIENT_ID")
LGESTAT_EMAIL = os.getenv("LGESTAT_EMAIL")
LGESTAT_PASSWORD = os.getenv("LGESTAT_PASSWORD")

from core.utils import read_table, normalize

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Chemin vers .numbers/.xlsx/.csv")
    parser.add_argument("--headless", type=int, default=1, help="1=headless, 0=visible")
    parser.add_argument("--out", default="rapport_verification.csv", help="CSV de sortie")
    args = parser.parse_args()

    # Verify environment variables
    if not all([LGESTAT_CLIENT_ID, LGESTAT_EMAIL, LGESTAT_PASSWORD]):
        raise ValueError(
            "Variables d'environnement manquantes. Créez un fichier .env avec:\n"
            "LGESTAT_CLIENT_ID=votre_id\n"
            "LGESTAT_EMAIL=votre_email\n"
            "LGESTAT_PASSWORD=votre_mot_de_passe"
        )

    # Read and normalize input data
    logger.info(f"Lecture du fichier: {args.input}")
    df = read_table(args.input)
    df = normalize(df)

    # Initialize LGEstat automation
    logger.info(f"{len(df)} lignes chargées. Démarrage du navigateur...")
    lgestat = LGEstatAutomation(
        client_id=LGESTAT_CLIENT_ID,
        email=LGESTAT_EMAIL,
        password=LGESTAT_PASSWORD,
        headless=bool(args.headless)
    )

    try:
        # Login to LGEstat
        if not lgestat.login():
            raise RuntimeError("Échec de connexion à LGEstat")

        # Process each person
        results = []
        for i, row in df.iterrows():
            numero = row["numero_personne"]
            groupe = row["groupe_attendu"]
            logger.info(f"[{i+1}/{len(df)}] Vérifier {numero} contre {groupe}")

            res = lgestat.find_person_and_check_group(numero, groupe)
            results.append(res)

        # Save results
        out_df = pd.DataFrame(results)
        out_df.to_csv(args.out, index=False)
        logger.success(f"Rapport généré: {args.out}")

    finally:
        lgestat.close()

if __name__ == "__main__":
    main()