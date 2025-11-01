"""
Utilities for LGEstat verification processes.
Provides common functions for file handling and data normalization.
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List
from loguru import logger

def read_table(input_path: str) -> pd.DataFrame:
    """Read input file (CSV, Excel, or Numbers) and return DataFrame."""
    p = Path(input_path)
    if not p.exists():
        raise FileNotFoundError(f"Fichier introuvable: {input_path}")

    ext = p.suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(p)
    elif ext in (".xlsx", ".xlsm"):
        df = pd.read_excel(p)
    elif ext == ".numbers":
        try:
            from numbers_parser import Document
        except Exception as e:
            raise RuntimeError("Le support .numbers nécessite `pip install numbers-parser`.") from e

        doc = Document(str(p))
        sheets = doc.sheets
        if not sheets:
            raise RuntimeError("Aucune feuille dans ce fichier .numbers")

        # Read first table from first sheet
        tbl = sheets[0].tables[0]
        data = [[cell.value for cell in row] for row in tbl.rows()]
        df = pd.DataFrame(data[1:], columns=[str(c) for c in data[0]])
    else:
        raise ValueError(f"Extension non supportée: {ext}")

    # Normalize headers
    df.columns = [c.strip().lower() for c in df.columns]
    return df

def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize and validate input DataFrame."""
    # Required columns check
    required = ["numero_personne", "groupe_attendu"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Colonne manquante: {col} (requis: {required})")

    # Clean up data
    df["numero_personne"] = df["numero_personne"].astype(str).str.strip()
    df["groupe_attendu"] = df["groupe_attendu"].astype(str).str.strip().str.upper()

    # Optional columns
    if "nom" in df.columns:
        df["nom"] = df["nom"].astype(str).str.strip().str.upper()
    if "prenom" in df.columns:
        df["prenom"] = df["prenom"].astype(str).str.strip().str.upper()

    return df