#!/usr/bin/env python3
"""
Excel Processor for FIA Automation
---------------------------------
Processes Excel files with participant data for FIA automation.
"""

import json
from datetime import datetime
from pathlib import Path
import re
from typing import Dict, List

import pandas as pd
from loguru import logger
from unidecode import unidecode

# Phone cleanup with optional E.164 using phonenumbers (if available)
try:
    import phonenumbers
    PHONENUMBERS_AVAILABLE = True
except Exception:
    PHONENUMBERS_AVAILABLE = False

# Constants
HEADER_SEARCH_ROWS = 60       # scan top rows to find header line
HEADER_SEARCH_COLS = 40       # scan left-most cols to speed up
TOP_SCAN_ROWS = 12            # rows to scan for "Numero du Groupe" in the top block
TOP_SCAN_COLS = 8

WANTED_LABELS = {
    "code_perso": [
        "code perso", "code-personnel", "codepersonnel", "code", "id", "code personne",
    ],
    "nom_prenom": [
        "nom et prénom", "nom et prenom", "nom & prénom", "nom & prenom",
        "nom, prénom", "nom, prenom", "nom et prénoms", "nom et prenoms",
        "nom et pr", "nom prenom", "nom/prénom", "nom/prenom",
        "nom",  # fallback single column
    ],
    "courriel": ["courriel", "email", "e-mail", "adresse email", "adresse courriel"],
    "telephone": ["téléphone", "telephone", "tel", "no de téléphone", "no de telephone", "numero de telephone"],
}

class ExcelProcessor:
    def __init__(self, input_path: str, sheet_name=0, output_dir: str = ""):
        """Initialize Excel processor with input file path and optional output directory."""
        self.input_path = Path(input_path)
        self.sheet_name = sheet_name
        self.ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.outdir = Path(output_dir) if output_dir else Path(f"run_{self.ts}")

    def norm_text(self, x: str) -> str:
        """Lowercase, strip, unaccent, collapse spaces."""
        if x is None:
            return ""
        s = unidecode(str(x)).lower().strip()
        s = re.sub(r"\\s+", " ", s)
        return s

    def find_header_row(self, df: pd.DataFrame) -> int:
        """Return the row index of the header that contains at least two wanted labels."""
        n_rows = min(HEADER_SEARCH_ROWS, len(df))
        n_cols = min(HEADER_SEARCH_COLS, df.shape[1])
        for r in range(n_rows):
            row = [self.norm_text(v) for v in df.iloc[r, :n_cols].tolist()]
            # test presence of at least two of our desired label families
            hits = 0
            for family in WANTED_LABELS.values():
                if any(any(lbl in cell for lbl in family) for cell in row):
                    hits += 1
            if hits >= 2:
                return r
        raise RuntimeError("Impossible de localiser la ligne d'en-tête (labels non trouvés dans les premières lignes).")

    def map_columns(self, header_row: pd.Series) -> dict:
        """Map raw header names to canonical keys."""
        mapping = {}
        for col_idx, name in enumerate(header_row.tolist()):
            n = self.norm_text(name)
            if not n:
                continue
            for canonical, family in WANTED_LABELS.items():
                if any(lbl == n or lbl in n for lbl in family):
                    # Keep first match only
                    mapping.setdefault(canonical, col_idx)
        return mapping

    def scan_top_for_group(self, df_raw: pd.DataFrame) -> str:
        """Scan the top-left block to find 'Numero du Groupe' and return the adjacent value."""
        n_rows = min(TOP_SCAN_ROWS, len(df_raw))
        n_cols = min(TOP_SCAN_COLS, df_raw.shape[1])
        for r in range(n_rows):
            for c in range(n_cols):
                v = self.norm_text(df_raw.iloc[r, c])
                if re.fullmatch(r"(numero|numero du|no du)\\s+groupe[: ]*", v):
                    if c + 1 < n_cols:
                        val = str(df_raw.iloc[r, c + 1]).strip()
                        if val and val.lower() != "nan":
                            return val
        # Fallback: fuzzy search
        for r in range(n_rows):
            for c in range(n_cols):
                v = self.norm_text(df_raw.iloc[r, c])
                if "numero du groupe" in v or "no du groupe" in v or "numero groupe" in v:
                    if c + 1 < n_cols:
                        val = str(df_raw.iloc[r, c + 1]).strip()
                        if val and val.lower() != "nan":
                            return val
        return ""

    def clean_phone(self, value: str, region="CA") -> str:
        """Return a cleaned phone number in E.164 format if possible."""
        if value is None:
            return ""
        s = str(value).strip()
        s = re.sub(r"[^\\d+]", "", s)  # keep digits and plus
        if not s:
            return ""
        if PHONENUMBERS_AVAILABLE:
            try:
                p = phonenumbers.parse(s, region)
                if phonenumbers.is_valid_number(p):
                    return phonenumbers.format_number(p, phonenumbers.PhoneNumberFormat.E164)
            except Exception:
                pass
        return s

    def process(self) -> Dict:
        """Process the Excel file and return summary of results."""
        if not self.input_path.exists():
            raise FileNotFoundError(f"File not found: {self.input_path}")

        # Create output directory
        self.outdir.mkdir(parents=True, exist_ok=True)

        # Read raw file without headers
        df_raw = pd.read_excel(self.input_path, header=None, sheet_name=self.sheet_name, engine="openpyxl")
        df_raw.iloc[:80, :20].to_csv(self.outdir / "input_snapshot.csv", index=False)

        # Extract group number and process headers
        numero_groupe = self.scan_top_for_group(df_raw)
        header_row_index = self.find_header_row(df_raw)
        header_series = df_raw.iloc[header_row_index]
        mapping = self.map_columns(header_series)

        # Read with proper headers
        df = pd.read_excel(self.input_path, header=header_row_index, sheet_name=self.sheet_name, engine="openpyxl")
        df.columns = [self.norm_text(c) for c in df.columns]

        # Map columns
        def get_col_idx_or_name(canon_key):
            if canon_key not in mapping:
                return None
            target_idx = mapping[canon_key]
            try:
                return df.columns[target_idx]
            except Exception:
                return None

        # Extract columns
        col_code = get_col_idx_or_name("code_perso")
        col_nom = get_col_idx_or_name("nom_prenom")
        col_mail = get_col_idx_or_name("courriel")
        col_tel = get_col_idx_or_name("telephone")

        # Build working dataframe
        work = pd.DataFrame()
        work["code_perso"] = df[col_code] if col_code and col_code in df.columns else ""
        work["nom_prenom"] = df[col_nom].astype(str).str.strip() if col_nom and col_nom in df.columns else ""
        work["courriel"] = df[col_mail].astype(str).str.strip().str.lower() if col_mail and col_mail in df.columns else ""
        work["telephone_raw"] = df[col_tel] if col_tel and col_tel in df.columns else ""

        # Clean data
        work["telephone"] = work["telephone_raw"].apply(lambda v: self.clean_phone(v, region="CA"))
        work.drop(columns=["telephone_raw"], inplace=True)
        work["numero_groupe"] = numero_groupe

        # Filter rows
        is_all_empty = (work[["code_perso","nom_prenom","courriel","telephone"]]
                       .astype(str).apply(lambda s: s.str.strip())
                       .replace({"": None}).isna().all(axis=1))
        work = work.loc[~is_all_empty].reset_index(drop=True)

        # Validate rows
        ok_mask = (work["telephone"].astype(str).str.len() > 0) | (work["nom_prenom"].astype(str).str.len() > 0)
        cleaned = work.loc[ok_mask].copy()
        errors = work.loc[~ok_mask].copy()
        if not errors.empty:
            errors["raison"] = "Manque phone ET nom_prenom"

        # Save outputs
        cleaned.to_csv(self.outdir / "cleaned_rows.csv", index=False)
        errors.to_csv(self.outdir / "errors.csv", index=False)

        # Create summary
        summary = {
            "input": str(self.input_path),
            "outdir": str(self.outdir),
            "numero_groupe_detecte": numero_groupe,
            "rows_total": int(len(work)),
            "rows_cleaned": int(len(cleaned)),
            "rows_errors": int(len(errors)),
            "columns_mapping": mapping,
            "header_row_index": int(header_row_index),
            "phonenumbers_available": PHONENUMBERS_AVAILABLE,
            "timestamp": self.ts,
        }

        with open(self.outdir / "run_summary.json", "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)

        logger.success(f"OK: cleaned_rows.csv ({len(cleaned)}), errors.csv ({len(errors)}), numero_groupe='{numero_groupe}'")
        logger.info(f"Résumé -> {self.outdir/'run_summary.json'}")

        return summary

    def get_processed_data(self) -> pd.DataFrame:
        """Return the cleaned data as a DataFrame for further processing."""
        cleaned_file = self.outdir / "cleaned_rows.csv"
        if cleaned_file.exists():
            return pd.read_csv(cleaned_file)
        raise FileNotFoundError("Processed data not found. Run process() first.")