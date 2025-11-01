"""
LGEstat automation module
------------------------
Handles interaction with LGEstat web interface for automated group verification
and other automation tasks.
"""

import os
import time
from typing import Optional, Dict, List
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from loguru import logger

@dataclass
class PersonData:
    """Structure to hold person information."""
    numero: str
    groupe_attendu: str
    nom: str = ""
    prenom: str = ""

class LGEstatAutomation:
    """Handles automated interactions with LGEstat web interface."""

    BASE_URL = "https://app.lgestat.com"
    LOGIN_URL = f"{BASE_URL}/fr/auth/login"
    SEARCH_URL = f"{BASE_URL}/fr/search"  # Add the actual search URL

    def __init__(self, client_id: str, email: str, password: str, headless: bool = True):
        """Initialize LGEstat automation with credentials."""
        self.client_id = client_id
        self.email = email
        self.password = password
        self.headless = headless
        self.driver = None

    def start_driver(self) -> None:
        """Initialize and configure Chrome WebDriver."""
        options = Options()
        # Add standard options for stability
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        if self.headless:
            options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--window-size=1280,1024")

        # Use ChromeDriverManager to get the correct driver version
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.set_page_load_timeout(60)
        self.driver.implicitly_wait(10)

    def login(self) -> bool:
        """
        Log into LGEstat using provided credentials.
        Returns True if login successful, False otherwise.
        """
        if not self.driver:
            self.start_driver()

        try:
            logger.info("Accessing LGEstat login page...")
            self.driver.get(self.LOGIN_URL)

            # Fill in login form
            self.driver.find_element(By.NAME, "id").send_keys(self.client_id)
            self.driver.find_element(By.NAME, "email").send_keys(self.email)
            self.driver.find_element(By.NAME, "password").send_keys(self.password)

            # Click login button
            submit_button = self.driver.find_element(By.CSS_SELECTOR, "input[type='submit'][value='Connexion']")
            submit_button.click()

            # Wait for redirect/login completion
            time.sleep(2)

            # Check if login was successful
            is_logged_in = "/auth/login" not in self.driver.current_url

            if is_logged_in:
                logger.success("✅ Connexion réussie à LGEstat!")
                # Try to get and log the user name or any welcome message if available
                try:
                    # Wait for the dashboard to load
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "dashboard"))
                    )
                    logger.info("📊 Dashboard chargé avec succès")
                except Exception:
                    pass  # Don't fail if we can't find these elements
            else:
                logger.error("❌ Échec de la connexion - Toujours sur la page de login")

            return is_logged_in

        except Exception as e:
            logger.error(f"❌ Échec de la connexion: {str(e)}")
            return False

    def search_person(self, numero: str) -> bool:
        """
        Navigate to search page and look for a person by their number.
        Returns True if person is found.
        """
        try:
            logger.info(f"🔍 Recherche de la personne {numero}...")

            # Navigate to search page
            self.driver.get(self.SEARCH_URL)
            logger.debug("Page de recherche chargée")

            # Wait for search input and enter person number
            search_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "search"))
            )
            search_input.clear()
            search_input.send_keys(numero)
            search_input.send_keys(Keys.RETURN)
            logger.debug(f"Numéro {numero} saisi et recherche lancée")

            # Wait for results and verify
            time.sleep(2)  # Allow time for results to load

            # Add visual confirmation
            if not self.headless:
                time.sleep(1)  # Additional pause to see results in non-headless mode

            logger.success(f"✅ Recherche effectuée pour {numero}")
            return True

        except Exception as e:
            logger.error(f"❌ Échec de la recherche pour {numero}: {str(e)}")
            return False

    def get_person_group(self) -> str:
        """Extract group information from person details page."""
        try:
            # Wait for and find the group element (adjust selector based on actual page)
            group_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-field='groupe']"))
            )
            return group_element.text.strip().upper()
        except Exception as e:
            logger.error(f"Failed to extract group info: {str(e)}")
            return ""

    def verify_person(self, person: PersonData) -> Dict:
        """
        Complete verification flow for a person.
        Returns verification results dictionary.
        """
        result = {
            "numero_personne": person.numero,
            "groupe_attendu": person.groupe_attendu,
            "nom": person.nom,
            "prenom": person.prenom,
            "est_dans_groupe": None,
            "groupe_trouve": "",
            "details": ""
        }

        try:
            # Search for person
            if not self.search_person(person.numero):
                result["details"] = "Personne non trouvée"
                return result

            # Get actual group
            groupe_trouve = self.get_person_group()
            result["groupe_trouve"] = groupe_trouve

            # Compare groups
            if groupe_trouve:
                result["est_dans_groupe"] = (groupe_trouve == person.groupe_attendu)
                result["details"] = (
                    "OK" if result["est_dans_groupe"]
                    else f"Groupe trouvé: {groupe_trouve}, Attendu: {person.groupe_attendu}"
                )
            else:
                result["details"] = "Impossible de trouver l'information de groupe"

        except Exception as e:
            result["details"] = f"Erreur: {str(e)}"
            logger.error(f"Verification failed for {person.numero}: {str(e)}")

        return result

    def process_verification_file(self, file_path: str, output_path: str = None) -> pd.DataFrame:
        """
        Process a verification file and return results as DataFrame.
        Args:
            file_path: Path to input file (.xlsx, .csv, .numbers)
            output_path: Optional path to save results CSV
        """
        from core.utils import read_table, normalize

        # Read and normalize input data
        df = read_table(file_path)
        df = normalize(df)

        if not self.driver or "/auth/login" in self.driver.current_url:
            if not self.login():
                raise RuntimeError("Failed to login to LGEstat")

        # Process each person
        results = []
        for i, row in df.iterrows():
            person = PersonData(
                numero=row["numero_personne"],
                groupe_attendu=row["groupe_attendu"],
                nom=row.get("nom", ""),
                prenom=row.get("prenom", "")
            )
            logger.info(f"[{i+1}/{len(df)}] Vérifier {person.numero} contre {person.groupe_attendu}")
            result = self.verify_person(person)
            results.append(result)

        # Create results DataFrame
        results_df = pd.DataFrame(results)

        # Save if output path provided
        if output_path:
            results_df.to_csv(output_path, index=False)
            logger.success(f"Results saved to {output_path}")

        return results_df

    def close(self):
        """Clean up and close the browser."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None