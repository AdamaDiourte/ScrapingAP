import logging
import os
import re
import time
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor


logger = logging.getLogger(__name__)


class AppelsProjetFinder:
    """Service de recherche et extraction d'appels à projets.

    Étapes clés:
    - lecture d'un fichier Excel d'entrée
    - recherche via IA selon un 'sujet'
    - scraping d'URL et extraction assistée par IA
    - génération d'un document Word
    """

    def __init__(self, api_key: str, api_provider: str = "openai") -> None:
        self.api_key = api_key
        self.api_provider = api_provider
        self.results: List[Dict[str, Any]] = []
        logger.info("Init AppelsProjetFinder provider=%s", api_provider)

    # -------------------- Lecture Excel --------------------
    def lire_fichier_excel(self, chemin_excel: str) -> Optional[pd.DataFrame]:
        try:
            logger.info("Lecture Excel: %s", chemin_excel)
            df = pd.read_excel(chemin_excel)
            logger.info("Excel lu: %d lignes", len(df))
            return df
        except Exception as exc:
            logger.exception("Erreur lecture Excel: %s", exc)
            return None

    # -------------------- Recherche IA --------------------
    def rechercher_avec_ia(self, sujet: str) -> List[Dict[str, Any]]:
        logger.info("Recherche IA sujet='%s' provider=%s", sujet, self.api_provider)
        if self.api_provider == "openai":
            return self._recherche_openai(sujet)
        if self.api_provider == "anthropic":
            return self._recherche_anthropic(sujet)
        logger.warning("Provider IA inconnu: %s", self.api_provider)
        return []

    def _call_openai_chat(self, messages: List[Dict[str, str]], model_legacy: str = "gpt-4", model_modern: str = "gpt-4o-mini") -> str:
        """Compatibilité OpenAI: utilise l'API 0.28 (ChatCompletion) ou 1.x (OpenAI client)."""
        try:
            import openai  # type: ignore
            version = getattr(openai, "__version__", "unknown")
            logger.info("OpenAI version=%s", version)

            # Cas SDK 0.28 (ancienne API)
            if hasattr(openai, "ChatCompletion"):
                openai.api_key = self.api_key
                logger.debug("OpenAI legacy ChatCompletion.create model=%s", model_legacy)
                response = openai.ChatCompletion.create(
                    model=model_legacy,
                    messages=messages,
                    temperature=0.3,
                )
                return response.choices[0].message.content  # type: ignore[attr-defined]

            # Cas SDK 1.x/2.x (nouvelle API)
            try:
                from openai import OpenAI  # type: ignore

                client = OpenAI(api_key=self.api_key)
                logger.debug("OpenAI modern chat.completions.create model=%s", model_modern)
                response = client.chat.completions.create(
                    model=model_modern,
                    messages=messages,
                    temperature=0.3,
                )
                return response.choices[0].message.content or ""
            except Exception as modern_exc:
                logger.exception("OpenAI modern client échec: %s", modern_exc)
                raise
        except Exception as exc:
            logger.exception("OpenAI indisponible: %s", exc)
            return ""

    def _recherche_openai(self, sujet: str) -> List[Dict[str, Any]]:
        prompt = (
            f"""Recherche les appels à projet actifs concernant: {sujet}

Pour chaque appel trouvé, extrais:
1. Titre de l'appel
2. Organisation responsable
3. Date de début des candidatures
4. Date de clôture des candidatures
5. Lien URL vers l'appel
6. Brève description (2-3 lignes)

Format de réponse en JSON:
[
  {{
    "titre": "...",
    "organisation": "...",
    "date_debut": "JJ/MM/AAAA",
    "date_cloture": "JJ/MM/AAAA",
    "url": "...",
    "description": "..."
  }}
]
"""
        )

        try:
            content = self._call_openai_chat(
                messages=[
                    {
                        "role": "system",
                        "content": "Tu es un assistant spécialisé dans la recherche d'appels à projet.",
                    },
                    {"role": "user", "content": prompt},
                ]
            )
            return self._parse_response(content)
        except Exception as exc:
            logger.exception("Erreur API OpenAI: %s", exc)
            return []

    def _recherche_anthropic(self, sujet: str) -> List[Dict[str, Any]]:
        try:
            import anthropic

            client = anthropic.Anthropic(api_key=self.api_key)
            prompt = (
                f"""Recherche les appels à projet actifs concernant: {sujet}

Pour chaque appel trouvé, extrais:
1. Titre de l'appel
2. Organisation responsable
3. Date de début des candidatures
4. Date de clôture des candidatures
5. Lien URL vers l'appel
6. Brève description (2-3 lignes)

Format de réponse en JSON:
[
  {{
    "titre": "...",
    "organisation": "...",
    "date_debut": "JJ/MM/AAAA",
    "date_cloture": "JJ/MM/AAAA",
    "url": "...",
    "description": "..."
  }}
]
"""
            )

            logger.debug("Appel Anthropic.messages.create")
            message = client.messages.create(
                model="claude-sonnet-4-5-20250929",
                max_tokens=4096,
                messages=[{"role": "user", "content": prompt}],
            )
            return self._parse_response(message.content[0].text)
        except Exception as exc:
            logger.exception("Erreur API Anthropic: %s", exc)
            return []

    # -------------------- Scraping --------------------
    def scraper_url(self, url: str) -> Optional[Dict[str, Any]]:
        try:
            logger.info("Scraping URL: %s", url)
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            }
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, "html.parser")
            text_content = soup.get_text(separator=" ", strip=True)
            return self._extraire_info_avec_ia(text_content, url)
        except Exception as exc:
            logger.exception("Erreur scraping %s", url)
            return None

    def _extraire_info_avec_ia(self, contenu: str, url: str) -> Optional[Dict[str, Any]]:
        logger.debug("Extraction IA depuis HTML url=%s contenu_len=%s", url, len(contenu))
        prompt = f"""Analyse ce contenu de page web et extrais les informations d'appel à projet:

{contenu[:4000]}

Extrais en JSON:
{{
  "titre": "...",
  "organisation": "...",
  "date_debut": "JJ/MM/AAAA",
  "date_cloture": "JJ/MM/AAAA",
  "description": "..."
}}

Si les informations ne sont pas disponibles, mets "Non spécifié".
"""

        if self.api_provider == "openai":
            try:
                content = self._call_openai_chat(messages=[{"role": "user", "content": prompt}])
                result = self._parse_response(content)
                if result:
                    result[0]["url"] = url
                    return result[0]
            except Exception as exc:
                logger.exception("Erreur extraction IA: %s", exc)

        return None

    # -------------------- Parsing --------------------
    def _parse_response(self, response_text: str) -> List[Dict[str, Any]]:
        try:
            import json

            logger.debug("Parsing réponse IA len=%s", len(response_text) if response_text else 0)
            json_match = re.search(r"\[.*\]", response_text, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())
            json_match = re.search(r"\{.*\}", response_text, re.DOTALL)
            if json_match:
                return [json.loads(json_match.group())]
        except Exception as exc:
            logger.exception("Erreur parsing JSON: %s", exc)
        return []

    # -------------------- Orchestration --------------------
    def traiter_fichier(self, chemin_excel: str) -> None:
        logger.info("Début traitement fichier: %s", chemin_excel)
        df = self.lire_fichier_excel(chemin_excel)
        if df is None:
            logger.error("Abandon: lecture Excel échouée")
            return

        # Normalise les noms de colonnes pour tolérer des synonymes
        original_to_lower = {c: c.strip().lower() for c in df.columns}
        df_proc = df.rename(columns=lambda c: original_to_lower[c])

        subject_candidates = [
            "sujet", "subject", "thème", "theme", "topic", "mots_clés", "mots-cles", "keywords"
        ]
        url_candidates = [
            "url", "lien", "link", "urls", "liens"
        ]

        subject_col = next((c for c in subject_candidates if c in df_proc.columns), None)
        url_col = next((c for c in url_candidates if c in df_proc.columns), None)

        if not subject_col and not url_col:
            logger.warning(
                "Aucune colonne reconnue. Attendu l'une de %s ou %s",
                subject_candidates,
                url_candidates,
            )

        self.results = []

        if subject_col:
            for idx, row in df_proc.iterrows():
                sujet = row[subject_col]
                if pd.notna(sujet) and str(sujet).strip():
                    logger.info("Recherche IA ligne=%s sujet=%s", idx, sujet)
                    self.results.extend(self.rechercher_avec_ia(str(sujet)))
                    time.sleep(1)

        if url_col:
            for idx, row in df_proc.iterrows():
                url = row[url_col]
                if pd.notna(url) and str(url).strip():
                    logger.info("Scraping ligne=%s url=%s", idx, url)
                    result = self.scraper_url(str(url))
                    if result:
                        self.results.append(result)
                    time.sleep(1)

        logger.info("Total résultats: %d", len(self.results))

    def generer_document_word(self, chemin_sortie: str = "appels_projet.docx") -> None:
        doc = Document()

        # Titre
        titre = doc.add_heading("Appels à Projet - Résultats de recherche", 0)
        titre.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        date_generation = doc.add_paragraph(
            f"Généré le: {datetime.now().strftime('%d/%m/%Y à %H:%M')}"
        )
        date_generation.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        doc.add_paragraph()

        if not self.results:
            logger.warning("Aucun résultat à exporter - génération d'un document vide informatif")
            p = doc.add_paragraph()
            p.add_run("Aucun résultat trouvé pour les données fournies. ")
            p.add_run(
                "Vérifiez que votre fichier Excel contient une colonne 'sujet' ou 'url' (ou leurs synonymes: 'thème', 'link', 'lien', ...)."
            )
        else:
            for idx, appel in enumerate(self.results, 1):
                doc.add_heading(f"{idx}. {appel.get('titre', 'Sans titre')}", level=2)

                if appel.get("organisation"):
                    p = doc.add_paragraph()
                    p.add_run("Organisation: ").bold = True
                    p.add_run(str(appel["organisation"]))

                p = doc.add_paragraph()
                p.add_run("Date de début: ").bold = True
                p.add_run(str(appel.get("date_debut", "Non spécifiée")))

                p = doc.add_paragraph()
                p.add_run("Date de clôture: ").bold = True
                run = p.add_run(str(appel.get("date_cloture", "Non spécifiée")))
                if appel.get("date_cloture") and appel["date_cloture"] != "Non spécifiée":
                    run.font.color.rgb = RGBColor(255, 0, 0)
                    run.bold = True

                if appel.get("url"):
                    p = doc.add_paragraph()
                    p.add_run("Lien: ").bold = True
                    p.add_run(str(appel["url"]))

                if appel.get("description"):
                    p = doc.add_paragraph()
                    p.add_run("Description: ").bold = True
                    p.add_run(str(appel["description"]))

                doc.add_paragraph("_" * 80)

        doc.save(chemin_sortie)
        logger.info("Document Word généré: %s", chemin_sortie)


