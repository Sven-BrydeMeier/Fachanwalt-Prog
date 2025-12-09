"""
Fachanwalt-Falllistenverwaltung - Streamlit App
================================================
Eine Anwendung zur Verwaltung von Fachanwaltsfällen mit FAO-Konformitätsprüfung.
Ermöglicht Upload, Sortierung und Export von Falllisten gemäß § 5 FAO.
Unterstützt PDF-Upload mit automatischer Fallauswertung via ChatGPT/OpenAI.

Verwendung: streamlit run app.py
Voraussetzungen: pip install streamlit pandas openpyxl openai pypdf2
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date
from typing import Dict, List, Tuple, Optional, Any
import json
import re

# PDF-Verarbeitung
try:
    from PyPDF2 import PdfReader
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# OpenAI-Integration
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# =============================================================================
# KONFIGURATION - FAO-Mindestanforderungen (§ 5 FAO, Stand 01.06.2022)
# =============================================================================

FAO_CONFIG: Dict[str, Dict[str, Any]] = {
    "Erbrecht": {
        "paragraph": "§ 5 Abs. 1 lit. m, § 14f FAO",
        "gesamt_min": 80,
        "gerichtlich_min": 20,
        "aussergerichtlich_min": 60,
        "fg_max": 15,
        "bereiche": {
            1: "Materielles Erbrecht und Bezüge zum Familien- und Gesellschaftsrecht",
            2: "Testamentsvollstreckung und Nachlassverwaltung",
            3: "Internationales Privatrecht und Erbschaftsteuerrecht",
            4: "Erbprozessrecht",
            5: "Vorweggenommene Erbfolge",
        },
        "bereiche_anforderung": {
            "min_bereiche": 3,
            "min_faelle_pro_bereich": 5,
            "beschreibung": "Mindestens 3 der Bereiche Nr. 1-5 mit je mind. 5 Fällen"
        },
        "hinweis": (
            "Mind. 80 Fälle gesamt, davon mind. 20 rechtsförmliche Verfahren "
            "(max. 15 fG-Verfahren) und mind. 60 außergerichtliche Fälle. "
            "Abdeckung von mind. 3 Bereichen (§ 14f Nr. 1-5) mit je mind. 5 Fällen."
        )
    },
    "Arbeitsrecht": {
        "paragraph": "§ 5 Abs. 1 lit. c, § 10 FAO",
        "gesamt_min": 100,
        "gerichtlich_min": 50,
        "aussergerichtlich_min": 0,
        "bereiche": {
            1: "Individualarbeitsrecht",
            2: "Kollektives Arbeitsrecht (Betriebsverfassung, Tarifrecht)",
            3: "Arbeitsgerichtliches Verfahren",
            4: "Sozialrecht im Arbeitsrecht",
        },
        "bereiche_anforderung": {
            "spezial_anforderung": "kollektiv",
            "kollektiv_bereich": 2,
            "kollektiv_min": 5,
            "beschreibung": "Mind. 5 Fälle aus dem kollektiven Arbeitsrecht (§ 10 Nr. 2)"
        },
        "hinweis": (
            "Mind. 100 Fälle gesamt, davon mind. die Hälfte (50) gerichtliche/"
            "rechtsförmliche Verfahren. Mind. 5 Fälle aus dem kollektiven Arbeitsrecht."
        )
    },
    "Miet- und Wohnungseigentumsrecht": {
        "paragraph": "§ 5 Abs. 1 lit. j, § 14c FAO",
        "gesamt_min": 120,
        "gerichtlich_min": 60,
        "aussergerichtlich_min": 0,
        "bereiche": {
            1: "Wohnraummiete",
            2: "Gewerbemiete und Pacht",
            3: "Wohnungseigentum",
        },
        "bereiche_anforderung": {
            "min_bereiche": 3,
            "min_faelle_pro_bereich": 5,
            "beschreibung": "Mind. je 5 Fälle zu Wohnraummiete, Gewerbemiete/Pacht und WEG"
        },
        "hinweis": (
            "Mind. 120 Fälle gesamt, davon mind. 60 gerichtliche Verfahren. "
            "Je mind. 5 Fälle zu Wohnraummiete, Gewerbemiete/Pacht und Wohnungseigentum."
        )
    },
    "Familienrecht": {
        "paragraph": "§ 5 Abs. 1 lit. e, § 12 FAO",
        "gesamt_min": 120,
        "gerichtlich_min": 60,
        "aussergerichtlich_min": 0,
        "verbund_doppelt": True,
        "bereiche": {
            1: "Materielles Familienrecht (Ehe, Unterhalt, Güterrecht, Kindschaft)",
            2: "Familiengerichtliches Verfahrensrecht",
            3: "Internationales Privatrecht im Familienrecht",
            4: "Vertragsgestaltung im Familienrecht",
        },
        "bereiche_anforderung": {
            "beschreibung": "Abdeckung der in § 12 genannten Bereiche erforderlich"
        },
        "hinweis": (
            "Mind. 120 Fälle gesamt, davon mind. 60 gerichtliche Verfahren. "
            "Gewillkürte/nötige Verbundverfahren zählen doppelt. "
            "Abdeckung von materiellem Familienrecht, Verfahrensrecht, IPR und Vertragsgestaltung."
        )
    },
    "Verkehrsrecht": {
        "paragraph": "§ 5 Abs. 1 lit. k, § 14d FAO",
        "gesamt_min": 160,
        "gerichtlich_min": 60,
        "aussergerichtlich_min": 0,
        "bereiche": {
            1: "Verkehrszivilrecht",
            2: "Versicherungsrecht im Verkehr",
            3: "Ordnungswidrigkeiten- und Strafrecht im Verkehr",
            4: "Verkehrsverwaltungsrecht",
        },
        "bereiche_anforderung": {
            "min_bereiche": 3,
            "min_faelle_pro_bereich": 5,
            "beschreibung": "Mind. 3 Bereiche mit je mind. 5 Fällen"
        },
        "hinweis": (
            "Mind. 160 Fälle gesamt, davon mind. 60 gerichtliche Verfahren. "
            "Mind. 3 der Bereiche (Verkehrszivilrecht, Versicherungsrecht, OWi/Strafrecht, "
            "Verwaltungsrecht) mit je mind. 5 Fällen."
        )
    },
    "Handels- und Gesellschaftsrecht": {
        "paragraph": "§ 5 Abs. 1 lit. p, § 14i FAO",
        "gesamt_min": 80,
        "gerichtlich_min": 40,
        "aussergerichtlich_min": 0,
        "bereiche": {
            1: "Streitverfahren (gerichtlich/Schiedsgerichtsbarkeit/Mediation)",
            2: "Gestaltung/Gründung/Umwandlung von Gesellschaften",
            3: "Handelsrecht",
        },
        "bereiche_anforderung": {
            "spezial_anforderung": "hgr",
            "streit_min": 10,
            "gestaltung_min": 10,
            "streit_bereich": 1,
            "gestaltung_bereich": 2,
            "beschreibung": "Mind. 10 Streitverfahren und mind. 10 Gestaltungs-/Gründungs-/Umwandlungsfälle"
        },
        "hinweis": (
            "Mind. 80 Fälle gesamt. Mind. 40 Fälle mit Streitverfahren/Schieds-/Mediationsverfahren "
            "und/oder Gestaltung/Gründung/Umwandlung. Davon mind. 10 Streitverfahren und "
            "mind. 10 Gestaltungs-/Gründungs-/Umwandlungsfälle."
        )
    },
}

# =============================================================================
# SPALTEN-MAPPING für unterschiedliche Import-Formate
# =============================================================================

COLUMN_MAPPING: Dict[str, List[str]] = {
    "kanzlei_az": ["Kanzlei-AZ", "kanzlei_az", "Az", "Aktenzeichen", "AZ", "Akte", "Kanzlei-Aktenzeichen"],
    "kurzrubrum": ["Kurzrubrum", "kurzrubrum", "Rubrum", "Parteien", "Beteiligte"],
    "sachverhalt": ["Sachverhalt", "sachverhalt", "Beschreibung", "Kurzbeschreibung", "Inhalt"],
    "zeitraum_von": ["Zeitraum_von", "zeitraum_von", "Von", "Beginn", "Start", "Anfang"],
    "zeitraum_bis": ["Zeitraum_bis", "zeitraum_bis", "Bis", "Ende", "Abschluss"],
    "gericht_az": ["Gerichts-AZ", "gericht_az", "Gerichts-Aktenzeichen", "Gericht-AZ", "GerichtsAZ"],
    "verfahrenstyp": ["Verfahrenstyp", "verfahrenstyp", "Typ", "Art_Verfahren"],
    "verfahrensart": ["Verfahrensart", "verfahrensart", "Art", "Verfahren"],
    "bereich_nr": ["Bereich_Nr", "bereich_nr", "Bereich-Nr", "BereichNr", "Bereich_Nummer"],
    "bereich_bezeichnung": ["Bereich_Bezeichnung", "bereich_bezeichnung", "Bereich", "Bereichsbezeichnung"],
    "bedeutung": ["Bedeutung", "bedeutung", "Schwierigkeit", "Komplexität"],
    "taetigkeitsbeschreibung": ["Taetigkeitsbeschreibung", "taetigkeitsbeschreibung", "Tätigkeit",
                                "Taetigkeit", "Tätigkeitsbeschreibung", "Beschreibung_Taetigkeit"],
    "stand": ["Stand", "stand", "Status"],
    "abschluss_art": ["Abschluss_Art", "abschluss_art", "Abschlussart", "Erledigung", "Erledigungsart"],
    "abschluss_datum": ["Abschluss_Datum", "abschluss_datum", "Abschlussdatum", "Datum_Abschluss"],
    "verbundener_fall": ["Verbundener_Fall", "verbundener_fall", "Verbunden", "Parallelfall", "Verbund"],
    "fachgebiet": ["Fachgebiet", "fachgebiet", "Rechtsgebiet", "Gebiet"],
}

TEMPLATE_COLUMNS = [
    "Kanzlei-AZ", "Kurzrubrum", "Sachverhalt", "Zeitraum_von", "Zeitraum_bis",
    "Gerichts-AZ", "Verfahrenstyp", "Verfahrensart", "Bereich_Nr", "Bereich_Bezeichnung",
    "Bedeutung", "Taetigkeitsbeschreibung", "Stand", "Abschluss_Art",
    "Abschluss_Datum", "Verbundener_Fall", "Fachgebiet"
]

# =============================================================================
# PDF-VERARBEITUNG
# =============================================================================

def extract_text_from_pdf(file) -> str:
    """Extrahiert Text aus einer PDF-Datei."""
    if not PDF_AVAILABLE:
        raise ImportError("PyPDF2 ist nicht installiert. Bitte führen Sie 'pip install pypdf2' aus.")

    reader = PdfReader(file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text


# =============================================================================
# OPENAI/CHATGPT-INTEGRATION
# =============================================================================

def get_openai_client(api_key: str) -> Optional[Any]:
    """Erstellt einen OpenAI-Client mit dem angegebenen API-Key."""
    if not OPENAI_AVAILABLE:
        return None
    return OpenAI(api_key=api_key)


def analyze_cases_with_gpt(
    client: Any,
    text: str,
    fachgebiet: str,
    model: str = "gpt-4o-mini"
) -> List[Dict[str, Any]]:
    """
    Analysiert Text mit ChatGPT und extrahiert Falldaten.

    Args:
        client: OpenAI-Client
        text: Zu analysierender Text (aus PDF oder anderem Dokument)
        fachgebiet: Gewähltes Fachgebiet für die Zuordnung
        model: OpenAI-Modell (default: gpt-4o-mini)

    Returns:
        Liste von Fällen als Dictionaries
    """
    bereiche = FAO_CONFIG[fachgebiet].get("bereiche", {})
    bereiche_text = "\n".join([f"  {nr}: {name}" for nr, name in bereiche.items()])

    system_prompt = f"""Du bist ein juristischer Assistent, der Falldaten für einen Fachanwaltsantrag extrahiert.
Das Fachgebiet ist: {fachgebiet}

Die Bereiche für dieses Fachgebiet sind:
{bereiche_text}

Deine Aufgabe ist es, aus dem gegebenen Text alle rechtlichen Fälle zu extrahieren und in ein strukturiertes Format zu bringen.

Für jeden Fall extrahiere folgende Informationen (falls vorhanden):
- kanzlei_az: Kanzlei-Aktenzeichen
- kurzrubrum: Anonymisiertes Rubrum (z.B. "A ./. B")
- sachverhalt: Kurze Sachverhaltsbeschreibung
- zeitraum_von: Beginn (Format: MM/YYYY)
- zeitraum_bis: Ende (Format: MM/YYYY)
- gericht_az: Gerichtsaktenzeichen (falls vorhanden)
- verfahrenstyp: "gerichtlich", "rechtsfoermlich" oder "aussergerichtlich"
- verfahrensart: z.B. "streitig", "fG", "Mahnverfahren", "Eilverfahren"
- bereich_nr: Nummer des Bereichs (1-{len(bereiche)})
- bereich_bezeichnung: Name des Bereichs
- bedeutung: "gering", "mittel" oder "hoch"
- taetigkeitsbeschreibung: Beschreibung der anwaltlichen Tätigkeit
- stand: "anhaengig" oder "abgeschlossen"
- abschluss_art: z.B. "Urteil", "Vergleich", "Beschluss"
- abschluss_datum: Datum (Format: TT.MM.YYYY)

Antworte NUR mit einem JSON-Array von Objekten. Keine zusätzliche Erklärung."""

    user_prompt = f"""Analysiere den folgenden Text und extrahiere alle Fälle im JSON-Format:

{text[:15000]}"""  # Begrenzen auf 15000 Zeichen

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.1,
            max_tokens=4000
        )

        content = response.choices[0].message.content.strip()

        # JSON aus der Antwort extrahieren
        # Manchmal ist die Antwort in Markdown-Code-Blöcken
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]

        cases = json.loads(content)

        if isinstance(cases, dict):
            cases = [cases]

        return cases

    except json.JSONDecodeError as e:
        st.error(f"Fehler beim Parsen der GPT-Antwort: {e}")
        return []
    except Exception as e:
        st.error(f"Fehler bei der GPT-Analyse: {e}")
        return []


def analyze_single_case_with_gpt(
    client: Any,
    text: str,
    fachgebiet: str,
    model: str = "gpt-4o-mini"
) -> Dict[str, Any]:
    """
    Analysiert einen einzelnen Fall-Text mit ChatGPT.
    Nützlich für manuelle Eingabe oder einzelne Dokumente.
    """
    bereiche = FAO_CONFIG[fachgebiet].get("bereiche", {})
    bereiche_text = "\n".join([f"  {nr}: {name}" for nr, name in bereiche.items()])

    system_prompt = f"""Du bist ein juristischer Assistent für Fachanwaltsanträge.
Fachgebiet: {fachgebiet}

Bereiche:
{bereiche_text}

Analysiere den Fall und gib ein JSON-Objekt mit diesen Feldern zurück:
- kanzlei_az, kurzrubrum, sachverhalt, zeitraum_von, zeitraum_bis
- gericht_az, verfahrenstyp, verfahrensart
- bereich_nr (Integer), bereich_bezeichnung
- bedeutung, taetigkeitsbeschreibung, stand
- abschluss_art, abschluss_datum

Antworte NUR mit JSON, keine Erklärung."""

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text}
            ],
            temperature=0.1,
            max_tokens=1000
        )

        content = response.choices[0].message.content.strip()

        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]

        return json.loads(content)

    except Exception as e:
        st.error(f"Fehler bei der Einzelfall-Analyse: {e}")
        return {}


# =============================================================================
# HILFSFUNKTIONEN
# =============================================================================

def map_column_name(col_name: str) -> Optional[str]:
    """Mappt einen Spaltennamen auf den internen Feldnamen."""
    col_normalized = col_name.strip()
    for internal_name, variants in COLUMN_MAPPING.items():
        if col_normalized in variants:
            return internal_name
    return None


def load_cases_from_file(file) -> pd.DataFrame:
    """Liest Fälle aus einer CSV- oder Excel-Datei."""
    filename = file.name.lower()

    if filename.endswith('.csv'):
        df = pd.read_csv(file, encoding='utf-8')
    elif filename.endswith('.xlsx') or filename.endswith('.xls'):
        df = pd.read_excel(file, engine='openpyxl')
    else:
        raise ValueError(f"Nicht unterstütztes Dateiformat: {filename}")

    return df


def cases_list_to_dataframe(cases: List[Dict[str, Any]]) -> pd.DataFrame:
    """Konvertiert eine Liste von Fall-Dictionaries in einen DataFrame."""
    if not cases:
        return pd.DataFrame()
    return pd.DataFrame(cases)


def normalize_case_df(df: pd.DataFrame, fachgebiet: str) -> pd.DataFrame:
    """Normalisiert einen DataFrame auf das interne Datenmodell."""
    # Spalten-Mapping anwenden
    new_columns = {}
    for col in df.columns:
        mapped = map_column_name(col)
        if mapped:
            new_columns[col] = mapped
        else:
            new_columns[col] = col.lower().replace("-", "_").replace(" ", "_")

    df = df.rename(columns=new_columns)

    # Standardfelder ergänzen
    default_values = {
        "fachgebiet": fachgebiet,
        "kanzlei_az": "",
        "kurzrubrum": "",
        "sachverhalt": "",
        "zeitraum_von": "",
        "zeitraum_bis": "",
        "gericht_az": "",
        "verfahrenstyp": "aussergerichtlich",
        "verfahrensart": "",
        "bereich_nr": 0,
        "bereich_bezeichnung": "",
        "bedeutung": "mittel",
        "taetigkeitsbeschreibung": "",
        "stand": "abgeschlossen",
        "abschluss_art": "",
        "abschluss_datum": "",
        "verbundener_fall": "",
    }

    for field, default in default_values.items():
        if field not in df.columns:
            df[field] = default

    df["fachgebiet"] = df["fachgebiet"].fillna(fachgebiet)
    df.loc[df["fachgebiet"] == "", "fachgebiet"] = fachgebiet

    # Verfahrenstyp normalisieren
    df["verfahrenstyp"] = df["verfahrenstyp"].astype(str).str.lower().str.strip()
    df["verfahrenstyp"] = df["verfahrenstyp"].replace({
        "gerichtlich": "gerichtlich",
        "rechtsförmlich": "rechtsfoermlich",
        "rechtsfoermlich": "rechtsfoermlich",
        "außergerichtlich": "aussergerichtlich",
        "aussergerichtlich": "aussergerichtlich",
    })

    df["verfahrensart"] = df["verfahrensart"].astype(str).str.lower().str.strip()

    df["bedeutung"] = df["bedeutung"].astype(str).str.lower().str.strip()
    df["bedeutung"] = df["bedeutung"].replace({
        "niedrig": "gering",
        "hoch": "hoch",
        "mittel": "mittel",
    })

    df["bereich_nr"] = pd.to_numeric(df["bereich_nr"], errors='coerce').fillna(0).astype(int)

    return df


def split_into_falllisten(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Trennt Fälle in gerichtliche/rechtsförmliche und außergerichtliche Verfahren."""
    mask_gerichtlich = df["verfahrenstyp"].isin(["gerichtlich", "rechtsfoermlich"])

    fl1 = df[mask_gerichtlich].copy()
    fl2 = df[~mask_gerichtlich].copy()

    fl1.insert(0, "FL1_Nr", range(1, len(fl1) + 1))
    fl2.insert(0, "FL2_Nr", range(1, len(fl2) + 1))

    return fl1, fl2


def count_by_bereich(df: pd.DataFrame) -> Dict[int, int]:
    """Zählt Fälle pro Bereich-Nummer."""
    return df.groupby("bereich_nr").size().to_dict()


def compute_summary(df: pd.DataFrame, fachgebiet: str) -> Dict[str, Any]:
    """Berechnet Ist-Zahlen und vergleicht mit FAO-Mindestanforderungen."""
    config = FAO_CONFIG.get(fachgebiet, {})

    gesamt = len(df)
    gerichtlich = len(df[df["verfahrenstyp"].isin(["gerichtlich", "rechtsfoermlich"])])
    aussergerichtlich = len(df[df["verfahrenstyp"] == "aussergerichtlich"])
    fg_verfahren = len(df[df["verfahrensart"].str.contains("fg|freiwillige", case=False, na=False)])
    bereich_counts = count_by_bereich(df)

    verbund_count = 0
    if fachgebiet == "Familienrecht" and config.get("verbund_doppelt"):
        verbund_count = len(df[df["verfahrensart"].str.contains("verbund", case=False, na=False)])
        gerichtlich_effektiv = gerichtlich + verbund_count
    else:
        gerichtlich_effektiv = gerichtlich

    checks = []
    warnungen = []

    # Gesamtfallzahl
    gesamt_min = config.get("gesamt_min", 0)
    if gesamt >= gesamt_min:
        checks.append(("Gesamtfallzahl", gesamt, gesamt_min, "erfuellt"))
    elif gesamt >= gesamt_min * 0.9:
        checks.append(("Gesamtfallzahl", gesamt, gesamt_min, "knapp"))
        warnungen.append(f"Gesamtfallzahl knapp unter Minimum: {gesamt}/{gesamt_min}")
    else:
        checks.append(("Gesamtfallzahl", gesamt, gesamt_min, "nicht_erfuellt"))
        warnungen.append(f"Gesamtfallzahl nicht erreicht: {gesamt}/{gesamt_min} (fehlen: {gesamt_min - gesamt})")

    # Gerichtliche Verfahren
    gerichtlich_min = config.get("gerichtlich_min", 0)
    if gerichtlich_min > 0:
        ist_wert = gerichtlich_effektiv if fachgebiet == "Familienrecht" else gerichtlich
        label = "Gerichtliche/rechtsförmliche Verfahren"
        if fachgebiet == "Familienrecht" and verbund_count > 0:
            label += f" (inkl. {verbund_count} Verbund-Bonus)"

        if ist_wert >= gerichtlich_min:
            checks.append((label, ist_wert, gerichtlich_min, "erfuellt"))
        elif ist_wert >= gerichtlich_min * 0.9:
            checks.append((label, ist_wert, gerichtlich_min, "knapp"))
            warnungen.append(f"Gerichtliche Verfahren knapp: {ist_wert}/{gerichtlich_min}")
        else:
            checks.append((label, ist_wert, gerichtlich_min, "nicht_erfuellt"))
            warnungen.append(f"Gerichtliche Verfahren fehlen: {ist_wert}/{gerichtlich_min} (fehlen: {gerichtlich_min - ist_wert})")

    # Außergerichtliche Verfahren
    aussergerichtlich_min = config.get("aussergerichtlich_min", 0)
    if aussergerichtlich_min > 0:
        if aussergerichtlich >= aussergerichtlich_min:
            checks.append(("Außergerichtliche Verfahren", aussergerichtlich, aussergerichtlich_min, "erfuellt"))
        elif aussergerichtlich >= aussergerichtlich_min * 0.9:
            checks.append(("Außergerichtliche Verfahren", aussergerichtlich, aussergerichtlich_min, "knapp"))
            warnungen.append(f"Außergerichtliche Verfahren knapp: {aussergerichtlich}/{aussergerichtlich_min}")
        else:
            checks.append(("Außergerichtliche Verfahren", aussergerichtlich, aussergerichtlich_min, "nicht_erfuellt"))
            warnungen.append(f"Außergerichtliche Verfahren fehlen: {aussergerichtlich}/{aussergerichtlich_min}")

    # fG-Maximum (Erbrecht)
    fg_max = config.get("fg_max")
    if fg_max is not None:
        if fg_verfahren <= fg_max:
            checks.append(("fG-Verfahren (max.)", fg_verfahren, fg_max, "erfuellt"))
        else:
            checks.append(("fG-Verfahren (max.)", fg_verfahren, fg_max, "nicht_erfuellt"))
            warnungen.append(f"Zu viele fG-Verfahren: {fg_verfahren}/{fg_max} (max. erlaubt)")

    # Bereichs-Anforderungen
    bereiche_anf = config.get("bereiche_anforderung", {})

    if "min_bereiche" in bereiche_anf:
        min_bereiche = bereiche_anf["min_bereiche"]
        min_faelle = bereiche_anf.get("min_faelle_pro_bereich", 5)
        bereiche_erfuellt = sum(1 for count in bereich_counts.values() if count >= min_faelle)

        if bereiche_erfuellt >= min_bereiche:
            checks.append((f"Bereiche mit ≥{min_faelle} Fällen", bereiche_erfuellt, min_bereiche, "erfuellt"))
        else:
            checks.append((f"Bereiche mit ≥{min_faelle} Fällen", bereiche_erfuellt, min_bereiche, "nicht_erfuellt"))
            warnungen.append(f"Zu wenige Bereiche mit ausreichend Fällen: {bereiche_erfuellt}/{min_bereiche}")

            for bereich_nr, bereich_name in config.get("bereiche", {}).items():
                count = bereich_counts.get(bereich_nr, 0)
                if count < min_faelle:
                    warnungen.append(f"Bereich {bereich_nr} ({bereich_name}): nur {count}/{min_faelle} Fälle")

    # Spezial-Anforderungen
    spezial = bereiche_anf.get("spezial_anforderung")

    if spezial == "kollektiv":
        kollektiv_bereich = bereiche_anf.get("kollektiv_bereich", 2)
        kollektiv_min = bereiche_anf.get("kollektiv_min", 5)
        kollektiv_count = bereich_counts.get(kollektiv_bereich, 0)

        if kollektiv_count >= kollektiv_min:
            checks.append(("Kollektives Arbeitsrecht", kollektiv_count, kollektiv_min, "erfuellt"))
        else:
            checks.append(("Kollektives Arbeitsrecht", kollektiv_count, kollektiv_min, "nicht_erfuellt"))
            warnungen.append(f"Kollektives Arbeitsrecht: {kollektiv_count}/{kollektiv_min} Fälle")

    elif spezial == "hgr":
        streit_bereich = bereiche_anf.get("streit_bereich", 1)
        gestaltung_bereich = bereiche_anf.get("gestaltung_bereich", 2)
        streit_min = bereiche_anf.get("streit_min", 10)
        gestaltung_min = bereiche_anf.get("gestaltung_min", 10)

        streit_count = bereich_counts.get(streit_bereich, 0)
        gestaltung_count = bereich_counts.get(gestaltung_bereich, 0)

        if streit_count >= streit_min:
            checks.append(("Streitverfahren", streit_count, streit_min, "erfuellt"))
        else:
            checks.append(("Streitverfahren", streit_count, streit_min, "nicht_erfuellt"))
            warnungen.append(f"Streitverfahren: {streit_count}/{streit_min}")

        if gestaltung_count >= gestaltung_min:
            checks.append(("Gestaltung/Gründung", gestaltung_count, gestaltung_min, "erfuellt"))
        else:
            checks.append(("Gestaltung/Gründung", gestaltung_count, gestaltung_min, "nicht_erfuellt"))
            warnungen.append(f"Gestaltung/Gründung: {gestaltung_count}/{gestaltung_min}")

    # Zeitraum-Prüfung (3 Jahre)
    heute = date.today()
    drei_jahre_zuvor = date(heute.year - 3, heute.month, heute.day)

    for idx, row in df.iterrows():
        try:
            abschluss = row.get("abschluss_datum")
            if pd.notna(abschluss) and abschluss != "":
                if isinstance(abschluss, str):
                    for fmt in ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"]:
                        try:
                            parsed = datetime.strptime(abschluss, fmt).date()
                            if parsed < drei_jahre_zuvor:
                                warnungen.append(f"Fall {row.get('kanzlei_az', idx)}: Abschluss außerhalb des 3-Jahres-Zeitraums ({abschluss})")
                            break
                        except ValueError:
                            continue
                elif isinstance(abschluss, (datetime, date)):
                    check_date = abschluss.date() if isinstance(abschluss, datetime) else abschluss
                    if check_date < drei_jahre_zuvor:
                        warnungen.append(f"Fall {row.get('kanzlei_az', idx)}: Abschluss außerhalb des 3-Jahres-Zeitraums")
        except Exception:
            pass

    # Warnung bei hohem Anteil geringer Bedeutung
    gering_count = len(df[df["bedeutung"] == "gering"])
    if gesamt > 0 and gering_count / gesamt > 0.5:
        warnungen.append(f"Hoher Anteil an Fällen mit geringer Bedeutung: {gering_count}/{gesamt} ({gering_count/gesamt*100:.0f}%)")

    return {
        "fachgebiet": fachgebiet,
        "paragraph": config.get("paragraph", ""),
        "gesamt": gesamt,
        "gerichtlich": gerichtlich,
        "gerichtlich_effektiv": gerichtlich_effektiv if fachgebiet == "Familienrecht" else gerichtlich,
        "aussergerichtlich": aussergerichtlich,
        "fg_verfahren": fg_verfahren,
        "verbund_count": verbund_count,
        "bereich_counts": bereich_counts,
        "bereiche_config": config.get("bereiche", {}),
        "checks": checks,
        "warnungen": warnungen,
    }


def format_zeitraum(von, bis) -> str:
    """Formatiert Zeitraum für Anzeige."""
    von_str = str(von) if pd.notna(von) else ""
    bis_str = str(bis) if pd.notna(bis) else ""

    if von_str and bis_str:
        return f"{von_str} - {bis_str}"
    elif von_str:
        return f"ab {von_str}"
    elif bis_str:
        return f"bis {bis_str}"
    return ""


def prepare_fl1_for_export(fl1: pd.DataFrame) -> pd.DataFrame:
    """Bereitet Fallliste 1 für Excel-Export vor."""
    export_df = pd.DataFrame({
        "FL1-Nr": fl1["FL1_Nr"],
        "Kurzrubrum": fl1["kurzrubrum"],
        "Kanzlei-AZ": fl1["kanzlei_az"],
        "Gerichts-AZ": fl1["gericht_az"],
        "Bereich-Nr": fl1["bereich_nr"],
        "Bereich-Bezeichnung": fl1["bereich_bezeichnung"],
        "Verfahrenstyp": fl1["verfahrenstyp"],
        "Verfahrensart": fl1["verfahrensart"],
        "Zeitraum": fl1.apply(lambda x: format_zeitraum(x["zeitraum_von"], x["zeitraum_bis"]), axis=1),
        "Bedeutung": fl1["bedeutung"],
        "Taetigkeit": fl1["taetigkeitsbeschreibung"],
        "Stand": fl1["stand"],
        "Abschluss-Art": fl1["abschluss_art"],
        "Abschluss-Datum": fl1["abschluss_datum"],
        "Verbundener-Fall": fl1["verbundener_fall"],
    })
    return export_df


def prepare_fl2_for_export(fl2: pd.DataFrame) -> pd.DataFrame:
    """Bereitet Fallliste 2 für Excel-Export vor."""
    export_df = pd.DataFrame({
        "FL2-Nr": fl2["FL2_Nr"],
        "Kurzrubrum": fl2["kurzrubrum"],
        "Kanzlei-AZ": fl2["kanzlei_az"],
        "Bereich-Nr": fl2["bereich_nr"],
        "Bereich-Bezeichnung": fl2["bereich_bezeichnung"],
        "Zeitraum": fl2.apply(lambda x: format_zeitraum(x["zeitraum_von"], x["zeitraum_bis"]), axis=1),
        "Bedeutung": fl2["bedeutung"],
        "Taetigkeit": fl2["taetigkeitsbeschreibung"],
        "Verbundener-Fall": fl2["verbundener_fall"],
    })
    return export_df


def create_summary_df(summary: Dict[str, Any]) -> pd.DataFrame:
    """Erstellt DataFrame für Summary-Tabellenblatt."""
    rows = []

    rows.append({"Kategorie": "FACHGEBIET", "Wert": summary["fachgebiet"], "Anforderung": summary["paragraph"]})
    rows.append({"Kategorie": "", "Wert": "", "Anforderung": ""})
    rows.append({"Kategorie": "=== KENNZAHLEN ===", "Wert": "", "Anforderung": ""})

    for check in summary["checks"]:
        kriterium, ist, soll, status = check
        status_symbol = {"erfuellt": "✓", "knapp": "⚠️", "nicht_erfuellt": "✗"}.get(status, "?")
        rows.append({
            "Kategorie": kriterium,
            "Wert": f"{ist}/{soll}",
            "Anforderung": status_symbol
        })

    rows.append({"Kategorie": "", "Wert": "", "Anforderung": ""})
    rows.append({"Kategorie": "=== BEREICHSVERTEILUNG ===", "Wert": "", "Anforderung": ""})

    for bereich_nr, bereich_name in summary.get("bereiche_config", {}).items():
        count = summary["bereich_counts"].get(bereich_nr, 0)
        rows.append({
            "Kategorie": f"Bereich {bereich_nr}",
            "Wert": count,
            "Anforderung": bereich_name
        })

    rows.append({"Kategorie": "", "Wert": "", "Anforderung": ""})

    if summary["warnungen"]:
        rows.append({"Kategorie": "=== WARNUNGEN ===", "Wert": "", "Anforderung": ""})
        for warnung in summary["warnungen"]:
            rows.append({"Kategorie": "⚠️", "Wert": warnung, "Anforderung": ""})
    else:
        rows.append({"Kategorie": "=== STATUS ===", "Wert": "Keine Warnungen", "Anforderung": "✓"})

    return pd.DataFrame(rows)


def create_excel(fl1_df: pd.DataFrame, fl2_df: pd.DataFrame,
                 summary: Dict[str, Any], fachgebiet: str) -> bytes:
    """Erstellt Excel-Arbeitsmappe mit Falllisten und Summary."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        fl1_export = prepare_fl1_for_export(fl1_df) if len(fl1_df) > 0 else pd.DataFrame()
        fl1_export.to_excel(writer, sheet_name="Fallliste_1", index=False)

        fl2_export = prepare_fl2_for_export(fl2_df) if len(fl2_df) > 0 else pd.DataFrame()
        fl2_export.to_excel(writer, sheet_name="Fallliste_2", index=False)

        summary_df = create_summary_df(summary)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    return output.getvalue()


def create_template_excel() -> bytes:
    """Erstellt eine leere Excel-Vorlage mit den erwarteten Spalten."""
    output = BytesIO()
    df = pd.DataFrame(columns=TEMPLATE_COLUMNS)

    beispiel = {
        "Kanzlei-AZ": "2024/001",
        "Kurzrubrum": "A ./. B (Nachlass)",
        "Sachverhalt": "Erbrechtliche Auseinandersetzung bzgl. Testament",
        "Zeitraum_von": "01/2024",
        "Zeitraum_bis": "06/2024",
        "Gerichts-AZ": "12 O 123/24",
        "Verfahrenstyp": "gerichtlich",
        "Verfahrensart": "streitig",
        "Bereich_Nr": 1,
        "Bereich_Bezeichnung": "Materielles Erbrecht",
        "Bedeutung": "mittel",
        "Taetigkeitsbeschreibung": "Beratung, Klageschrift, Verhandlung, Vergleich",
        "Stand": "abgeschlossen",
        "Abschluss_Art": "Vergleich",
        "Abschluss_Datum": "15.06.2024",
        "Verbundener_Fall": "",
        "Fachgebiet": "Erbrecht"
    }
    df = pd.concat([df, pd.DataFrame([beispiel])], ignore_index=True)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Faelle", index=False)

        erklaerungen = pd.DataFrame({
            "Spalte": TEMPLATE_COLUMNS,
            "Beschreibung": [
                "Internes Aktenzeichen der Kanzlei",
                "Anonymisiertes, aber individuelles Rubrum",
                "Kurze Sachverhaltsbeschreibung",
                "Beginn des Mandats (MM/JJJJ)",
                "Ende des Mandats (MM/JJJJ)",
                "Aktenzeichen des Gerichts (falls vorhanden)",
                "gerichtlich / rechtsfoermlich / aussergerichtlich",
                "streitig / fG / Mahnverfahren / Eilverfahren / etc.",
                "Nummer des Bereichs gem. FAO (1-5 je nach Fachgebiet)",
                "Bezeichnung des Bereichs",
                "gering / mittel / hoch",
                "Individuelle Beschreibung der Tätigkeit",
                "anhaengig / abgeschlossen",
                "Urteil / Beschluss / Vergleich / Klagerücknahme / etc.",
                "Datum des Abschlusses (TT.MM.JJJJ)",
                "Kanzlei-AZ des verbundenen Falls (falls vorhanden)",
                "Erbrecht / Arbeitsrecht / etc. (optional wenn alle Fälle gleich)"
            ]
        })
        erklaerungen.to_excel(writer, sheet_name="Anleitung", index=False)

    return output.getvalue()


def get_status_symbol(status: str) -> str:
    """Gibt Symbol für Status zurück."""
    return {"erfuellt": "✓", "knapp": "⚠️", "nicht_erfuellt": "✗"}.get(status, "?")


# =============================================================================
# STREAMLIT APP
# =============================================================================

def main():
    """Hauptfunktion der Streamlit-App."""

    st.set_page_config(
        page_title="Fachanwalt Falllistenverwaltung",
        page_icon="⚖️",
        layout="wide"
    )

    st.title("⚖️ Fachanwalt-Falllistenverwaltung")
    st.markdown("**Erstellen Sie FAO-konforme Falllisten für Ihren Fachanwaltsantrag**")

    # Session State initialisieren
    if "cases_df" not in st.session_state:
        st.session_state.cases_df = pd.DataFrame()
    if "openai_api_key" not in st.session_state:
        st.session_state.openai_api_key = ""

    # Sidebar
    with st.sidebar:
        st.header("⚙️ Einstellungen")

        # Fachgebiet-Auswahl
        fachgebiet = st.selectbox(
            "Fachgebiet auswählen",
            options=list(FAO_CONFIG.keys()),
            help="Wählen Sie das Fachgebiet für die Falllistenerstellung"
        )

        # OpenAI API-Key Eingabe
        st.markdown("---")
        st.subheader("🤖 ChatGPT-Verbindung")

        api_key = st.text_input(
            "OpenAI API-Key",
            type="password",
            value=st.session_state.openai_api_key,
            help="Für die automatische Fallanalyse. Key unter platform.openai.com erstellen."
        )

        if api_key:
            st.session_state.openai_api_key = api_key
            st.success("✓ API-Key aktiv")

        # Modell-Auswahl
        gpt_model = st.selectbox(
            "GPT-Modell",
            options=["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-3.5-turbo"],
            help="gpt-4o-mini: schnell & günstig | gpt-4o: beste Qualität"
        )

        # Hinweis zu Mindestanforderungen
        st.markdown("---")
        st.subheader("📌 FAO-Anforderungen")
        config = FAO_CONFIG[fachgebiet]
        st.markdown(f"**{config['paragraph']}**")
        st.info(config["hinweis"])

        if "bereiche" in config:
            with st.expander("Bereiche anzeigen"):
                for nr, name in config["bereiche"].items():
                    st.markdown(f"**{nr}.** {name}")

        # Optionale Muster-Vorlage
        st.markdown("---")
        st.subheader("📋 Muster-Vorlage")
        st.caption("Optional: Excel-Muster zum Ansehen des erwarteten Formats")
        template_bytes = create_template_excel()
        st.download_button(
            label="Muster-Excel ansehen",
            data=template_bytes,
            file_name="muster_fallliste.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Zeigt das Format, das ChatGPT automatisch aus Ihren PDFs erstellt"
        )

    # =========================================================================
    # HAUPTBEREICH
    # =========================================================================

    # Kurze Erklärung
    st.markdown("""
    ### So funktioniert's:
    1. **PDF-Dateien hochladen** (Mandatsübersichten, Fallberichte, Aktenverzeichnisse)
    2. **ChatGPT analysiert** automatisch und extrahiert die Falldaten
    3. **FAO-Check** zeigt, ob alle Anforderungen erfüllt sind
    4. **Fertige Fallliste herunterladen** zur Einreichung bei der Kammer
    """)

    st.markdown("---")

    # Tabs für Eingabemethoden - PDF zuerst!
    tab_pdf, tab_manual = st.tabs([
        "📕 PDF hochladen (empfohlen)",
        "✏️ Fall manuell eingeben"
    ])

    all_cases = []

    # Tab 1: PDF Upload mit KI-Analyse (Hauptmethode)
    with tab_pdf:
        st.subheader("PDF-Dokumente mit KI analysieren")

        if not PDF_AVAILABLE:
            st.error("⚠️ PyPDF2 fehlt. Installation: `pip install pypdf2`")
        elif not OPENAI_AVAILABLE:
            st.error("⚠️ OpenAI fehlt. Installation: `pip install openai`")
        elif not st.session_state.openai_api_key:
            st.warning("⚠️ Bitte geben Sie Ihren OpenAI API-Key in der Sidebar ein, um PDFs zu analysieren.")
            st.markdown("""
            **So erhalten Sie einen API-Key:**
            1. Gehen Sie zu [platform.openai.com](https://platform.openai.com)
            2. Erstellen Sie ein Konto oder melden Sie sich an
            3. Navigieren Sie zu API Keys → Create new secret key
            4. Kopieren Sie den Key in das Feld links in der Sidebar
            """)
        else:
            st.markdown("""
            Laden Sie Ihre **Mandatsübersichten, Fallberichte oder Aktenverzeichnisse** hoch.
            ChatGPT liest die Dokumente und extrahiert automatisch alle relevanten Falldaten.
            """)

            pdf_files = st.file_uploader(
                "PDF-Dateien auswählen",
                type=["pdf"],
                accept_multiple_files=True,
                key="pdf_upload",
                help="Sie können mehrere PDFs gleichzeitig hochladen"
            )

            if pdf_files:
                if st.button("🤖 PDFs mit ChatGPT analysieren", type="primary", use_container_width=True):
                    client = get_openai_client(st.session_state.openai_api_key)

                    if client:
                        progress_bar = st.progress(0)
                        for i, pdf_file in enumerate(pdf_files):
                            with st.spinner(f"Analysiere {pdf_file.name}..."):
                                try:
                                    pdf_text = extract_text_from_pdf(pdf_file)

                                    if len(pdf_text.strip()) < 50:
                                        st.warning(f"⚠️ {pdf_file.name}: Wenig Text gefunden. Gescanntes PDF?")
                                        continue

                                    cases = analyze_cases_with_gpt(
                                        client,
                                        pdf_text,
                                        fachgebiet,
                                        model=gpt_model
                                    )

                                    if cases:
                                        df = cases_list_to_dataframe(cases)
                                        df = normalize_case_df(df, fachgebiet)
                                        all_cases.append(df)
                                        st.success(f"✓ {pdf_file.name}: **{len(cases)} Fälle** erkannt")

                                        with st.expander(f"Details: {pdf_file.name}"):
                                            st.dataframe(
                                                df[["kurzrubrum", "verfahrenstyp", "bereich_nr", "bedeutung"]],
                                                use_container_width=True,
                                                hide_index=True
                                            )
                                    else:
                                        st.warning(f"⚠️ {pdf_file.name}: Keine Fälle erkannt")

                                except Exception as e:
                                    st.error(f"✗ {pdf_file.name}: {str(e)}")

                            progress_bar.progress((i + 1) / len(pdf_files))

    # Tab 2: Manueller Eintrag
    with tab_manual:
        st.subheader("Einzelnen Fall eingeben")

        manual_method = st.radio(
            "Eingabeart",
            ["Freitext (KI-Analyse)", "Strukturierte Eingabe"],
            horizontal=True,
            help="Freitext: Beschreiben Sie den Fall, ChatGPT füllt die Felder aus"
        )

        if manual_method == "Freitext (KI-Analyse)":
            if not st.session_state.openai_api_key:
                st.warning("⚠️ Für die KI-Analyse bitte OpenAI API-Key in der Sidebar eingeben.")
            else:
                freitext = st.text_area(
                    "Fallbeschreibung",
                    placeholder="""Beschreiben Sie den Fall frei, z.B.:

Im Januar 2024 beauftragte mich Mandant A mit der Durchsetzung seiner Erbansprüche gegen die Erbengemeinschaft B, C und D. Es ging um ein Nachlassvermögen von ca. 500.000 EUR. Nach außergerichtlicher Korrespondenz und gescheiterten Vergleichsverhandlungen erhob ich Klage beim LG München (Az. 12 O 4567/24). Das Verfahren endete im Juni 2024 durch Vergleich.""",
                    height=200
                )

                if st.button("🤖 Mit ChatGPT analysieren", type="primary"):
                    if freitext:
                        client = get_openai_client(st.session_state.openai_api_key)
                        if client:
                            with st.spinner("ChatGPT analysiert..."):
                                case_data = analyze_single_case_with_gpt(
                                    client, freitext, fachgebiet, gpt_model
                                )

                                if case_data:
                                    new_df = pd.DataFrame([case_data])
                                    new_df = normalize_case_df(new_df, fachgebiet)
                                    all_cases.append(new_df)
                                    st.success("✓ Fall erkannt und hinzugefügt!")

                                    with st.expander("Erkannte Daten"):
                                        st.json(case_data)
                    else:
                        st.warning("Bitte geben Sie eine Fallbeschreibung ein.")

        else:  # Strukturierte Eingabe
            st.markdown("Füllen Sie die Felder manuell aus:")

            col1, col2 = st.columns(2)

            with col1:
                m_kanzlei_az = st.text_input("Kanzlei-AZ", placeholder="2024/001")
                m_kurzrubrum = st.text_input("Kurzrubrum", placeholder="A ./. B")
                m_sachverhalt = st.text_area("Sachverhalt", placeholder="Kurze Beschreibung...", height=100)
                m_zeitraum_von = st.text_input("Zeitraum von", placeholder="01/2024")
                m_zeitraum_bis = st.text_input("Zeitraum bis", placeholder="06/2024")
                m_gericht_az = st.text_input("Gerichts-AZ", placeholder="12 O 123/24")

            with col2:
                m_verfahrenstyp = st.selectbox("Verfahrenstyp", ["gerichtlich", "rechtsfoermlich", "aussergerichtlich"])
                m_verfahrensart = st.text_input("Verfahrensart", placeholder="streitig / fG / Mahnverfahren")
                m_bereich_nr = st.selectbox("Bereich-Nr", list(config["bereiche"].keys()))
                m_bereich_bez = config["bereiche"].get(m_bereich_nr, "")
                st.text_input("Bereich-Bezeichnung", value=m_bereich_bez, disabled=True)
                m_bedeutung = st.selectbox("Bedeutung", ["gering", "mittel", "hoch"])
                m_stand = st.selectbox("Stand", ["abgeschlossen", "anhaengig"])

            m_taetigkeit = st.text_area("Tätigkeitsbeschreibung", placeholder="Beratung, Klageschrift, Verhandlung...", height=80)

            col3, col4 = st.columns(2)
            with col3:
                m_abschluss_art = st.text_input("Abschluss-Art", placeholder="Vergleich / Urteil / ...")
            with col4:
                m_abschluss_datum = st.text_input("Abschluss-Datum", placeholder="15.06.2024")

            if st.button("➕ Fall hinzufügen", type="primary"):
                new_case = {
                    "kanzlei_az": m_kanzlei_az,
                    "kurzrubrum": m_kurzrubrum,
                    "sachverhalt": m_sachverhalt,
                    "zeitraum_von": m_zeitraum_von,
                    "zeitraum_bis": m_zeitraum_bis,
                    "gericht_az": m_gericht_az,
                    "verfahrenstyp": m_verfahrenstyp,
                    "verfahrensart": m_verfahrensart,
                    "bereich_nr": m_bereich_nr,
                    "bereich_bezeichnung": m_bereich_bez,
                    "bedeutung": m_bedeutung,
                    "taetigkeitsbeschreibung": m_taetigkeit,
                    "stand": m_stand,
                    "abschluss_art": m_abschluss_art,
                    "abschluss_datum": m_abschluss_datum,
                    "fachgebiet": fachgebiet,
                    "verbundener_fall": ""
                }

                new_df = pd.DataFrame([new_case])
                all_cases.append(new_df)
                st.success("✓ Fall hinzugefügt!")

    # =========================================================================
    # FÄLLE VERARBEITEN UND ANZEIGEN
    # =========================================================================

    # Alle neuen Fälle mit bestehenden kombinieren
    if all_cases:
        new_combined = pd.concat(all_cases, ignore_index=True)
        if not st.session_state.cases_df.empty:
            st.session_state.cases_df = pd.concat(
                [st.session_state.cases_df, new_combined],
                ignore_index=True
            )
        else:
            st.session_state.cases_df = new_combined

    # Wenn keine Fälle vorhanden
    if st.session_state.cases_df.empty:
        st.markdown("---")
        st.info("👆 Laden Sie PDFs hoch oder geben Sie Fälle manuell ein, um Ihre Fallliste zu erstellen.")
        return

    # Nach Fachgebiet filtern
    filtered_df = st.session_state.cases_df[
        st.session_state.cases_df["fachgebiet"] == fachgebiet
    ].copy()

    if len(filtered_df) == 0:
        filtered_df = st.session_state.cases_df.copy()

    # =========================================================================
    # ERGEBNISSE ANZEIGEN
    # =========================================================================

    st.markdown("---")

    # Falllisten aufteilen
    fl1, fl2 = split_into_falllisten(filtered_df)
    summary = compute_summary(filtered_df, fachgebiet)

    # Dashboard mit Kennzahlen
    st.header("📊 Ihre Fallliste - Übersicht")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        delta_gesamt = summary["gesamt"] - FAO_CONFIG[fachgebiet]['gesamt_min']
        st.metric(
            label="Gesamtfälle",
            value=summary["gesamt"],
            delta=f"{delta_gesamt:+d} zum Soll" if delta_gesamt != 0 else "Soll erreicht",
            delta_color="normal" if delta_gesamt >= 0 else "inverse"
        )

    with col2:
        soll_ger = FAO_CONFIG[fachgebiet].get('gerichtlich_min', 0)
        delta_ger = summary["gerichtlich"] - soll_ger
        st.metric(
            label="Gerichtlich",
            value=summary["gerichtlich"],
            delta=f"{delta_ger:+d}" if soll_ger > 0 else None,
            delta_color="normal" if delta_ger >= 0 else "inverse"
        )

    with col3:
        soll_ag = FAO_CONFIG[fachgebiet].get('aussergerichtlich_min', 0)
        if soll_ag > 0:
            delta_ag = summary["aussergerichtlich"] - soll_ag
            st.metric(
                label="Außergerichtlich",
                value=summary["aussergerichtlich"],
                delta=f"{delta_ag:+d}",
                delta_color="normal" if delta_ag >= 0 else "inverse"
            )
        else:
            st.metric(label="Außergerichtlich", value=summary["aussergerichtlich"])

    with col4:
        erfuellt_count = sum(1 for c in summary["checks"] if c[3] == "erfuellt")
        total_checks = len(summary["checks"])
        st.metric(
            label="FAO-Kriterien",
            value=f"{erfuellt_count}/{total_checks}",
            delta="Alle erfüllt ✓" if erfuellt_count == total_checks else f"{total_checks - erfuellt_count} offen"
        )

    st.markdown("---")

    # FAO-Konformitätscheck
    st.header("✅ FAO-Konformitätscheck")

    check_cols = st.columns(2)

    for i, check in enumerate(summary["checks"]):
        kriterium, ist, soll, status = check
        col = check_cols[i % 2]

        with col:
            symbol = get_status_symbol(status)
            if status == "erfuellt":
                st.success(f"{symbol} **{kriterium}**: {ist}/{soll}")
            elif status == "knapp":
                st.warning(f"{symbol} **{kriterium}**: {ist}/{soll}")
            else:
                st.error(f"{symbol} **{kriterium}**: {ist}/{soll}")

    # Warnungen (falls vorhanden)
    if summary["warnungen"]:
        st.markdown("---")
        st.subheader("⚠️ Hinweise")
        for warnung in summary["warnungen"]:
            st.warning(warnung)

    # Bereichsverteilung
    st.markdown("---")
    st.header("📈 Bereichsverteilung")

    bereiche_config = FAO_CONFIG[fachgebiet].get("bereiche", {})
    bereich_data = []

    for bereich_nr, bereich_name in bereiche_config.items():
        count = summary["bereich_counts"].get(bereich_nr, 0)
        bereich_data.append({
            "Bereich": f"{bereich_nr}. {bereich_name}",
            "Anzahl": count
        })

    if bereich_data:
        bereich_df = pd.DataFrame(bereich_data)
        st.bar_chart(bereich_df.set_index("Bereich"))

    # Falllisten-Vorschau
    st.markdown("---")
    st.header("📋 Vorschau der Falllisten")

    preview_tab1, preview_tab2 = st.tabs([
        f"Fallliste 1 - Gerichtlich/Rechtsförmlich ({len(fl1)})",
        f"Fallliste 2 - Außergerichtlich ({len(fl2)})"
    ])

    with preview_tab1:
        if len(fl1) > 0:
            display_cols = ["FL1_Nr", "kurzrubrum", "kanzlei_az", "gericht_az",
                          "bereich_nr", "verfahrenstyp", "bedeutung", "stand"]
            display_cols = [c for c in display_cols if c in fl1.columns]
            st.dataframe(fl1[display_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Keine gerichtlichen/rechtsförmlichen Verfahren vorhanden.")

    with preview_tab2:
        if len(fl2) > 0:
            display_cols = ["FL2_Nr", "kurzrubrum", "kanzlei_az", "bereich_nr", "bedeutung"]
            display_cols = [c for c in display_cols if c in fl2.columns]
            st.dataframe(fl2[display_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Keine außergerichtlichen Verfahren vorhanden.")

    # =========================================================================
    # EXPORT - FERTIGE FALLLISTE ZUM EINREICHEN
    # =========================================================================

    st.markdown("---")
    st.header("📥 Fertige Fallliste herunterladen")

    st.markdown("""
    Die Excel-Datei enthält Ihre **einreichungsfertige Fallliste** mit:
    - **Fallliste 1**: Gerichtliche/rechtsförmliche Verfahren
    - **Fallliste 2**: Außergerichtliche Verfahren
    - **Summary**: Übersicht mit FAO-Check und Statistiken
    """)

    excel_bytes = create_excel(fl1, fl2, summary, fachgebiet)
    filename = f"Fallliste_{fachgebiet.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"

    col1, col2, col3 = st.columns([2, 1, 1])

    with col1:
        st.download_button(
            label="📥 FALLLISTE HERUNTERLADEN (Excel)",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    with col2:
        st.metric("Fälle gesamt", summary["gesamt"])

    with col3:
        if st.button("🗑️ Zurücksetzen", help="Alle Fälle löschen und neu beginnen"):
            st.session_state.cases_df = pd.DataFrame()
            st.rerun()


if __name__ == "__main__":
    main()
