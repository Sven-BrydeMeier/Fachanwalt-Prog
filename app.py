"""
Fachanwalt-Falllistenverwaltung - Streamlit App
================================================
Eine Anwendung zur Verwaltung von Fachanwaltsfällen mit FAO-Konformitätsprüfung.
Ermöglicht Upload, Sortierung und Export von Falllisten gemäß § 5 FAO.
Unterstützt PDF-Upload mit automatischer Fallauswertung via ChatGPT/OpenAI.

Verwendung: streamlit run app.py
Voraussetzungen: pip install streamlit pandas openpyxl openai pypdf2
"""

# =============================================================================
# VERSION
# =============================================================================
APP_VERSION = "25.12.11-15:17"

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date
from typing import Dict, List, Tuple, Optional, Any
import json
import re
import time
import requests
from urllib.parse import urlparse, parse_qs

# =============================================================================
# UPLOAD-LIMITS FÜR STREAMLIT CLOUD
# =============================================================================

# Maximale Dateigröße pro Upload in Bytes (100 MB pro Datei)
MAX_FILE_SIZE_MB = 100
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

# Maximale Gesamtgröße aller Uploads in einer Session (500 MB)
MAX_TOTAL_UPLOAD_MB = 500
MAX_TOTAL_UPLOAD_BYTES = MAX_TOTAL_UPLOAD_MB * 1024 * 1024

# Chunk-Größe für Cloud-Downloads (5 MB)
CLOUD_CHUNK_SIZE = 5 * 1024 * 1024

# PDF-Verarbeitung
try:
    from PyPDF2 import PdfReader
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# OpenAI-Integration
try:
    from openai import OpenAI, APIError, APIConnectionError, RateLimitError, APITimeoutError
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


def check_file_size(file) -> Tuple[bool, str]:
    """
    Prüft ob eine Datei das Upload-Limit überschreitet.

    Returns:
        Tuple[bool, str]: (ist_ok, fehlermeldung)
    """
    try:
        # Dateigröße ermitteln
        file.seek(0, 2)  # Ans Ende springen
        file_size = file.tell()
        file.seek(0)  # Zurück zum Anfang

        if file_size > MAX_FILE_SIZE_BYTES:
            size_mb = file_size / (1024 * 1024)
            return False, f"Datei zu groß: {size_mb:.1f} MB (max. {MAX_FILE_SIZE_MB} MB erlaubt)"

        return True, ""
    except Exception as e:
        return False, f"Fehler bei Größenprüfung: {str(e)}"


def check_total_upload_size(files: List) -> Tuple[bool, str, float]:
    """
    Prüft ob die Gesamtgröße aller Dateien das Limit überschreitet.

    Returns:
        Tuple[bool, str, float]: (ist_ok, fehlermeldung, gesamtgroesse_mb)
    """
    total_size = 0
    for f in files:
        try:
            f.seek(0, 2)
            total_size += f.tell()
            f.seek(0)
        except:
            pass

    total_mb = total_size / (1024 * 1024)

    if total_size > MAX_TOTAL_UPLOAD_BYTES:
        return False, f"Gesamtgröße zu hoch: {total_mb:.1f} MB (max. {MAX_TOTAL_UPLOAD_MB} MB)", total_mb

    return True, "", total_mb


# =============================================================================
# CLOUD-DRIVE INTEGRATION (iCloud, Google Drive)
# =============================================================================

def parse_cloud_link(url: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Erkennt und parst Cloud-Drive Freigabe-Links.

    Returns:
        Tuple[Optional[str], Optional[str]]: (provider, download_url oder file_id)
        Provider kann sein: "google_drive", "google_drive_folder", "dropbox", "icloud"
    """
    url = url.strip()
    parsed = urlparse(url)

    # Google Drive
    if "drive.google.com" in parsed.netloc:
        # Ordner-Format: https://drive.google.com/drive/folders/FOLDER_ID
        if "/folders/" in url:
            parts = url.split("/folders/")
            if len(parts) > 1:
                folder_id = parts[1].split("/")[0].split("?")[0]
                return "google_drive_folder", folder_id
        # Datei-Format: https://drive.google.com/file/d/FILE_ID/view
        if "/file/d/" in url:
            parts = url.split("/file/d/")
            if len(parts) > 1:
                file_id = parts[1].split("/")[0].split("?")[0]
                return "google_drive", file_id
        # Format: https://drive.google.com/open?id=FILE_ID
        elif "id=" in url:
            query = parse_qs(parsed.query)
            if "id" in query:
                return "google_drive", query["id"][0]

    # iCloud
    if "icloud.com" in parsed.netloc:
        # iCloud-Links müssen über die API aufgelöst werden
        return "icloud", url

    # Dropbox
    if "dropbox.com" in parsed.netloc:
        # Dropbox-Links auf direkten Download umwandeln
        if "dl=0" in url:
            direct_url = url.replace("dl=0", "dl=1")
        elif "dl=1" not in url:
            direct_url = url + ("&dl=1" if "?" in url else "?dl=1")
        else:
            direct_url = url
        return "dropbox", direct_url

    return None, None


def list_google_drive_folder(folder_id: str) -> List[Dict[str, str]]:
    """
    Listet alle PDF-Dateien in einem öffentlichen Google Drive Ordner auf.

    Returns:
        Liste von Dictionaries mit 'id', 'name' für jede PDF-Datei
    """
    # Google Drive API für öffentliche Ordner
    # Verwendet die eingebettete Ansicht, um Dateiliste zu erhalten
    api_url = f"https://drive.google.com/embeddedfolderview?id={folder_id}"

    try:
        response = requests.get(api_url, timeout=30)
        if response.status_code != 200:
            return []

        # HTML parsen um Datei-IDs zu extrahieren
        content = response.text
        files = []

        # Suche nach Datei-Links im Format /file/d/FILE_ID
        file_pattern = r'/file/d/([a-zA-Z0-9_-]+)'
        file_ids = set(re.findall(file_pattern, content))

        for file_id in file_ids:
            files.append({
                'id': file_id,
                'name': f'Datei_{file_id[:8]}.pdf'  # Placeholder-Name
            })

        # Wenn keine Dateien gefunden, versuche alternative Methode
        if not files:
            # Versuche über die Google Drive API (ohne Auth für öffentliche Ordner)
            api_url2 = f"https://www.googleapis.com/drive/v3/files?q='{folder_id}'+in+parents&fields=files(id,name,mimeType)"
            try:
                resp2 = requests.get(api_url2, timeout=30)
                if resp2.status_code == 200:
                    data = resp2.json()
                    for f in data.get('files', []):
                        if f.get('mimeType') == 'application/pdf' or f.get('name', '').lower().endswith('.pdf'):
                            files.append({'id': f['id'], 'name': f['name']})
            except:
                pass

        return files

    except Exception as e:
        st.error(f"Fehler beim Laden des Ordners: {str(e)}")
        return []


def get_google_drive_folder_files_via_webpage(folder_id: str) -> List[Dict[str, str]]:
    """
    Alternative Methode: Lädt Ordnerinhalt über die Webseite.
    """
    files = []

    # Versuche über die normale Drive-Ansicht
    url = f"https://drive.google.com/drive/folders/{folder_id}"

    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=30)

        if response.status_code == 200:
            # Suche nach Datei-IDs und Namen
            # Google Drive speichert Daten in JavaScript-Objekten
            content = response.text

            # Muster für Datei-Einträge
            file_pattern = r'\["([a-zA-Z0-9_-]{25,})","([^"]+\.pdf)"'
            matches = re.findall(file_pattern, content, re.IGNORECASE)

            for file_id, file_name in matches:
                files.append({'id': file_id, 'name': file_name})

            # Fallback: Nur IDs extrahieren
            if not files:
                id_pattern = r'/file/d/([a-zA-Z0-9_-]{25,})'
                ids = set(re.findall(id_pattern, content))
                for idx, fid in enumerate(ids):
                    files.append({'id': fid, 'name': f'Datei_{idx+1}.pdf'})

    except Exception as e:
        pass

    return files


def download_from_google_drive(file_id: str, progress_callback=None) -> Optional[BytesIO]:
    """
    Lädt eine Datei von Google Drive herunter (öffentliche Freigabe).
    Unterstützt Chunk-basiertes Laden für große Dateien.
    """
    # Direkt-Download-URL
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    try:
        session = requests.Session()

        # Erste Anfrage - prüft auf Virenwarnungsseite bei großen Dateien
        response = session.get(download_url, stream=True, timeout=30)

        # Google Drive Virenscan-Warnung umgehen
        for key, value in response.cookies.items():
            if key.startswith('download_warning'):
                download_url = f"https://drive.google.com/uc?export=download&confirm={value}&id={file_id}"
                response = session.get(download_url, stream=True, timeout=30)
                break

        if response.status_code != 200:
            return None

        # Content-Length prüfen
        content_length = response.headers.get('content-length')
        if content_length:
            file_size = int(content_length)
            if file_size > MAX_TOTAL_UPLOAD_BYTES:
                raise ValueError(f"Datei zu groß: {file_size / (1024*1024):.1f} MB")

        # Chunk-basiertes Laden
        buffer = BytesIO()
        downloaded = 0

        for chunk in response.iter_content(chunk_size=CLOUD_CHUNK_SIZE):
            if chunk:
                buffer.write(chunk)
                downloaded += len(chunk)

                # Größenlimit prüfen
                if downloaded > MAX_TOTAL_UPLOAD_BYTES:
                    raise ValueError(f"Download überschreitet Limit von {MAX_TOTAL_UPLOAD_MB} MB")

                # Progress-Callback
                if progress_callback and content_length:
                    progress_callback(downloaded / int(content_length))

        buffer.seek(0)
        return buffer

    except requests.RequestException as e:
        st.error(f"Download-Fehler: {str(e)}")
        return None
    except ValueError as e:
        st.error(str(e))
        return None


def download_from_dropbox(url: str, progress_callback=None) -> Optional[BytesIO]:
    """
    Lädt eine Datei von Dropbox herunter.
    """
    try:
        response = requests.get(url, stream=True, timeout=60)

        if response.status_code != 200:
            return None

        content_length = response.headers.get('content-length')
        if content_length and int(content_length) > MAX_TOTAL_UPLOAD_BYTES:
            raise ValueError(f"Datei zu groß: {int(content_length) / (1024*1024):.1f} MB")

        buffer = BytesIO()
        downloaded = 0

        for chunk in response.iter_content(chunk_size=CLOUD_CHUNK_SIZE):
            if chunk:
                buffer.write(chunk)
                downloaded += len(chunk)

                if downloaded > MAX_TOTAL_UPLOAD_BYTES:
                    raise ValueError(f"Download überschreitet Limit von {MAX_TOTAL_UPLOAD_MB} MB")

                if progress_callback and content_length:
                    progress_callback(downloaded / int(content_length))

        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"Dropbox-Download-Fehler: {str(e)}")
        return None


def download_from_cloud(url: str, progress_callback=None) -> Tuple[Optional[BytesIO], str]:
    """
    Universelle Funktion zum Download von Cloud-Speichern.

    Returns:
        Tuple[Optional[BytesIO], str]: (datei_buffer, fehlermeldung)
    """
    provider, identifier = parse_cloud_link(url)

    if provider is None:
        return None, "Unbekannter Cloud-Anbieter. Unterstützt: Google Drive, Dropbox, iCloud"

    if provider == "google_drive":
        buffer = download_from_google_drive(identifier, progress_callback)
        if buffer:
            return buffer, ""
        return None, "Google Drive Download fehlgeschlagen. Ist die Datei öffentlich freigegeben?"

    if provider == "dropbox":
        buffer = download_from_dropbox(identifier, progress_callback)
        if buffer:
            return buffer, ""
        return None, "Dropbox Download fehlgeschlagen. Prüfen Sie die Freigabe-Einstellungen."

    if provider == "icloud":
        return None, "iCloud-Links werden derzeit nicht direkt unterstützt. Bitte laden Sie die Datei herunter und nutzen Sie den normalen Upload."

    return None, "Download nicht möglich"


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
- sachverhalt: PFLICHTFELD! Schreibe IMMER eine inhaltliche Kurzbeschreibung des Falls in GENAU 4-6 Sätzen.
  Beschreibe konkret: 1) Ausgangssituation/Mandant, 2) Streitgegenstand/Problem, 3) Vorgehensweise/Verfahren,
  4) Ergebnis/Ausgang. KEINE Standardfloskeln wie "Beratung und Vertretung". Jeder Fall braucht eine
  individuelle, aussagekräftige Beschreibung des rechtlichen Sachverhalts!
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

WICHTIG: Das Feld "sachverhalt" ist das wichtigste Feld! Es MUSS für jeden Fall 4-6 aussagekräftige Sätze enthalten!

Antworte NUR mit einem JSON-Array von Objekten. Keine zusätzliche Erklärung."""

    user_prompt = f"""Analysiere den folgenden Text und extrahiere alle Fälle im JSON-Format:

{text[:12000]}"""  # Begrenzen auf 12000 Zeichen für stabilere Verarbeitung

    # Retry-Logik für API-Aufrufe
    max_retries = 3
    retry_delay = 2  # Sekunden

    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.1,
                max_tokens=4000,
                timeout=120  # 2 Minuten Timeout
            )

            content = response.choices[0].message.content.strip()

            # JSON aus der Antwort extrahieren
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
            error_str = str(e).lower()
            is_retryable = any(x in error_str for x in ["503", "502", "500", "timeout", "overloaded", "rate", "connection"])

            if is_retryable and attempt < max_retries - 1:
                wait_time = retry_delay * (attempt + 1)
                st.warning(f"⏳ API-Fehler, wiederhole in {wait_time}s... (Versuch {attempt + 2}/{max_retries})")
                time.sleep(wait_time)
                continue
            else:
                st.error(f"Fehler bei der GPT-Analyse: {e}")
                return []

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
- kanzlei_az: Kanzlei-Aktenzeichen (generiere eines falls nicht vorhanden)
- kurzrubrum: Anonymisiertes Rubrum (z.B. "A ./. B")
- sachverhalt: PFLICHTFELD! Schreibe GENAU 4-6 Sätze mit konkreter Sachverhaltsbeschreibung:
  1) Wer ist der Mandant und was war die Ausgangslage?
  2) Was war der Streitgegenstand/das rechtliche Problem?
  3) Wie wurde vorgegangen (außergerichtlich/gerichtlich)?
  4) Wie endete der Fall?
  KEINE Standardfloskeln! Jeder Fall braucht eine individuelle Beschreibung!
- zeitraum_von, zeitraum_bis: Format MM/YYYY
- gericht_az: Gerichtsaktenzeichen (falls vorhanden)
- verfahrenstyp: "gerichtlich", "rechtsfoermlich" oder "aussergerichtlich"
- verfahrensart: z.B. "streitig", "fG", "Mahnverfahren"
- bereich_nr (Integer), bereich_bezeichnung
- bedeutung: "gering", "mittel" oder "hoch"
- taetigkeitsbeschreibung: Konkrete anwaltliche Tätigkeiten
- stand: "anhaengig" oder "abgeschlossen"
- abschluss_art, abschluss_datum

Das Feld "sachverhalt" ist das WICHTIGSTE - es MUSS 4-6 aussagekräftige Sätze enthalten!

Antworte NUR mit JSON, keine Erklärung."""

    # Retry-Logik
    max_retries = 3
    retry_delay = 2

    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": text}
                ],
                temperature=0.1,
                max_tokens=1500,
                timeout=60
            )

            content = response.choices[0].message.content.strip()

            if "```json" in content:
                content = content.split("```json")[1].split("```")[0]
            elif "```" in content:
                content = content.split("```")[1].split("```")[0]

            return json.loads(content)

        except Exception as e:
            error_str = str(e).lower()
            is_retryable = any(x in error_str for x in ["503", "502", "500", "timeout", "overloaded", "rate", "connection"])

            if is_retryable and attempt < max_retries - 1:
                wait_time = retry_delay * (attempt + 1)
                st.warning(f"⏳ API-Fehler, wiederhole in {wait_time}s... (Versuch {attempt + 2}/{max_retries})")
                time.sleep(wait_time)
                continue
            else:
                st.error(f"Fehler bei der Einzelfall-Analyse: {e}")
                return {}

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
        "Sachverhalt": fl1["sachverhalt"],
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
        "Sachverhalt": fl2["sachverhalt"],
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
                 summary: Dict[str, Any], fachgebiet: str,
                 unprocessed_files: List[Dict] = None,
                 unrecognized_texts: List[Dict] = None) -> bytes:
    """Erstellt Excel-Arbeitsmappe mit Falllisten, Summary und Problemfällen."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    output = BytesIO()

    # Farben definieren
    header_fill_blue = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    header_fill_orange = PatternFill(start_color="E65100", end_color="E65100", fill_type="solid")
    header_fill_red = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    normal_font = Font(size=10)
    border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    def format_worksheet(ws, header_fill, column_widths=None):
        """Formatiert ein Arbeitsblatt mit Header-Farben, Filtern und Spaltenbreiten."""
        if ws.max_row < 1:
            return

        # Header formatieren (erste Zeile)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border

        # Autofilter aktivieren (Dropdown-Sortierung)
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions

        # Datenzeilen formatieren
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.font = normal_font
                cell.border = border
                cell.alignment = Alignment(vertical='top', wrap_text=False)

        # Spaltenbreiten setzen
        if column_widths:
            for col_idx, width in column_widths.items():
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = width
        else:
            # Automatische Spaltenbreiten
            for col_idx in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col_idx)
                max_length = 0
                column_header = ws.cell(row=1, column=col_idx).value

                # Sachverhalt-Spalte breiter und mit Zeilenumbruch
                if column_header and 'Sachverhalt' in str(column_header):
                    ws.column_dimensions[col_letter].width = 50
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                        for cell in row:
                            cell.alignment = Alignment(vertical='top', wrap_text=True)
                    continue

                # Andere Spalten automatisch anpassen
                for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass

                # Breite begrenzen
                adjusted_width = min(max_length + 2, 30)
                adjusted_width = max(adjusted_width, 10)
                ws.column_dimensions[col_letter].width = adjusted_width

        # Zeile fixieren (Header bleibt sichtbar beim Scrollen)
        ws.freeze_panes = 'A2'

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Fallliste 1 (gerichtlich) - Blau
        fl1_export = prepare_fl1_for_export(fl1_df) if len(fl1_df) > 0 else pd.DataFrame()
        fl1_export.to_excel(writer, sheet_name="Fallliste_1", index=False)

        # Fallliste 2 (außergerichtlich) - Grün
        fl2_export = prepare_fl2_for_export(fl2_df) if len(fl2_df) > 0 else pd.DataFrame()
        fl2_export.to_excel(writer, sheet_name="Fallliste_2", index=False)

        # Summary - Blau
        summary_df = create_summary_df(summary)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # Nicht verarbeitete Dateien - Orange
        if unprocessed_files:
            unprocessed_df = pd.DataFrame(unprocessed_files)
            unprocessed_df.to_excel(writer, sheet_name="Nicht_verarbeitet", index=False)

        # Nicht erkannte Fälle - Rot
        if unrecognized_texts:
            unrecognized_df = pd.DataFrame(unrecognized_texts)
            unrecognized_df.to_excel(writer, sheet_name="Nicht_erkannt", index=False)

        # Workbook holen und Formatierung anwenden
        workbook = writer.book

        # Fallliste 1 formatieren
        if "Fallliste_1" in workbook.sheetnames:
            format_worksheet(workbook["Fallliste_1"], header_fill_blue)

        # Fallliste 2 formatieren
        if "Fallliste_2" in workbook.sheetnames:
            format_worksheet(workbook["Fallliste_2"], header_fill_green)

        # Summary formatieren
        if "Summary" in workbook.sheetnames:
            ws_summary = workbook["Summary"]
            format_worksheet(ws_summary, header_fill_blue, {1: 35, 2: 20, 3: 40})

        # Nicht verarbeitet formatieren
        if "Nicht_verarbeitet" in workbook.sheetnames:
            format_worksheet(workbook["Nicht_verarbeitet"], header_fill_orange)

        # Nicht erkannt formatieren
        if "Nicht_erkannt" in workbook.sheetnames:
            format_worksheet(workbook["Nicht_erkannt"], header_fill_red)

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


def check_duplicates(new_cases: pd.DataFrame, existing_cases: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Prüft auf Duplikate anhand des Kanzlei-Aktenzeichens.

    Returns:
        Tuple[pd.DataFrame, pd.DataFrame]: (neue_faelle_ohne_duplikate, gefundene_duplikate)
    """
    if existing_cases.empty or new_cases.empty:
        return new_cases, pd.DataFrame()

    if "kanzlei_az" not in new_cases.columns or "kanzlei_az" not in existing_cases.columns:
        return new_cases, pd.DataFrame()

    # Bestehende Aktenzeichen sammeln
    existing_az = set(existing_cases["kanzlei_az"].dropna().astype(str).str.strip().str.lower())
    existing_az.discard("")  # Leere entfernen

    # Duplikate finden
    duplicates_mask = new_cases["kanzlei_az"].fillna("").astype(str).str.strip().str.lower().isin(existing_az)

    duplicates = new_cases[duplicates_mask].copy()
    non_duplicates = new_cases[~duplicates_mask].copy()

    return non_duplicates, duplicates


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

    # Header mit gut sichtbarer Versionsnummer
    col_title, col_version = st.columns([4, 1])
    with col_title:
        st.title("⚖️ Fachanwalt-Falllistenverwaltung")
        st.markdown("**Erstellen Sie FAO-konforme Falllisten für Ihren Fachanwaltsantrag**")
    with col_version:
        st.markdown(f"""
        <div style="background-color: #1f77b4; color: white; padding: 10px 15px;
                    border-radius: 8px; text-align: center; margin-top: 10px;">
            <strong>Version</strong><br>
            <span style="font-size: 1.2em;">{APP_VERSION}</span>
        </div>
        """, unsafe_allow_html=True)

    # Session State initialisieren
    if "cases_df" not in st.session_state:
        st.session_state.cases_df = pd.DataFrame()
    if "openai_api_key" not in st.session_state:
        st.session_state.openai_api_key = ""
    if "upload_key" not in st.session_state:
        st.session_state.upload_key = 0
    if "pending_duplicates" not in st.session_state:
        st.session_state.pending_duplicates = []
    if "unprocessed_files" not in st.session_state:
        st.session_state.unprocessed_files = []  # Dateien die nicht verarbeitet werden konnten
    if "unrecognized_texts" not in st.session_state:
        st.session_state.unrecognized_texts = []  # Texte ohne erkannte Fälle
    if "last_processed_file" not in st.session_state:
        st.session_state.last_processed_file = None  # Zuletzt verarbeitete Datei

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

        # Aktuelle Fallanzahl anzeigen
        st.markdown("---")
        st.subheader("📊 Aktuelle Fallliste")
        current_count = len(st.session_state.cases_df) if not st.session_state.cases_df.empty else 0
        st.metric("Erfasste Fälle", current_count)
        if current_count > 0:
            st.caption("Die Fallliste wird bei jedem Upload erweitert.")

        # Zuletzt verarbeitete Datei anzeigen
        if st.session_state.last_processed_file:
            st.markdown("---")
            st.subheader("📄 Zuletzt verarbeitet")
            st.info(f"**{st.session_state.last_processed_file}**")
            st.caption("Diese Akte wurde zuletzt analysiert.")

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
    tab_pdf, tab_cloud, tab_manual = st.tabs([
        "📕 PDF hochladen (empfohlen)",
        "☁️ Cloud-Link (Google Drive/Dropbox)",
        "✏️ Fall manuell eingeben"
    ])

    all_cases = []

    # Tab 1: PDF Upload mit KI-Analyse (Hauptmethode)
    with tab_pdf:
        st.subheader("PDF-Dokumente mit KI analysieren")

        # Upload-Limit-Hinweis
        st.info(f"📦 **Upload-Limit:** Max. {MAX_FILE_SIZE_MB} MB pro Datei, {MAX_TOTAL_UPLOAD_MB} MB gesamt")

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

            **Bei größeren Datenmengen:** Nutzen Sie den Cloud-Link Tab für Google Drive oder Dropbox.
            """)

            pdf_files = st.file_uploader(
                "PDF-Dateien auswählen",
                type=["pdf"],
                accept_multiple_files=True,
                key=f"pdf_upload_{st.session_state.upload_key}",
                help=f"Sie können mehrere PDFs hochladen (max. {MAX_FILE_SIZE_MB} MB pro Datei)"
            )

            if pdf_files:
                # Dateigrößen prüfen
                files_to_process = []
                rejected_files = []

                for pdf_file in pdf_files:
                    is_ok, error_msg = check_file_size(pdf_file)
                    if is_ok:
                        files_to_process.append(pdf_file)
                    else:
                        rejected_files.append((pdf_file.name, error_msg))
                        # Zu unprocessed_files hinzufügen
                        st.session_state.unprocessed_files.append({
                            "Dateiname": pdf_file.name,
                            "Grund": error_msg,
                            "Typ": "Größenlimit überschritten",
                            "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                        })

                # Gesamtgröße prüfen
                if files_to_process:
                    total_ok, total_error, total_mb = check_total_upload_size(files_to_process)
                    st.caption(f"Gesamtgröße: {total_mb:.1f} MB von {MAX_TOTAL_UPLOAD_MB} MB")

                    if not total_ok:
                        st.error(f"⚠️ {total_error}")
                        st.warning("Tipp: Nutzen Sie den **Cloud-Link Tab** für große Datenmengen!")
                        # Alle Dateien als nicht verarbeitet markieren
                        for f in files_to_process:
                            st.session_state.unprocessed_files.append({
                                "Dateiname": f.name,
                                "Grund": total_error,
                                "Typ": "Gesamtlimit überschritten",
                                "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                            })
                        files_to_process = []

                # Abgelehnte Dateien anzeigen
                if rejected_files:
                    st.error("❌ Folgende Dateien wurden abgelehnt (zu groß):")
                    for filename, error in rejected_files:
                        st.markdown(f"- **{filename}**: {error}")
                    st.warning("Tipp: Nutzen Sie den **Cloud-Link Tab** für Google Drive oder Dropbox!")

                if files_to_process:
                    if st.button("🤖 PDFs mit ChatGPT analysieren", type="primary", use_container_width=True):
                        client = get_openai_client(st.session_state.openai_api_key)

                        if client:
                            progress_bar = st.progress(0)
                            for i, pdf_file in enumerate(files_to_process):
                                with st.spinner(f"Analysiere {pdf_file.name}..."):
                                    try:
                                        pdf_text = extract_text_from_pdf(pdf_file)

                                        if len(pdf_text.strip()) < 50:
                                            st.warning(f"⚠️ {pdf_file.name}: Wenig Text gefunden. Gescanntes PDF?")
                                            # Als nicht erkannt dokumentieren
                                            st.session_state.unrecognized_texts.append({
                                                "Dateiname": pdf_file.name,
                                                "Grund": "Wenig oder kein Text extrahiert (möglicherweise gescanntes PDF)",
                                                "Textlänge": len(pdf_text.strip()),
                                                "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                            })
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
                                            st.session_state.last_processed_file = pdf_file.name
                                            st.success(f"✓ {pdf_file.name}: **{len(cases)} Fälle** erkannt")

                                            with st.expander(f"Details: {pdf_file.name}"):
                                                st.dataframe(
                                                    df[["kurzrubrum", "sachverhalt", "verfahrenstyp", "bereich_nr"]],
                                                    use_container_width=True,
                                                    hide_index=True
                                                )
                                        else:
                                            st.warning(f"⚠️ {pdf_file.name}: Keine Fälle erkannt")
                                            # Als nicht erkannt dokumentieren mit Textauszug
                                            st.session_state.unrecognized_texts.append({
                                                "Dateiname": pdf_file.name,
                                                "Grund": "Keine Fälle von GPT erkannt",
                                                "Textauszug": pdf_text[:500] + "..." if len(pdf_text) > 500 else pdf_text,
                                                "Textlänge": len(pdf_text),
                                                "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                            })

                                    except Exception as e:
                                        st.error(f"✗ {pdf_file.name}: {str(e)}")
                                        # Fehler dokumentieren
                                        st.session_state.unprocessed_files.append({
                                            "Dateiname": pdf_file.name,
                                            "Grund": str(e),
                                            "Typ": "Verarbeitungsfehler",
                                            "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                        })

                                progress_bar.progress((i + 1) / len(files_to_process))

                            # Nach erfolgreicher Analyse: Upload-Bereich zurücksetzen
                            if all_cases:
                                st.session_state.upload_key += 1
                                st.success("✅ Analyse abgeschlossen! Upload-Bereich wurde geleert.")
                                st.rerun()

    # Tab 2: Cloud-Link Upload (Google Drive, Dropbox, iCloud)
    with tab_cloud:
        st.subheader("PDF von Cloud-Speicher laden")

        st.markdown("""
        **Ideal für große Datenmengen!** Die Dateien werden in Paketen heruntergeladen und verarbeitet.

        **Unterstützte Anbieter:**
        - **Google Drive**: Datei **oder Ordner** freigeben → "Jeder mit dem Link" → Link kopieren
        - **Dropbox**: Datei freigeben → Link kopieren
        - **iCloud**: *Derzeit nicht direkt unterstützt* (bitte Datei herunterladen und normal hochladen)

        **Tipp:** Bei Google Drive können Sie auch einen **ganzen Ordner** mit mehreren PDF-Dateien freigeben.
        Die App verarbeitet dann alle PDF-Dateien nacheinander.
        """)

        if not st.session_state.openai_api_key:
            st.warning("⚠️ Bitte geben Sie Ihren OpenAI API-Key in der Sidebar ein.")
        else:
            cloud_url = st.text_input(
                "Freigabe-Link eingeben",
                placeholder="https://drive.google.com/file/d/... oder .../folders/... oder https://www.dropbox.com/...",
                help="Der Link muss öffentlich zugänglich sein (Freigabe für 'Jeder mit dem Link')"
            )

            if cloud_url:
                # Link prüfen
                provider, identifier = parse_cloud_link(cloud_url)

                if provider:
                    # Anzeige anpassen für Ordner
                    if provider == "google_drive_folder":
                        st.success(f"✓ Erkannt: **Google Drive Ordner**")
                        st.info("📁 Ordner-Link erkannt. Alle PDF-Dateien werden nacheinander verarbeitet.")
                    else:
                        st.success(f"✓ Erkannt: **{provider.replace('_', ' ').title()}**")

                    # Google Drive Ordner: Liste Dateien auf
                    if provider == "google_drive_folder":
                        if st.button("📁 Ordner-Inhalt laden", type="secondary"):
                            with st.spinner("Lade Ordner-Inhalt..."):
                                # Versuche beide Methoden
                                folder_files = list_google_drive_folder(identifier)
                                if not folder_files:
                                    folder_files = get_google_drive_folder_files_via_webpage(identifier)

                                if folder_files:
                                    st.session_state.folder_files = folder_files
                                    st.success(f"✓ {len(folder_files)} PDF-Datei(en) gefunden!")
                                else:
                                    st.warning("⚠️ Keine PDF-Dateien gefunden oder Ordner nicht zugänglich. "
                                             "Stellen Sie sicher, dass der Ordner öffentlich freigegeben ist.")

                        # Wenn Dateien gefunden, anzeigen und verarbeiten
                        if "folder_files" in st.session_state and st.session_state.folder_files:
                            folder_files = st.session_state.folder_files

                            st.markdown(f"**Gefundene Dateien ({len(folder_files)}):**")
                            for idx, f in enumerate(folder_files[:20]):  # Max 20 anzeigen
                                st.text(f"  {idx+1}. {f['name']}")
                            if len(folder_files) > 20:
                                st.text(f"  ... und {len(folder_files) - 20} weitere")

                            if st.button("☁️ Alle Dateien verarbeiten", type="primary", use_container_width=True):
                                client = get_openai_client(st.session_state.openai_api_key)
                                if not client:
                                    st.error("OpenAI-Client konnte nicht initialisiert werden.")
                                else:
                                    total_progress = st.progress(0, text="Verarbeite Dateien...")
                                    all_folder_cases = []

                                    for file_idx, file_info in enumerate(folder_files):
                                        file_id = file_info['id']
                                        file_name = file_info['name']

                                        st.write(f"📄 Verarbeite: **{file_name}** ({file_idx+1}/{len(folder_files)})")

                                        try:
                                            # Download
                                            buffer = download_from_google_drive(file_id)

                                            if buffer:
                                                # Text extrahieren
                                                try:
                                                    pdf_text = extract_text_from_pdf(buffer)

                                                    if len(pdf_text.strip()) < 50:
                                                        st.warning(f"⚠️ {file_name}: Wenig Text gefunden")
                                                        st.session_state.unrecognized_texts.append({
                                                            "Dateiname": file_name,
                                                            "Grund": "Wenig Text extrahiert (< 50 Zeichen)",
                                                            "Textauszug": pdf_text[:200] if pdf_text else "Kein Text",
                                                            "Textlänge": len(pdf_text),
                                                            "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                                        })
                                                    else:
                                                        # GPT-Analyse
                                                        cases = analyze_cases_with_gpt(client, pdf_text, fachgebiet, model=gpt_model)

                                                        if cases:
                                                            df = cases_list_to_dataframe(cases)
                                                            df = normalize_case_df(df, fachgebiet)
                                                            all_folder_cases.append(df)
                                                            st.session_state.last_processed_file = file_name
                                                            st.success(f"✓ {len(cases)} Fälle aus {file_name}")
                                                        else:
                                                            st.warning(f"⚠️ {file_name}: Keine Fälle erkannt")
                                                            st.session_state.unrecognized_texts.append({
                                                                "Dateiname": file_name,
                                                                "Grund": "Keine Fälle von GPT erkannt",
                                                                "Textauszug": pdf_text[:500] + "..." if len(pdf_text) > 500 else pdf_text,
                                                                "Textlänge": len(pdf_text),
                                                                "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                                            })

                                                except Exception as e:
                                                    st.error(f"❌ {file_name}: Fehler bei Verarbeitung - {str(e)}")
                                                    st.session_state.unprocessed_files.append({
                                                        "Dateiname": file_name,
                                                        "Grund": f"Verarbeitungsfehler: {str(e)}",
                                                        "Typ": "PDF-Fehler",
                                                        "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                                    })
                                            else:
                                                st.warning(f"⚠️ {file_name}: Download fehlgeschlagen")
                                                st.session_state.unprocessed_files.append({
                                                    "Dateiname": file_name,
                                                    "Grund": "Download fehlgeschlagen",
                                                    "Typ": "Download-Fehler",
                                                    "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                                })

                                        except Exception as e:
                                            st.error(f"❌ {file_name}: {str(e)}")
                                            st.session_state.unprocessed_files.append({
                                                "Dateiname": file_name,
                                                "Grund": str(e),
                                                "Typ": "Allgemeiner Fehler",
                                                "Zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                            })

                                        total_progress.progress((file_idx + 1) / len(folder_files))

                                    # Alle Fälle zu all_cases hinzufügen (normaler Ablauf)
                                    if all_folder_cases:
                                        total_count = sum(len(df) for df in all_folder_cases)
                                        all_cases.extend(all_folder_cases)

                                        # Ordner-Cache leeren
                                        st.session_state.folder_files = []

                                        st.success(f"✅ **Ordner-Verarbeitung abgeschlossen!** "
                                                 f"Insgesamt {total_count} Fälle aus {len(all_folder_cases)} Dateien extrahiert.")
                                    else:
                                        st.warning("⚠️ Keine Fälle in den Dateien gefunden.")

                    else:
                        # Einzelne Datei (wie bisher)
                        if st.button("☁️ Von Cloud laden und analysieren", type="primary", use_container_width=True):
                            with st.spinner("Lade Datei von Cloud..."):
                                progress_placeholder = st.empty()

                                def update_progress(progress):
                                    progress_placeholder.progress(progress, text=f"Download: {progress*100:.0f}%")

                                buffer, error = download_from_cloud(cloud_url, update_progress)

                                if buffer:
                                    progress_placeholder.empty()
                                    st.success("✓ Download abgeschlossen!")

                                    # Prüfen ob PDF
                                    try:
                                        with st.spinner("Extrahiere Text aus PDF..."):
                                            pdf_text = extract_text_from_pdf(buffer)

                                        if len(pdf_text.strip()) < 50:
                                            st.warning("⚠️ Wenig Text gefunden. Gescanntes PDF?")
                                        else:
                                            client = get_openai_client(st.session_state.openai_api_key)

                                            if client:
                                                # Text in Chunks aufteilen für große Dokumente
                                                text_chunks = []
                                                chunk_size = 10000  # Zeichen pro Chunk

                                                if len(pdf_text) > chunk_size:
                                                    st.info(f"📄 Großes Dokument erkannt ({len(pdf_text)} Zeichen). Verarbeite in Paketen...")

                                                    for i in range(0, len(pdf_text), chunk_size):
                                                        text_chunks.append(pdf_text[i:i+chunk_size])
                                                else:
                                                    text_chunks = [pdf_text]

                                                total_cases = []
                                                chunk_progress = st.progress(0)

                                                for idx, chunk in enumerate(text_chunks):
                                                    with st.spinner(f"Analysiere Paket {idx+1}/{len(text_chunks)}..."):
                                                        cases = analyze_cases_with_gpt(
                                                            client,
                                                            chunk,
                                                            fachgebiet,
                                                            model=gpt_model
                                                        )
                                                        if cases:
                                                            total_cases.extend(cases)

                                                    chunk_progress.progress((idx + 1) / len(text_chunks))

                                                if total_cases:
                                                    df = cases_list_to_dataframe(total_cases)
                                                    df = normalize_case_df(df, fachgebiet)
                                                    all_cases.append(df)
                                                    st.session_state.last_processed_file = f"Cloud-Dokument ({provider})"
                                                    st.success(f"✓ **{len(total_cases)} Fälle** aus Cloud-Dokument extrahiert!")

                                                    with st.expander("Erkannte Fälle anzeigen"):
                                                        st.dataframe(
                                                            df[["kurzrubrum", "sachverhalt", "verfahrenstyp", "bereich_nr"]],
                                                            use_container_width=True,
                                                            hide_index=True
                                                        )
                                                else:
                                                    st.warning("⚠️ Keine Fälle im Dokument erkannt")

                                    except Exception as e:
                                        st.error(f"Fehler bei der Verarbeitung: {str(e)}")
                                else:
                                    st.error(f"❌ {error}")
                else:
                    st.warning("⚠️ Link nicht erkannt. Bitte einen gültigen Google Drive (Datei oder Ordner) oder Dropbox Link eingeben.")

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

    # Alle neuen Fälle mit bestehenden kombinieren (mit Duplikat-Prüfung)
    if all_cases:
        new_combined = pd.concat(all_cases, ignore_index=True)

        # Duplikat-Prüfung durchführen
        non_duplicates, duplicates = check_duplicates(new_combined, st.session_state.cases_df)

        # Duplikate zur manuellen Freigabe speichern
        if not duplicates.empty:
            st.session_state.pending_duplicates = duplicates.to_dict('records')

        # Nur nicht-doppelte Fälle hinzufügen
        if not non_duplicates.empty:
            if not st.session_state.cases_df.empty:
                st.session_state.cases_df = pd.concat(
                    [st.session_state.cases_df, non_duplicates],
                    ignore_index=True
                )
            else:
                st.session_state.cases_df = non_duplicates

    # Duplikat-Warnung anzeigen und manuelle Freigabe ermöglichen
    if st.session_state.pending_duplicates:
        st.markdown("---")
        st.warning(f"⚠️ **{len(st.session_state.pending_duplicates)} Duplikat(e) gefunden!** Diese Aktenzeichen existieren bereits:")

        for idx, dup in enumerate(st.session_state.pending_duplicates):
            with st.expander(f"Duplikat: {dup.get('kanzlei_az', 'Unbekannt')} - {dup.get('kurzrubrum', '')}"):
                st.json(dup)
                col1, col2 = st.columns(2)
                with col1:
                    if st.button(f"✅ Trotzdem hinzufügen", key=f"add_dup_{idx}"):
                        new_df = pd.DataFrame([dup])
                        if not st.session_state.cases_df.empty:
                            st.session_state.cases_df = pd.concat(
                                [st.session_state.cases_df, new_df],
                                ignore_index=True
                            )
                        else:
                            st.session_state.cases_df = new_df
                        st.session_state.pending_duplicates.pop(idx)
                        st.rerun()
                with col2:
                    if st.button(f"❌ Verwerfen", key=f"skip_dup_{idx}"):
                        st.session_state.pending_duplicates.pop(idx)
                        st.rerun()

        if st.button("🗑️ Alle Duplikate verwerfen"):
            st.session_state.pending_duplicates = []
            st.rerun()

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
            display_cols = ["FL1_Nr", "kurzrubrum", "sachverhalt", "kanzlei_az", "gericht_az",
                          "bereich_nr", "verfahrenstyp", "stand"]
            display_cols = [c for c in display_cols if c in fl1.columns]
            st.dataframe(fl1[display_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Keine gerichtlichen/rechtsförmlichen Verfahren vorhanden.")

    with preview_tab2:
        if len(fl2) > 0:
            display_cols = ["FL2_Nr", "kurzrubrum", "sachverhalt", "kanzlei_az", "bereich_nr"]
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
    - **Nicht_verarbeitet**: Dateien die nicht verarbeitet werden konnten (z.B. zu groß)
    - **Nicht_erkannt**: Texte aus denen keine Fälle extrahiert werden konnten
    """)

    excel_bytes = create_excel(
        fl1, fl2, summary, fachgebiet,
        unprocessed_files=st.session_state.unprocessed_files,
        unrecognized_texts=st.session_state.unrecognized_texts
    )
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
