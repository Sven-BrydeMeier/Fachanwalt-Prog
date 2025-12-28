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
APP_VERSION = "25.12.13-23:33"

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

# OCR-Integration für gescannte PDFs
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# PDF-Report Generation
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Word-Dokument Generation
try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

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

def extract_text_from_pdf(file, use_ocr: bool = True) -> Tuple[str, str]:
    """
    Extrahiert Text aus einer PDF-Datei.
    Versucht zuerst normale Textextraktion, dann OCR falls wenig Text gefunden.

    Returns:
        Tuple[str, str]: (extrahierter_text, methode: 'text' oder 'ocr')
    """
    if not PDF_AVAILABLE:
        raise ImportError("PyPDF2 ist nicht installiert. Bitte führen Sie 'pip install pypdf2' aus.")

    # Zuerst normale Textextraktion versuchen
    file.seek(0)
    reader = PdfReader(file)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"

    # Wenn genug Text gefunden wurde, zurückgeben
    if len(text.strip()) >= 100:
        return text, "text"

    # OCR versuchen, wenn wenig Text und OCR verfügbar
    if use_ocr and OCR_AVAILABLE and len(text.strip()) < 100:
        try:
            file.seek(0)
            pdf_bytes = file.read()
            images = convert_from_bytes(pdf_bytes, dpi=200)

            ocr_text = ""
            for i, image in enumerate(images):
                # Tesseract mit deutscher Sprache
                page_text = pytesseract.image_to_string(image, lang='deu')
                ocr_text += page_text + "\n"

            if len(ocr_text.strip()) > len(text.strip()):
                return ocr_text, "ocr"
        except Exception as e:
            # OCR fehlgeschlagen, ursprünglichen Text zurückgeben
            pass

    return text, "text"


def extract_text_from_pdf_simple(file) -> str:
    """Legacy-Funktion für Rückwärtskompatibilität."""
    text, _ = extract_text_from_pdf(file, use_ocr=True)
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
# PLAUSIBILITÄTSPRÜFUNG
# =============================================================================

def check_case_plausibility(case: Dict) -> List[Dict]:
    """
    Prüft einen Fall auf Plausibilität und gibt Warnungen zurück.

    Returns:
        List[Dict]: Liste von Warnungen mit 'typ', 'feld', 'meldung'
    """
    warnings = []

    # Sachverhalt-Prüfung (4-6 Sätze gefordert)
    sachverhalt = case.get("sachverhalt", "")
    if sachverhalt:
        # Sätze zählen (einfache Heuristik)
        sentences = len([s for s in sachverhalt.replace("!", ".").replace("?", ".").split(".") if s.strip()])
        if sentences < 4:
            warnings.append({
                "typ": "warnung",
                "feld": "sachverhalt",
                "meldung": f"Nur {sentences} Sätze - FAO verlangt 4-6 Sätze"
            })
        elif sentences > 8:
            warnings.append({
                "typ": "hinweis",
                "feld": "sachverhalt",
                "meldung": f"{sentences} Sätze - evtl. kürzen auf 4-6 Sätze"
            })
    else:
        warnings.append({
            "typ": "fehler",
            "feld": "sachverhalt",
            "meldung": "Sachverhalt fehlt - Pflichtfeld!"
        })

    # Pflichtfelder prüfen
    pflichtfelder = {
        "kurzrubrum": "Kurzrubrum",
        "verfahrenstyp": "Verfahrenstyp",
        "bereich_nr": "Bereich-Nr"
    }

    for feld, bezeichnung in pflichtfelder.items():
        wert = case.get(feld, "")
        if not wert or str(wert).strip() == "":
            warnings.append({
                "typ": "fehler",
                "feld": feld,
                "meldung": f"{bezeichnung} fehlt - Pflichtfeld!"
            })

    # Aktenzeichen-Prüfung
    kanzlei_az = case.get("kanzlei_az", "")
    if not kanzlei_az or str(kanzlei_az).strip() == "":
        warnings.append({
            "typ": "warnung",
            "feld": "kanzlei_az",
            "meldung": "Kanzlei-AZ fehlt - empfohlen für Zuordnung"
        })

    # Zeitraum-Prüfung
    zeitraum_von = case.get("zeitraum_von", "")
    zeitraum_bis = case.get("zeitraum_bis", "")
    if not zeitraum_von and not zeitraum_bis:
        warnings.append({
            "typ": "hinweis",
            "feld": "zeitraum",
            "meldung": "Kein Zeitraum angegeben"
        })

    # Gerichtsaktenzeichen bei gerichtlichen Verfahren
    verfahrenstyp = case.get("verfahrenstyp", "").lower()
    if verfahrenstyp in ["gerichtlich", "rechtsfoermlich"]:
        gericht_az = case.get("gericht_az", "")
        if not gericht_az or str(gericht_az).strip() == "":
            warnings.append({
                "typ": "warnung",
                "feld": "gericht_az",
                "meldung": "Gerichts-AZ fehlt bei gerichtlichem Verfahren"
            })

    return warnings


def check_all_cases_plausibility(df: pd.DataFrame) -> Dict:
    """
    Prüft alle Fälle auf Plausibilität.

    Returns:
        Dict mit 'fehler', 'warnungen', 'hinweise' und 'details'
    """
    result = {
        "fehler": 0,
        "warnungen": 0,
        "hinweise": 0,
        "details": []  # Liste von (index, fall_info, warnings)
    }

    if df.empty:
        return result

    for idx, row in df.iterrows():
        case = row.to_dict()
        warnings = check_case_plausibility(case)

        if warnings:
            fall_info = f"{case.get('kanzlei_az', 'Unbekannt')} - {case.get('kurzrubrum', '')}"
            result["details"].append((idx, fall_info, warnings))

            for w in warnings:
                if w["typ"] == "fehler":
                    result["fehler"] += 1
                elif w["typ"] == "warnung":
                    result["warnungen"] += 1
                else:
                    result["hinweise"] += 1

    return result


# =============================================================================
# SESSION SPEICHERN/LADEN
# =============================================================================

def save_session_to_json(cases_df: pd.DataFrame, fachgebiet: str,
                         unprocessed_files: List, unrecognized_texts: List,
                         upload_history: List) -> str:
    """
    Speichert die aktuelle Session als JSON-String.
    """
    session_data = {
        "version": APP_VERSION,
        "timestamp": datetime.now().isoformat(),
        "fachgebiet": fachgebiet,
        "cases": cases_df.to_dict(orient="records") if not cases_df.empty else [],
        "unprocessed_files": unprocessed_files,
        "unrecognized_texts": unrecognized_texts,
        "upload_history": upload_history
    }
    return json.dumps(session_data, ensure_ascii=False, indent=2, default=str)


def load_session_from_json(json_str: str) -> Dict:
    """
    Lädt eine Session aus einem JSON-String.

    Returns:
        Dict mit 'cases_df', 'fachgebiet', 'unprocessed_files', etc.
    """
    try:
        data = json.loads(json_str)

        cases_df = pd.DataFrame(data.get("cases", []))

        return {
            "success": True,
            "cases_df": cases_df,
            "fachgebiet": data.get("fachgebiet", "Erbrecht"),
            "unprocessed_files": data.get("unprocessed_files", []),
            "unrecognized_texts": data.get("unrecognized_texts", []),
            "upload_history": data.get("upload_history", []),
            "timestamp": data.get("timestamp", ""),
            "version": data.get("version", "")
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }


# =============================================================================
# PDF-REPORT GENERIERUNG
# =============================================================================

def create_pdf_report(fl1_df: pd.DataFrame, fl2_df: pd.DataFrame,
                      summary: Dict, fachgebiet: str) -> bytes:
    """
    Erstellt einen PDF-Report der Fallliste.
    """
    if not REPORTLAB_AVAILABLE:
        raise ImportError("ReportLab ist nicht installiert. Bitte 'pip install reportlab' ausführen.")

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                           leftMargin=1.5*cm, rightMargin=1.5*cm,
                           topMargin=2*cm, bottomMargin=2*cm)

    styles = getSampleStyleSheet()
    story = []

    # Titel-Styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=30
    )

    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceBefore=20,
        spaceAfter=10
    )

    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=10,
        spaceAfter=6
    )

    # Titelseite
    story.append(Paragraph(f"Fallliste {fachgebiet}", title_style))
    story.append(Paragraph(f"Erstellt am: {datetime.now().strftime('%d.%m.%Y %H:%M')}", normal_style))
    story.append(Spacer(1, 20))

    # Zusammenfassung
    story.append(Paragraph("Übersicht", heading_style))
    summary_data = [
        ["Kriterium", "Wert"],
        ["Fälle gesamt", str(summary.get("gesamt", 0))],
        ["Gerichtliche Verfahren", str(summary.get("gerichtlich", 0))],
        ["Außergerichtliche Verfahren", str(summary.get("aussergerichtlich", 0))],
        ["Rechtsförmliche Verfahren", str(summary.get("rechtsfoermlich", 0))],
    ]

    summary_table = Table(summary_data, colWidths=[10*cm, 5*cm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a5f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f0f0')]),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 20))

    # FAO-Check
    story.append(Paragraph("FAO-Konformitätscheck", heading_style))
    check_data = [["Kriterium", "Ist", "Soll", "Status"]]
    for check in summary.get("checks", []):
        kriterium, ist, soll, status = check
        status_text = {"erfuellt": "✓ Erfüllt", "knapp": "⚠ Knapp", "nicht_erfuellt": "✗ Nicht erfüllt"}.get(status, status)
        check_data.append([kriterium, str(ist), str(soll), status_text])

    check_table = Table(check_data, colWidths=[7*cm, 3*cm, 3*cm, 4*cm])
    check_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a5f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (2, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    story.append(check_table)

    # Fallliste 1 (gekürzt für Übersicht)
    if not fl1_df.empty:
        story.append(PageBreak())
        story.append(Paragraph(f"Fallliste 1 - Gerichtliche/Rechtsförmliche Verfahren ({len(fl1_df)} Fälle)", heading_style))

        fl1_data = [["Nr.", "Kurzrubrum", "Bereich", "Typ"]]
        for _, row in fl1_df.head(50).iterrows():  # Max 50 für Übersicht
            fl1_data.append([
                str(row.get("FL1_Nr", "")),
                str(row.get("kurzrubrum", ""))[:40],
                str(row.get("bereich_nr", "")),
                str(row.get("verfahrenstyp", ""))[:15]
            ])

        if len(fl1_df) > 50:
            fl1_data.append(["...", f"(weitere {len(fl1_df) - 50} Fälle)", "", ""])

        fl1_table = Table(fl1_data, colWidths=[1.5*cm, 10*cm, 2*cm, 3*cm])
        fl1_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2e7d32')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(fl1_table)

    # Fallliste 2 (gekürzt)
    if not fl2_df.empty:
        story.append(Spacer(1, 20))
        story.append(Paragraph(f"Fallliste 2 - Außergerichtliche Verfahren ({len(fl2_df)} Fälle)", heading_style))

        fl2_data = [["Nr.", "Kurzrubrum", "Bereich"]]
        for _, row in fl2_df.head(50).iterrows():
            fl2_data.append([
                str(row.get("FL2_Nr", "")),
                str(row.get("kurzrubrum", ""))[:50],
                str(row.get("bereich_nr", ""))
            ])

        if len(fl2_df) > 50:
            fl2_data.append(["...", f"(weitere {len(fl2_df) - 50} Fälle)", ""])

        fl2_table = Table(fl2_data, colWidths=[1.5*cm, 12*cm, 2*cm])
        fl2_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565c0')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(fl2_table)

    # Footer-Info
    story.append(Spacer(1, 30))
    story.append(Paragraph(
        f"Generiert mit FAO-Falllisten-Generator v{APP_VERSION}",
        ParagraphStyle('Footer', parent=styles['Normal'], fontSize=8, textColor=colors.grey)
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# =============================================================================
# ANTRAGSSCHREIBEN GENERIERUNG
# =============================================================================

def create_antragsschreiben(
    antragsteller: Dict,
    fachgebiet: str,
    summary: Dict,
    klausuren: List[Dict],
    bereich_verteilung: Dict,
    fl1_count: int,
    fl2_count: int
) -> bytes:
    """
    Erstellt ein formelles Antragsschreiben als Word-Dokument.

    Args:
        antragsteller: Dict mit Name, Anschrift, Zulassungsdatum, etc.
        fachgebiet: Das beantragte Fachgebiet
        summary: Zusammenfassung der Fallliste
        klausuren: Liste der bestandenen Klausuren
        bereich_verteilung: Verteilung der Fälle auf Bereiche
        fl1_count: Anzahl Fälle in Fallliste 1
        fl2_count: Anzahl Fälle in Fallliste 2

    Returns:
        bytes: Word-Dokument als Bytes
    """
    if not DOCX_AVAILABLE:
        raise ImportError("python-docx ist nicht installiert. Bitte 'pip install python-docx' ausführen.")

    doc = Document()

    # Seitenränder
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2)

    # Absender (rechtsbündig)
    absender = doc.add_paragraph()
    absender.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    absender_run = absender.add_run(
        f"{antragsteller.get('name', '[Name]')}\n"
        f"{antragsteller.get('kanzlei', '[Kanzlei]')}\n"
        f"{antragsteller.get('strasse', '[Straße]')}\n"
        f"{antragsteller.get('plz_ort', '[PLZ Ort]')}\n"
        f"Tel.: {antragsteller.get('telefon', '[Telefon]')}\n"
        f"E-Mail: {antragsteller.get('email', '[E-Mail]')}"
    )
    absender_run.font.size = Pt(10)

    doc.add_paragraph()  # Leerzeile

    # Empfänger
    empfaenger = doc.add_paragraph()
    empfaenger_run = empfaenger.add_run(
        f"{antragsteller.get('kammer_name', 'Rechtsanwaltskammer [Ort]')}\n"
        f"{antragsteller.get('kammer_strasse', '[Straße]')}\n"
        f"{antragsteller.get('kammer_plz_ort', '[PLZ Ort]')}"
    )
    empfaenger_run.font.size = Pt(11)
    empfaenger_run.bold = True

    doc.add_paragraph()  # Leerzeile

    # Datum
    datum = doc.add_paragraph()
    datum.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    datum_run = datum.add_run(f"{antragsteller.get('ort', '[Ort]')}, den {datetime.now().strftime('%d.%m.%Y')}")
    datum_run.font.size = Pt(11)

    doc.add_paragraph()  # Leerzeile

    # Betreff
    betreff = doc.add_paragraph()
    betreff_run = betreff.add_run(f'Antrag auf Verleihung der Bezeichnung "Fachanwalt für {fachgebiet}"')
    betreff_run.bold = True
    betreff_run.font.size = Pt(12)

    doc.add_paragraph()  # Leerzeile

    # Anrede
    anrede = doc.add_paragraph()
    anrede.add_run("Sehr geehrte Damen und Herren,").font.size = Pt(11)

    doc.add_paragraph()  # Leerzeile

    # Einleitung
    config = FAO_CONFIG.get(fachgebiet, {})
    paragraph_ref = config.get("paragraph", "§ 5 FAO")

    einleitung = doc.add_paragraph()
    einleitung_text = (
        f'hiermit beantrage ich die Verleihung der Bezeichnung "Fachanwalt für {fachgebiet}" '
        f"gemäß {paragraph_ref}.\n\n"
        f"Ich bin seit dem {antragsteller.get('zulassung_datum', '[Datum]')} zur Rechtsanwaltschaft zugelassen "
        f"und bei der {antragsteller.get('kammer_name', 'Rechtsanwaltskammer [Ort]')} als Rechtsanwalt/Rechtsanwältin "
        f"eingetragen."
    )
    einleitung.add_run(einleitung_text).font.size = Pt(11)

    # Theoretische Kenntnisse
    doc.add_paragraph()
    theo_header = doc.add_paragraph()
    theo_header.add_run("I. Nachweis der besonderen theoretischen Kenntnisse").bold = True

    theo_text = doc.add_paragraph()
    lehrgang_info = antragsteller.get('lehrgang_info', '[Lehrgangsbezeichnung]')
    lehrgang_datum = antragsteller.get('lehrgang_datum', '[Datum]')

    theo_content = (
        f'Ich habe den Fachanwaltslehrgang "{lehrgang_info}" erfolgreich absolviert. '
        f"Der Lehrgang wurde am {lehrgang_datum} abgeschlossen.\n\n"
        f"Folgende Klausuren wurden bestanden:"
    )
    theo_text.add_run(theo_content).font.size = Pt(11)

    # Klausurtabelle
    if klausuren:
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Header
        header_cells = table.rows[0].cells
        header_cells[0].text = "Klausur"
        header_cells[1].text = "Datum"
        header_cells[2].text = "Ergebnis"

        for cell in header_cells:
            cell.paragraphs[0].runs[0].bold = True

        # Klausuren eintragen
        for klausur in klausuren:
            row = table.add_row()
            row.cells[0].text = klausur.get('bezeichnung', '')
            row.cells[1].text = klausur.get('datum', '')
            row.cells[2].text = klausur.get('ergebnis', 'bestanden')

    doc.add_paragraph()

    # Praktische Erfahrungen
    prak_header = doc.add_paragraph()
    prak_header.add_run("II. Nachweis der besonderen praktischen Erfahrungen").bold = True

    gesamt = summary.get("gesamt", 0)
    gerichtlich = summary.get("gerichtlich", 0)
    aussergerichtlich = summary.get("aussergerichtlich", 0)
    config_gesamt_min = config.get("gesamt_min", 0)
    config_gerichtlich_min = config.get("gerichtlich_min", 0)

    prak_text = doc.add_paragraph()
    prak_content = (
        f"In den letzten drei Jahren vor Antragstellung habe ich insgesamt {gesamt} Fälle "
        f"aus dem Bereich {fachgebiet} persönlich und weisungsfrei bearbeitet.\n\n"
        f"Die beigefügten Falllisten weisen nach:\n"
        f"• Fallliste 1 (gerichtliche/rechtsförmliche Verfahren): {fl1_count} Fälle\n"
        f"• Fallliste 2 (außergerichtliche Verfahren): {fl2_count} Fälle\n\n"
        f"Die Anforderungen gemäß {paragraph_ref} (mindestens {config_gesamt_min} Fälle, "
        f"davon mindestens {config_gerichtlich_min} gerichtliche/rechtsförmliche Verfahren) "
        f"sind damit erfüllt."
    )
    prak_text.add_run(prak_content).font.size = Pt(11)

    # Bereichsverteilung
    if bereich_verteilung:
        doc.add_paragraph()
        bereich_header = doc.add_paragraph()
        bereich_header.add_run("Verteilung auf die Bereiche:").italic = True

        bereiche_config = config.get("bereiche", {})
        bereich_table = doc.add_table(rows=1, cols=3)
        bereich_table.style = 'Table Grid'

        header_cells = bereich_table.rows[0].cells
        header_cells[0].text = "Nr."
        header_cells[1].text = "Bereich"
        header_cells[2].text = "Anzahl"

        for cell in header_cells:
            cell.paragraphs[0].runs[0].bold = True

        for bereich_nr, count in sorted(bereich_verteilung.items()):
            bereich_name = bereiche_config.get(bereich_nr, f"Bereich {bereich_nr}")
            row = bereich_table.add_row()
            row.cells[0].text = str(bereich_nr)
            row.cells[1].text = bereich_name
            row.cells[2].text = str(count)

    doc.add_paragraph()

    # Anlagen
    anlagen_header = doc.add_paragraph()
    anlagen_header.add_run("III. Anlagen").bold = True

    anlagen_text = doc.add_paragraph()
    anlagen_content = (
        "Dem Antrag füge ich bei:\n\n"
        "1. Fallliste 1 – Gerichtliche und rechtsförmliche Verfahren (Excel)\n"
        "2. Fallliste 2 – Außergerichtliche Verfahren (Excel)\n"
        "3. Zusammenfassende Übersicht mit FAO-Konformitätsprüfung\n"
        "4. Zeugnis über den erfolgreichen Abschluss des Fachanwaltslehrgangs\n"
        "5. Klausurzeugnisse (soweit separat ausgestellt)\n"
        "6. Nachweis der Berufshaftpflichtversicherung\n"
        "7. Erklärung gemäß § 7 FAO (Eigenverantwortliche Bearbeitung)"
    )
    anlagen_text.add_run(anlagen_content).font.size = Pt(11)

    doc.add_paragraph()

    # Schluss
    schluss = doc.add_paragraph()
    schluss_content = (
        "Für Rückfragen stehe ich Ihnen jederzeit gerne zur Verfügung.\n\n"
        "Mit freundlichen kollegialen Grüßen"
    )
    schluss.add_run(schluss_content).font.size = Pt(11)

    doc.add_paragraph()
    doc.add_paragraph()

    # Unterschrift
    unterschrift = doc.add_paragraph()
    unterschrift.add_run(antragsteller.get('name', '[Name]')).font.size = Pt(11)
    unterschrift.add_run("\nRechtsanwalt/Rechtsanwältin").font.size = Pt(10)

    # Dokument speichern
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


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
    if "upload_history" not in st.session_state:
        st.session_state.upload_history = []  # Historie der hochgeladenen Dateien
    if "editing_case_index" not in st.session_state:
        st.session_state.editing_case_index = None  # Index des aktuell bearbeiteten Falls
    if "selected_fachgebiet" not in st.session_state:
        st.session_state.selected_fachgebiet = None  # Für Fachgebiet-Wechsel

    # Antragsteller-Daten für Antragsschreiben
    if "antragsteller" not in st.session_state:
        st.session_state.antragsteller = {
            "name": "",
            "kanzlei": "",
            "strasse": "",
            "plz_ort": "",
            "telefon": "",
            "email": "",
            "ort": "",
            "zulassung_datum": "",
            "kammer_name": "",
            "kammer_strasse": "",
            "kammer_plz_ort": "",
            "lehrgang_info": "",
            "lehrgang_datum": ""
        }
    if "klausuren" not in st.session_state:
        st.session_state.klausuren = []  # Liste der bestandenen Klausuren

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

        # Session speichern/laden
        st.markdown("---")
        st.subheader("💾 Session verwalten")

        # Session speichern
        if not st.session_state.cases_df.empty:
            session_json = save_session_to_json(
                st.session_state.cases_df,
                fachgebiet,
                st.session_state.unprocessed_files,
                st.session_state.unrecognized_texts,
                st.session_state.upload_history
            )
            st.download_button(
                label="💾 Session speichern",
                data=session_json,
                file_name=f"fallliste_session_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                help="Speichern Sie Ihre Arbeit, um später fortzufahren"
            )

        # Session laden
        uploaded_session = st.file_uploader(
            "Session laden",
            type=["json"],
            help="Laden Sie eine gespeicherte Session",
            key="session_upload"
        )

        if uploaded_session:
            if st.button("📂 Session laden", type="secondary"):
                session_data = load_session_from_json(uploaded_session.read().decode("utf-8"))
                if session_data["success"]:
                    st.session_state.cases_df = session_data["cases_df"]
                    st.session_state.unprocessed_files = session_data["unprocessed_files"]
                    st.session_state.unrecognized_texts = session_data["unrecognized_texts"]
                    st.session_state.upload_history = session_data["upload_history"]
                    st.success(f"✓ Session geladen ({len(session_data['cases_df'])} Fälle)")
                    st.caption(f"Gespeichert am: {session_data['timestamp'][:16] if session_data['timestamp'] else 'Unbekannt'}")
                    st.rerun()
                else:
                    st.error(f"Fehler: {session_data['error']}")

        # Upload-Historie anzeigen
        if st.session_state.upload_history:
            st.markdown("---")
            st.subheader("📜 Upload-Historie")
            with st.expander(f"{len(st.session_state.upload_history)} Dateien verarbeitet"):
                for entry in st.session_state.upload_history[-10:]:  # Letzte 10
                    st.caption(f"• {entry.get('datei', 'Unbekannt')} ({entry.get('zeitpunkt', '')})")
            if st.button("🗑️ Historie löschen", key="clear_history"):
                st.session_state.upload_history = []
                st.rerun()

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
                                        # OCR-Unterstützung: Gibt (text, methode) zurück
                                        pdf_text, extraction_method = extract_text_from_pdf(pdf_file, use_ocr=OCR_AVAILABLE)

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

                                        # OCR-Hinweis
                                        if extraction_method == "ocr":
                                            st.info(f"🔍 {pdf_file.name}: Text via OCR extrahiert (gescanntes PDF)")

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

                                            # Upload-Historie aktualisieren
                                            st.session_state.upload_history.append({
                                                "datei": pdf_file.name,
                                                "faelle": len(cases),
                                                "methode": extraction_method,
                                                "zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                            })

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

                            # Nach erfolgreicher Analyse: Fälle speichern und Upload-Bereich zurücksetzen
                            if all_cases:
                                # Fälle zusammenführen
                                new_combined = pd.concat(all_cases, ignore_index=True)
                                total_count = len(new_combined)

                                # Duplikat-Prüfung
                                non_duplicates, duplicates = check_duplicates(new_combined, st.session_state.cases_df)

                                # Fälle in Session State speichern
                                if not non_duplicates.empty:
                                    if not st.session_state.cases_df.empty:
                                        st.session_state.cases_df = pd.concat(
                                            [st.session_state.cases_df, non_duplicates],
                                            ignore_index=True
                                        )
                                    else:
                                        st.session_state.cases_df = non_duplicates

                                # Status-Nachricht
                                if not duplicates.empty:
                                    st.info(f"ℹ️ {len(duplicates)} Duplikate wurden übersprungen.")

                                st.session_state.upload_key += 1
                                st.success(f"✅ Analyse abgeschlossen! {len(non_duplicates)} neue Fälle gespeichert.")
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
                                                # Text extrahieren (mit OCR-Unterstützung)
                                                try:
                                                    pdf_text, extraction_method = extract_text_from_pdf(buffer, use_ocr=OCR_AVAILABLE)

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
                                                        # OCR-Hinweis
                                                        if extraction_method == "ocr":
                                                            st.info(f"🔍 {file_name}: Text via OCR extrahiert")

                                                        # GPT-Analyse
                                                        cases = analyze_cases_with_gpt(client, pdf_text, fachgebiet, model=gpt_model)

                                                        if cases:
                                                            df = cases_list_to_dataframe(cases)
                                                            df = normalize_case_df(df, fachgebiet)
                                                            all_folder_cases.append(df)
                                                            st.session_state.last_processed_file = file_name

                                                            # Upload-Historie aktualisieren
                                                            st.session_state.upload_history.append({
                                                                "datei": file_name,
                                                                "faelle": len(cases),
                                                                "methode": extraction_method,
                                                                "quelle": "Google Drive Ordner",
                                                                "zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                                            })

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

                                    # Fälle speichern und Auswertung anzeigen
                                    if all_folder_cases:
                                        total_count = sum(len(df) for df in all_folder_cases)
                                        new_combined = pd.concat(all_folder_cases, ignore_index=True)

                                        # Duplikat-Prüfung
                                        non_duplicates, duplicates = check_duplicates(new_combined, st.session_state.cases_df)

                                        if not duplicates.empty:
                                            st.session_state.pending_duplicates.extend(duplicates.to_dict('records'))

                                        # Fälle in Session State speichern
                                        if not non_duplicates.empty:
                                            if not st.session_state.cases_df.empty:
                                                st.session_state.cases_df = pd.concat(
                                                    [st.session_state.cases_df, non_duplicates],
                                                    ignore_index=True
                                                )
                                            else:
                                                st.session_state.cases_df = non_duplicates

                                        # Ordner-Cache leeren
                                        st.session_state.folder_files = []

                                        st.success(f"✅ **{total_count} Fälle** aus {len(all_folder_cases)} Dateien extrahiert!")
                                        # Seite neu laden um Auswertung anzuzeigen
                                        st.rerun()
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

                                    # Prüfen ob PDF (mit OCR-Unterstützung)
                                    try:
                                        with st.spinner("Extrahiere Text aus PDF..."):
                                            pdf_text, extraction_method = extract_text_from_pdf(buffer, use_ocr=OCR_AVAILABLE)

                                        if extraction_method == "ocr":
                                            st.info("🔍 Text via OCR extrahiert (gescanntes PDF)")

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
                                                    st.session_state.last_processed_file = f"Cloud-Dokument ({provider})"

                                                    # Upload-Historie aktualisieren
                                                    st.session_state.upload_history.append({
                                                        "datei": f"Cloud-Dokument ({provider})",
                                                        "faelle": len(total_cases),
                                                        "methode": extraction_method,
                                                        "quelle": provider,
                                                        "zeitpunkt": datetime.now().strftime("%d.%m.%Y %H:%M")
                                                    })

                                                    # Duplikat-Prüfung
                                                    non_duplicates, duplicates = check_duplicates(df, st.session_state.cases_df)

                                                    if not duplicates.empty:
                                                        st.session_state.pending_duplicates.extend(duplicates.to_dict('records'))

                                                    # Fälle in Session State speichern
                                                    if not non_duplicates.empty:
                                                        if not st.session_state.cases_df.empty:
                                                            st.session_state.cases_df = pd.concat(
                                                                [st.session_state.cases_df, non_duplicates],
                                                                ignore_index=True
                                                            )
                                                        else:
                                                            st.session_state.cases_df = non_duplicates

                                                    st.success(f"✓ **{len(total_cases)} Fälle** aus Cloud-Dokument extrahiert!")
                                                    # Seite neu laden um Auswertung anzuzeigen
                                                    st.rerun()
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
    # PLAUSIBILITÄTSPRÜFUNG
    # =========================================================================

    st.markdown("---")
    st.header("🔍 Plausibilitätsprüfung")

    plausibility_result = check_all_cases_plausibility(filtered_df)

    # Zusammenfassung
    plaus_col1, plaus_col2, plaus_col3 = st.columns(3)

    with plaus_col1:
        if plausibility_result["fehler"] > 0:
            st.error(f"❌ **{plausibility_result['fehler']}** Fehler")
        else:
            st.success("✓ Keine Fehler")

    with plaus_col2:
        if plausibility_result["warnungen"] > 0:
            st.warning(f"⚠️ **{plausibility_result['warnungen']}** Warnungen")
        else:
            st.success("✓ Keine Warnungen")

    with plaus_col3:
        if plausibility_result["hinweise"] > 0:
            st.info(f"ℹ️ **{plausibility_result['hinweise']}** Hinweise")
        else:
            st.success("✓ Keine Hinweise")

    # Details anzeigen
    if plausibility_result["details"]:
        with st.expander(f"📋 Details zu {len(plausibility_result['details'])} Fällen mit Anmerkungen"):
            for idx, fall_info, warnings in plausibility_result["details"][:20]:  # Max 20 anzeigen
                st.markdown(f"**{fall_info}**")
                for w in warnings:
                    if w["typ"] == "fehler":
                        st.markdown(f"  - ❌ {w['meldung']}")
                    elif w["typ"] == "warnung":
                        st.markdown(f"  - ⚠️ {w['meldung']}")
                    else:
                        st.markdown(f"  - ℹ️ {w['meldung']}")
                st.markdown("---")

            if len(plausibility_result["details"]) > 20:
                st.caption(f"... und {len(plausibility_result['details']) - 20} weitere Fälle")

    # =========================================================================
    # FÄLLE BEARBEITEN
    # =========================================================================

    st.markdown("---")
    st.header("✏️ Fälle bearbeiten")

    with st.expander("Fall bearbeiten oder löschen"):
        if not filtered_df.empty:
            # Fall zum Bearbeiten auswählen
            case_options = []
            for idx, row in filtered_df.iterrows():
                case_options.append(f"{idx}: {row.get('kanzlei_az', 'Ohne AZ')} - {row.get('kurzrubrum', '')[:30]}")

            selected_case = st.selectbox("Fall auswählen", case_options, key="edit_case_select")

            if selected_case:
                case_idx = int(selected_case.split(":")[0])
                case_data = st.session_state.cases_df.loc[case_idx].to_dict()

                # Bearbeitungsformular
                edit_col1, edit_col2 = st.columns(2)

                with edit_col1:
                    new_kanzlei_az = st.text_input("Kanzlei-AZ", value=str(case_data.get("kanzlei_az", "")), key="edit_kanzlei_az")
                    new_kurzrubrum = st.text_input("Kurzrubrum", value=str(case_data.get("kurzrubrum", "")), key="edit_kurzrubrum")
                    new_gericht_az = st.text_input("Gerichts-AZ", value=str(case_data.get("gericht_az", "")), key="edit_gericht_az")

                with edit_col2:
                    new_verfahrenstyp = st.selectbox(
                        "Verfahrenstyp",
                        ["gerichtlich", "rechtsfoermlich", "aussergerichtlich"],
                        index=["gerichtlich", "rechtsfoermlich", "aussergerichtlich"].index(
                            str(case_data.get("verfahrenstyp", "aussergerichtlich")).lower()
                        ) if str(case_data.get("verfahrenstyp", "")).lower() in ["gerichtlich", "rechtsfoermlich", "aussergerichtlich"] else 2,
                        key="edit_verfahrenstyp"
                    )
                    bereiche_list = list(FAO_CONFIG[fachgebiet].get("bereiche", {}).keys())
                    current_bereich = case_data.get("bereich_nr", 1)
                    try:
                        bereich_index = bereiche_list.index(int(current_bereich)) if current_bereich else 0
                    except (ValueError, TypeError):
                        bereich_index = 0
                    new_bereich_nr = st.selectbox("Bereich-Nr", bereiche_list, index=bereich_index, key="edit_bereich_nr")

                new_sachverhalt = st.text_area(
                    "Sachverhalt (4-6 Sätze empfohlen)",
                    value=str(case_data.get("sachverhalt", "")),
                    height=150,
                    key="edit_sachverhalt"
                )

                # Buttons
                btn_col1, btn_col2, btn_col3 = st.columns(3)

                with btn_col1:
                    if st.button("💾 Änderungen speichern", type="primary", key="save_edit"):
                        st.session_state.cases_df.at[case_idx, "kanzlei_az"] = new_kanzlei_az
                        st.session_state.cases_df.at[case_idx, "kurzrubrum"] = new_kurzrubrum
                        st.session_state.cases_df.at[case_idx, "gericht_az"] = new_gericht_az
                        st.session_state.cases_df.at[case_idx, "verfahrenstyp"] = new_verfahrenstyp
                        st.session_state.cases_df.at[case_idx, "bereich_nr"] = new_bereich_nr
                        st.session_state.cases_df.at[case_idx, "sachverhalt"] = new_sachverhalt
                        st.session_state.cases_df.at[case_idx, "bereich_bezeichnung"] = FAO_CONFIG[fachgebiet]["bereiche"].get(new_bereich_nr, "")
                        st.success("✓ Änderungen gespeichert!")
                        st.rerun()

                with btn_col2:
                    if st.button("🗑️ Fall löschen", type="secondary", key="delete_case"):
                        st.session_state.cases_df = st.session_state.cases_df.drop(case_idx).reset_index(drop=True)
                        st.success("✓ Fall gelöscht!")
                        st.rerun()

                with btn_col3:
                    st.caption(f"Index: {case_idx}")

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

    # Download-Buttons
    download_col1, download_col2 = st.columns(2)

    with download_col1:
        st.download_button(
            label="📥 EXCEL herunterladen",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    with download_col2:
        # PDF-Report (falls verfügbar)
        if REPORTLAB_AVAILABLE:
            try:
                pdf_bytes = create_pdf_report(fl1, fl2, summary, fachgebiet)
                pdf_filename = f"Fallliste_{fachgebiet.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
                st.download_button(
                    label="📄 PDF-Report herunterladen",
                    data=pdf_bytes,
                    file_name=pdf_filename,
                    mime="application/pdf",
                    type="secondary",
                    use_container_width=True
                )
            except Exception as e:
                st.warning(f"PDF-Erstellung nicht möglich: {str(e)}")
        else:
            st.info("PDF-Export benötigt: `pip install reportlab`")

    # Statistik und Reset
    stat_col1, stat_col2, stat_col3 = st.columns([1, 1, 1])

    with stat_col1:
        st.metric("Fälle gesamt", summary["gesamt"])

    with stat_col2:
        erfuellt = sum(1 for c in summary["checks"] if c[3] == "erfuellt")
        st.metric("FAO-Kriterien erfüllt", f"{erfuellt}/{len(summary['checks'])}")

    with stat_col3:
        if st.button("🗑️ Zurücksetzen", help="Alle Fälle löschen und neu beginnen", use_container_width=True):
            st.session_state.cases_df = pd.DataFrame()
            st.session_state.upload_history = []
            st.session_state.unprocessed_files = []
            st.session_state.unrecognized_texts = []
            st.rerun()

    # =========================================================================
    # ANTRAGSSCHREIBEN GENERIEREN
    # =========================================================================

    st.markdown("---")
    st.header("📝 Antragsschreiben erstellen")

    st.markdown("""
    Erstellen Sie ein formelles **Antragsschreiben an die Rechtsanwaltskammer** mit:
    - Bezugnahme auf Ihre Falllisten und deren Verteilung
    - Auflistung der bestandenen Klausuren
    - Alle erforderlichen Anlagen
    """)

    with st.expander("📋 Antragsteller-Daten eingeben", expanded=False):
        st.subheader("Ihre Angaben")

        antrag_col1, antrag_col2 = st.columns(2)

        with antrag_col1:
            st.markdown("**Persönliche Daten:**")
            st.session_state.antragsteller["name"] = st.text_input(
                "Name (mit Titel)",
                value=st.session_state.antragsteller.get("name", ""),
                placeholder="Rechtsanwalt Max Mustermann",
                key="antrag_name"
            )
            st.session_state.antragsteller["kanzlei"] = st.text_input(
                "Kanzlei",
                value=st.session_state.antragsteller.get("kanzlei", ""),
                placeholder="Mustermann & Partner Rechtsanwälte",
                key="antrag_kanzlei"
            )
            st.session_state.antragsteller["strasse"] = st.text_input(
                "Straße",
                value=st.session_state.antragsteller.get("strasse", ""),
                placeholder="Musterstraße 123",
                key="antrag_strasse"
            )
            st.session_state.antragsteller["plz_ort"] = st.text_input(
                "PLZ Ort",
                value=st.session_state.antragsteller.get("plz_ort", ""),
                placeholder="12345 Musterstadt",
                key="antrag_plz_ort"
            )
            st.session_state.antragsteller["telefon"] = st.text_input(
                "Telefon",
                value=st.session_state.antragsteller.get("telefon", ""),
                placeholder="030 123456789",
                key="antrag_telefon"
            )
            st.session_state.antragsteller["email"] = st.text_input(
                "E-Mail",
                value=st.session_state.antragsteller.get("email", ""),
                placeholder="ra.mustermann@kanzlei.de",
                key="antrag_email"
            )
            st.session_state.antragsteller["ort"] = st.text_input(
                "Ort (für Datum)",
                value=st.session_state.antragsteller.get("ort", ""),
                placeholder="Berlin",
                key="antrag_ort"
            )
            st.session_state.antragsteller["zulassung_datum"] = st.text_input(
                "Zulassungsdatum",
                value=st.session_state.antragsteller.get("zulassung_datum", ""),
                placeholder="01.01.2020",
                key="antrag_zulassung"
            )

        with antrag_col2:
            st.markdown("**Rechtsanwaltskammer:**")
            st.session_state.antragsteller["kammer_name"] = st.text_input(
                "Name der Kammer",
                value=st.session_state.antragsteller.get("kammer_name", ""),
                placeholder="Rechtsanwaltskammer Berlin",
                key="antrag_kammer_name"
            )
            st.session_state.antragsteller["kammer_strasse"] = st.text_input(
                "Straße der Kammer",
                value=st.session_state.antragsteller.get("kammer_strasse", ""),
                placeholder="Littenstraße 9",
                key="antrag_kammer_strasse"
            )
            st.session_state.antragsteller["kammer_plz_ort"] = st.text_input(
                "PLZ Ort der Kammer",
                value=st.session_state.antragsteller.get("kammer_plz_ort", ""),
                placeholder="10179 Berlin",
                key="antrag_kammer_plz"
            )

            st.markdown("**Fachanwaltslehrgang:**")
            st.session_state.antragsteller["lehrgang_info"] = st.text_input(
                "Lehrgangsbezeichnung",
                value=st.session_state.antragsteller.get("lehrgang_info", ""),
                placeholder=f"Fachanwaltslehrgang {fachgebiet} der DAA",
                key="antrag_lehrgang"
            )
            st.session_state.antragsteller["lehrgang_datum"] = st.text_input(
                "Abschlussdatum Lehrgang",
                value=st.session_state.antragsteller.get("lehrgang_datum", ""),
                placeholder="15.10.2024",
                key="antrag_lehrgang_datum"
            )

        # Klausuren-Bereich
        st.markdown("---")
        st.subheader("Bestandene Klausuren")

        # Bestehende Klausuren anzeigen
        if st.session_state.klausuren:
            for i, klausur in enumerate(st.session_state.klausuren):
                klausur_col1, klausur_col2, klausur_col3, klausur_col4 = st.columns([3, 2, 2, 1])
                with klausur_col1:
                    st.text(f"📝 {klausur['bezeichnung']}")
                with klausur_col2:
                    st.text(klausur['datum'])
                with klausur_col3:
                    st.text(klausur['ergebnis'])
                with klausur_col4:
                    if st.button("🗑️", key=f"del_klausur_{i}"):
                        st.session_state.klausuren.pop(i)
                        st.rerun()

        # Neue Klausur hinzufügen
        st.markdown("**Neue Klausur hinzufügen:**")
        new_klausur_col1, new_klausur_col2, new_klausur_col3 = st.columns([3, 2, 2])

        with new_klausur_col1:
            new_klausur_bez = st.text_input(
                "Klausurbezeichnung",
                placeholder=f"Klausur 1: Materielles {fachgebiet}",
                key="new_klausur_bez"
            )
        with new_klausur_col2:
            new_klausur_datum = st.text_input(
                "Datum",
                placeholder="15.09.2024",
                key="new_klausur_datum"
            )
        with new_klausur_col3:
            new_klausur_ergebnis = st.selectbox(
                "Ergebnis",
                ["bestanden", "gut bestanden", "sehr gut bestanden", "mit Auszeichnung bestanden"],
                key="new_klausur_ergebnis"
            )

        if st.button("➕ Klausur hinzufügen", key="add_klausur"):
            if new_klausur_bez and new_klausur_datum:
                st.session_state.klausuren.append({
                    "bezeichnung": new_klausur_bez,
                    "datum": new_klausur_datum,
                    "ergebnis": new_klausur_ergebnis
                })
                st.success(f"✓ Klausur '{new_klausur_bez}' hinzugefügt!")
                st.rerun()
            else:
                st.warning("Bitte Bezeichnung und Datum eingeben.")

    # Antragsschreiben generieren
    if DOCX_AVAILABLE:
        # Bereichsverteilung berechnen
        bereich_counts = summary.get("bereich_counts", {})

        # Prüfen ob genug Daten vorhanden
        antragsteller_komplett = all([
            st.session_state.antragsteller.get("name"),
            st.session_state.antragsteller.get("kammer_name"),
            st.session_state.antragsteller.get("zulassung_datum")
        ])

        if antragsteller_komplett and st.session_state.klausuren:
            try:
                antrags_bytes = create_antragsschreiben(
                    antragsteller=st.session_state.antragsteller,
                    fachgebiet=fachgebiet,
                    summary=summary,
                    klausuren=st.session_state.klausuren,
                    bereich_verteilung=bereich_counts,
                    fl1_count=len(fl1),
                    fl2_count=len(fl2)
                )

                antrags_filename = f"Antrag_Fachanwalt_{fachgebiet.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx"

                st.download_button(
                    label="📝 ANTRAGSSCHREIBEN HERUNTERLADEN (Word)",
                    data=antrags_bytes,
                    file_name=antrags_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )

                st.success("✓ Antragsschreiben bereit zum Download!")

            except Exception as e:
                st.error(f"Fehler bei der Erstellung: {str(e)}")
        else:
            missing = []
            if not st.session_state.antragsteller.get("name"):
                missing.append("Name")
            if not st.session_state.antragsteller.get("kammer_name"):
                missing.append("Rechtsanwaltskammer")
            if not st.session_state.antragsteller.get("zulassung_datum"):
                missing.append("Zulassungsdatum")
            if not st.session_state.klausuren:
                missing.append("Klausuren")

            st.info(f"📋 Bitte ergänzen Sie: {', '.join(missing)}")
            st.caption("Öffnen Sie den Bereich 'Antragsteller-Daten eingeben' oben, um die Daten zu erfassen.")
    else:
        st.warning("Word-Export benötigt: `pip install python-docx`")


if __name__ == "__main__":
    main()
