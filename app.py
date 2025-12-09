"""
Fachanwalt-Falllistenverwaltung - Streamlit App
================================================
Eine Anwendung zur Verwaltung von Fachanwaltsfällen mit FAO-Konformitätsprüfung.
Ermöglicht Upload, Sortierung und Export von Falllisten gemäß § 5 FAO.

Verwendung: streamlit run app.py
Voraussetzungen: pip install streamlit pandas openpyxl
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date
from typing import Dict, List, Tuple, Optional, Any
from enum import Enum

# =============================================================================
# KONFIGURATION - FAO-Mindestanforderungen (§ 5 FAO, Stand 01.06.2022)
# =============================================================================

FAO_CONFIG: Dict[str, Dict[str, Any]] = {
    "Erbrecht": {
        "paragraph": "§ 5 Abs. 1 lit. m, § 14f FAO",
        "gesamt_min": 80,
        "gerichtlich_min": 20,
        "aussergerichtlich_min": 60,
        "fg_max": 15,  # Max. Verfahren der freiwilligen Gerichtsbarkeit
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
        "gerichtlich_min": 50,  # mindestens die Hälfte
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
        "gerichtlich_min": 60,  # gewillkürte/nötige Verbundverfahren zählen doppelt
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
        "gerichtlich_min": 40,  # Streit-/Schieds-/Mediationsverfahren + Gestaltung/Gründung
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

# Standard-Spalten für die Template-Excel
TEMPLATE_COLUMNS = [
    "Kanzlei-AZ", "Kurzrubrum", "Sachverhalt", "Zeitraum_von", "Zeitraum_bis",
    "Gerichts-AZ", "Verfahrenstyp", "Verfahrensart", "Bereich_Nr", "Bereich_Bezeichnung",
    "Bedeutung", "Taetigkeitsbeschreibung", "Stand", "Abschluss_Art",
    "Abschluss_Datum", "Verbundener_Fall", "Fachgebiet"
]

# =============================================================================
# HILFSFUNKTIONEN
# =============================================================================

def map_column_name(col_name: str) -> Optional[str]:
    """
    Mappt einen Spaltennamen auf den internen Feldnamen.

    Args:
        col_name: Ursprünglicher Spaltenname aus der Datei

    Returns:
        Interner Feldname oder None wenn keine Zuordnung gefunden
    """
    col_normalized = col_name.strip()
    for internal_name, variants in COLUMN_MAPPING.items():
        if col_normalized in variants:
            return internal_name
    return None


def load_cases_from_file(file) -> pd.DataFrame:
    """
    Liest Fälle aus einer CSV- oder Excel-Datei.

    Args:
        file: Hochgeladene Datei (Streamlit UploadedFile)

    Returns:
        DataFrame mit den eingelesenen Fällen
    """
    filename = file.name.lower()

    if filename.endswith('.csv'):
        df = pd.read_csv(file, encoding='utf-8')
    elif filename.endswith('.xlsx') or filename.endswith('.xls'):
        df = pd.read_excel(file, engine='openpyxl')
    else:
        raise ValueError(f"Nicht unterstütztes Dateiformat: {filename}")

    return df


def normalize_case_df(df: pd.DataFrame, fachgebiet: str) -> pd.DataFrame:
    """
    Normalisiert einen DataFrame auf das interne Datenmodell.
    Mappt Spaltennamen und ergänzt fehlende Standardfelder.

    Args:
        df: Roher DataFrame aus der Datei
        fachgebiet: Gewähltes Fachgebiet für Default-Zuordnung

    Returns:
        Normalisierter DataFrame mit einheitlichen Spaltennamen
    """
    # Spalten-Mapping anwenden
    new_columns = {}
    for col in df.columns:
        mapped = map_column_name(col)
        if mapped:
            new_columns[col] = mapped
        else:
            # Behalte Original-Spalte mit lowercase
            new_columns[col] = col.lower().replace("-", "_").replace(" ", "_")

    df = df.rename(columns=new_columns)

    # Standardfelder ergänzen falls nicht vorhanden
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

    # Wenn kein Fachgebiet in der Datei, das gewählte verwenden
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

    # Verfahrensart normalisieren
    df["verfahrensart"] = df["verfahrensart"].astype(str).str.lower().str.strip()

    # Bedeutung normalisieren
    df["bedeutung"] = df["bedeutung"].astype(str).str.lower().str.strip()
    df["bedeutung"] = df["bedeutung"].replace({
        "niedrig": "gering",
        "hoch": "hoch",
        "mittel": "mittel",
    })

    # Bereich_Nr als Integer
    df["bereich_nr"] = pd.to_numeric(df["bereich_nr"], errors='coerce').fillna(0).astype(int)

    return df


def split_into_falllisten(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Trennt Fälle in gerichtliche/rechtsförmliche und außergerichtliche Verfahren.

    Args:
        df: Normalisierter DataFrame mit allen Fällen

    Returns:
        Tuple aus (FL1: gerichtlich/rechtsförmlich, FL2: außergerichtlich)
    """
    mask_gerichtlich = df["verfahrenstyp"].isin(["gerichtlich", "rechtsfoermlich"])

    fl1 = df[mask_gerichtlich].copy()
    fl2 = df[~mask_gerichtlich].copy()

    # Fortlaufende Nummern vergeben
    fl1.insert(0, "FL1_Nr", range(1, len(fl1) + 1))
    fl2.insert(0, "FL2_Nr", range(1, len(fl2) + 1))

    return fl1, fl2


def count_by_bereich(df: pd.DataFrame) -> Dict[int, int]:
    """Zählt Fälle pro Bereich-Nummer."""
    return df.groupby("bereich_nr").size().to_dict()


def compute_summary(df: pd.DataFrame, fachgebiet: str) -> Dict[str, Any]:
    """
    Berechnet Ist-Zahlen und vergleicht mit FAO-Mindestanforderungen.

    Args:
        df: Normalisierter DataFrame (nur für das gewählte Fachgebiet)
        fachgebiet: Name des Fachgebiets

    Returns:
        Dictionary mit Summary-Daten und Warnungen
    """
    config = FAO_CONFIG.get(fachgebiet, {})

    # Grundzählung
    gesamt = len(df)
    gerichtlich = len(df[df["verfahrenstyp"].isin(["gerichtlich", "rechtsfoermlich"])])
    aussergerichtlich = len(df[df["verfahrenstyp"] == "aussergerichtlich"])

    # fG-Verfahren zählen (für Erbrecht relevant)
    fg_verfahren = len(df[df["verfahrensart"].str.contains("fg|freiwillige", case=False, na=False)])

    # Bereichsverteilung
    bereich_counts = count_by_bereich(df)

    # Verbundverfahren (für Familienrecht - zählen doppelt)
    verbund_count = 0
    if fachgebiet == "Familienrecht" and config.get("verbund_doppelt"):
        verbund_count = len(df[df["verfahrensart"].str.contains("verbund", case=False, na=False)])
        gerichtlich_effektiv = gerichtlich + verbund_count  # Verbundverfahren zählen doppelt
    else:
        gerichtlich_effektiv = gerichtlich

    # FAO-Prüfung
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

    # Außergerichtliche Verfahren (Erbrecht-spezifisch)
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

            # Detaillierte Bereichs-Warnung
            for bereich_nr, bereich_name in config.get("bereiche", {}).items():
                count = bereich_counts.get(bereich_nr, 0)
                if count < min_faelle:
                    warnungen.append(f"Bereich {bereich_nr} ({bereich_name}): nur {count}/{min_faelle} Fälle")

    # Spezial-Anforderungen (Arbeitsrecht, HGR)
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

    # Zusätzliche Warnungen
    # Zeitraum-Prüfung (3 Jahre)
    heute = date.today()
    drei_jahre_zuvor = date(heute.year - 3, heute.month, heute.day)

    # Prüfe auf Fälle außerhalb des 3-Jahres-Zeitraums
    for idx, row in df.iterrows():
        try:
            abschluss = row.get("abschluss_datum")
            if pd.notna(abschluss) and abschluss != "":
                if isinstance(abschluss, str):
                    # Versuche verschiedene Datumsformate
                    for fmt in ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"]:
                        try:
                            parsed = datetime.strptime(abschluss, fmt).date()
                            if parsed < drei_jahre_zuvor:
                                warnungen.append(f"Fall {row.get('kanzlei_az', idx)}: Abschluss außerhalb des 3-Jahres-Zeitraums ({abschluss})")
                            break
                        except ValueError:
                            continue
                elif isinstance(abschluss, (datetime, date)):
                    if abschluss.date() if isinstance(abschluss, datetime) else abschluss < drei_jahre_zuvor:
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

    # Kopfzeile mit Fachgebiet
    rows.append({"Kategorie": "FACHGEBIET", "Wert": summary["fachgebiet"], "Anforderung": summary["paragraph"]})
    rows.append({"Kategorie": "", "Wert": "", "Anforderung": ""})

    # Kennzahlen-Überschrift
    rows.append({"Kategorie": "=== KENNZAHLEN ===", "Wert": "", "Anforderung": ""})

    # FAO-Checks
    for check in summary["checks"]:
        kriterium, ist, soll, status = check
        status_symbol = {"erfuellt": "✓", "knapp": "⚠️", "nicht_erfuellt": "✗"}.get(status, "?")
        rows.append({
            "Kategorie": kriterium,
            "Wert": f"{ist}/{soll}",
            "Anforderung": status_symbol
        })

    rows.append({"Kategorie": "", "Wert": "", "Anforderung": ""})

    # Bereichsverteilung
    rows.append({"Kategorie": "=== BEREICHSVERTEILUNG ===", "Wert": "", "Anforderung": ""})

    for bereich_nr, bereich_name in summary.get("bereiche_config", {}).items():
        count = summary["bereich_counts"].get(bereich_nr, 0)
        rows.append({
            "Kategorie": f"Bereich {bereich_nr}",
            "Wert": count,
            "Anforderung": bereich_name
        })

    rows.append({"Kategorie": "", "Wert": "", "Anforderung": ""})

    # Warnungen
    if summary["warnungen"]:
        rows.append({"Kategorie": "=== WARNUNGEN ===", "Wert": "", "Anforderung": ""})
        for warnung in summary["warnungen"]:
            rows.append({"Kategorie": "⚠️", "Wert": warnung, "Anforderung": ""})
    else:
        rows.append({"Kategorie": "=== STATUS ===", "Wert": "Keine Warnungen", "Anforderung": "✓"})

    return pd.DataFrame(rows)


def create_excel(fl1_df: pd.DataFrame, fl2_df: pd.DataFrame,
                 summary: Dict[str, Any], fachgebiet: str) -> bytes:
    """
    Erstellt Excel-Arbeitsmappe mit Falllisten und Summary.

    Args:
        fl1_df: DataFrame für Fallliste 1 (gerichtlich/rechtsförmlich)
        fl2_df: DataFrame für Fallliste 2 (außergerichtlich)
        summary: Summary-Dictionary aus compute_summary()
        fachgebiet: Name des Fachgebiets

    Returns:
        Excel-Datei als Bytes
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Fallliste 1
        fl1_export = prepare_fl1_for_export(fl1_df) if len(fl1_df) > 0 else pd.DataFrame()
        fl1_export.to_excel(writer, sheet_name="Fallliste_1", index=False)

        # Fallliste 2
        fl2_export = prepare_fl2_for_export(fl2_df) if len(fl2_df) > 0 else pd.DataFrame()
        fl2_export.to_excel(writer, sheet_name="Fallliste_2", index=False)

        # Summary
        summary_df = create_summary_df(summary)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    return output.getvalue()


def create_template_excel() -> bytes:
    """Erstellt eine leere Excel-Vorlage mit den erwarteten Spalten."""
    output = BytesIO()

    # Leeres DataFrame mit Spalten
    df = pd.DataFrame(columns=TEMPLATE_COLUMNS)

    # Eine Beispielzeile hinzufügen
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

        # Zusätzliches Blatt mit Erklärungen
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


def get_status_color(status: str) -> str:
    """Gibt Farbe für Status zurück."""
    return {
        "erfuellt": "green",
        "knapp": "orange",
        "nicht_erfuellt": "red"
    }.get(status, "gray")


def get_status_symbol(status: str) -> str:
    """Gibt Symbol für Status zurück."""
    return {
        "erfuellt": "✓",
        "knapp": "⚠️",
        "nicht_erfuellt": "✗"
    }.get(status, "?")


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
    st.markdown("**FAO-konforme Falllisten für Fachanwaltsanträge**")

    # Sidebar
    with st.sidebar:
        st.header("📋 Einstellungen")

        # Fachgebiet-Auswahl
        fachgebiet = st.selectbox(
            "Fachgebiet auswählen",
            options=list(FAO_CONFIG.keys()),
            help="Wählen Sie das Fachgebiet für die Falllistenerstellung"
        )

        # Hinweis zu Mindestanforderungen
        st.markdown("---")
        st.subheader("📌 Anforderungen")
        config = FAO_CONFIG[fachgebiet]
        st.markdown(f"**{config['paragraph']}**")
        st.info(config["hinweis"])

        # Bereiche anzeigen
        if "bereiche" in config:
            st.markdown("**Bereiche:**")
            for nr, name in config["bereiche"].items():
                st.markdown(f"- {nr}. {name}")

        st.markdown("---")

        # Template-Download
        st.subheader("📥 Vorlage")
        template_bytes = create_template_excel()
        st.download_button(
            label="Excel-Vorlage herunterladen",
            data=template_bytes,
            file_name="fallliste_vorlage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Laden Sie eine leere Excel-Vorlage mit den erwarteten Spalten herunter"
        )

    # Hauptbereich
    st.header("📤 Dateien hochladen")

    uploaded_files = st.file_uploader(
        "Fall-Dateien hochladen (CSV oder Excel)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        help="Laden Sie eine oder mehrere Dateien mit Ihren Fällen hoch"
    )

    if not uploaded_files:
        st.info(
            "👈 Bitte laden Sie Ihre Fall-Dateien hoch oder laden Sie die Vorlage herunter, "
            "um das erwartete Format zu sehen."
        )

        # Anleitung anzeigen
        with st.expander("📖 Anleitung zur Verwendung"):
            st.markdown("""
            ### So verwenden Sie diese App:

            1. **Fachgebiet wählen**: Wählen Sie in der Seitenleiste das gewünschte Fachgebiet aus.

            2. **Vorlage herunterladen**: Laden Sie die Excel-Vorlage herunter, um das erwartete Dateiformat zu sehen.

            3. **Fälle eintragen**: Füllen Sie die Vorlage mit Ihren Fällen aus. Jede Zeile entspricht einem Fall.

            4. **Dateien hochladen**: Laden Sie Ihre ausgefüllte(n) Datei(en) hier hoch.

            5. **Auswertung prüfen**: Die App zeigt Ihnen automatisch:
               - Die Aufteilung in gerichtliche und außergerichtliche Verfahren
               - Den FAO-Konformitätscheck mit allen Mindestanforderungen
               - Warnungen bei fehlenden Fällen oder Problemen

            6. **Export**: Laden Sie die fertige Fallliste als Excel-Datei herunter.

            ### Wichtige Hinweise:

            - **Verfahrenstyp**: Verwenden Sie `gerichtlich`, `rechtsfoermlich` oder `aussergerichtlich`
            - **Bereich_Nr**: Tragen Sie die Nummer des Fachbereichs gem. FAO ein (siehe Seitenleiste)
            - **Zeitraum**: Fälle sollten innerhalb der letzten 3 Jahre abgeschlossen sein
            """)
        return

    # Dateien verarbeiten
    all_cases = []

    with st.spinner("Dateien werden verarbeitet..."):
        for file in uploaded_files:
            try:
                df = load_cases_from_file(file)
                df = normalize_case_df(df, fachgebiet)
                all_cases.append(df)
                st.success(f"✓ {file.name}: {len(df)} Fälle geladen")
            except Exception as e:
                st.error(f"✗ {file.name}: Fehler beim Laden - {str(e)}")

    if not all_cases:
        st.warning("Keine Fälle konnten geladen werden. Bitte prüfen Sie das Dateiformat.")
        return

    # Alle Fälle zusammenführen
    combined_df = pd.concat(all_cases, ignore_index=True)

    # Nach Fachgebiet filtern
    if "fachgebiet" in combined_df.columns:
        filtered_df = combined_df[combined_df["fachgebiet"] == fachgebiet].copy()
        if len(filtered_df) == 0:
            st.warning(f"Keine Fälle für das Fachgebiet '{fachgebiet}' gefunden. "
                      f"Es werden alle {len(combined_df)} Fälle verwendet.")
            filtered_df = combined_df.copy()
    else:
        filtered_df = combined_df.copy()

    st.markdown("---")

    # Falllisten aufteilen
    fl1, fl2 = split_into_falllisten(filtered_df)

    # Summary berechnen
    summary = compute_summary(filtered_df, fachgebiet)

    # Dashboard-Bereich
    st.header("📊 Dashboard")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            label="Gesamtfälle",
            value=summary["gesamt"],
            delta=f"Soll: {FAO_CONFIG[fachgebiet]['gesamt_min']}"
        )

    with col2:
        st.metric(
            label="Gerichtlich/Rechtsförmlich",
            value=summary["gerichtlich"],
            delta=f"Soll: {FAO_CONFIG[fachgebiet].get('gerichtlich_min', 0)}"
        )

    with col3:
        st.metric(
            label="Außergerichtlich",
            value=summary["aussergerichtlich"],
            delta=f"Soll: {FAO_CONFIG[fachgebiet].get('aussergerichtlich_min', 0)}" if FAO_CONFIG[fachgebiet].get('aussergerichtlich_min', 0) > 0 else None
        )

    with col4:
        erfuellt_count = sum(1 for c in summary["checks"] if c[3] == "erfuellt")
        total_checks = len(summary["checks"])
        st.metric(
            label="FAO-Kriterien erfüllt",
            value=f"{erfuellt_count}/{total_checks}",
            delta="Alle erfüllt" if erfuellt_count == total_checks else f"{total_checks - erfuellt_count} offen"
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
            color = get_status_color(status)

            if status == "erfuellt":
                st.success(f"{symbol} **{kriterium}**: {ist}/{soll}")
            elif status == "knapp":
                st.warning(f"{symbol} **{kriterium}**: {ist}/{soll}")
            else:
                st.error(f"{symbol} **{kriterium}**: {ist}/{soll}")

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

    # Warnungen
    if summary["warnungen"]:
        st.markdown("---")
        st.header("⚠️ Warnungen")

        for warnung in summary["warnungen"]:
            st.warning(warnung)

    # Falllisten-Vorschau
    st.markdown("---")
    st.header("📋 Falllisten-Vorschau")

    tab1, tab2 = st.tabs([
        f"Fallliste 1 - Gerichtlich/Rechtsförmlich ({len(fl1)} Fälle)",
        f"Fallliste 2 - Außergerichtlich ({len(fl2)} Fälle)"
    ])

    with tab1:
        if len(fl1) > 0:
            display_cols = ["FL1_Nr", "kurzrubrum", "kanzlei_az", "gericht_az",
                          "bereich_nr", "verfahrenstyp", "verfahrensart", "bedeutung", "stand"]
            display_cols = [c for c in display_cols if c in fl1.columns]
            st.dataframe(fl1[display_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Keine gerichtlichen/rechtsförmlichen Verfahren vorhanden.")

    with tab2:
        if len(fl2) > 0:
            display_cols = ["FL2_Nr", "kurzrubrum", "kanzlei_az",
                          "bereich_nr", "bedeutung"]
            display_cols = [c for c in display_cols if c in fl2.columns]
            st.dataframe(fl2[display_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Keine außergerichtlichen Verfahren vorhanden.")

    # Excel-Download
    st.markdown("---")
    st.header("📥 Export")

    excel_bytes = create_excel(fl1, fl2, summary, fachgebiet)

    filename = f"fallliste_{fachgebiet.lower().replace(' ', '_').replace('-', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"

    col1, col2 = st.columns([1, 3])

    with col1:
        st.download_button(
            label="📥 Excel-Datei herunterladen",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    with col2:
        st.info(f"Die Excel-Datei enthält 3 Tabellenblätter: Fallliste_1, Fallliste_2 und Summary.")


if __name__ == "__main__":
    main()
