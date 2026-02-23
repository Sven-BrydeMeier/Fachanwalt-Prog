# Fachanwalt-Falllistenverwaltung - Vollständige Spezifikation für React-Migration

**Version:** 25.12.13-23:33
**Repository:** https://github.com/Sven-BrydeMeier/Fachanwalt-Prog
**Aktuelle Implementierung:** Streamlit (Python, Single-File `app.py`, 3.403 Zeilen)

---

## 1. PROJEKTÜBERSICHT

### 1.1 Zweck
Web-Anwendung für Rechtsanwälte zur Erstellung FAO-konformer Falllisten (§ 5 Fachanwaltsordnung) für den Antrag auf Verleihung einer Fachanwaltsbezeichnung bei der Rechtsanwaltskammer.

### 1.2 Kern-Workflow
```
PDF-Akten hochladen → ChatGPT analysiert → Fälle extrahiert → FAO-Check → Excel/Word/PDF Export
```

### 1.3 Technologie-Stack (Aktuell → Ziel)
| Aktuell (Streamlit) | Ziel (React) |
|---|---|
| Python/Streamlit | React + TypeScript |
| Session State | Redux / Zustand / Context |
| OpenAI Python SDK | OpenAI REST API / Node.js SDK |
| Pandas DataFrame | JavaScript Arrays/Objects |
| openpyxl | SheetJS (xlsx) |
| python-docx | docx.js |
| ReportLab | jsPDF / react-pdf |
| PyPDF2 + Tesseract | pdf.js + Tesseract.js |

---

## 2. ROLLEN UND BERECHTIGUNGEN

### 2.1 Rollen-Definition

| Rolle | Beschreibung | Berechtigung |
|---|---|---|
| **Anwalt (Antragsteller)** | Hauptnutzer - erstellt Fallliste für FAO-Antrag | Alle Funktionen: Upload, Analyse, Bearbeitung, Export |
| **Kanzlei-Admin** *(Erweiterung)* | Verwaltet mehrere Anwälte in einer Kanzlei | Kann Sessions aller Anwälte einsehen, Vorlagen verwalten |
| **Prüfer** *(Erweiterung)* | Prüft Falllisten vor Einreichung (Vier-Augen-Prinzip) | Nur Lesen + Kommentare + Freigabe/Ablehnung |
| **System** | Automatische Prozesse | GPT-Analyse, FAO-Validierung, Export-Generierung |

### 2.2 Rollenspezifische Funktionen

#### Anwalt (Antragsteller)
```
├── Fälle erfassen
│   ├── PDF-Upload (einzeln/mehrfach)
│   ├── Cloud-Upload (Google Drive Ordner/Datei, Dropbox)
│   ├── Manuelle Eingabe (Freitext + KI oder Strukturiert)
│   └── Session laden (JSON-Import)
├── Fälle verwalten
│   ├── Fälle bearbeiten (alle Felder)
│   ├── Fälle löschen (einzeln)
│   ├── Duplikate prüfen/annehmen/verwerfen
│   └── Plausibilitätsprüfung einsehen
├── Auswertung
│   ├── Dashboard (Kennzahlen)
│   ├── FAO-Konformitätscheck
│   ├── Bereichsverteilung (Balkendiagramm)
│   └── Falllisten-Vorschau (FL1/FL2)
├── Export
│   ├── Excel-Fallliste herunterladen
│   ├── PDF-Report herunterladen
│   ├── Antragsschreiben (Word) generieren
│   └── Session speichern (JSON-Export)
└── Antragsverwaltung
    ├── Persönliche Daten eingeben
    ├── Kammer-Daten eingeben
    ├── Lehrgang-Daten eingeben
    └── Klausuren verwalten (hinzufügen/löschen)
```

---

## 3. ALLE FUNKTIONEN (32 Funktionen)

### 3.1 PDF-Verarbeitung

#### `extractTextFromPdf(file, useOcr)`
- **Input:** PDF-Datei (File/Blob), Boolean
- **Output:** `{ text: string, method: 'text' | 'ocr' }`
- **Logik:**
  1. Versuche Text mit pdf.js extrahieren
  2. Wenn < 100 Zeichen und OCR aktiv: Konvertiere Seiten zu Bildern
  3. Tesseract OCR mit Sprache `deu` auf jede Seite
  4. Gib längeren Text zurück
- **Aufgerufen von:** PDF-Upload, Cloud-Upload

#### `checkFileSize(file)`
- **Input:** File
- **Output:** `{ ok: boolean, error: string }`
- **Logik:** Prüft `file.size <= MAX_FILE_SIZE_BYTES` (100 MB)

#### `checkTotalUploadSize(files)`
- **Input:** File[]
- **Output:** `{ ok: boolean, error: string, totalMb: number }`
- **Logik:** Summe aller Dateigrößen <= 500 MB

### 3.2 Cloud-Integration

#### `parseCloudLink(url)`
- **Input:** URL-String
- **Output:** `{ provider: string | null, identifier: string | null }`
- **Provider-Erkennung:**
  - `google_drive_file`: `/file/d/{ID}/` oder `?id={ID}`
  - `google_drive_folder`: `/folders/{ID}` oder `?id={ID}` mit `folders`
  - `dropbox`: `dropbox.com` URLs
- **Regex-Patterns:**
  ```
  Google File: /\/file\/d\/([a-zA-Z0-9_-]+)/
  Google Folder: /\/folders\/([a-zA-Z0-9_-]+)/
  ```

#### `listGoogleDriveFolder(folderId)`
- **Input:** Google Drive Folder-ID
- **Output:** `Array<{ name: string, id: string }>`
- **Logik:**
  1. Versuch 1: Google Drive API v3 (öffentlich)
     `GET https://www.googleapis.com/drive/v3/files?q='{folderId}'+in+parents+and+mimeType='application/pdf'`
  2. Versuch 2: Webpage Scraping über `embeddedfolderview`
  3. Regex: Dateien aus HTML extrahieren

#### `downloadFromGoogleDrive(fileId, progressCallback)`
- **Input:** File-ID, Progress-Callback
- **Output:** `Blob | null`
- **Logik:** Chunk-basierter Download (5 MB Chunks) mit Virus-Scan-Handling
  - URL: `https://drive.google.com/uc?export=download&id={fileId}`
  - Wenn Virus-Scan-Warning: Confirm-Token extrahieren und erneut laden

#### `downloadFromDropbox(url, progressCallback)`
- **Input:** Dropbox-URL
- **Output:** `Blob | null`
- **Logik:** `dl=0` → `dl=1` ersetzen für Direktdownload

#### `downloadFromCloud(url, progressCallback)`
- **Input:** Cloud-URL
- **Output:** `{ buffer: Blob | null, error: string }`
- **Logik:** Ruft `parseCloudLink` → dann den passenden Download-Provider

### 3.3 OpenAI/GPT-Integration

#### `getOpenAIClient(apiKey)`
- **Input:** API-Key String
- **Output:** OpenAI Client oder null
- **Logik:** Erstellt OpenAI-Client, fängt Auth-Fehler ab

#### `analyzeCasesWithGPT(client, text, fachgebiet, model)`
- **Input:** Client, PDF-Text (max 12.000 Zeichen), Fachgebiet, Modell
- **Output:** `Array<CaseData>`
- **GPT-Prompt (Zusammenfassung):**
  ```
  "Du bist ein Experte für deutsches Recht, spezialisiert auf {fachgebiet}.
  Analysiere den Text und extrahiere ALLE erkennbaren Fälle.
  Für jeden Fall erstelle einen Sachverhalt von 4-6 Sätzen.
  Antworte NUR als JSON-Array."
  ```
- **JSON-Schema pro Fall:**
  ```json
  {
    "kanzlei_az": "string",
    "kurzrubrum": "string (Mandant A ./. Mandant B)",
    "sachverhalt": "string (4-6 Sätze, detailliert)",
    "zeitraum_von": "string (MM/YYYY)",
    "zeitraum_bis": "string (MM/YYYY)",
    "gericht_az": "string",
    "verfahrenstyp": "gerichtlich | rechtsfoermlich | aussergerichtlich",
    "verfahrensart": "string",
    "bereich_nr": "number (1-5)",
    "bereich_bezeichnung": "string",
    "bedeutung": "gering | mittel | hoch",
    "taetigkeitsbeschreibung": "string",
    "stand": "anhaengig | abgeschlossen",
    "abschluss_art": "string",
    "abschluss_datum": "string",
    "verbundener_fall": "string"
  }
  ```
- **Retry-Logik:** 3 Versuche mit 2s, 5s, 10s Pause
- **Fehlerbehandlung:** APIError, RateLimitError, Timeout, JSON-Parse-Fehler

#### `analyzeSingleCaseWithGPT(client, text, fachgebiet, model)`
- **Input:** Client, Freitext, Fachgebiet, Modell
- **Output:** `CaseData | null`
- **Logik:** Gleich wie `analyzeCasesWithGPT`, aber für einzelnen Fall

### 3.4 Daten-Transformation

#### `mapColumnName(colName)`
- **Input:** Beliebiger Spaltenname (z.B. "Aktenzeichen")
- **Output:** Interner Feldname (z.B. "kanzlei_az")
- **Mapping-Tabelle:**
  ```javascript
  const COLUMN_MAPPING = {
    kanzlei_az: ["kanzlei_az", "aktenzeichen", "az", "kanzlei-az", "kanzleiaz", "int. az"],
    kurzrubrum: ["kurzrubrum", "rubrum", "parteien", "beteilige", "mandant"],
    sachverhalt: ["sachverhalt", "beschreibung", "sachverhaltsdarstellung", "zusammenfassung"],
    zeitraum_von: ["zeitraum_von", "von", "beginn", "start", "bearbeitungsbeginn"],
    zeitraum_bis: ["zeitraum_bis", "bis", "ende", "abschluss_datum"],
    gericht_az: ["gericht_az", "gerichtsaktenzeichen", "gerichts-az", "az_gericht"],
    verfahrenstyp: ["verfahrenstyp", "typ", "art_des_verfahrens", "verfahren"],
    verfahrensart: ["verfahrensart", "verfahren_art"],
    bereich_nr: ["bereich_nr", "bereich", "bereichsnummer", "nr"],
    bereich_bezeichnung: ["bereich_bezeichnung", "bereichsbezeichnung", "bereich_name"],
    bedeutung: ["bedeutung", "schwierigkeit", "gewicht"],
    taetigkeitsbeschreibung: ["taetigkeitsbeschreibung", "taetigkeit", "tätigkeit"],
    stand: ["stand", "status", "verfahrensstand"],
    abschluss_art: ["abschluss_art", "abschlussart", "ergebnis"],
    abschluss_datum: ["abschluss_datum", "abschlussdatum"],
    verbundener_fall: ["verbundener_fall", "verbunden", "zusammenhang"],
    fachgebiet: ["fachgebiet", "rechtsgebiet", "gebiet"]
  }
  ```

#### `casesListToDataFrame(cases)`
- **Input:** `Array<CaseData>`
- **Output:** Normalisiertes Array mit allen Pflichtfeldern

#### `normalizeCaseDF(cases, fachgebiet)`
- **Input:** Case-Array, Fachgebiet
- **Output:** Normalisiertes Array mit:
  - Spalten-Mapping angewendet
  - Fehlende Spalten mit Defaults gefüllt
  - `bereich_nr` als Number konvertiert
  - `fachgebiet` gesetzt

#### `splitIntoFalllisten(cases)`
- **Input:** Case-Array
- **Output:** `{ fl1: Array, fl2: Array }`
- **Logik:**
  - **FL1** (Fallliste 1): `verfahrenstyp IN ('gerichtlich', 'rechtsfoermlich')` → bekommt `FL1_Nr`
  - **FL2** (Fallliste 2): `verfahrenstyp = 'aussergerichtlich'` → bekommt `FL2_Nr`

#### `countByBereich(cases)`
- **Input:** Case-Array
- **Output:** `{ [bereich_nr: number]: count }`

### 3.5 FAO-Compliance

#### `computeSummary(cases, fachgebiet)`
- **Input:** Case-Array, Fachgebiet
- **Output:**
  ```typescript
  {
    gesamt: number,
    gerichtlich: number,
    aussergerichtlich: number,
    rechtsfoermlich: number,
    fg_count: number,
    bereich_counts: { [nr: number]: number },
    checks: Array<[kriterium: string, ist: number, soll: number, status: string]>,
    warnungen: string[]
  }
  ```
- **Check-Status:** `"erfuellt"` | `"knapp"` | `"nicht_erfuellt"`
- **Knapp-Schwelle:** ist >= soll * 0.8
- **Spezial-Logik pro Fachgebiet:**
  - **Erbrecht:** max 15 fG-Verfahren, 3+ Bereiche mit je 5+ Fällen
  - **Arbeitsrecht:** 5+ kollektive Arbeitsrecht-Fälle (Bereich 2)
  - **Familienrecht:** 30+ Verbund-Doppelzählungen, 10+ Gewaltschutz
  - **Handels-/Gesellschaftsrecht:** 40+ streitig, 10+ gestaltend

### 3.6 Plausibilitätsprüfung

#### `checkCasePlausibility(case)`
- **Input:** Einzelner Fall
- **Output:** `Array<{ typ: 'fehler'|'warnung'|'hinweis', feld: string, meldung: string }>`
- **Prüfungen:**
  1. Sachverhalt: < 4 Sätze → Warnung, 0 Sätze → Fehler, > 8 Sätze → Hinweis
  2. Pflichtfelder (kurzrubrum, verfahrenstyp, bereich_nr): leer → Fehler
  3. Kanzlei-AZ: leer → Warnung
  4. Zeitraum: beide leer → Hinweis
  5. Gerichts-AZ: leer bei gerichtlichem Verfahren → Warnung

#### `checkAllCasesPlausibility(cases)`
- **Input:** Case-Array
- **Output:** `{ fehler: number, warnungen: number, hinweise: number, details: Array }`

### 3.7 Duplikat-Prüfung

#### `checkDuplicates(newCases, existingCases)`
- **Input:** Neue Fälle, Bestehende Fälle
- **Output:** `{ nonDuplicates: Array, duplicates: Array }`
- **Logik:** Case-insensitive Vergleich von `kanzlei_az` (trimmed, lowercase)

### 3.8 Export-Funktionen

#### `createExcel(fl1, fl2, summary, fachgebiet, unprocessedFiles, unrecognizedTexts)`
- **Output:** Excel-Datei (Blob)
- **Worksheets:**
  1. **Fallliste_1** - Gerichtliche Verfahren (grüner Header #2e7d32)
  2. **Fallliste_2** - Außergerichtliche Verfahren (blauer Header #1565c0)
  3. **Summary** - FAO-Check + Statistiken (dunkelblauer Header #1e3a5f)
  4. **Nicht_verarbeitet** - Fehlgeschlagene Dateien (roter Header #c62828)
  5. **Nicht_erkannt** - Texte ohne Fälle (oranger Header #e65100)
  6. **Anleitung** - Spalten-Erklärungen
- **Styling:** AutoFilter, Spaltenbreiten angepasst, Zeilenfarben alternierend

#### `createPdfReport(fl1, fl2, summary, fachgebiet)`
- **Output:** PDF-Datei (Blob)
- **Inhalt:** Zusammenfassung, FAO-Check-Tabelle, FL1/FL2-Übersicht (max 50 Fälle)

#### `createAntragsschreiben(antragsteller, fachgebiet, summary, klausuren, bereichVerteilung, fl1Count, fl2Count)`
- **Output:** Word-Dokument (Blob)
- **Struktur:** Siehe Abschnitt 6.2

### 3.9 Session-Management

#### `saveSessionToJSON(casesDF, fachgebiet, unprocessedFiles, unrecognizedTexts, uploadHistory)`
- **Output:** JSON-String
- **Schema:**
  ```json
  {
    "version": "25.12.13-23:33",
    "timestamp": "2025-12-13T23:33:00",
    "fachgebiet": "Erbrecht",
    "cases": [...],
    "unprocessed_files": [...],
    "unrecognized_texts": [...],
    "upload_history": [...]
  }
  ```

#### `loadSessionFromJSON(jsonString)`
- **Output:** `{ success: boolean, casesDF, fachgebiet, ... } | { success: false, error: string }`

---

## 4. DATENMODELL

### 4.1 Case (Fall)
```typescript
interface Case {
  kanzlei_az: string;           // "2024/001"
  kurzrubrum: string;           // "A ./. B"
  sachverhalt: string;          // 4-6 Sätze, detailliert
  zeitraum_von: string;         // "01/2024"
  zeitraum_bis: string;         // "06/2024"
  gericht_az: string;           // "12 O 123/24"
  verfahrenstyp: 'gerichtlich' | 'rechtsfoermlich' | 'aussergerichtlich';
  verfahrensart: string;        // "streitig" / "fG" / "Mahnverfahren"
  bereich_nr: number;           // 1-5 (je nach Fachgebiet)
  bereich_bezeichnung: string;  // Text des Bereichs
  bedeutung: 'gering' | 'mittel' | 'hoch';
  taetigkeitsbeschreibung: string;
  stand: 'anhaengig' | 'abgeschlossen';
  abschluss_art: string;        // "Vergleich" / "Urteil" / ...
  abschluss_datum: string;      // "15.06.2024"
  verbundener_fall: string;     // Kanzlei-AZ des verbundenen Falls
  fachgebiet: string;           // "Erbrecht"
}
```

### 4.2 Antragsteller
```typescript
interface Antragsteller {
  name: string;              // "Rechtsanwalt Max Mustermann"
  kanzlei: string;           // "Mustermann & Partner"
  strasse: string;           // "Musterstraße 123"
  plz_ort: string;           // "12345 Berlin"
  telefon: string;           // "030 123456789"
  email: string;             // "ra@kanzlei.de"
  ort: string;               // "Berlin" (für Datumzeile)
  zulassung_datum: string;   // "01.01.2020"
  kammer_name: string;       // "Rechtsanwaltskammer Berlin"
  kammer_strasse: string;    // "Littenstraße 9"
  kammer_plz_ort: string;    // "10179 Berlin"
  lehrgang_info: string;     // "Fachanwaltslehrgang Erbrecht der DAA"
  lehrgang_datum: string;    // "15.10.2024"
}
```

### 4.3 Klausur
```typescript
interface Klausur {
  bezeichnung: string;  // "Klausur 1: Materielles Erbrecht"
  datum: string;        // "15.09.2024"
  ergebnis: 'bestanden' | 'gut bestanden' | 'sehr gut bestanden' | 'mit Auszeichnung bestanden';
}
```

### 4.4 Upload-Historie
```typescript
interface UploadHistoryEntry {
  datei: string;             // Dateiname
  faelle: number;            // Anzahl erkannter Fälle
  methode: 'text' | 'ocr';  // Extraktionsmethode
  quelle?: string;           // "Google Drive Ordner" / "Dropbox"
  zeitpunkt: string;         // "13.12.2025 23:33"
}
```

### 4.5 Unprocessed File
```typescript
interface UnprocessedFile {
  Dateiname: string;
  Grund: string;
  Typ: string;       // "Größenlimit" / "Verarbeitungsfehler" / "Download-Fehler"
  Zeitpunkt: string;
}
```

### 4.6 Unrecognized Text
```typescript
interface UnrecognizedText {
  Dateiname: string;
  Grund: string;
  Textauszug?: string;   // Erste 500 Zeichen
  Textlänge: number;
  Zeitpunkt: string;
}
```

---

## 5. FAO-KONFIGURATION (6 Fachgebiete)

### 5.1 Erbrecht
```json
{
  "paragraph": "§ 5 Abs. 1 lit. m, § 14f FAO",
  "gesamt_min": 80,
  "gerichtlich_min": 20,
  "aussergerichtlich_min": 60,
  "fg_max": 15,
  "bereiche": {
    "1": "Materielles Erbrecht und Bezüge zum Familien- und Gesellschaftsrecht",
    "2": "Testamentsvollstreckung und Nachlassverwaltung",
    "3": "Internationales Privatrecht und Erbschaftsteuerrecht",
    "4": "Erbprozessrecht",
    "5": "Vorweggenommene Erbfolge"
  },
  "bereiche_anforderung": {
    "min_bereiche": 3,
    "min_faelle_pro_bereich": 5
  }
}
```

### 5.2 Arbeitsrecht
```json
{
  "paragraph": "§ 5 Abs. 1 lit. c, § 10 FAO",
  "gesamt_min": 100,
  "gerichtlich_min": 50,
  "bereiche": {
    "1": "Individualarbeitsrecht",
    "2": "Kollektives Arbeitsrecht",
    "3": "Arbeitsgerichtliches Verfahren",
    "4": "Sozialrecht im Arbeitsrecht"
  },
  "spezial": "mind. 5 Fälle kollektives Arbeitsrecht (Bereich 2)"
}
```

### 5.3 Miet- und Wohnungseigentumsrecht
```json
{
  "paragraph": "§ 5 Abs. 1 lit. j, § 14c FAO",
  "gesamt_min": 120,
  "gerichtlich_min": 60,
  "bereiche": {
    "1": "Wohnraummiete",
    "2": "Gewerbemiete und Pacht",
    "3": "Wohnungseigentum"
  },
  "bereiche_anforderung": { "min_bereiche": 3, "min_faelle_pro_bereich": 5 }
}
```

### 5.4 Familienrecht
```json
{
  "paragraph": "§ 5 Abs. 1 lit. f, § 12 FAO",
  "gesamt_min": 120,
  "gerichtlich_min": 60,
  "bereiche": {
    "1": "Eherecht und Lebenspartnerschaftsrecht",
    "2": "Kindschaftsrecht",
    "3": "Unterhaltsrecht",
    "4": "Güterrecht und Vermögensauseinandersetzung",
    "5": "Versorgungsausgleich"
  },
  "spezial": "30+ Verbund-Doppelzählungen, 10+ Gewaltschutz"
}
```

### 5.5 Verkehrsrecht
```json
{
  "paragraph": "§ 5 Abs. 1 lit. n, § 14g FAO",
  "gesamt_min": 160,
  "gerichtlich_min": 60,
  "bereiche": {
    "1": "Verkehrszivilrecht (Schadensregulierung)",
    "2": "Verkehrsstraf- und OWi-Recht",
    "3": "Verkehrsverwaltungsrecht (Fahrerlaubnisrecht)",
    "4": "Versicherungsrecht im Verkehr"
  },
  "bereiche_anforderung": { "min_bereiche": 3, "min_faelle_pro_bereich": 5 }
}
```

### 5.6 Handels- und Gesellschaftsrecht
```json
{
  "paragraph": "§ 5 Abs. 1 lit. h, § 14a FAO",
  "gesamt_min": 100,
  "gerichtlich_min": 30,
  "bereiche": {
    "1": "Recht der Personen- und Kapitalgesellschaften",
    "2": "Handelsrecht und Handelsvertreterrecht",
    "3": "Konzern- und Umwandlungsrecht",
    "4": "Gesellschaftsrechtliche Streitigkeiten"
  },
  "spezial": "40+ streitig, 10+ gestaltend"
}
```

---

## 6. UI-STRUKTUR UND SEITENAUFBAU

### 6.1 Seitenlayout
```
┌─────────────────────────────────────────────────────────────────┐
│ ⚖️ Fachanwalt-Falllistenverwaltung              [Version X.X]  │
│ Erstellen Sie FAO-konforme Falllisten...                       │
├─────────────┬───────────────────────────────────────────────────┤
│  SIDEBAR    │  HAUPTBEREICH                                    │
│             │                                                   │
│  Fachgebiet │  "So funktioniert's" (4 Schritte)               │
│  [Dropdown] │  ─────────────────────────────────                │
│             │  TABS:                                            │
│  API-Key    │  ┌──────────┬──────────┬──────────┐              │
│  [Password] │  │📕 PDF    │☁️ Cloud  │✏️ Manuell│              │
│             │  │(empfohl.)│(GDrive/  │          │              │
│  📊 Fälle:  │  │          │Dropbox)  │          │              │
│  [42]       │  └──────────┴──────────┴──────────┘              │
│             │                                                   │
│  📄 Zuletzt:│  [Upload-/Eingabe-Bereich je Tab]               │
│  akte_01.pdf│  ─────────────────────────────────                │
│             │  📊 ÜBERSICHT (Metriken)                         │
│  GPT-Modell │  ─────────────────────────────────                │
│  [Dropdown] │  ✅ FAO-KONFORMITÄTSCHECK                       │
│             │  ─────────────────────────────────                │
│  FAO-Info   │  📈 BEREICHSVERTEILUNG (Chart)                  │
│  [Expander] │  ─────────────────────────────────                │
│             │  📋 FALLLISTEN-VORSCHAU (FL1/FL2 Tabs)          │
│  📋 Muster  │  ─────────────────────────────────                │
│  [Download] │  🔍 PLAUSIBILITÄTSPRÜFUNG                       │
│             │  ─────────────────────────────────                │
│  💾 Session │  ✏️ FÄLLE BEARBEITEN [Expander]                 │
│  [Save/Load]│  ─────────────────────────────────                │
│             │  📥 EXPORT (Excel + PDF)                         │
│  📜 Historie│  ─────────────────────────────────                │
│  [Expander] │  📝 ANTRAGSSCHREIBEN                             │
│             │     [Daten-Eingabe + Word-Download]               │
└─────────────┴───────────────────────────────────────────────────┘
```

### 6.2 Antragsschreiben-Struktur (Word)
```
┌─────────────────────────────────────────────────┐
│                    [Absender rechtsbündig]       │
│                    Name, Kanzlei, Adresse        │
│                                                  │
│  [Empfänger fett]                                │
│  Rechtsanwaltskammer ...                         │
│                                                  │
│                    [Ort], den [Datum]             │
│                                                  │
│  BETREFF: Antrag auf Verleihung der              │
│  Bezeichnung "Fachanwalt für [Fachgebiet]"       │
│                                                  │
│  Sehr geehrte Damen und Herren,                  │
│                                                  │
│  hiermit beantrage ich ... gemäß [§ FAO] ...     │
│  Zugelassen seit [Datum] bei [Kammer].           │
│                                                  │
│  I. Theoretische Kenntnisse                      │
│  Lehrgang "[Name]" am [Datum] abgeschlossen.     │
│  ┌─────────────┬──────────┬──────────┐           │
│  │ Klausur     │ Datum    │ Ergebnis │           │
│  ├─────────────┼──────────┼──────────┤           │
│  │ Klausur 1   │ 15.09.24 │ bestanden│           │
│  │ Klausur 2   │ 16.09.24 │ bestanden│           │
│  └─────────────┴──────────┴──────────┘           │
│                                                  │
│  II. Praktische Erfahrungen                      │
│  [X] Fälle bearbeitet.                           │
│  • FL1: [X] gerichtliche Verfahren               │
│  • FL2: [X] außergerichtliche Verfahren           │
│  Anforderungen [§] erfüllt.                      │
│  ┌─────┬──────────────────────┬───────┐          │
│  │ Nr. │ Bereich              │Anzahl │          │
│  ├─────┼──────────────────────┼───────┤          │
│  │  1  │ Materielles Erbrecht │  25   │          │
│  │  2  │ Testamentsvollstr.   │  18   │          │
│  └─────┴──────────────────────┴───────┘          │
│                                                  │
│  III. Anlagen                                    │
│  1. Fallliste 1 (Excel)                          │
│  2. Fallliste 2 (Excel)                          │
│  3. Übersicht FAO-Check                          │
│  4. Lehrgangszeugnis                             │
│  5. Klausurzeugnisse                             │
│  6. Berufshaftpflicht                            │
│  7. Erklärung § 7 FAO                            │
│                                                  │
│  Mit freundlichen kollegialen Grüßen             │
│                                                  │
│  [Name]                                          │
│  Rechtsanwalt/Rechtsanwältin                     │
└──────────────────────────────────────────────────┘
```

---

## 7. ALLE VERKNÜPFUNGEN (Datenfluss)

### 7.1 Haupt-Datenfluss
```
                    ┌──────────────┐
                    │   PDF-Upload │
                    └──────┬───────┘
                           │
            ┌──────────────┼──────────────┐
            ▼              ▼              ▼
    ┌───────────┐  ┌─────────────┐  ┌──────────┐
    │PDF Upload │  │Cloud Upload │  │ Manuell  │
    │(lokal)    │  │(GDrive/DB) │  │(Freitext/│
    │           │  │            │  │Struktur) │
    └─────┬─────┘  └──────┬─────┘  └────┬─────┘
          │               │              │
          ▼               ▼              ▼
    ┌─────────────────────────────────────────┐
    │          extractTextFromPdf()            │
    │          (Text oder OCR)                 │
    └──────────────────┬──────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────┐
    │       analyzeCasesWithGPT()             │
    │       (OpenAI API → JSON Array)         │
    └──────────────────┬──────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────┐
    │    casesListToDataFrame()               │
    │    normalizeCaseDF()                    │
    └──────────────────┬──────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────┐
    │       checkDuplicates()                 │
    │    ┌──────────────────────────┐         │
    │    │  Duplikat? → pending     │         │
    │    │  Neu?      → cases_df    │         │
    │    └──────────────────────────┘         │
    └──────────────────┬──────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────┐
    │          SESSION STATE                  │
    │  ┌──────────────────────────────────┐   │
    │  │ cases_df (alle Fälle)           │   │
    │  │ upload_history                   │   │
    │  │ unprocessed_files               │   │
    │  │ unrecognized_texts              │   │
    │  │ antragsteller / klausuren       │   │
    │  └──────────────────────────────────┘   │
    └──────┬─────────┬──────────┬─────────────┘
           │         │          │
           ▼         ▼          ▼
    ┌──────────┐ ┌────────┐ ┌───────────┐
    │compute   │ │split   │ │checkAll   │
    │Summary() │ │Into    │ │Cases      │
    │          │ │Fall    │ │Plausi     │
    │          │ │listen()│ │bility()   │
    └────┬─────┘ └───┬────┘ └─────┬─────┘
         │           │            │
         ▼           ▼            ▼
    ┌─────────────────────────────────────────┐
    │              ANZEIGE                    │
    │  • Dashboard (Metriken)                 │
    │  • FAO-Check (Ampel)                    │
    │  • Bereichsverteilung (Chart)           │
    │  • Vorschau FL1/FL2 (Tabellen)         │
    │  • Plausibilitätsprüfung               │
    │  • Fall-Editor                          │
    └────────────────┬────────────────────────┘
                     │
                     ▼
    ┌─────────────────────────────────────────┐
    │              EXPORT                     │
    │  ┌───────────┐ ┌────────┐ ┌─────────┐  │
    │  │createExcel│ │create  │ │create   │  │
    │  │()        │ │Pdf     │ │Antrags  │  │
    │  │          │ │Report()│ │schreiben│  │
    │  └─────┬─────┘ └───┬────┘ └────┬────┘  │
    │        ▼           ▼           ▼       │
    │    .xlsx        .pdf        .docx      │
    └─────────────────────────────────────────┘
```

### 7.2 Session-Save/Load-Zyklus
```
    cases_df + history + files
                │
                ▼
        saveSessionToJSON()
                │
                ▼
        fallliste_session_*.json
                │
        (Download / Später Upload)
                │
                ▼
        loadSessionFromJSON()
                │
                ▼
    cases_df + history + files
        (Session wiederhergestellt)
```

---

## 8. REACT-KOMPONENTENSTRUKTUR (Empfehlung)

```
src/
├── components/
│   ├── Layout/
│   │   ├── AppHeader.tsx             # Titel + Version
│   │   ├── Sidebar.tsx               # Einstellungen, Session, Historie
│   │   └── MainContent.tsx           # Hauptbereich-Container
│   ├── Input/
│   │   ├── PdfUploadTab.tsx          # PDF hochladen + Analyse
│   │   ├── CloudUploadTab.tsx        # Google Drive / Dropbox
│   │   ├── ManualInputTab.tsx        # Freitext + Strukturiert
│   │   └── InputTabs.tsx             # Tab-Container
│   ├── Analysis/
│   │   ├── Dashboard.tsx             # 4 KPI-Metriken
│   │   ├── FaoComplianceCheck.tsx    # FAO-Prüfung mit Ampeln
│   │   ├── BereichChart.tsx          # Balkendiagramm
│   │   └── CasePreview.tsx           # FL1/FL2 Tabellen
│   ├── Management/
│   │   ├── PlausibilityCheck.tsx     # Plausibilitätsprüfung
│   │   ├── CaseEditor.tsx            # Fall bearbeiten/löschen
│   │   └── DuplicateHandler.tsx      # Duplikat-Entscheidung
│   ├── Export/
│   │   ├── ExcelExport.tsx           # Excel-Download
│   │   ├── PdfExport.tsx             # PDF-Download
│   │   └── ApplicationLetter.tsx     # Antragsschreiben
│   ├── Application/
│   │   ├── ApplicantForm.tsx         # Antragsteller-Daten
│   │   ├── ExamManager.tsx           # Klausuren verwalten
│   │   └── ChamberForm.tsx           # Kammer-Daten
│   └── Session/
│       ├── SessionSave.tsx           # JSON-Export
│       ├── SessionLoad.tsx           # JSON-Import
│       └── UploadHistory.tsx         # Historie-Anzeige
├── hooks/
│   ├── useCases.ts                   # Case CRUD-Operationen
│   ├── useFaoConfig.ts              # FAO-Konfiguration
│   ├── useOpenAI.ts                 # GPT-Integration
│   ├── useCloudUpload.ts            # Cloud-Download
│   ├── usePdfExtraction.ts          # PDF-Text-Extraktion
│   └── useSession.ts               # Session-Management
├── services/
│   ├── openai.service.ts            # OpenAI API Calls
│   ├── googleDrive.service.ts       # Google Drive API
│   ├── dropbox.service.ts           # Dropbox API
│   ├── pdfExtractor.service.ts      # PDF + OCR
│   ├── excelGenerator.service.ts    # Excel-Erstellung
│   ├── wordGenerator.service.ts     # Word-Erstellung
│   └── pdfReportGenerator.service.ts # PDF-Report
├── store/
│   ├── casesSlice.ts                # Cases State
│   ├── applicantSlice.ts            # Antragsteller State
│   ├── sessionSlice.ts              # Session State
│   └── store.ts                     # Redux/Zustand Store
├── config/
│   ├── faoConfig.ts                 # FAO-Mindestanforderungen
│   ├── columnMapping.ts             # Spalten-Mapping
│   └── constants.ts                 # MAX_FILE_SIZE etc.
├── utils/
│   ├── plausibility.ts              # Plausibilitätsprüfung
│   ├── duplicateCheck.ts            # Duplikat-Erkennung
│   ├── caseNormalizer.ts            # Daten-Normalisierung
│   ├── faoCompliance.ts             # FAO-Check-Logik
│   └── formatters.ts               # Zeitraum-Formatierung etc.
└── types/
    ├── case.types.ts                # Case, Antragsteller, Klausur
    ├── fao.types.ts                 # FAO-Config Typen
    └── session.types.ts             # Session-Typen
```

---

## 9. API-ENDPUNKTE (Falls Backend geplant)

| Method | Endpoint | Beschreibung |
|--------|----------|--------------|
| POST | `/api/analyze` | PDF-Text an GPT senden |
| POST | `/api/upload` | PDF hochladen |
| GET | `/api/cloud/google/{id}` | Google Drive Datei laden |
| GET | `/api/cloud/google/folder/{id}` | Google Drive Ordner listen |
| GET | `/api/cloud/dropbox` | Dropbox Datei laden |
| POST | `/api/export/excel` | Excel generieren |
| POST | `/api/export/pdf` | PDF-Report generieren |
| POST | `/api/export/docx` | Antragsschreiben generieren |
| POST | `/api/session/save` | Session speichern |
| POST | `/api/session/load` | Session laden |
| POST | `/api/validate` | Plausibilitätsprüfung |
| POST | `/api/fao/check` | FAO-Compliance prüfen |

---

## 10. KONSTANTEN

```typescript
const APP_VERSION = "25.12.13-23:33";
const MAX_FILE_SIZE_MB = 100;
const MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024;
const MAX_TOTAL_UPLOAD_MB = 500;
const MAX_TOTAL_UPLOAD_BYTES = 500 * 1024 * 1024;
const CLOUD_CHUNK_SIZE = 5 * 1024 * 1024;
const GPT_MAX_TEXT_LENGTH = 12000;
const GPT_RETRY_ATTEMPTS = 3;
const GPT_RETRY_DELAYS = [2000, 5000, 10000]; // ms
const GPT_MODELS = ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-3.5-turbo"];
```

---

## 11. ABHÄNGIGKEITEN (npm-Äquivalente)

| Python-Paket | npm-Paket | Verwendung |
|---|---|---|
| streamlit | react + react-dom | UI-Framework |
| pandas | - (native Arrays) | Datenverarbeitung |
| openpyxl | xlsx (SheetJS) | Excel-Export |
| openai | openai | GPT-Integration |
| PyPDF2 | pdfjs-dist | PDF-Textextraktion |
| pytesseract + pdf2image | tesseract.js | OCR |
| reportlab | jspdf + jspdf-autotable | PDF-Report |
| python-docx | docx (docx.js) | Word-Export |
| requests | fetch / axios | HTTP-Requests |
| - | recharts / chart.js | Diagramme |
| - | @tanstack/react-table | Datentabellen |
| - | react-dropzone | Datei-Upload |
| - | zustand / redux-toolkit | State Management |
| - | tailwindcss / chakra-ui | Styling |
