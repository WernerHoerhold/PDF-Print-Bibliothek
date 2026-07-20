# PDF-Print-Bibliothek - Dokumentation

## Projektübersicht

**Projektname:** PDF-Print-Bibliothek  
**Version:** 1.1  
**Beschreibung:** ActiveX-DLL für VB6 zur Erstellung und zum Druck von einfachen PDF-Dokumenten (Messschriebe)  
**Hersteller:** Kapp Niles GmbH & Co, KG  
**Copyright:** 2026  

### Projektzweck
Die PDF-Print-Bibliothek bietet VB6-Anwendungen die Möglichkeit, PDF-Dokumente programmgesteuert zu erstellen und diese über Adobe Reader oder einen Drucker auszudrucken.

---

## Systemanforderungen

- **Windows-Version:** XP oder höher
- **VB6 Runtime:** Erforderlich
- **Datenbankverbindung:** Keine
- **Externe Libraries/Referenzen:**
  - OLE Automation (stdole2.tlb)
  - Adobe Reader (optional, zum Drucken erforderlich)
- **Benötigte DLLs:** PdfLib.dll (wird mit dem Projekt bereitgestellt)

---

## Klassifizierung der 41 Klassen

### KATEGORIE A - HAUPTKLASSEN (Umfassend dokumentiert)
Diese Klassen bilden die Kern-API und werden von Anwendern direkt verwendet:

| Klasse | Aufgabe | Sichtbarkeit |
|--------|---------|-------------|
| **PdfDocument** | Verwaltung des PDF-Dokuments, Seiten, Schriften | Public |
| **PdfPage** | Repräsentation einer einzelnen PDF-Seite | Public |
| **PdfCanvas2D** | 2D Zeichenfläche für Grafiken und Text | Public |
| **PdfWriter** | Kern-Engine zum Schreiben von PDF-Dateien | Internal |
| **PdfContentBuilder** | Erstellung des PDF-Content-Streams | Internal |
| **PdfFile** | Datei-Handling für PDF-Speicherung | Internal |
| **PdfPrinter** | Drucker-Integration und Verwaltung | Public |
| **Pdf** | Einstiegspunkt der Bibliothek | Public |

### KATEGORIE B - UNTERSTÜTZUNGSKLASSEN (Moderate Dokumentation)
Wichtige Hilfsfunktionen, mehrfach verwendet:

| Klasse | Aufgabe |
|--------|---------|
| PdfBuffer | Puffer für kompletten PDF-Inhalt |
| PdfSerializer | Serialisierung/Deserialisierung für PDF-Strukturen |
| PdfObjectIndex | Verwaltung aller PDF-Objekte |
| PdfGraphicsState | Aktueller Zeichenzustand (Stift, Pinsel, Text) |
| PdfTextState | Textzustand (Schrift, Größe, Farbe) |
| PdfFormat | Format-Handling und Zahlenformatierung |
| PdfPageObjects | Verwaltung von Seiten-Objekten |
| PdfImage | Bilder in PDF-Dokumenten |
| PdfBinaryFile | Binär-Datei-Handling |

### KATEGORIE C - INTERNE/UTILITY-KLASSEN (Minimale Dokumentation)
Spezifische Helper-Funktionen, interne Nutzung:

**Zeichenbefehle (7 Klassen):**
- PdfCmdLine - Linienzeichnung
- PdfCmdRectangle - Rechtecke
- PdfCmdEllipse - Ellipsen
- PdfCmdPolygon - Polygone
- PdfCmdPolyline - Polylinien
- PdfCmdText - Textausgabe
- PdfCmdImage - Bildausgabe

**Stil & Farbe (5 Klassen):**
- PdfPen - Zeichenstift-Eigenschaften
- PdfBrush - Füllpinsel-Eigenschaften
- PdfColor - Farb-Umrechnung (VB RGB ↔ PDF RGB)
- PdfLineCaps - Linienenden-Stile
- PdfLineJoins - Linienverbindungs-Stile

**Geometrie (3 Klassen):**
- PdfGeometry - Geometrische Berechnungen
- PdfPoint - Punkt-Darstellung
- PdfPointCollection - Sammlung von Punkten

**Konfiguration (5 Klassen):**
- PdfPageSizes - Vordefinierte Seitengröße (A4, A3, Letter, etc.)
- PdfOrientations - Seitenorientierung (Portrait/Landscape)
- PdfPrintOptions - Druckoptionen
- PdfPrinterInfo - Drucker-Informationen
- PdfIniFile - INI-Datei-Handling

**PDF-Intern (4 Klassen):**
- PdfObject - Basis-Klasse für PDF-Objekte
- PdfObjectNumbers - Verwaltung von Objektnummern
- PdfBinaryFile - Binärdaten-Handling
- PdfVersion - Versioning-Informationen

**Module (3 Module):**
- PdfEnums - Enumerationen (PdfObjectType, PdfPageSize, etc.)
- PdfGlobals - Globale Variablen und Singleton-Pattern
- PdfVersion - Versionsinformationen

---

## Projektstruktur

```
PDF-Print-Bibliothek/
├── PdfCanvas/
│   ├── Pdflib.vbp                    # VB6 Projektdatei
│   ├── Pdflib.vbw                    # VB6 Workspace
│   ├── PdfLib.dll                    # Kompilierte DLL
│   │
│   ├── KATEGORIE A - Hauptklassen:
│   ├── Pdf.cls                       # Einstiegspunkt
│   ├── PdfDocument.cls               # Hauptklasse
│   ├── PdfPage.cls                   # Seiten-Verwaltung
│   ├── PdfCanvas2D.cls               # Zeichenfläche
│   ├── PdfWriter.cls                 # PDF-Schreib-Engine
│   ├── PdfPrinter.cls                # Druckerverwaltung
│   │
│   ├── KATEGORIE B - Unterstützungsklassen:
│   ├── PdfBuffer.cls                 # Puffer
│   ├── PdfSerializer.cls             # Serialisierung
│   ├── PdfObjectIndex.cls            # Objektverwaltung
│   ├── PdfGraphicsState.cls          # Zeichenzustand
│   ├── PdfContentBuilder.cls         # Content-Erstellung
│   │
│   └── [weitere 26 Klassen...]       # Kategorie C
│
└── Printer.ini                        # Drucker-Konfiguration
```

---

## Verwendungsbeispiel

```vb
' Einstieg über Pdf-Klasse
Dim Pdf As New Pdf
Dim Doc As PdfDocument

' Neues Dokument erstellen
Set Doc = Pdf.CreateDocument()

' Seite hinzufügen
Dim Page As PdfPage
Set Page = Doc.AddPage()

' Text zeichnen
Page.Canvas.DrawText 100, 100, "Hallo PDF!"

' Rechteck zeichnen
Page.Canvas.DrawRectangle 50, 50, 200, 100

' PDF speichern
Doc.Save "C:\Output\Test.pdf"
```

---

## Datenbankschema

Diese Bibliothek verwendet keine Datenbank.

---

## Fehlerbehandlung

Fehler werden über die Properties `LastErrorNumber` und `LastErrorText` der `PdfDocument`-Klasse gemeldet:

```vb
If Not Doc.Save("C:\Output\test.pdf") Then
    MsgBox "Fehler: " & Doc.LastErrorText
End If
```

---

## Module & Klassen Übersicht

### Module (3)
1. **PdfEnums.bas** - Enumerationen (PdfObjectType, PdfPageSize, PdfOrientation)
2. **PdfGlobals.bas** - Globale Singleton-Instanzen (PageSizes, Orientations)
3. **PdfVersion.bas** - Versioning-Konstanten

### Hauptklassen (8 - KATEGORIE A)
- Pdf
- PdfDocument
- PdfPage
- PdfCanvas2D
- PdfWriter
- PdfContentBuilder
- PdfFile
- PdfPrinter

### Unterstützungsklassen (9 - KATEGORIE B)
- PdfBuffer
- PdfSerializer
- PdfObjectIndex
- PdfGraphicsState
- PdfTextState
- PdfFormat
- PdfPageObjects
- PdfImage
- PdfBinaryFile

### Utility-Klassen (24 - KATEGORIE C)
- Zeichenbefehle: PdfCmdLine, PdfCmdRectangle, PdfCmdEllipse, PdfCmdPolygon, PdfCmdPolyline, PdfCmdText, PdfCmdImage
- Stil: PdfPen, PdfBrush, PdfColor, PdfLineCaps, PdfLineJoins
- Geometrie: PdfGeometry, PdfPoint, PdfPointCollection
- Konfiguration: PdfPageSizes, PdfOrientations, PdfPrintOptions, PdfPrinterInfo, PdfIniFile
- Intern: PdfObject, PdfObjectNumbers

---

## Installation & Setup

1. PdfLib.dll registrieren (als Administrator):
   ```
   regsvr32 PdfLib.dll
   ```

2. In VB6 hinzufügen:
   - Projekt > Referenzen
   - Browse zu PdfLib.dll
   - OK

3. Verwendung:
   ```vb
   Dim Pdf As New Pdf
   Dim Doc As PdfDocument
   Set Doc = Pdf.CreateDocument()
   ```

---

## Changelog

### Version 1.1 (2026-07-20)
- Erste öffentliche Release
- Dokumentation erstellt
- 41 Klassen + 3 Module

---

## Support & Kontakt

Bei Fragen oder Problemen: Dokumentation beachten oder Support kontaktieren.

---

**Dokumentation erstellt: 20.07.2026**  
**Dokumentations-Version: 1.0**
