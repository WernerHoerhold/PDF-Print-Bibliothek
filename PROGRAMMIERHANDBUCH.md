# PDF-Print-Bibliothek - PROGRAMMIERHANDBUCH

**Autor:** Kapp Niles GmbH & Co, KG  
**Version:** 1.1  
**Datum:** 20.07.2026  
**Status:** Fertiggestellt

---

## INHALTSVERZEICHNIS

1. [Einführung](#einführung)
2. [Systemanforderungen](#systemanforderungen)
3. [Installation & Setup](#installation--setup)
4. [Schnellstart](#schnellstart)
5. [Hauptklassen & API](#hauptklassen--api)
6. [Verwendungsbeispiele](#verwendungsbeispiele)
7. [Klassifizierung & Architektur](#klassifizierung--architektur)
8. [Fehlerbehandlung](#fehlerbehandlung)
9. [FAQ & Häufige Probleme](#faq--häufige-probleme)
10. [Changelog](#changelog)

---

## 1. Einführung

### Was ist die PDF-Print-Bibliothek?

Die **PDF-Print-Bibliothek** ist eine ActiveX-DLL für Visual Basic 6.0, die es ermöglicht, PDF-Dokumente programmgesteuert zu erstellen und zu drucken. Sie wurde von Kapp Niles GmbH & Co, KG entwickelt und richtet sich an VB6-Entwickler, die einfache PDF-Dokumente (z.B. Messschriebe) erzeugen müssen.

### Hauptmerkmale

✅ **PDF-Erstellung:** Programmgesteuert neue PDFs erstellen  
✅ **Zeichnungen:** Linien, Rechtecke, Ellipsen, Polygone zeichnen  
✅ **Text:** Text mit verschiedenen Schriftarten und Größen  
✅ **Bilder:** JPEG-Bilder einbinden  
✅ **Druck:** Über Adobe Reader drucken  
✅ **Einfach zu nutzen:** Intuitive API

---

## 2. Systemanforderungen

### Hardware
- **CPU:** Pentium II oder höher
- **RAM:** Mindestens 128 MB
- **Festplatte:** 5 MB für DLL + Output

### Software
- **Windows:** XP, Vista, 7, 8, 10, 11
- **VB6:** Runtime installiert
- **Adobe Reader:** Optional (zum Drucken erforderlich)

### Abhängigkeiten
- `stdole2.tlb` - OLE Automation (Windows-Standard)
- `PdfLib.dll` - Im Projektordner enthalten

---

## 3. Installation & Setup

### Schritt 1: DLL registrieren

```bash
# Als Administrator öffnen und ausführen:
regsvr32 PdfLib.dll

# Erfolgsmeldung:
# "DllRegisterServer in PdfLib.dll erfolgreich ausgeführt"
```

### Schritt 2: In VB6 Projekt einbinden

1. **Projekt** → **Referenzen...**
2. **Browse** klicken
3. Navigate zu `PdfLib.dll`
4. **Öffnen** klicken
5. **OK**

### Schritt 3: Test

```vb
Private Sub Form_Load()
    Dim Pdf As New Pdf
    MsgBox Pdf.About
End Sub
```

---

## 4. Schnellstart

### Minimales Beispiel: "Hello PDF"

```vb
Public Sub CreateSimplePDF()
    ' Objekte deklarieren
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page As PdfPage
    
    ' Neues Dokument
    Set Doc = Pdf.CreateDocument()
    
    ' Seite hinzufügen
    Set Page = Doc.AddPage()
    
    ' Text zeichnen
    Page.Canvas.DrawText 100, 100, "Hallo PDF-Welt!"
    
    ' Speichern
    If Doc.Save("C:\Output\test.pdf") Then
        MsgBox "PDF erfolgreich erstellt!"
    Else
        MsgBox "Fehler: " & Doc.LastErrorText
    End If
End Sub
```

### Etwas komplexer: Mit Shapes & Bildern

```vb
Public Sub CreateAdvancedPDF()
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page As PdfPage
    Dim Font As PdfFont
    
    Set Doc = Pdf.CreateDocument()
    Set Page = Doc.AddPage()
    
    ' Schrift konfigurieren
    Set Font = Doc.AddFont("Helvetica")
    Font.Size = 14
    Page.Canvas.Font = Font
    
    ' Rechteck zeichnen
    Page.Canvas.DrawRectangle 50, 50, 200, 100
    
    ' Text zeichnen
    Page.Canvas.DrawText 60, 80, "Rechteck mit Text"
    
    ' Bild laden und zeichnen
    Dim Img As PdfImage
    Set Img = Pdf.LoadImage("C:\Images\logo.jpg")
    Doc.AddImage Img
    Page.Canvas.DrawImage Img, 50, 200
    
    ' Speichern
    Doc.Save "C:\Output\advanced.pdf"
End Sub
```

---

## 5. Hauptklassen & API

### 5.1 Klasse: Pdf (Einstiegspunkt)

**Aufgabe:** Factory-Methoden zur Objekterstellung

```vb
Dim Pdf As New Pdf

' Neue PDF erstellen
Dim Doc As PdfDocument
Set Doc = Pdf.CreateDocument()

' Bild laden
Dim Img As PdfImage
Set Img = Pdf.LoadImage("C:\Images\test.jpg")

' Punkt-Sammlung erstellen
Dim Points As PdfPointCollection
Set Points = Pdf.CreatePointCollection()

' Versionsinformation
Debug.Print Pdf.Version        ' "1.1"
Debug.Print Pdf.Build          ' Build-Nummer
Debug.Print Pdf.Name           ' "PDFLib"
Debug.Print Pdf.Copyright      ' Copyright-Text
Debug.Print Pdf.About          ' Vollständige Info
```

**Methoden:**
| Methode | Rückgabe | Beschreibung |
|---------|----------|-------------|
| `CreateDocument()` | PdfDocument | Neues leeres PDF-Dokument |
| `LoadImage(FileName)` | PdfImage | Lädt JPEG-Bild aus Datei |
| `CreatePointCollection()` | PdfPointCollection | Neue Punkt-Sammlung |

---

### 5.2 Klasse: PdfDocument (Dokumentverwaltung)

**Aufgabe:** Verwaltung des gesamten PDF-Dokuments

```vb
Dim Doc As New PdfDocument

' Seite hinzufügen
Dim Page As PdfPage
Set Page = Doc.AddPage()

' Auf Seiten zugreifen
Dim FirstPage As PdfPage
Set FirstPage = Doc.Page(1)

' Schrift registrieren
Dim Font As PdfFont
Set Font = Doc.AddFont("Arial")

' Bild registrieren
Dim Img As PdfImage
Doc.AddImage Img

' Speichern
If Not Doc.Save("C:\Output\test.pdf") Then
    MsgBox "Fehler: " & Doc.LastErrorText
End If

' Properties
Debug.Print Doc.PageCount      ' Anzahl Seiten
Debug.Print Doc.FontCount      ' Anzahl Schriften
Debug.Print Doc.ImageCount     ' Anzahl Bilder
```

**Eigenschaften:**
| Property | Typ | Beschreibung |
|----------|-----|-------------|
| `PageCount` | Long | Anzahl der Seiten (READ-ONLY) |
| `Page(Index)` | PdfPage | Zugriff auf Seite mit Index |
| `FontCount` | Long | Anzahl registrierter Schriften |
| `Font(Index)` | PdfFont | Zugriff auf Schrift |
| `ImageCount` | Long | Anzahl registrierter Bilder |
| `Image(Index)` | PdfImage | Zugriff auf Bild |
| `DefaultFont` | PdfFont | Standard-Schrift (Helvetica 12pt) |
| `LastErrorNumber` | Long | Fehlernummer des letzten Fehlers |
| `LastErrorText` | String | Fehlerbeschreibung |

---

### 5.3 Klasse: PdfPage (Seite)

**Aufgabe:** Einzelne PDF-Seite mit Zeichenfläche

```vb
Dim Page As PdfPage
Set Page = Doc.AddPage()

' Seitengröße und Orientierung
Page.PageSize = psA4              ' A4, A3, A5, Letter, Legal
Page.Orientation = poPortrait     ' Portrait oder Landscape

' Abmessungen (in PostScript-Punkten)
Debug.Print Page.Width            ' 595 für A4
Debug.Print Page.Height           ' 842 für A4

' Canvas (Zeichnenfläche)
Dim Canvas As PdfCanvas2D
Set Canvas = Page.Canvas

' Zeichnen
Canvas.DrawLine 100, 100, 200, 200
Canvas.DrawText 100, 300, "Hallo!"
```

**Eigenschaften:**
| Property | Typ | Standard | Beschreibung |
|----------|-----|----------|-------------|
| `Width` | Long | 595 | Seitenbreite in Punkten |
| `Height` | Long | 842 | Seitenh öhe in Punkten |
| `PageSize` | Long | psA4 | Seitengröße (A4, Letter, etc.) |
| `Orientation` | Long | poPortrait | Portrait/Landscape |
| `Canvas` | PdfCanvas2D | - | Zeichenfläche |

**Seitengröße-Enumerationen (PageSize):**
```
psA5 (420 x 595)
psA4 (595 x 842)        ← Standard
psA3 (842 x 1191)
psLetter (612 x 792)
psLegal (612 x 1008)
```

---

### 5.4 Klasse: PdfCanvas2D (Zeichenfläche)

**Aufgabe:** 2D-Zeichenbefehle für die Seite

```vb
Dim Canvas As PdfCanvas2D
Set Canvas = Page.Canvas

' Linien
Canvas.DrawLine 50, 50, 200, 200

' Rechtecke
Canvas.DrawRectangle 50, 50, 100, 100      ' Umriss
Canvas.FillRectangle 50, 50, 100, 100      ' Gefüllt

' Ellipsen
Canvas.DrawEllipse 50, 50, 100, 100        ' Umriss
Canvas.FillEllipse 50, 50, 100, 100        ' Gefüllt

' Text
Page.Canvas.Font.Size = 14
Canvas.DrawText 100, 300, "Beispieltext"

' Stift-Eigenschaften
Canvas.Pen.Width = 2                        ' Liniendicke
Canvas.Pen.Color = vbBlue                   ' Farbe
Canvas.Pen.Solid()                          ' Vollständige Linie

' Füllpinsel
Canvas.Brush.Color = vbYellow               ' Füllfarbe

' Schrift
Canvas.Font.Size = 12                       ' Größe
```

**Zeichnungs-Methoden:**
```vb
' Linien
DrawLine(X1, Y1, X2, Y2)

' Rechtecke
DrawRectangle(X, Y, Width, Height)      ' Umriss
FillRectangle(X, Y, Width, Height)      ' Gefüllt

' Ellipsen
DrawEllipse(X, Y, Width, Height)        ' Umriss
FillEllipse(X, Y, Width, Height)        ' Gefüllt

' Polygone
DrawPolygon(Points)                      ' Umriss
FillPolygon(Points)                      ' Gefüllt

' Text
DrawText(X, Y, Text)

' Bilder
DrawImage(Image, X, Y)                   ' Originalgröße
DrawImageEx(Image, X, Y, W, H)          ' Skaliert
DrawImageFit(Image, X, Y, MaxW, MaxH)   ' Proportional angepasst
```

---

### 5.5 Klasse: PdfFont (Schriftart)

**Aufgabe:** Schriftart-Verwaltung

```vb
Dim Font As PdfFont
Set Font = Doc.AddFont("Helvetica")

' Eigenschaften
Font.Size = 12                         ' Größe in Punkten
Debug.Print Font.Name                  ' "Helvetica"
Debug.Print Font.Alias                 ' "F1" (automatisch)
Debug.Print Font.ObjectNumber          ' Interne Nummer

' Verwendung
Page.Canvas.Font = Font
Page.Canvas.DrawText 100, 100, "Text in dieser Schrift"
```

**Verfügbare Schriftarten:**
- Helvetica (Standard)
- Times-Roman
- Courier
- Symbol
- ZapfDingbats

---

### 5.6 Klasse: PdfPen (Zeichenstift)

**Aufgabe:** Eigenschaften für Linienzüge

```vb
Dim Pen As PdfPen
Set Pen = Canvas.Pen

' Eigenschaften
Pen.Width = 2                           ' Liniendicke
Pen.Color = vbRed                       ' Farbe (OLE_COLOR)
Pen.LineCap = 0                         ' 0=butt, 1=round, 2=projecting
Pen.LineJoin = 0                        ' 0=miter, 1=round, 2=bevel
Pen.MiterLimit = 10

' Linientyp
Pen.Solid()                             ' ————————
Pen.Dash()                              ' — — — —
Pen.Dot()                               ' · · · · ·
Pen.DashDot()                           ' —·—·—·—
Pen.DashDotDot()                        ' —··—··—
```

---

### 5.7 Klasse: PdfBrush (Füllpinsel)

**Aufgabe:** Füllfarben-Verwaltung

```vb
Dim Brush As PdfBrush
Set Brush = Canvas.Brush

' Füllung
Brush.Color = vbGreen                   ' Füllfarbe (OLE_COLOR)
Canvas.FillRectangle 50, 50, 100, 100   ' Mit dieser Farbe füllen
```

---

## 6. Verwendungsbeispiele

### Beispiel 1: Einfaches Dokument mit Text & Zeichnungen

```vb
Public Sub BeispielEinfach()
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page As PdfPage
    
    Set Doc = Pdf.CreateDocument()
    Set Page = Doc.AddPage()
    
    ' Überschrift
    Dim FontHeader As PdfFont
    Set FontHeader = Doc.AddFont("Helvetica")
    FontHeader.Size = 20
    Page.Canvas.Font = FontHeader
    Page.Canvas.DrawText 50, 750, "Rechnungsdokument"
    
    ' Linie unter Überschrift
    Page.Canvas.Pen.Width = 1
    Page.Canvas.DrawLine 50, 730, 545, 730
    
    ' Text in normaler Größe
    Dim FontNormal As PdfFont
    Set FontNormal = Doc.AddFont("Helvetica")
    FontNormal.Size = 10
    Page.Canvas.Font = FontNormal
    Page.Canvas.DrawText 50, 700, "Rechnungsnummer: INV-2026-001"
    Page.Canvas.DrawText 50, 680, "Datum: 20.07.2026"
    Page.Canvas.DrawText 50, 660, "Kunde: Muster GmbH"
    
    ' Box für Summe
    Page.Canvas.Pen.Width = 2
    Page.Canvas.DrawRectangle 350, 100, 150, 50
    Page.Canvas.DrawText 360, 125, "Summe: 1.234,56 EUR"
    
    ' Speichern
    If Doc.Save("C:\Output\rechnung.pdf") Then
        MsgBox "Rechnung erstellt!"
    End If
End Sub
```

### Beispiel 2: Mehrere Seiten

```vb
Public Sub BeispielMehrereSeiten()
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page1 As PdfPage, Page2 As PdfPage
    
    Set Doc = Pdf.CreateDocument()
    
    ' Erste Seite
    Set Page1 = Doc.AddPage()
    Page1.Canvas.DrawText 100, 100, "Das ist Seite 1"
    
    ' Zweite Seite
    Set Page2 = Doc.AddPage()
    Page2.Canvas.DrawText 100, 100, "Das ist Seite 2"
    Page2.PageSize = psA3              ' Andere Größe
    
    Doc.Save "C:\Output\multi.pdf"
End Sub
```

### Beispiel 3: Mit Bildern

```vb
Public Sub BeispielMitBildern()
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page As PdfPage
    Dim Img As PdfImage
    
    Set Doc = Pdf.CreateDocument()
    Set Page = Doc.AddPage()
    
    ' Bild laden
    Set Img = Pdf.LoadImage("C:\Images\logo.jpg")
    
    ' Zu Dokument hinzufügen
    Doc.AddImage Img
    
    ' Zeichnen in Originalgröße
    Page.Canvas.DrawImage Img, 50, 50
    
    ' Oder skaliert
    Page.Canvas.DrawImageEx Img, 50, 300, 200, 100
    
    ' Oder proportional angepasst
    Page.Canvas.DrawImageFit Img, 50, 500, 100, 100
    
    Doc.Save "C:\Output\mit_bild.pdf"
End Sub
```

### Beispiel 4: Verschiedene Formen

```vb
Public Sub BeispielFormen()
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page As PdfPage
    
    Set Doc = Pdf.CreateDocument()
    Set Page = Doc.AddPage()
    
    ' Rechteck-Umriss
    Page.Canvas.Pen.Width = 2
    Page.Canvas.DrawRectangle 50, 50, 100, 50
    
    ' Gefülltes Rechteck
    Page.Canvas.Brush.Color = vbYellow
    Page.Canvas.FillRectangle 50, 150, 100, 50
    
    ' Rechteck mit Umriss und Füllung
    Page.Canvas.DrawRectangle 50, 250, 100, 50
    Page.Canvas.FillRectangle 50, 250, 100, 50
    
    ' Ellipse
    Page.Canvas.Pen.Color = vbBlue
    Page.Canvas.DrawEllipse 200, 50, 80, 60
    
    ' Gefüllte Ellipse
    Page.Canvas.Brush.Color = vbCyan
    Page.Canvas.FillEllipse 200, 150, 80, 60
    
    ' Linie mit verschiedenen Typen
    Page.Canvas.Pen.Width = 1
    Page.Canvas.Pen.Solid()
    Page.Canvas.DrawLine 50, 350, 200, 350
    
    Page.Canvas.Pen.Dash()
    Page.Canvas.DrawLine 50, 370, 200, 370
    
    Page.Canvas.Pen.Dot()
    Page.Canvas.DrawLine 50, 390, 200, 390
    
    Doc.Save "C:\Output\formen.pdf"
End Sub
```

---

## 7. Klassifizierung & Architektur

### Architektur-Übersicht

```
┌─────────────────────────────────────────┐
│  ANWENDER-APPLIKATION (VB6 Form)        │
├─────────────────────────────────────────┤
│  KATEGORIE A - HAUPTKLASSEN             │
│  ┌─────────────────────────────────────┐│
│  │ Pdf → CreateDocument()              ││
│  │ PdfDocument → AddPage(), Save()     ││
│  │ PdfPage → Canvas, PageSize          ││
│  │ PdfCanvas2D → DrawText(), DrawLine()││
│  │ PdfWriter → BuildDocument()         ││
│  └─────────────────────────────────────┘│
├─────────────────────────────────────────┤
│  KATEGORIE B - SUPPORT-KLASSEN          │
│  ┌─────────────────────────────────────┐│
│  │ PdfBuffer → Text-Speicherung        ││
│  │ PdfSerializer → PDF-Syntax          ││
│  │ PdfContentBuilder → Drawing Cmds    ││
│  │ PdfObjectIndex → Object Mgmt        ││
│  │ PdfFont, PdfPen, PdfBrush          ││
│  └─────────────────────────────────────┘│
├─────────────────────────────────────────┤
│  KATEGORIE C - UTILITY-KLASSEN          │
│  ┌─────────────────────────────────────┐│
│  │ PdfCmdLine, PdfCmdRectangle, ...    ││
│  │ PdfGeometry, PdfPoint, PdfColor     ││
│  │ PdfPageSizes, PdfOrientations       ││
│  │ und 14 weitere Utility-Klassen      ││
│  └─────────────────────────────────────┘│
├─────────────────────────────────────────┤
│  DEPENDENCIES                            │
│  - PdfLib.dll (externe Bibliothek)      │
│  - stdole2.tlb (Windows-Automation)     │
│  - Adobe Reader (optional, Druck)       │
└─────────────────────────────────────────┘
```

### Klassen-Übersicht

**KATEGORIE A (8):** Pdf, PdfDocument, PdfPage, PdfCanvas2D, PdfWriter, PdfContentBuilder, PdfFile, PdfPrinter

**KATEGORIE B (9):** PdfBuffer, PdfSerializer, PdfObjectIndex, PdfGraphicsState, PdfTextState, PdfFormat, PdfPageObjects, PdfImage, PdfBinaryFile

**KATEGORIE C (24):** Alle Zeichenbefehle, Stil-Klassen, Geometrie, Konfiguration, PDF-interne Klassen

**MODULE (3):** PdfEnums, PdfGlobals, PdfVersion

---

## 8. Fehlerbehandlung

### Fehler abfangen

```vb
Dim Doc As PdfDocument
Set Doc = Pdf.CreateDocument()

' Versuchen zu speichern
If Not Doc.Save("C:\Output\test.pdf") Then
    ' Fehler aufgetreten
    Dim ErrorNum As Long
    Dim ErrorMsg As String
    
    ErrorNum = Doc.LastErrorNumber
    ErrorMsg = Doc.LastErrorText
    
    MsgBox "Fehler " & ErrorNum & ": " & ErrorMsg
End If
```

### Häufige Fehler

| Fehler | Ursache | Lösung |
|--------|--------|--------|
| "Datei nicht gefunden" | Ausgabeverzeichnis existiert nicht | Verzeichnis erstellen |
| "Zugriff verweigert" | Keine Schreibberechtigung | Pfad/Berechtigungen prüfen |
| "Objekt nicht gesetzt" | Vergessenes `New` oder `Set` | Initialisierung prüfen |
| "Ungültige Seitengröße" | Falscher PageSize-Wert | psA4, psLetter, etc. verwenden |

---

## 9. FAQ & Häufige Probleme

### F: Wie drucke ich das PDF automatisch?

**A:** PDFs werden über Adobe Reader gedruckt:
```vb
Dim Shell As Object
Set Shell = CreateObject("WScript.Shell")

' PDF speichern
Doc.Save "C:\Output\test.pdf"

' Über Adobe Reader drucken
Shell.Run "AcroRd32.exe /p /h """ & "C:\Output\test.pdf" & """"
```

### F: Kann ich PDF-Dateien ändern?

**A:** Nein, die Bibliothek kann nur neue PDFs erstellen, nicht bearbeiten.

### F: Welche Schriftarten werden unterstützt?

**A:** Nur die StandardPDF-Fonts: Helvetica, Times-Roman, Courier, Symbol, ZapfDingbats

### F: Kann ich Text in Farbe zeichnen?

**A:** Ja, über das Pen/Brush-System:
```vb
Page.Canvas.Pen.Color = vbRed
Page.Canvas.DrawText 100, 100, "Roter Text"
```

### F: Was ist die maximale Dateigröße?

**A:** Abhängig vom Speicher. Mit Bildern: ~50 MB praktisches Maximum.

### F: Funktioniert es auf 64-Bit Windows?

**A:** Ja, aber VB6 muss im Compatibility-Modus laufen.

---

## 10. Changelog

### Version 1.1 (20.07.2026)
- ✅ Erste öffentliche Version
- ✅ 41 Klassen + 3 Module
- ✅ Vollständige Dokumentation
- ✅ Code-Kommentierung
- ✅ Beispiele & Handbuch

### Version 1.0 (interner Build)
- Erste Implementierung

---

## Zusätzliche Ressourcen

**PDF-Spezifikation:** [Adobe PDF Reference](https://www.adobe.com)  
**VB6 Dokumentation:** [MSDN VB6](https://docs.microsoft.com/en-us/office/vba/api/overview)  
**Support:** Kontakt mit Kapp Niles GmbH & Co, KG

---

## Index der Klassen

### Nach Kategorie

**KATEGORIE A (Hauptklassen):**
- Pdf - Einstiegspunkt
- PdfDocument - Dokumentverwaltung
- PdfPage - Seite
- PdfCanvas2D - Zeichenfläche
- PdfWriter - PDF-Engine
- PdfContentBuilder - Content-Stream
- PdfFile - Datei-Speicherung
- PdfPrinter - Drucker-Integration

**KATEGORIE B (Support):**
- PdfBuffer - Text-Puffer
- PdfSerializer - PDF-Syntax
- PdfObjectIndex - Objekt-Verwaltung
- PdfGraphicsState - Zeichenzustand
- PdfTextState - Text-Zustand
- PdfFormat - Formatierung
- PdfPageObjects - Seiten-Objekte
- PdfImage - Bildverwaltung
- PdfBinaryFile - Binär-Datei

**KATEGORIE C (Utility):**
- 24 weitere Klassen für spezifische Aufgaben

### Alphabetische Liste

A: Pdf, PdfBrush, PdfBuffer  
C: PdfCanvas2D, PdfCmdEllipse, PdfCmdImage, PdfCmdLine, PdfCmdPolygon, PdfCmdPolyline, PdfCmdRectangle, PdfCmdText, PdfColor, PdfContentBuilder  
D: PdfDocument  
E: PdfEnums  
F: PdfFile, PdfFont, PdfFormat  
G: PdfGeometry, PdfGlobals, PdfGraphicsState  
I: PdfImage, PdfIniFile  
L: PdfLineCaps, PdfLineJoins  
O: PdfObject, PdfObjectIndex, PdfObjectNumbers, PdfOrientations  
P: PdfPage, PdfPageObjects, PdfPageSizes, PdfPen, PdfPoint, PdfPointCollection, PdfPrinter, PdfPrinterInfo, PdfPrintOptions  
S: PdfSerializer  
T: PdfTextState  
V: PdfVersion, PdfWriter

---

**Dokumentation Version 1.0 - 20.07.2026**  
**© 2026 Kapp Niles GmbH & Co, KG**
