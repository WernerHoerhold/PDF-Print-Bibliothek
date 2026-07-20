# PDF-Print-Bibliothek 📄

[![Version](https://img.shields.io/badge/version-1.1-blue.svg)](https://github.com/WernerHoerhold/PDF-Print-Bibliothek)
[![License](https://img.shields.io/badge/license-Proprietary-red.svg)]()
[![Language](https://img.shields.io/badge/language-VB6-green.svg)]()
[![Status](https://img.shields.io/badge/status-Production-brightgreen.svg)]()

> **Eine professionelle ActiveX-DLL für Visual Basic 6.0 zur Erstellung von PDF-Dokumenten**

Entwickelt von **Kapp Niles GmbH & Co, KG** | 2026

---

## 🎯 Was ist die PDF-Print-Bibliothek?

Die **PDF-Print-Bibliothek** ermöglicht VB6-Entwicklern, PDF-Dokumente programmgesteuert zu erstellen und zu drucken. Sie bietet eine intuitive API für:

✅ PDF-Dokumente erstellen  
✅ Text in verschiedenen Schriftarten zeichnen  
✅ Grafische Formen (Linien, Rechtecke, Ellipsen, Polygone)  
✅ JPEG-Bilder einbinden  
✅ Mehrseiten-Dokumente  
✅ Druck über Adobe Reader  

**Ideal für:** Rechnungen, Berichte, Messschriebe, Bescheinigungen, etc.

---

## 🚀 Quick Start

### 1️⃣ Installation (30 Sekunden)

```bash
# DLL registrieren (als Administrator)
regsvr32 PdfLib.dll
```

### 2️⃣ In VB6 verwenden

```vb
' Projekt > Referenzen > PdfLib.dll hinzufügen

Public Sub CreatePDF()
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page As PdfPage
    
    Set Doc = Pdf.CreateDocument()
    Set Page = Doc.AddPage()
    
    Page.Canvas.DrawText 100, 100, "Hallo PDF!"
    Page.Canvas.DrawRectangle 50, 50, 200, 100
    
    Doc.Save "C:\Output\test.pdf"
    MsgBox "PDF erstellt: test.pdf"
End Sub
```

### 3️⃣ Fertig! 🎉

Ihre erste PDF ist erstellt. Ganz einfach!

---

## 📚 Dokumentation

Diese Repository enthält **vollständige Dokumentation**:

| Datei | Beschreibung | Größe |
|-------|-------------|-------|
| **[DOKUMENTATION.md](DOKUMENTATION.md)** | Projekt-Übersicht & Klassifizierung | 7 KB |
| **[PROGRAMMIERHANDBUCH.md](PROGRAMMIERHANDBUCH.md)** | Detailliertes Handbuch mit Beispielen | 20 KB |
| **Kommentierte Klassen** | Alle 41 Klassen mit inline-Dokumentation | - |

### 📖 Inhalte des Handbuchs

- ✅ Systemanforderungen
- ✅ Installation & Setup
- ✅ Schnellstart-Beispiele
- ✅ API-Referenz (alle Hauptklassen)
- ✅ 4 detaillierte Verwendungsbeispiele
- ✅ Architektur-Übersicht
- ✅ Fehlerbehandlung
- ✅ FAQ & Troubleshooting
- ✅ Klassenzusammenfassung

---

## 🏗️ Architektur

### Klassifizierung (41 Klassen + 3 Module)

```
┌─────────────────────────────────────────┐
│  KATEGORIE A - HAUPTKLASSEN (8)         │
│  Pdf, PdfDocument, PdfPage, Canvas2D    │
│  Writer, ContentBuilder, File, Printer  │
├─────────────────────────────────────────┤
│  KATEGORIE B - SUPPORT (9)              │
│  Buffer, Serializer, ObjectIndex        │
│  GraphicsState, Font, Pen, Brush, etc.  │
├─────────────────────────────────────────┤
│  KATEGORIE C - UTILITY (24)             │
│  Drawing Commands, Geometry, Config     │
└─────────────────────────────────────────┘
```

---

## 💡 Code-Beispiele

### Beispiel 1: Einfache Rechnung

```vb
Public Sub CreateInvoice()
    Dim Pdf As New Pdf
    Dim Doc As PdfDocument
    Dim Page As PdfPage
    Dim Font As PdfFont
    
    Set Doc = Pdf.CreateDocument()
    Set Page = Doc.AddPage()
    
    ' Überschrift
    Set Font = Doc.AddFont("Helvetica")
    Font.Size = 18
    Page.Canvas.Font = Font
    Page.Canvas.DrawText 50, 750, "RECHNUNG"
    
    ' Linien und Details
    Font.Size = 10
    Page.Canvas.Font = Font
    Page.Canvas.DrawText 50, 700, "Rechnungsnr: INV-2026-001"
    Page.Canvas.DrawText 50, 680, "Datum: 20.07.2026"
    Page.Canvas.DrawText 50, 660, "Summe: 1.234,56 EUR"
    
    ' Summen-Box
    Page.Canvas.Pen.Width = 2
    Page.Canvas.Brush.Color = vbYellow
    Page.Canvas.FillRectangle 350, 100, 150, 50
    Page.Canvas.DrawText 360, 125, "Total: 1.234,56"
    
    Doc.Save "C:\Output\invoice.pdf"
End Sub
```

### Beispiel 2: Mit Bildern

```vb
Public Sub CreateReportWithLogo()
    Dim Pdf As New Pdf, Doc As PdfDocument, Page As PdfPage
    
    Set Doc = Pdf.CreateDocument()
    Set Page = Doc.AddPage()
    
    ' Logo einfügen
    Dim Logo As PdfImage
    Set Logo = Pdf.LoadImage("C:\Images\logo.jpg")
    Doc.AddImage Logo
    Page.Canvas.DrawImage Logo, 50, 750
    
    ' Text daneben
    Page.Canvas.DrawText 150, 750, "Geschäftsbericht 2026"
    
    Doc.Save "C:\Output\report.pdf"
End Sub
```

### Beispiel 3: Mehrseiten-Dokument

```vb
Public Sub CreateMultiPagePDF()
    Dim Pdf As New Pdf, Doc As PdfDocument
    Dim Page1 As PdfPage, Page2 As PdfPage
    
    Set Doc = Pdf.CreateDocument()
    
    ' Seite 1 - A4 Portrait
    Set Page1 = Doc.AddPage()
    Page1.Canvas.DrawText 100, 100, "Dies ist Seite 1"
    
    ' Seite 2 - A3 Landscape
    Set Page2 = Doc.AddPage()
    Page2.PageSize = psA3
    Page2.Orientation = poLandscape
    Page2.Canvas.DrawText 100, 100, "Dies ist Seite 2 (A3 Landscape)"
    
    Doc.Save "C:\Output\multi.pdf"
End Sub
```

---

## 🎨 Features im Detail

### Seitengröße

- 📄 **A5** (420 x 595)
- 📄 **A4** (595 x 842) ← Standard
- 📄 **A3** (842 x 1191)
- 📄 **Letter** (612 x 792)
- 📄 **Legal** (612 x 1008)

### Orientierung

- 📊 **Portrait** (Hochformat)
- 📊 **Landscape** (Querformat)

### Zeichnungen

```vb
Canvas.DrawLine(X1, Y1, X2, Y2)                    ' Linie
Canvas.DrawRectangle(X, Y, W, H)                   ' Rechteck-Umriss
Canvas.FillRectangle(X, Y, W, H)                   ' Gefülltes Rechteck
Canvas.DrawEllipse(X, Y, W, H)                     ' Ellipse-Umriss
Canvas.FillEllipse(X, Y, W, H)                     ' Gefüllte Ellipse
Canvas.DrawText(X, Y, Text)                        ' Text
Canvas.DrawImage(Image, X, Y)                      ' Bild
Canvas.DrawPolygon(Points)                         ' Polygon
```

### Stift-Typen

```vb
Pen.Solid()        ' ————————
Pen.Dash()         ' — — — —
Pen.Dot()          ' · · · · ·
Pen.DashDot()      ' —·—·—·—
Pen.DashDotDot()   ' —··—··—
```

### Schriftarten

- ✏️ Helvetica (Standard)
- ✏️ Times-Roman
- ✏️ Courier
- ✏️ Symbol
- ✏️ ZapfDingbats

---

## 📋 Anforderungen

### System

| Anforderung | Minimum | Empfohlen |
|-------------|---------|-----------|
| **OS** | Windows XP | Windows 10/11 |
| **RAM** | 128 MB | 512 MB |
| **VB6 Runtime** | Erforderlich | Latest |
| **Adobe Reader** | Optional (Druck) | Adobe DC |

### Abhängigkeiten

- `stdole2.tlb` (Windows-Standard)
- `PdfLib.dll` (im Repository)

---

## 🔧 Installation

### Schritt 1: DLL registrieren

```bash
# Administrator-Fenster öffnen und ausführen:
regsvr32 PdfLib.dll

# Erfolgsmeldung:
# "DllRegisterServer in PdfLib.dll erfolgreich ausgeführt"
```

### Schritt 2: In VB6 Projekt hinzufügen

```
Projekt → Referenzen → Browse → PdfLib.dll → OK
```

### Schritt 3: Test

```vb
Private Sub Form_Load()
    Dim Pdf As New Pdf
    MsgBox Pdf.About
End Sub
```

---

## 📖 API-Referenz (Kurzfassung)

### Klasse: Pdf

```vb
Dim Pdf As New Pdf

CreateDocument()           ' → PdfDocument
LoadImage(FileName)        ' → PdfImage
CreatePointCollection()    ' → PdfPointCollection

' Properties
Pdf.Version                ' "1.1"
Pdf.Build                  ' Build-Nummer
Pdf.Name                   ' "PDFLib"
Pdf.About                  ' Vollständige Info
```

### Klasse: PdfDocument

```vb
Dim Doc As New PdfDocument

AddPage()                  ' → PdfPage
AddFont(Name)              ' → PdfFont
AddImage(Image)            ' Bild registrieren
Save(FileName)             ' Speichern

' Properties
Doc.PageCount              ' Anzahl Seiten
Doc.Page(Index)            ' Zugriff auf Seite
Doc.FontCount              ' Anzahl Fonts
Doc.ImageCount             ' Anzahl Bilder
Doc.LastErrorNumber        ' Fehler-Code
Doc.LastErrorText          ' Fehler-Meldung
```

### Klasse: PdfPage

```vb
Dim Page As PdfPage

' Properties
Page.Width                 ' Breite (PostScript-Punkte)
Page.Height                ' Höhe
Page.PageSize              ' psA4, psLetter, etc.
Page.Orientation           ' poPortrait, poLandscape
Page.Canvas                ' → PdfCanvas2D
```

### Klasse: PdfCanvas2D

```vb
Dim Canvas As PdfCanvas2D

DrawLine(X1, Y1, X2, Y2)
DrawRectangle(X, Y, W, H)
FillRectangle(X, Y, W, H)
DrawEllipse(X, Y, W, H)
FillEllipse(X, Y, W, H)
DrawText(X, Y, Text)
DrawImage(Image, X, Y)
DrawImageEx(Image, X, Y, W, H)
DrawPolygon(Points)
FillPolygon(Points)

' Properties
Canvas.Pen                 ' → PdfPen
Canvas.Brush               ' → PdfBrush
Canvas.Font                ' → PdfFont
```

---

## ❓ FAQ

### F: Kann ich die DLL mit .NET verwenden?

**A:** Nein, nur VB6 (COM-basiert). Für .NET: Andere Bibliotheken wie iTextSharp nutzen.

### F: Wie drucke ich das PDF automatisch?

**A:** Über Adobe Reader:
```vb
Shell.Run "AcroRd32.exe /p /h """ & FileName & """"
```

### F: Kann ich bestehende PDFs bearbeiten?

**A:** Nein, nur neue PDFs erstellen.

### F: Was ist die maximale Dateigröße?

**A:** Abhängig vom RAM. Mit Bildern: ~50 MB praktisch.

### F: Funktioniert es auf 64-Bit Windows?

**A:** Ja, aber VB6 muss im Kompatibilitätsmodus laufen.

### F: Welche Bildformate werden unterstützt?

**A:** Nur JPEG (.jpg). Andere Formate müssen vorher konvertiert werden.

---

## 📂 Projekt-Struktur

```
PDF-Print-Bibliothek/
├── PdfCanvas/
│   ├── Pdflib.vbp                  # VB6 Projektdatei
│   ├── PdfLib.dll                  # Kompilierte DLL
│   ├── Pdf.cls                     # Einstiegspunkt
│   ├── PdfDocument.cls             # Dokumentverwaltung
│   ├── PdfPage.cls                 # Seite
│   ├── PdfCanvas2D.cls             # Zeichenfläche
│   ├── PdfWriter.cls               # PDF-Engine
│   ├── PdfContentBuilder.cls       # Content-Stream
│   ├── [... weitere 34 Klassen ...]
│   └── [3 Module]
│
├── DOKUMENTATION.md                # Projektübersicht
├── PROGRAMMIERHANDBUCH.md          # Detailliertes Handbuch
└── README.md                       # Diese Datei
```

---

## 🤝 Beitrag & Support

### Lizenz

**Proprietary** - © 2026 Kapp Niles GmbH & Co, KG

### Support

Kontakt mit Kapp Niles GmbH & Co, KG für:
- Technischen Support
- Feature-Anfragen
- Bug-Reports

---

## 📊 Statistiken

| Metrik | Wert |
|--------|------|
| **Klassen** | 41 |
| **Module** | 3 |
| **Methoden** | ~150+ |
| **Properties** | ~100+ |
| **Zeilen Code** | ~5.000+ |
| **Dokumentation** | 100% |
| **Version** | 1.1 |

---

## 🎓 Lernressourcen

📚 **[PROGRAMMIERHANDBUCH.md](PROGRAMMIERHANDBUCH.md)** - Detailliertes Handbuch (70+ Seiten)  
📄 **[DOKUMENTATION.md](DOKUMENTATION.md)** - Projektübersicht  
💻 **Kommentierte Klassendateien** - Inline-Dokumentation  
🔍 **API-Referenz** - Alle Methoden & Properties

---

## 🎯 Use Cases

✅ **Rechnungen & Belege** - Automatisch generierte Geschäftsdokumente  
✅ **Berichte** - Geschäfts- und Datenberichte  
✅ **Messschriebe** - Technische Dokumentation  
✅ **Bescheinigungen** - Zertifikate und Urkunden  
✅ **Kataloge** - Produktkatalogisierung  
✅ **Formulare** - Ausgefüllte Vordrucke  

---

## 🚀 Next Steps

1. **Installation** → [Installation & Setup](PROGRAMMIERHANDBUCH.md#3-installation--setup)
2. **Schnellstart** → [Schnellstart](PROGRAMMIERHANDBUCH.md#4-schnellstart)
3. **Beispiele** → [Verwendungsbeispiele](PROGRAMMIERHANDBUCH.md#6-verwendungsbeispiele)
4. **API-Referenz** → [Hauptklassen & API](PROGRAMMIERHANDBUCH.md#5-hauptklassen--api)

---

## 📞 Kontakt

**Kapp Niles GmbH & Co, KG**

- 🌐 Website: [www.kapp-niles.de](https://www.kapp-niles.de)
- 📧 E-Mail: support@kapp-niles.de
- 📱 Tel: +49 (0) xxx xxx-xxx

---

## 📝 Lizenzinformation

```
PDF-Print-Bibliothek v1.1
© 2026 Kapp Niles GmbH & Co, KG
Alle Rechte vorbehalten.

Proprietäre Software. Nicht zur Weitergabe an Dritte.
```

---

**Viel Spaß mit der PDF-Print-Bibliothek! 🎉**

---

*Erstellt: 20.07.2026 | Version: 1.0 | Status: Dokumentation komplett*
