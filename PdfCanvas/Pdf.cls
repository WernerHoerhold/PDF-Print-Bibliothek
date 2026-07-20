VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================
' KLASSE......: Pdf
' KATEGORIE...: A - HAUPTKLASSE (Einstiegspunkt)
' VERSION.....: 1.0
'
' BESCHREIBUNG:
' Einstiegspunkt der PDF-Bibliothek. Diese Klasse stellt
' Factory-Methoden zur Verfügung, um alle anderen
' Objekte der Bibliothek zu erstellen.
'
' VERWENDUNG:
'   Dim Pdf As New Pdf
'   Dim Doc As PdfDocument
'   Set Doc = Pdf.CreateDocument()
'
' ABHÄNGIGKEITEN: Keine externen
' AUTOR: Kapp Niles GmbH & Co, KG
' DATUM: 2026
'===============================================================

'---------------------------------------------------------------
' METHODE: CreateDocument
' RÜCKGABE: PdfDocument - Neues PDF-Dokument
' BESCHREIBUNG: Erstellt ein neues, leeres PDF-Dokument
'---------------------------------------------------------------
Public Function CreateDocument() As PdfDocument
    Set CreateDocument = New PdfDocument
End Function

'---------------------------------------------------------------
' METHODE: LoadImage
' PARAMETER: FileName (String) - Pfad zur Bilddatei (JPEG)
' RÜCKGABE: PdfImage - Geladenesbildobjekt
' BESCHREIBUNG: Lädt ein JPEG-Bild aus Datei
'---------------------------------------------------------------
Public Function LoadImage(ByVal FileName As String) As PdfImage
    Dim Img As PdfImage
    Set Img = New PdfImage
    Img.LoadFromFile FileName
    Set LoadImage = Img
End Function

'---------------------------------------------------------------
' METHODE: CreateImage (STUB - nicht implementiert)
'---------------------------------------------------------------
Public Function CreateImage() As PdfImage
End Function

'---------------------------------------------------------------
' METHODE: CreateFont (STUB - nicht implementiert)
'---------------------------------------------------------------
Public Function CreateFont() As PdfFont
End Function

'---------------------------------------------------------------
' METHODE: CreatePointCollection
' RÜCKGABE: PdfPointCollection - Leere Punkt-Sammlung
' BESCHREIBUNG: Erstellt eine Sammlung von Punkten für
'               Polylinien und Polygone
'---------------------------------------------------------------
Public Function CreatePointCollection() As PdfPointCollection
    Set CreatePointCollection = New PdfPointCollection
End Function

'---------------------------------------------------------------
' METHODE: CreateBrush (STUB - nicht implementiert)
'---------------------------------------------------------------
Public Function CreateBrush() As PdfBrush
End Function

'---------------------------------------------------------------
' METHODE: CreatePen (STUB - nicht implementiert)
'---------------------------------------------------------------
Public Function CreatePen() As PdfPen
End Function

'---------------------------------------------------------------
' PROPERTY: Version
' RÜCKGABE: String - Versionsnummer (z.B. "1.1")
'---------------------------------------------------------------
Public Property Get Version() As String
    Version = PDF_VERSION
End Property

'---------------------------------------------------------------
' PROPERTY: Build
' RÜCKGABE: String - Build-Nummer
'---------------------------------------------------------------
Public Property Get Build() As String
    Build = PDF_BUILD
End Property

'---------------------------------------------------------------
' PROPERTY: Name
' RÜCKGABE: String - Bibliotheks-Name ("PDFLib")
'---------------------------------------------------------------
Public Property Get Name() As String
    Name = PDF_NAME
End Property

'---------------------------------------------------------------
' PROPERTY: Copyright
' RÜCKGABE: String - Copyright-Information
'---------------------------------------------------------------
Public Property Get Copyright() As String
    Copyright = PDF_COPYRIGHT
End Property

'---------------------------------------------------------------
' PROPERTY: About
' RÜCKGABE: String - Vollständige Informationen über
'                    die Bibliothek (Name, Version, Build, Copyright)
'---------------------------------------------------------------
Public Property Get About() As String
    About = PDF_NAME & vbCrLf & _
            " Version " & PDF_VERSION & vbCrLf & _
            " (Build " & PDF_BUILD & ")" & vbCrLf & vbCrLf & _
            PDF_COPYRIGHT
End Property
