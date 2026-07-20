VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================
' KLASSE......: PdfDocument
' KATEGORIE...: A - HAUPTKLASSE (Kern der Bibliothek)
' VERSION.....: 1.0
'
' BESCHREIBUNG:
' Repräsentiert ein PDF-Dokument. Dies ist die zentrale Klasse,
' mit der Anwender hauptsächlich arbeiten. Sie verwaltet:
' - Seiten (Collection)
' - Schriftarten (Fonts)
' - Bilder
' - Standard-Einstellungen
'
' VERWENDUNG:
'   Dim Doc As New PdfDocument
'   Dim Page As PdfPage
'   Set Page = Doc.AddPage()
'   Page.Canvas.DrawText 100, 100, "Hallo PDF!"
'   Doc.Save "C:\Output\Test.pdf"
'
' ABHÄNGIGKEITEN: PdfPage, PdfFont, PdfImage, PdfWriter, PdfFile
' AUTOR: Kapp Niles GmbH & Co, KG
'===============================================================

Private mPages As Collection
Private mFonts As Collection
Private mDefaultFont As PdfFont
Private mImages As Collection
Private mPageSizes As PdfPageSizes
Private mOrientations As PdfOrientations
Private mLineCaps As PdfLineCaps
Private mLineJoins As PdfLineJoins
Private mLastErrorNumber As Long
Private mLastErrorText As String

'---------------------------------------------------------------
' EVENT: Class_Initialize
' BESCHREIBUNG: Initialisiert das Dokument mit Standard-Werten
'               - Leere Seiten-Sammlung
'               - Helvetica-Schrift (Standard)
'               - A4-Format, Portrait-Orientierung
'---------------------------------------------------------------
Private Sub Class_Initialize()
    Set mPages = New Collection
    Set mFonts = New Collection
    Set mImages = New Collection
    
    Set mDefaultFont = AddFont("Helvetica")
    mDefaultFont.Size = 12
    
    Set mPageSizes = New PdfPageSizes
    Set mOrientations = New PdfOrientations
    Set mLineCaps = New PdfLineCaps
    Set mLineJoins = New PdfLineJoins
End Sub

'---------------------------------------------------------------
' METHODE: AddPage
' RÜCKGABE: PdfPage - Neu erstellte Seite
' BESCHREIBUNG: Fügt eine neue leere Seite zum Dokument hinzu.
'               Die neue Seite erhält automatisch die Standard-
'               Schrift und wird zur Sammlung hinzugefügt.
'---------------------------------------------------------------
Public Function AddPage() As PdfPage
    Dim P As PdfPage
    Set P = New PdfPage
    Set P.Canvas.Font = mDefaultFont
    mPages.Add P
    Set AddPage = P
End Function

'---------------------------------------------------------------
' PROPERTY: PageCount (READ-ONLY)
' RÜCKGABE: Long - Anzahl der Seiten im Dokument
'---------------------------------------------------------------
Public Property Get PageCount() As Long
    PageCount = mPages.Count
End Property

'---------------------------------------------------------------
' PROPERTY: Page(Index)
' PARAMETER: Index (Long) - 1-basierter Index (1 bis PageCount)
' RÜCKGABE: PdfPage - Die Seite mit diesem Index
'---------------------------------------------------------------
Public Property Get Page(ByVal Index As Long) As PdfPage
    Set Page = mPages(Index)
End Property

'---------------------------------------------------------------
' PROPERTY: Pages (INTERNAL)
' RÜCKGABE: Collection - Alle Seiten (nur für interne Nutzung)
'---------------------------------------------------------------
Friend Property Get Pages() As Collection
    Set Pages = mPages
End Property

'---------------------------------------------------------------
' METHODE: Save
' PARAMETER: FileName (String) - Zielverzeichnis und Dateiname
' RÜCKGABE: Boolean - True wenn erfolgreich, False bei Fehler
' BESCHREIBUNG: Speichert das gesamte Dokument als PDF-Datei.
'               Bei Fehler können Fehlerdetails über
'               LastErrorNumber und LastErrorText abgerufen werden.
'
' BEISPIEL:
'   If Not Doc.Save("C:\Output\test.pdf") Then
'     MsgBox "Fehler: " & Doc.LastErrorText
'   End If
'---------------------------------------------------------------
Public Function Save(ByVal FileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim Writer As New PdfWriter
    Dim File As New PdfFile
    
    ClearLastError
    
    File.Save FileName, Writer.BuildDocument(Me)
    Save = True
    
    Exit Function
ErrorHandler:
    SetLastError Err.Number, Err.Description
    Save = False
End Function

'---------------------------------------------------------------
' PROPERTY: DefaultFont
' RÜCKGABE: PdfFont - Die Standard-Schrift des Dokuments
' BESCHREIBUNG: Liefert die Standard-Schrift, die neuen Seiten
'               automatisch zugewiesen wird.
'---------------------------------------------------------------
Public Property Get DefaultFont() As PdfFont
    Set DefaultFont = mDefaultFont
End Property

'---------------------------------------------------------------
' METHODE: AddImage
' PARAMETER: Image (PdfImage) - Bilder-Objekt
' BESCHREIBUNG: Registriert ein Bild im Dokument.
'               Das Bild erhält automatisch einen eindeutigen
'               Alias (Im1, Im2, etc.).
'---------------------------------------------------------------
Public Sub AddImage(ByVal Image As PdfImage)
    Image.Alias = "Im" & CStr(mImages.Count + 1)
    mImages.Add Image
End Sub

'---------------------------------------------------------------
' PROPERTY: Image(Index)
' RÜCKGABE: PdfImage - Das Bild mit dem angegebenen Index
'---------------------------------------------------------------
Public Property Get Image(ByVal Index As Long) As PdfImage
    Set Image = mImages(Index)
End Property

'---------------------------------------------------------------
' PROPERTY: ImageCount (READ-ONLY)
' RÜCKGABE: Long - Anzahl der registrierten Bilder
'---------------------------------------------------------------
Public Property Get ImageCount() As Long
    ImageCount = mImages.Count
End Property

'---------------------------------------------------------------
' METHODE: AddFont
' PARAMETER: Name (String) - PDF-Schriftname (z.B. "Helvetica")
' RÜCKGABE: PdfFont - Neu erstelltes Schrift-Objekt
' BESCHREIBUNG: Registriert eine neue Schriftart.
'               Die Schrift erhält automatisch einen Alias (F1, F2, etc.)
'               und wird mit der Standard-Größe 10 initialisiert.
'---------------------------------------------------------------
Public Function AddFont(ByVal Name As String) As PdfFont
    Dim F As PdfFont
    Set F = New PdfFont
    F.Name = Name
    F.Alias = "F" & CStr(mFonts.Count + 1)
    F.Size = 10
    mFonts.Add F
    Set AddFont = F
End Function

'---------------------------------------------------------------
' PROPERTY: FontCount (READ-ONLY)
' RÜCKGABE: Long - Anzahl der registrierten Schriften
'---------------------------------------------------------------
Public Property Get FontCount() As Long
    FontCount = mFonts.Count
End Property

'---------------------------------------------------------------
' PROPERTY: Font(Index)
' RÜCKGABE: PdfFont - Die Schrift mit dem angegebenen Index
'---------------------------------------------------------------
Public Property Get Font(ByVal Index As Long) As PdfFont
    Set Font = mFonts(Index)
End Property

'---------------------------------------------------------------
' PROPERTY: PageSizes
' RÜCKGABE: PdfPageSizes - Singleton für Seitengrößen
'---------------------------------------------------------------
Public Property Get PageSizes() As PdfPageSizes
    Set PageSizes = PdfGlobals.PageSizes
End Property

'---------------------------------------------------------------
' PROPERTY: Orientations
' RÜCKGABE: PdfOrientations - Singleton für Orientierungen
'---------------------------------------------------------------
Public Property Get Orientations() As PdfOrientations
    Set Orientations = PdfGlobals.Orientations
End Property

'---------------------------------------------------------------
' PROPERTY: LineCaps
' RÜCKGABE: PdfLineCaps - Linienenden-Stile
'---------------------------------------------------------------
Public Property Get LineCaps() As PdfLineCaps
    Set LineCaps = mLineCaps
End Property

'---------------------------------------------------------------
' PROPERTY: LineJoins
' RÜCKGABE: PdfLineJoins - Linienverbindungs-Stile
'---------------------------------------------------------------
Public Property Get LineJoins() As PdfLineJoins
    Set LineJoins = mLineJoins
End Property

'---------------------------------------------------------------
' INTERNER HELPER: ClearLastError
' BESCHREIBUNG: Setzt die Fehler-Variablen zurück
'---------------------------------------------------------------
Private Sub ClearLastError()
    mLastErrorNumber = 0
    mLastErrorText = ""
End Sub

'---------------------------------------------------------------
' INTERNER HELPER: SetLastError
' BESCHREIBUNG: Speichert Fehlerinformationen
'---------------------------------------------------------------
Private Sub SetLastError(ByVal Number As Long, ByVal Description As String)
    mLastErrorNumber = Number
    mLastErrorText = Description
End Sub

'---------------------------------------------------------------
' PROPERTY: LastErrorNumber
' RÜCKGABE: Long - Nummer des letzten Fehlers (0 = kein Fehler)
'---------------------------------------------------------------
Public Property Get LastErrorNumber() As Long
    LastErrorNumber = mLastErrorNumber
End Property

'---------------------------------------------------------------
' PROPERTY: LastErrorText
' RÜCKGABE: String - Beschreibung des letzten Fehlers
'---------------------------------------------------------------
Public Property Get LastErrorText() As String
    LastErrorText = mLastErrorText
End Property
