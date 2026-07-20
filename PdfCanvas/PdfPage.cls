VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================
' KLASSE......: PdfPage
' KATEGORIE...: A - HAUPTKLASSE
' VERSION.....: 1.0
'
' BESCHREIBUNG:
' Repräsentiert eine einzelne PDF-Seite innerhalb eines
' Dokuments. Jede Seite hat:
' - Eigene Abmessungen (Breite x Höhe)
' - Größe (A4, Letter, etc.)
' - Orientierung (Portrait/Landscape)
' - Canvas (PdfCanvas2D) für Zeichnungen
'
' VERWENDUNG:
'   Dim Doc As New PdfDocument
'   Dim Page As PdfPage
'   Set Page = Doc.AddPage()
'   Page.Canvas.DrawLine 50, 50, 200, 200
'
' STANDARD-WERT: A4 Portrait (595 x 842 Punkte)
' ABHÄNGIGKEITEN: PdfCanvas2D
'===============================================================

Private mWidth As Long
Private mHeight As Long
Private mCanvas As PdfCanvas2D
Private mPageSize As Long
Private mOrientation As Long

'---------------------------------------------------------------
' EVENT: Class_Initialize
' BESCHREIBUNG: Initialisiert die Seite mit A4 Portrait
'               Seitengröße in PostScript-Punkten (72 dpi)
'---------------------------------------------------------------
Private Sub Class_Initialize()
    ' A4 in PostScript-Punkten: 595 x 842
    mPageSize = psA4
    mOrientation = poPortrait
    UpdatePageSize
    Set mCanvas = New PdfCanvas2D
End Sub

'---------------------------------------------------------------
' PROPERTY: Width
' RÜCKGABE: Long - Seitenbreite in Punkten
'---------------------------------------------------------------
Public Property Get Width() As Long
    Width = mWidth
End Property

Public Property Let Width(ByVal Value As Long)
    mWidth = Value
End Property

'---------------------------------------------------------------
' PROPERTY: Height
' RÜCKGABE: Long - Seitenhöhe in Punkten
'---------------------------------------------------------------
Public Property Get Height() As Long
    Height = mHeight
End Property

Public Property Let Height(ByVal Value As Long)
    mHeight = Value
End Property

'---------------------------------------------------------------
' PROPERTY: Canvas
' RÜCKGABE: PdfCanvas2D - Zeichenfläche für diese Seite
' BESCHREIBUNG: Über das Canvas können Grafiken und Text
'               gezeichnet werden.
'---------------------------------------------------------------
Public Property Get Canvas() As PdfCanvas2D
    Set Canvas = mCanvas
End Property

'---------------------------------------------------------------
' PROPERTY: PageSize
' PARAMETER: Value (Long) - Eine der Enumerationen:
'            psA5 (420x595), psA4 (595x842), psA3 (842x1191),
'            psLetter (612x792), psLegal (612x1008)
' BESCHREIBUNG: Setzt die Seitengröße und passt Breite/Höhe
'               entsprechend an
'---------------------------------------------------------------
Public Property Get PageSize() As Long
    PageSize = mPageSize
End Property

Public Property Let PageSize(ByVal Value As Long)
    mPageSize = Value
    UpdatePageSize
End Property

'---------------------------------------------------------------
' PROPERTY: Orientation
' PARAMETER: Value (Long) - poPortrait (0) oder poLandscape (1)
' BESCHREIBUNG: Setzt die Seiten-Orientierung und vertauscht
'               bei Bedarf Breite und Höhe
'---------------------------------------------------------------
Public Property Get Orientation() As Long
    Orientation = mOrientation
End Property

Public Property Let Orientation(ByVal Value As Long)
    mOrientation = Value
    UpdatePageSize
End Property

'---------------------------------------------------------------
' INTERNER HELPER: UpdatePageSize
' BESCHREIBUNG: Aktualisiert Breite und Höhe basierend auf
'               Seitengröße und Orientierung
'---------------------------------------------------------------
Private Sub UpdatePageSize()
    Select Case mPageSize
        Case psA5
            mWidth = 420
            mHeight = 595
        Case psA4
            mWidth = 595
            mHeight = 842
        Case psA3
            mWidth = 842
            mHeight = 1191
        Case psLetter
            mWidth = 612
            mHeight = 792
        Case psLegal
            mWidth = 612
            mHeight = 1008
    End Select
    
    ' Bei Landscape: Breite und Höhe vertauschen
    If mOrientation = poLandscape Then
        Dim T As Long
        T = mWidth
        mWidth = mHeight
        mHeight = T
    End If
End Sub
