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

'=========================================================
' Klasse.....: PdfDocument
' Version....: 1.0
'
' Beschreibung
' -------------
' Repräsentiert ein PDF-Dokument.
'
' Der Anwender arbeitet hauptsächlich mit dieser Klasse.
'=========================================================
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

Private Sub Class_Initialize()

    Set mPages = New Collection
    Set mFonts = New Collection
    Set mImages = New Collection
    
'    Set mDefaultFont = New PdfFont
'
'    With mDefaultFont
'        .Name = "Helvetica"
'        .Alias = "F1"
'        .Size = 12
'    End With
    
    Set mDefaultFont = AddFont("Helvetica")
    mDefaultFont.Size = 12
    
    Set mPageSizes = New PdfPageSizes
    Set mOrientations = New PdfOrientations
    Set mLineCaps = New PdfLineCaps
    Set mLineJoins = New PdfLineJoins
    
End Sub

'---------------------------------------------------------
' Neue Seite anlegen
'---------------------------------------------------------
Public Function AddPage() As PdfPage

    Dim P As PdfPage

    Set P = New PdfPage

    Set P.Canvas.Font = mDefaultFont
    
    mPages.Add P

    Set AddPage = P

End Function

'---------------------------------------------------------
' Anzahl Seiten
'---------------------------------------------------------
Public Property Get PageCount() As Long

    PageCount = mPages.Count

End Property

'---------------------------------------------------------
' Zugriff auf eine Seite
'---------------------------------------------------------
Public Property Get Page(ByVal Index As Long) As PdfPage

    Set Page = mPages(Index)

End Property

'---------------------------------------------------------
' Interner Zugriff
'---------------------------------------------------------
Friend Property Get Pages() As Collection

    Set Pages = mPages

End Property

'---------------------------------------------------------
' PDF speichern
'---------------------------------------------------------
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

Public Property Get DefaultFont() As PdfFont

    Set DefaultFont = mDefaultFont

End Property

'---------------------------------------------------------
' Bild hinzufügen
'---------------------------------------------------------
Public Sub AddImage(ByVal Image As PdfImage)

    Image.Alias = "Im" & CStr(mImages.Count + 1)
    
    mImages.Add Image

End Sub

'---------------------------------------------------------
' Bild zurückgeben
'---------------------------------------------------------
Public Property Get Image( _
        ByVal Index As Long) As PdfImage

    Set Image = mImages(Index)

End Property

'---------------------------------------------------------
' Anzahl Bilder
'---------------------------------------------------------
Public Property Get ImageCount() As Long

    ImageCount = mImages.Count

End Property


'---------------------------------------------------------
' Neue Schrift registrieren
'---------------------------------------------------------
Public Function AddFont(ByVal Name As String) As PdfFont

    Dim F As PdfFont

    Set F = New PdfFont

    F.Name = Name
    F.Alias = "F" & CStr(mFonts.Count + 1)
    F.Size = 10

    mFonts.Add F

    Set AddFont = F

End Function

Public Property Get FontCount() As Long

    FontCount = mFonts.Count

End Property

Public Property Get Font(ByVal Index As Long) As PdfFont

    Set Font = mFonts(Index)

End Property

Public Property Get PageSizes() As PdfPageSizes

    Set PageSizes = PdfGlobals.PageSizes
    
End Property

Public Property Get Orientations() As PdfOrientations

    Set Orientations = PdfGlobals.Orientations
    

End Property

Public Property Get LineCaps() As PdfLineCaps

    Set LineCaps = mLineCaps

End Property

Public Property Get LineJoins() As PdfLineJoins

    Set LineJoins = mLineJoins

End Property

Private Sub ClearLastError()

    mLastErrorNumber = 0
    mLastErrorText = ""

End Sub

Private Sub SetLastError( _
        ByVal Number As Long, _
        ByVal Description As String)

    mLastErrorNumber = Number
    mLastErrorText = Description

End Sub
Public Property Get LastErrorNumber() As Long

    LastErrorNumber = mLastErrorNumber

End Property

Public Property Get LastErrorText() As String

    LastErrorText = mLastErrorText

End Property
