VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================
' KLASSE......: PdfWriter
' KATEGORIE...: A - HAUPTKLASSE (Interne PDF-Engine)
' VERSION.....: 1.0
'
' BESCHREIBUNG:
' Kern-Engine zum Schreiben von PDF-Dateien. Diese interne
' Klasse verwaltet:
' - PDF-Struktur (Header, Catalog, Pages)
' - Objekt-Numbering
' - Ressourcen (Fonts, Images)
' - Cross-Reference-Tabelle (XRef)
'
' VERWENDUNG: Intern - nicht direkt vom Anwender
' DEPENDENCIES: PdfBuffer, PdfSerializer, PdfObjectIndex
' AUTOR: Kapp Niles GmbH & Co, KG
'===============================================================

Private mBuffer As PdfBuffer
Private mSerializer As PdfSerializer
Private mIndex As PdfObjectIndex
Private mStartXRef As Long
Private mNumbers As PdfObjectNumbers
Private mCatalogObject As Long
Private mPagesObject As Long
Private mPageObjects As Collection

Private Sub Class_Initialize()
   Set mBuffer = New PdfBuffer
   Set mSerializer = New PdfSerializer
   mSerializer.SetBuffer mBuffer
   Set mIndex = New PdfObjectIndex
   Set mNumbers = New PdfObjectNumbers
End Sub

'---------------------------------------------------------------
' METHODE: BuildDocument
' PARAMETER: Doc (PdfDocument) - Zu speicherndes Dokument
' RÜCKGABE: String - Kompletter PDF-Dateiinhalt
' BESCHREIBUNG: Erstellt die komplette PDF-Datei mit allen
'               Objekten, Seiten, Fonts und Bildern
'---------------------------------------------------------------
Public Function BuildDocument(ByVal Doc As PdfDocument) As String
    mBuffer.Clear
    mIndex.Clear
    mNumbers.Clear
    Set mPageObjects = New Collection
    
    ' Nummern zuweisen
    mCatalogObject = mNumbers.NextNumber
    mPagesObject = mNumbers.NextNumber
    
    Dim I As Long
    Dim P As PdfPageObjects
    
    ' Nummern für Seiten
    For I = 1 To Doc.PageCount
        Set P = New PdfPageObjects
        P.PageObject = mNumbers.NextNumber
        P.ContentObject = mNumbers.NextNumber
        mPageObjects.Add P
    Next
    
    ' Nummern für Fonts
    For I = 1 To Doc.FontCount
        Doc.Font(I).ObjectNumber = mNumbers.NextNumber
    Next
    
    ' Nummern für Bilder
    For I = 1 To Doc.ImageCount
        Doc.Image(I).ObjectNumber = mNumbers.NextNumber
    Next
    
    ' PDF schreiben
    WriteHeader
    WriteCatalog
    WritePages
    
    For I = 1 To Doc.PageCount
        WritePage Doc, Doc.Page(I), mPageObjects(I)
    Next
    
    For I = 1 To Doc.PageCount
        WriteContent Doc.Page(I), mPageObjects(I)
    Next
    
    For I = 1 To Doc.FontCount
        WriteFont Doc.Font(I)
    Next
    
    For I = 1 To Doc.ImageCount
        WriteImage Doc.Image(I)
    Next
    
    WriteXRef
    WriteTrailer
    
    BuildDocument = mBuffer.Text
End Function

Private Sub WriteHeader()
    mSerializer.AddRaw "%PDF-1.4"
    mSerializer.AddRaw "%"
End Sub

Private Sub WriteCatalog()
    BeginObject mCatalogObject, CLng(potCatalog)
    mSerializer.DictionaryBegin
    mSerializer.AddName "Type", "/Catalog"
    mSerializer.AddReference "Pages", mPagesObject
    mSerializer.DictionaryEnd
    mSerializer.EndObject
End Sub

Private Sub WritePages()
    BeginObject mPagesObject, CLng(potpages)
    mSerializer.DictionaryBegin
    mSerializer.AddName "Type", "/Pages"
    mSerializer.AddInteger "Count", mPageObjects.Count
    
    Dim I As Long
    Dim P As PdfPageObjects
    Dim S As String
    S = "/Kids ["
    For I = 1 To mPageObjects.Count
        Set P = mPageObjects(I)
        S = S & P.PageObject & " 0 R " & vbCrLf
    Next
    S = S & "]"
    mSerializer.AddRaw S
    mSerializer.DictionaryEnd
    mSerializer.EndObject
End Sub

Private Sub WritePage(ByVal Doc As PdfDocument, ByVal Page As PdfPage, ByVal Obj As PdfPageObjects)
    BeginObject Obj.PageObject, CLng(potPage)
    mSerializer.DictionaryBegin
    mSerializer.AddName "Type", "/Page"
    mSerializer.AddReference "Parent", mPagesObject
    mSerializer.AddValue "MediaBox", "[0 0 " & CStr(Page.Width) & " " & CStr(Page.Height) & "]"
    mSerializer.AddReference "Contents", Obj.ContentObject
    WriteResources Doc
    mSerializer.DictionaryEnd
    mSerializer.EndObject
End Sub

Private Sub WriteContent(ByVal Page As PdfPage, ByVal Obj As PdfPageObjects)
    Dim Builder As New PdfContentBuilder
    Dim Inhalt As String
    Inhalt = Builder.Build(Page.Canvas)
    BeginObject Obj.ContentObject, CLng(potContent)
    mSerializer.DictionaryBegin
    mSerializer.AddInteger "Length", Len(Inhalt)
    mSerializer.DictionaryEnd
    mSerializer.BeginStream
    mSerializer.AddRaw Inhalt
    mSerializer.EndStream
    mSerializer.EndObject
End Sub

Private Sub BeginObject(ByVal Number As Long, ByVal Kind As PdfObjectType)
    mIndex.Add Number, mBuffer.Position, Kind
    mSerializer.BeginObject Number
End Sub

Private Sub WriteTrailer()
    mSerializer.AddRaw "trailer" & vbCrLf
    mSerializer.DictionaryBegin
    mSerializer.AddInteger "Size", mIndex.Count + 1
    mSerializer.AddReference "Root", 1
    mSerializer.DictionaryEnd
    mSerializer.AddRaw "startxref" & vbCrLf
    mSerializer.AddRaw CStr(mStartXRef) & vbCrLf
    mSerializer.AddRaw "%%EOF" & vbCrLf
End Sub

Private Sub WriteXRef()
    Dim I As Long
    mStartXRef = mBuffer.Position
    mSerializer.AddRaw "xref" & vbCrLf
    mSerializer.AddRaw "0 " & (mIndex.Count + 1) & vbCrLf
    mSerializer.AddRaw "0000000000 65535 f " & vbCrLf
    For I = 1 To mIndex.Count
        mSerializer.AddRaw Right$("0000000000" & CStr(mIndex.Position(I)), 10) & " 00000 n " & vbCrLf
    Next I
End Sub

Private Sub WriteFont(ByVal Font As PdfFont)
    BeginObject Font.ObjectNumber, CLng(potFont)
    mSerializer.DictionaryBegin
    mSerializer.AddName "Type", "/Font"
    mSerializer.AddName "Subtype", "/Type1"
    mSerializer.AddName "BaseFont", "/" & Font.Name
    mSerializer.DictionaryEnd
    mSerializer.EndObject
End Sub

Private Sub WriteImage(ByVal Img As PdfImage)
    BeginObject Img.ObjectNumber, CLng(potImage)
    mSerializer.DictionaryBegin
    mSerializer.AddName "Type", "/XObject"
    mSerializer.AddName "Subtype", "/Image"
    mSerializer.AddInteger "Width", Img.Width
    mSerializer.AddInteger "Height", Img.Height
    
    Select Case Img.Components
        Case 1
            mSerializer.AddName "ColorSpace", "/DeviceGray"
        Case 3
            mSerializer.AddName "ColorSpace", "/DeviceRGB"
        Case 4
            mSerializer.AddName "ColorSpace", "/DeviceCMYK"
        Case Else
            Err.Raise vbObjectError + 1001, "PdfWriter", "Unbekannter JPEG-Farbraum"
    End Select
    
    mSerializer.AddInteger "BitsPerComponent", 8
    mSerializer.AddName "Filter", "/DCTDecode"
    
    Dim Data() As Byte
    Dim I As Long
    Data = Img.Data
    mSerializer.AddInteger "Length", UBound(Data) + 1
    mSerializer.DictionaryEnd
    mSerializer.BeginStream
    
    For I = 0 To UBound(Data)
        mBuffer.AppendByte Data(I)
    Next
    
    mSerializer.EndStream
    mSerializer.EndObject
End Sub

Private Sub WriteResources(ByVal Doc As PdfDocument)
    mSerializer.AddRaw "/Resources <<" & vbCrLf
    Dim I As Long
    Dim F As PdfFont
    
    mSerializer.AddRaw "/Font <<" & vbCrLf
    For I = 1 To Doc.FontCount
        Set F = Doc.Font(I)
        mSerializer.AddRaw "/" & F.Alias & " " & F.ObjectNumber & " 0 R" & vbCrLf
    Next
    mSerializer.AddRaw ">>" & vbCrLf
    
    If Doc.ImageCount > 0 Then
        Dim Img As PdfImage
        mSerializer.AddRaw "/XObject <<"
        For I = 1 To Doc.ImageCount
            Set Img = Doc.Image(I)
            mSerializer.AddRaw "/" & Img.Alias & " " & Img.ObjectNumber & " 0 R" & vbCrLf
        Next
        mSerializer.AddRaw ">>"
    End If
    
    mSerializer.AddRaw ">>" & vbCrLf
End Sub
