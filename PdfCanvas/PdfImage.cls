VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfImage
' Aufgabe....: Beschreibt ein Bild im PDF-Dokument
'
' Version....: 0.9.5
'=========================================================

Private mName As String

Private mWidth As Long
Private mHeight As Long

Private mData() As Byte
Private mComponents As Long
Private mAlias As String
Private mObjectNumber As Long

'---------------------------------------------------------
' Name des Bildes
'---------------------------------------------------------
Public Property Get Name() As String

    Name = mName

End Property

Public Property Let Name(ByVal Value As String)

    mName = Value

End Property

'---------------------------------------------------------
' Bildbreite in Pixel
'---------------------------------------------------------
Public Property Get Width() As Long

    Width = mWidth

End Property

Public Property Let Width(ByVal Value As Long)

    mWidth = Value

End Property

'---------------------------------------------------------
' Bildhöhe in Pixel
'---------------------------------------------------------
Public Property Get Height() As Long

    Height = mHeight

End Property

Public Property Let Height(ByVal Value As Long)

    mHeight = Value

End Property

'---------------------------------------------------------
' Bilddaten (JPEG-Stream)
'---------------------------------------------------------
Public Property Get Data() As Byte()

    Data = mData

End Property

Public Property Let Data(ByRef Value() As Byte)

    mData = Value

End Property

'---------------------------------------------------------
' JPEG-Datei laden
'---------------------------------------------------------
Public Sub LoadFromFile(ByVal FileName As String)

    Dim BF As New PdfBinaryFile
    Dim P As Long

    mData = BF.Load(FileName)

    P = InStrRev(FileName, "\")

    If P > 0 Then
        mName = Mid$(FileName, P + 1)
    Else
        mName = FileName
    End If
    
    ReadJpegHeader
    
End Sub

'---------------------------------------------------------
' Ein Byte lesen
'---------------------------------------------------------
Private Function GetByte( _
        ByVal Position As Long) As Long

    'GetByte = AscB(MidB$(mData, Position, 1))
    GetByte = mData(Position - 1)

End Function

'---------------------------------------------------------
' 16-Bit-Wert (Big Endian) lesen
'---------------------------------------------------------
Private Function GetWord( _
        ByVal Position As Long) As Long

    GetWord = _
        GetByte(Position) * &H100& + _
        GetByte(Position + 1)

End Function

'---------------------------------------------------------
' JPEG-Marker suchen
'---------------------------------------------------------
Private Function FindMarker( _
        ByVal Marker As Long) As Long

    Dim P As Long
    Dim L As Long

    L = UBound(mData)

    P = 1

    Do While P < L

        If GetByte(P) = &HFF Then

            If GetByte(P + 1) = Marker Then

                FindMarker = P

                Exit Function

            End If

            If GetByte(P + 1) <> &HD8 _
            And GetByte(P + 1) <> &HD9 Then

                P = P + GetWord(P + 2) + 2

            Else

                P = P + 2

            End If

        Else

            P = P + 1

        End If

    Loop

End Function

'---------------------------------------------------------
' JPEG-Informationen lesen
'---------------------------------------------------------
Private Sub ReadJpegHeader()

    Dim P As Long

    If GetByte(1) <> &HFF Or GetByte(2) <> &HD8 Then
        Err.Raise vbObjectError + 1002, "PdfImage", "Ungültige JPEG-Datei."
    End If
    
    ' Baseline JPEG
    P = FindMarker(&HC0)

    ' Progressive JPEG
    If P = 0 Then
        P = FindMarker(&HC2)
    End If

    If P = 0 Then
        Exit Sub
    End If

    mHeight = GetWord(P + 5)

    mWidth = GetWord(P + 7)
    
    mComponents = GetByte(P + 9)

End Sub

'---------------------------------------------------------
' Anzahl Farbkomponenten
'---------------------------------------------------------
Public Property Get Components() As Long

    Components = mComponents

End Property

Public Property Get Alias() As String

    Alias = mAlias

End Property

Public Property Let Alias(ByVal Value As String)

    mAlias = Value

End Property

Public Property Get ObjectNumber() As Long

    ObjectNumber = mObjectNumber

End Property

Public Property Let ObjectNumber(ByVal Value As Long)

    mObjectNumber = Value

End Property
