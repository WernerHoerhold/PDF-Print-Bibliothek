VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfColor
' Aufgabe....: Umrechnung von VB-Farben nach PDF-RGB
'
' Version....: 0.9.0
'=========================================================

Private mFormat As PdfFormat

Private Sub Class_Initialize()

    Set mFormat = New PdfFormat

End Sub

'---------------------------------------------------------
' Erzeugt den PDF-RGB-String
' Beispiel:
' vbRed -> "1 0 0"
'---------------------------------------------------------
Public Function ToPdfRGB(ByVal Color As OLE_COLOR) As String

    Dim R As Long
    Dim G As Long
    Dim B As Long

    ' VB speichert Farben als BGR
    R = Color Mod 256
    G = (Color \ 256) Mod 256
    B = (Color \ 65536) Mod 256

    ToPdfRGB = _
        mFormat.Number(R / 255#) & " " & _
        mFormat.Number(G / 255#) & " " & _
        mFormat.Number(B / 255#)

End Function
