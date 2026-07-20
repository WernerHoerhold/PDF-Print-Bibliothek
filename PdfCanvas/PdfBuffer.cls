VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfBuffer
' Aufgabe....: Puffer für den kompletten PDF-Inhalt
'
' Version....: 0.6.0
'=========================================================

Private mText As String

'---------------------------------------------------------
' Puffer löschen
'---------------------------------------------------------
Public Sub Clear()

    mText = ""

End Sub

'---------------------------------------------------------
' Text anhängen
'---------------------------------------------------------
Public Sub Add(ByVal Text As String)

    mText = mText & Text

End Sub

'---------------------------------------------------------
' Zeile anhängen (CR/LF)
'---------------------------------------------------------
Public Sub AddLine(ByVal Text As String)

    mText = mText & Text & vbCrLf

End Sub

'---------------------------------------------------------
' Gesamten Inhalt liefern
'---------------------------------------------------------
Public Property Get Text() As String

    Text = mText

End Property

'---------------------------------------------------------
' Aktuelle Länge
'---------------------------------------------------------
Public Property Get Length() As Long

    Length = LenB(StrConv(mText, vbFromUnicode))

End Property

'---------------------------------------------------------
' Aktuelle Byteposition
'---------------------------------------------------------
Public Property Get Position() As Long

    Position = Length

End Property

'---------------------------------------------------------
' Ein Byte an den Puffer anhängen
'---------------------------------------------------------
Public Sub AppendByte(ByVal Value As Byte)

    mText = mText & Chr$(Value)

End Sub
