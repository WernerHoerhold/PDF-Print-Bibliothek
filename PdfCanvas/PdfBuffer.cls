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

'===============================================================
' KLASSE......: PdfBuffer
' KATEGORIE...: B - UNTERSTÜTZUNGSKLASSE
' VERSION.....: 0.6.0
'
' BESCHREIBUNG:
' Puffer für den kompletten PDF-Inhalt. Diese Klasse verwaltet
' die PDF-Datei im Speicher als String und bietet Methoden
' zum Hinzufügen von Text und zur Positionsverwaltung.
'
' VERWENDUNG: Intern durch PdfSerializer und PdfWriter
' AUTOR: Kapp Niles GmbH & Co, KG
'===============================================================

Private mText As String

Public Sub Clear()
    mText = ""
End Sub

Public Sub Add(ByVal Text As String)
    mText = mText & Text
End Sub

Public Sub AddLine(ByVal Text As String)
    mText = mText & Text & vbCrLf
End Sub

Public Property Get Text() As String
    Text = mText
End Property

Public Property Get Length() As Long
    Length = LenB(StrConv(mText, vbFromUnicode))
End Property

Public Property Get Position() As Long
    Position = Length
End Property

Public Sub AppendByte(ByVal Value As Byte)
    mText = mText & Chr$(Value)
End Sub
