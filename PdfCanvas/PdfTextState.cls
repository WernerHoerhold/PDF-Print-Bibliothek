VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfTextState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfTextState
' Aufgabe....: Beschreibt den aktuellen Textzustand
'
' Version....: 0.9.1
'=========================================================

Private mFont As PdfFont
Private mSize As Double
Private mColor As OLE_COLOR

Private Sub Class_Initialize()

    'Set mFont = New PdfFont
    'mFont.Name = "Helvetica"
    'mFont.Alias = "F1"

    Set mFont = Nothing

    mSize = 12
    mColor = vbBlack

End Sub

'---------------------------------------------------------
' Schrift
'---------------------------------------------------------
Public Property Get Font() As PdfFont

    Set Font = mFont

End Property

Public Property Set Font(ByVal Value As PdfFont)

    Set mFont = Value

End Property


'---------------------------------------------------------
' Schriftgröße
'---------------------------------------------------------
Public Property Get Size() As Double

    Size = mSize

End Property

Public Property Let Size(ByVal Value As Double)

    If Value < 1 Then Value = 1

    mSize = Value

End Property

'---------------------------------------------------------
' Textfarbe
'---------------------------------------------------------
Public Property Get Color() As OLE_COLOR

    Color = mColor

End Property

Public Property Let Color(ByVal Value As OLE_COLOR)

    mColor = Value

End Property
