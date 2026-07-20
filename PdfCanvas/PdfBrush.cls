VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfBrush
' Aufgabe....: Beschreibt den Füllpinsel
'
' Version....: 0.9.2
'=========================================================

Private mColor As OLE_COLOR

Private Sub Class_Initialize()

    mColor = vbWhite

End Sub

'---------------------------------------------------------
' Füllfarbe
'---------------------------------------------------------
Public Property Get Color() As OLE_COLOR

    Color = mColor

End Property

Public Property Let Color(ByVal Value As OLE_COLOR)

    mColor = Value

End Property
