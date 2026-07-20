VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfPoint
' Aufgabe....: Beschreibt einen zweidimensionalen Punkt
'
' Version....: 0.9.4
'=========================================================

Private mX As Double
Private mY As Double

'---------------------------------------------------------
' X-Koordinate
'---------------------------------------------------------
Public Property Get X() As Double

    X = mX

End Property

Public Property Let X(ByVal Value As Double)

    mX = Value

End Property

'---------------------------------------------------------
' Y-Koordinate
'---------------------------------------------------------
Public Property Get Y() As Double

    Y = mY

End Property

Public Property Let Y(ByVal Value As Double)

    mY = Value

End Property

'---------------------------------------------------------
' Koordinaten setzen
'---------------------------------------------------------
Public Sub Assign( _
        ByVal X As Double, _
        ByVal Y As Double)

    mX = X
    mY = Y

End Sub

'---------------------------------------------------------
' Punkt kopieren
'---------------------------------------------------------
Public Function Clone() As PdfPoint

    Dim P As PdfPoint

    Set P = New PdfPoint

    P.Assign mX, mY

    Set Clone = P

End Function
