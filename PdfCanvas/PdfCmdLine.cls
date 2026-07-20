VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfCmdLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfCmdLine
' Aufgabe....: Linienbefehl
'=========================================================

Private mX1 As Double
Private mY1 As Double
Private mX2 As Double
Private mY2 As Double
Private mState As PdfGraphicsState

Public Property Get X1() As Double
    X1 = mX1
End Property

Public Property Let X1(ByVal Value As Double)
    mX1 = Value
End Property

Public Property Get Y1() As Double
    Y1 = mY1
End Property

Public Property Let Y1(ByVal Value As Double)
    mY1 = Value
End Property

Public Property Get X2() As Double
    X2 = mX2
End Property

Public Property Let X2(ByVal Value As Double)
    mX2 = Value
End Property

Public Property Get Y2() As Double
    Y2 = mY2
End Property

Public Property Let Y2(ByVal Value As Double)
    mY2 = Value
End Property

'---------------------------------------------------------
' Zeichenzustand
'---------------------------------------------------------
Public Property Get State() As PdfGraphicsState

    Set State = mState

End Property

Public Property Set State(ByVal Value As PdfGraphicsState)

    Set mState = Value

End Property
