VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfCmdText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfCmdText
' Aufgabe....: Textbefehl
'
' Version....: 0.8.0
'=========================================================

Private mX As Double
Private mY As Double
Private mText As String
Private mState As PdfGraphicsState

Public Property Get X() As Double
    X = mX
End Property

Public Property Let X(ByVal Value As Double)
    mX = Value
End Property

Public Property Get Y() As Double
    Y = mY
End Property

Public Property Let Y(ByVal Value As Double)
    mY = Value
End Property

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByVal Value As String)
    mText = Value
End Property

Public Property Get State() As PdfGraphicsState

    Set State = mState

End Property

Public Property Set State(ByVal Value As PdfGraphicsState)

    Set mState = Value

End Property
