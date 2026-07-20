VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfCmdEllipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mX As Double
Private mY As Double
Private mWidth As Double
Private mHeight As Double

Private mStroke As Boolean
Private mFill As Boolean

Private mState As PdfGraphicsState

'----------------------------
' Position
'----------------------------
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

'----------------------------
' Größe
'----------------------------
Public Property Get Width() As Double
    Width = mWidth
End Property

Public Property Let Width(ByVal Value As Double)
    mWidth = Value
End Property

Public Property Get Height() As Double
    Height = mHeight
End Property

Public Property Let Height(ByVal Value As Double)
    mHeight = Value
End Property

'----------------------------
' Zeichnen
'----------------------------
Public Property Get Stroke() As Boolean
    Stroke = mStroke
End Property

Public Property Let Stroke(ByVal Value As Boolean)
    mStroke = Value
End Property

'----------------------------
' Füllen
'----------------------------
Public Property Get Fill() As Boolean
    Fill = mFill
End Property

Public Property Let Fill(ByVal Value As Boolean)
    mFill = Value
End Property

'----------------------------
' Graphics State
'----------------------------
Public Property Get State() As PdfGraphicsState

    Set State = mState

End Property

Public Property Set State(ByVal Value As PdfGraphicsState)

    Set mState = Value

End Property

