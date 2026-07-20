VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfCmdPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfCmdPolygon
' Aufgabe....: Polygon
'=========================================================

Private mPoints As PdfPointCollection

Private mState As PdfGraphicsState

Private mStroke As Boolean
Private mFill As Boolean

Private Sub Class_Initialize()

    Set mPoints = New PdfPointCollection

End Sub

Public Property Get Points() As PdfPointCollection

    Set Points = mPoints

End Property

Public Property Set Points(ByVal Value As PdfPointCollection)

    Set mPoints = Value

End Property

Public Property Get State() As PdfGraphicsState

    Set State = mState

End Property

Public Property Set State(ByVal Value As PdfGraphicsState)

    Set mState = Value

End Property

Public Property Get Stroke() As Boolean

    Stroke = mStroke

End Property

Public Property Let Stroke(ByVal Value As Boolean)

    mStroke = Value

End Property
Public Property Get Fill() As Boolean

    Fill = mFill

End Property

Public Property Let Fill(ByVal Value As Boolean)

    mFill = Value

End Property
