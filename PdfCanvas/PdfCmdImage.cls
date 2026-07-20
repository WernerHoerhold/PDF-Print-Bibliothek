VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfCmdImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfCmdImage
' Aufgabe....: Bild auf einer PDF-Seite platzieren
'
' Version....: 0.9.5
'=========================================================

Private mImage As PdfImage

Private mX As Double
Private mY As Double

Private mWidth As Double
Private mHeight As Double

'---------------------------------------------------------
' Bild
'---------------------------------------------------------
Public Property Get Image() As PdfImage

    Set Image = mImage

End Property

Public Property Set Image(ByVal Value As PdfImage)

    Set mImage = Value

End Property

'---------------------------------------------------------
' X-Position
'---------------------------------------------------------
Public Property Get X() As Double

    X = mX

End Property

Public Property Let X(ByVal Value As Double)

    mX = Value

End Property

'---------------------------------------------------------
' Y-Position
'---------------------------------------------------------
Public Property Get Y() As Double

    Y = mY

End Property

Public Property Let Y(ByVal Value As Double)

    mY = Value

End Property

'---------------------------------------------------------
' Breite
'---------------------------------------------------------
Public Property Get Width() As Double

    Width = mWidth

End Property

Public Property Let Width(ByVal Value As Double)

    mWidth = Value

End Property

'---------------------------------------------------------
' Höhe
'---------------------------------------------------------
Public Property Get Height() As Double

    Height = mHeight

End Property

Public Property Let Height(ByVal Value As Double)

    mHeight = Value

End Property
