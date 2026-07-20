VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfGraphicsState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfGraphicsState
' Aufgabe....: Aktueller Zeichenzustand
'
' Version....: 0.9.0
'=========================================================

Private mPen As PdfPen
Private mText As PdfTextState
Private mBrush As PdfBrush

Private Sub Class_Initialize()

    Set mPen = New PdfPen
    Set mText = New PdfTextState
    Set mBrush = New PdfBrush

End Sub

'---------------------------------------------------------
' Zeichenstift
'---------------------------------------------------------
Public Property Get Pen() As PdfPen

    Set Pen = mPen

End Property

'---------------------------------------------------------
' Zustand kopieren
'---------------------------------------------------------
Public Function Clone() As PdfGraphicsState

    Dim GS As PdfGraphicsState

    Set GS = New PdfGraphicsState

    GS.Pen.Assign mPen
    
    Set GS.Text.Font = mText.Font
    
    GS.Text.Size = mText.Size
    GS.Text.Color = mText.Color
    GS.Brush.Color = mBrush.Color

    Set Clone = GS

End Function

'---------------------------------------------------------
' Textzustand
'---------------------------------------------------------
Public Property Get Text() As PdfTextState

    Set Text = mText

End Property

'---------------------------------------------------------
' Füllpinsel
'---------------------------------------------------------
Public Property Get Brush() As PdfBrush

    Set Brush = mBrush

End Property
