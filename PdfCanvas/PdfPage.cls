VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfPage
' Version....: 1.0
'
' Beschreibung
' -------------
' Repräsentiert eine einzelne PDF-Seite.
'=========================================================

Private mWidth As Long
Private mHeight As Long

Private mCanvas As PdfCanvas2D
Private mPageSize As Long
Private mOrientation As Long

Private Sub Class_Initialize()

    'A4 in PostScript-Punkten (72 dpi)
'    mWidth = 595
'    mHeight = 842
    mPageSize = psA4
    mOrientation = poPortrait

    UpdatePageSize

    Set mCanvas = New PdfCanvas2D

End Sub

'---------------------------------------------------------
' Seitenbreite
'---------------------------------------------------------
Public Property Get Width() As Long

    Width = mWidth

End Property

Public Property Let Width(ByVal Value As Long)

    mWidth = Value

End Property

'---------------------------------------------------------
' Seitenhöhe
'---------------------------------------------------------
Public Property Get Height() As Long

    Height = mHeight

End Property

Public Property Let Height(ByVal Value As Long)

    mHeight = Value

End Property

'---------------------------------------------------------
' Canvas
'---------------------------------------------------------
Public Property Get Canvas() As PdfCanvas2D

    Set Canvas = mCanvas

End Property

Public Property Get PageSize() As Long

    PageSize = mPageSize

End Property

Public Property Let PageSize(ByVal Value As Long)

    mPageSize = Value

    UpdatePageSize

End Property

Public Property Get Orientation() As Long

    Orientation = mOrientation

End Property

Public Property Let Orientation(ByVal Value As Long)

    mOrientation = Value

    UpdatePageSize

End Property

Private Sub UpdatePageSize()

    Select Case mPageSize

        Case psA5

            mWidth = 420
            mHeight = 595

        Case psA4

            mWidth = 595
            mHeight = 842

        Case psA3

            mWidth = 842
            mHeight = 1191

        Case psLetter

            mWidth = 612
            mHeight = 792

        Case psLegal

            mWidth = 612
            mHeight = 1008

    End Select

    If mOrientation = poLandscape Then

        Dim T As Long

        T = mWidth
        mWidth = mHeight
        mHeight = T

    End If

End Sub
