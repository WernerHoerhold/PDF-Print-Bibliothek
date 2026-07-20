VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfPen
' Aufgabe....: Beschreibt den Zeichenstift
'
' Version....: 0.9.0
'=========================================================

Private mColor As OLE_COLOR
Private mWidth As Double
Private mLineCap As Long
Private mLineJoin As Long
Private mMiterLimit As Double
Private mDashArray() As Double
Private mDashCount As Long
Private mDashPhase As Double

Private Sub Class_Initialize()

    mColor = vbBlack
    mWidth = 1
    mLineCap = 0
    mLineJoin = 0
    mMiterLimit = 10
    mDashCount = 0
    mDashPhase = 0
    
End Sub

'---------------------------------------------------------
' Linienfarbe
'---------------------------------------------------------
Public Property Get Color() As OLE_COLOR

    Color = mColor

End Property

Public Property Let Color(ByVal Value As OLE_COLOR)

    mColor = Value

End Property

'---------------------------------------------------------
' Linienstärke
'---------------------------------------------------------
Public Property Get Width() As Double

    Width = mWidth

End Property

Public Property Let Width(ByVal Value As Double)

    If Value <= 0 Then
        Value = 1
    End If

    mWidth = Value

End Property

Public Property Get LineCap() As Long

    LineCap = mLineCap

End Property

Public Property Let LineCap(ByVal Value As Long)

    If Value < 0 Then Value = 0
    If Value > 2 Then Value = 2

    mLineCap = Value

End Property

Public Property Get LineJoin() As Long

    LineJoin = mLineJoin

End Property

Public Property Let LineJoin(ByVal Value As Long)

    If Value < 0 Then Value = 0
    If Value > 2 Then Value = 2

    mLineJoin = Value

End Property

Public Property Get MiterLimit() As Double

    MiterLimit = mMiterLimit

End Property

Public Property Let MiterLimit(ByVal Value As Double)

    If Value < 1 Then Value = 1

    mMiterLimit = Value

End Property

Public Property Get DashPhase() As Double

    DashPhase = mDashPhase

End Property

Public Property Let DashPhase(ByVal Value As Double)

    If Value < 0 Then Value = 0

    mDashPhase = Value

End Property

Public Sub SetDashPattern(ByVal Pattern As Variant)

    Dim I As Long

    If IsEmpty(Pattern) Then

        mDashCount = 0
        Exit Sub

    End If

    ReDim mDashArray(LBound(Pattern) To UBound(Pattern))

    For I = LBound(Pattern) To UBound(Pattern)

        mDashArray(I) = CDbl(Pattern(I))

    Next

    mDashCount = UBound(Pattern) - LBound(Pattern) + 1

End Sub

Public Sub ClearDashPattern()

    mDashCount = 0

End Sub

Public Sub Solid()

    ClearDashPattern

End Sub

Public Sub Dash()

    SetDashPattern Array(3 * Width, 3 * Width)

End Sub

Public Sub Dot()

    SetDashPattern Array(1, 6 * Width)

End Sub

Public Sub DashDot()

    SetDashPattern Array(3 * Width, 3 * Width, 1, 3 * Width)

End Sub

Public Sub DashDotDot()

    SetDashPattern Array(3 * Width, 3 * Width, 1, 3 * Width, 1, 3 * Width)

End Sub

Public Property Get DashCount() As Long

    DashCount = mDashCount

End Property

Public Property Get DashValue(ByVal Index As Long) As Double

    DashValue = mDashArray(Index)

End Property

Public Sub Assign(ByVal Source As PdfPen)

    Dim I As Long

    Width = Source.Width
    Color = Source.Color

    LineCap = Source.LineCap
    LineJoin = Source.LineJoin
    MiterLimit = Source.MiterLimit

    DashPhase = Source.DashPhase

    If Source.DashCount = 0 Then

        ClearDashPattern

    Else

        ReDim mDashArray(0 To Source.DashCount - 1)

        For I = 0 To Source.DashCount - 1

            mDashArray(I) = Source.DashValue(I)

        Next

        mDashCount = Source.DashCount

    End If

End Sub
