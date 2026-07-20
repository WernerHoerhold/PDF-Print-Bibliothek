VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfGeometry
' Aufgabe....: Geometrische Hilfsfunktionen
'
' Version....: 0.9.3
'=========================================================

' Bézier-Konstante für Kreis/Ellipse
Private Const Kappa As Double = 0.552284749830794
Private mFormat As PdfFormat

Private Sub Class_Initialize()

    Set mFormat = New PdfFormat

End Sub

Private Function K() As Double

    K = Kappa

End Function

Private Function F(ByVal Value As Double) As String

    F = mFormat.Number(Value)

End Function

Public Function BuildEllipsePath( _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Width As Double, _
        ByVal Height As Double) As String

    Dim RX As Double
    Dim RY As Double

    Dim CX As Double
    Dim CY As Double

    Dim DX As Double
    Dim DY As Double

    Dim S As String

    RX = Width / 2#
    RY = Height / 2#

    CX = X + RX
    CY = Y + RY

    DX = RX * K()
    DY = RY * K()

    '------------------------------------------------------
    ' Startpunkt
    '------------------------------------------------------
    S = _
        F(CX + RX) & " " & _
        F(CY) & " m" & vbCrLf

    '------------------------------------------------------
    ' Rechts -> Oben
    '------------------------------------------------------
    S = S & _
        F(CX + RX) & " " & F(CY + DY) & " " & _
        F(CX + DX) & " " & F(CY + RY) & " " & _
        F(CX) & " " & F(CY + RY) & " c" & vbCrLf

    '------------------------------------------------------
    ' Oben -> Links
    '------------------------------------------------------
    S = S & _
        F(CX - DX) & " " & F(CY + RY) & " " & _
        F(CX - RX) & " " & F(CY + DY) & " " & _
        F(CX - RX) & " " & F(CY) & " c" & vbCrLf

    '------------------------------------------------------
    ' Links -> Unten
    '------------------------------------------------------
    S = S & _
        F(CX - RX) & " " & F(CY - DY) & " " & _
        F(CX - DX) & " " & F(CY - RY) & " " & _
        F(CX) & " " & F(CY - RY) & " c" & vbCrLf

    '------------------------------------------------------
    ' Unten -> Rechts
    '------------------------------------------------------
    S = S & _
        F(CX + DX) & " " & F(CY - RY) & " " & _
        F(CX + RX) & " " & F(CY - DY) & " " & _
        F(CX + RX) & " " & F(CY) & " c" & vbCrLf

    BuildEllipsePath = S

End Function

Public Function BuildLinePath( _
        ByVal X1 As Double, _
        ByVal Y1 As Double, _
        ByVal X2 As Double, _
        ByVal Y2 As Double) As String

    BuildLinePath = _
        F(X1) & " " & F(Y1) & " m" & vbCrLf & _
        F(X2) & " " & F(Y2) & " l"

End Function

Public Function BuildRectanglePath( _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Width As Double, _
        ByVal Height As Double) As String

    BuildRectanglePath = _
        F(X) & " " & _
        F(Y) & " " & _
        F(Width) & " " & _
        F(Height) & " re"

End Function

Public Function BuildPolylinePath( _
        ByVal Points As PdfPointCollection) As String

    Dim I As Long
    Dim Pt As PdfPoint
    Dim S As String

    If Points.Count = 0 Then Exit Function

    Set Pt = Points.Item(1)

    S = F(Pt.X) & " " & F(Pt.Y) & " m" & vbCrLf

    For I = 2 To Points.Count

        Set Pt = Points.Item(I)

        S = S _
          & F(Pt.X) & " " _
          & F(Pt.Y) & " l" _
          & vbCrLf

    Next

    BuildPolylinePath = S

End Function

Public Function BuildPolygonPath( _
        ByVal Points As PdfPointCollection) As String

    Dim S As String

    S = BuildPolylinePath(Points)

    If Len(S) > 0 Then

        S = S & "h" & vbCrLf

    End If

    BuildPolygonPath = S

End Function

