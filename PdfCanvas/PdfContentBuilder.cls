VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfContentBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================
' KLASSE......: PdfContentBuilder
' KATEGORIE...: A - HAUPTKLASSE (Content-Stream Erstellung)
' VERSION.....: 0.6.0
'
' BESCHREIBUNG:
' Erstellt den PDF-Content-Stream aus den Zeichenbefehlen.
' Diese Klasse wandelt die Canvas-Befehle (DrawLine, DrawText, etc.)
' in PDF-Operatoren um.
'
' PDF-OPERATOREN:
' - "m" / "l" / "S": Linien (moveto, lineto, stroke)
' - "BT" / "ET": Text-Block (begintext, endtext)
' - "Tj": Text ausgeben
' - "re" / "f" / "S" / "B": Rechtecke (rectangle, fill, stroke, fillstroke)
' - "Do": Bild ausgeben
'
' VERWENDUNG: Intern durch PdfWriter
'===============================================================

Dim F As New PdfFormat
Dim Col As New PdfColor
Dim Geo As New PdfGeometry

'---------------------------------------------------------------
' METHODE: Build
' PARAMETER: Canvas (PdfCanvas2D) - Die zu konvertierendde Zeichenfläche
' RÜCKGABE: String - PDF-Content-Stream
' BESCHREIBUNG: Konvertiert alle Canvas-Befehle in PDF-Operatoren
'---------------------------------------------------------------
Public Function Build(ByVal Canvas As PdfCanvas2D) As String
    Dim Cmd As Object
    Dim S As String
    
    For Each Cmd In Canvas.Commands
        Select Case TypeName(Cmd)
            Case "PdfCmdLine"
                S = S & BuildPenCommands(Cmd.State.Pen) & _
                    EnsureCrLf(Geo.BuildLinePath(Cmd.X1, Cmd.Y1, Cmd.X2, Cmd.Y2)) & "S" & vbCrLf
            
            Case "PdfCmdText"
                S = S & Col.ToPdfRGB(Cmd.State.Text.Color) & " rg" & vbCrLf & _
                    "BT" & vbCrLf & "/" & Cmd.State.Text.Font.Alias & " " & _
                    F.Number(Cmd.State.Text.Size) & " Tf" & vbCrLf & _
                    F.Number(Cmd.X) & " " & F.Number(Cmd.Y) & " Td" & vbCrLf & _
                    "(" & EscapeText(Cmd.Text) & ") Tj" & vbCrLf & "ET" & vbCrLf
            
            Case "PdfCmdRectangle"
                S = S & BuildPenCommands(Cmd.State.Pen) & _
                    Col.ToPdfRGB(Cmd.State.Brush.Color) & " rg" & vbCrLf & _
                    EnsureCrLf(Geo.BuildRectanglePath(Cmd.X, Cmd.Y, Cmd.Width, Cmd.Height))
                If Cmd.Stroke And Cmd.Fill Then
                    S = S & "B" & vbCrLf
                ElseIf Cmd.Stroke Then
                    S = S & "S" & vbCrLf
                ElseIf Cmd.Fill Then
                    S = S & "f" & vbCrLf
                End If
            
            Case "PdfCmdEllipse"
                S = S & BuildPenCommands(Cmd.State.Pen) & _
                    Col.ToPdfRGB(Cmd.State.Brush.Color) & " rg" & vbCrLf & _
                    EnsureCrLf(Geo.BuildEllipsePath(Cmd.X, Cmd.Y, Cmd.Width, Cmd.Height))
                If Cmd.Stroke And Cmd.Fill Then
                    S = S & "B" & vbCrLf
                ElseIf Cmd.Stroke Then
                    S = S & "S" & vbCrLf
                ElseIf Cmd.Fill Then
                    S = S & "f" & vbCrLf
                End If
            
            Case "PdfCmdPolyline"
                S = S & BuildPenCommands(Cmd.State.Pen) & _
                    EnsureCrLf(Geo.BuildPolylinePath(Cmd.Points)) & "S" & vbCrLf
            
            Case "PdfCmdPolygon"
                S = S & BuildPenCommands(Cmd.State.Pen) & _
                    Col.ToPdfRGB(Cmd.State.Brush.Color) & " rg" & vbCrLf & _
                    EnsureCrLf(Geo.BuildPolygonPath(Cmd.Points))
                If Cmd.Stroke And Cmd.Fill Then
                    S = S & "B" & vbCrLf
                ElseIf Cmd.Stroke Then
                    S = S & "S" & vbCrLf
                ElseIf Cmd.Fill Then
                    S = S & "f" & vbCrLf
                End If
            
            Case "PdfCmdImage"
                S = S & "q" & vbCrLf & _
                    F.Matrix(Cmd.Width, 0, 0, Cmd.Height, Cmd.X, Cmd.Y) & " cm" & vbCrLf & _
                    "/" & Cmd.Image.Alias & " Do" & vbCrLf & "Q" & vbCrLf
        End Select
    Next
    
    Build = S
End Function

Private Function EscapeText(ByVal Text As String) As String
    Text = Replace$(Text, "\\", "\\\\")
    Text = Replace$(Text, "(", "\\(")
    Text = Replace$(Text, ")", "\\)")
    EscapeText = Text
End Function

Private Function EnsureCrLf(ByVal Text As String) As String
    If Len(Text) = 0 Then
        Exit Function
    End If
    If Right$(Text, 2) <> vbCrLf Then
        Text = Text & vbCrLf
    End If
    EnsureCrLf = Text
End Function

Private Function BuildPenCommands(ByVal Pen As PdfPen) As String
    Dim F As New PdfFormat
    Dim S As String
    Dim I As Long
    
    S = ""
    S = S & F.Number(Pen.Width) & " w" & vbCrLf
    S = S & F.Number(Pen.LineCap) & " J" & vbCrLf
    S = S & F.Number(Pen.LineJoin) & " j" & vbCrLf
    S = S & F.Number(Pen.MiterLimit) & " M" & vbCrLf
    S = S & "["
    For I = 0 To Pen.DashCount - 1
        If I > 0 Then S = S & " "
        S = S & F.Number(Pen.DashValue(I))
    Next
    S = S & "] " & F.Number(Pen.DashPhase) & " d" & vbCrLf
    S = S & Col.ToPdfRGB(Pen.Color) & " RG" & vbCrLf
    
    BuildPenCommands = S
End Function
