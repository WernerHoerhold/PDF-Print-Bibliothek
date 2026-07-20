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

'=========================================================
' Klasse.....: PdfContentBuilder
' Aufgabe....: Erzeugt den PDF-Content-Stream
'
' Version....: 0.6.0
'=========================================================

Dim F As New PdfFormat
Dim Col As New PdfColor
Dim Geo As New PdfGeometry

Public Function Build(ByVal Canvas As PdfCanvas2D) As String

    Dim Cmd As Object
    Dim S As String

    For Each Cmd In Canvas.Commands

        Select Case TypeName(Cmd)

            Case "PdfCmdLine"

                S = S _
                    & BuildPenCommands(Cmd.State.Pen) _
                    & EnsureCrLf(Geo.BuildLinePath( _
                                        Cmd.X1, Cmd.Y1, _
                                        Cmd.X2, Cmd.Y2)) _
                    & "S" & vbCrLf
  
            Case "PdfCmdText"
            
                S = S _
                    & Col.ToPdfRGB(Cmd.State.Text.Color) & " rg" & vbCrLf _
                    & "BT" & vbCrLf _
                    & "/" & Cmd.State.Text.Font.Alias & " " & F.Number(Cmd.State.Text.Size) & " Tf" & vbCrLf _
                    & F.Number(Cmd.X) & " " _
                    & F.Number(Cmd.Y) & " Td" & vbCrLf _
                    & "(" & EscapeText(Cmd.Text) & ") Tj" & vbCrLf _
                    & "ET" & vbCrLf
                    
            Case "PdfCmdRectangle"
            
                S = S _
                    & BuildPenCommands(Cmd.State.Pen) _
                    & Col.ToPdfRGB(Cmd.State.Brush.Color) & " rg" & vbCrLf _
                    & EnsureCrLf(Geo.BuildRectanglePath( _
                                            Cmd.X, _
                                            Cmd.Y, _
                                            Cmd.Width, _
                                            Cmd.Height))
                
                If Cmd.Stroke And Cmd.Fill Then
                    S = S & "B" & vbCrLf
                ElseIf Cmd.Stroke Then
                    S = S & "S" & vbCrLf
                ElseIf Cmd.Fill Then
                    S = S & "f" & vbCrLf
                End If
                    
            Case "PdfCmdEllipse"
            
            
                S = S _
                    & BuildPenCommands(Cmd.State.Pen) _
                    & Col.ToPdfRGB(Cmd.State.Brush.Color) & " rg" & vbCrLf _
                    & EnsureCrLf(Geo.BuildEllipsePath( _
                                            Cmd.X, _
                                            Cmd.Y, _
                                            Cmd.Width, _
                                            Cmd.Height))
            
                If Cmd.Stroke And Cmd.Fill Then
                    S = S & "B" & vbCrLf
                ElseIf Cmd.Stroke Then
                    S = S & "S" & vbCrLf
                ElseIf Cmd.Fill Then
                    S = S & "f" & vbCrLf
                End If
                    
            Case "PdfCmdPolyline"
            
                    
                S = S _
                    & BuildPenCommands(Cmd.State.Pen) _
                    & EnsureCrLf(Geo.BuildPolylinePath(Cmd.Points)) _
                    & "S" & vbCrLf
                    
            Case "PdfCmdPolygon"

                S = S _
                    & BuildPenCommands(Cmd.State.Pen) _
                    & Col.ToPdfRGB(Cmd.State.Brush.Color) & " rg" & vbCrLf _
                    & EnsureCrLf(Geo.BuildPolygonPath(Cmd.Points))
            
                If Cmd.Stroke And Cmd.Fill Then
                    S = S & "B" & vbCrLf
                ElseIf Cmd.Stroke Then
                    S = S & "S" & vbCrLf
                ElseIf Cmd.Fill Then
                    S = S & "f" & vbCrLf
                End If
                    
                    
            Case "PdfCmdImage"

                S = S _
                    & "q" & vbCrLf _
                    & F.Matrix( _
                        Cmd.Width, _
                        0, _
                        0, _
                        Cmd.Height, _
                        Cmd.X, _
                        Cmd.Y) _
                    & " cm" & vbCrLf _
                    & "/" & Cmd.Image.Alias & " Do" & vbCrLf _
                    & "Q" & vbCrLf
                    
        End Select

    Next

    Build = S

End Function


'---------------------------------------------------------
' Text für PDF maskieren
'---------------------------------------------------------
Private Function EscapeText(ByVal Text As String) As String

    Text = Replace$(Text, "\", "\\")
    Text = Replace$(Text, "(", "\(")
    Text = Replace$(Text, ")", "\)")

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

    'Linienbreite
    S = S & F.Number(Pen.Width) & " w" & vbCrLf

    'LineCap
    S = S & F.Number(Pen.LineCap) & " J" & vbCrLf

    'LineJoin
    S = S & F.Number(Pen.LineJoin) & " j" & vbCrLf

    'MiterLimit
    S = S & F.Number(Pen.MiterLimit) & " M" & vbCrLf

    'DashPattern
    S = S & "["

    For I = 0 To Pen.DashCount - 1

        If I > 0 Then S = S & " "

        S = S & F.Number(Pen.DashValue(I))

    Next

    S = S & "] "

    S = S & F.Number(Pen.DashPhase)

    S = S & " d" & vbCrLf

    'Farbe
    S = S & Col.ToPdfRGB(Pen.Color) & " RG" & vbCrLf

    BuildPenCommands = S

End Function
