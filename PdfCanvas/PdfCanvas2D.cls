VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfCanvas2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfCanvas2D
' Version....: 1.0
'
' Aufgabe:
' Sammelt alle Zeichenbefehle einer Seite.
'=========================================================

Private mCommands As Collection
Private mState As PdfGraphicsState

Private Sub Class_Initialize()

    Set mCommands = New Collection
    Set mState = New PdfGraphicsState
    
End Sub

'---------------------------------------------------------
' Linie
'---------------------------------------------------------
Public Sub DrawLine(ByVal X1 As Double, _
                    ByVal Y1 As Double, _
                    ByVal X2 As Double, _
                    ByVal Y2 As Double)

    Dim Cmd As PdfCmdLine

    Set Cmd = New PdfCmdLine
    Set Cmd.State = mState.Clone
    
    Cmd.X1 = X1
    Cmd.Y1 = Y1

    Cmd.X2 = X2
    Cmd.Y2 = Y2

    mCommands.Add Cmd

End Sub


'---------------------------------------------------------
' Anzahl der Befehle
'---------------------------------------------------------
Public Property Get CommandCount() As Long

    CommandCount = mCommands.Count

End Property

'---------------------------------------------------------
' Zugriff auf einen Zeichenbefehl
'---------------------------------------------------------
Public Property Get Command(ByVal Index As Long) As Object

    Set Command = mCommands(Index)

End Property
'---------------------------------------------------------
' Alle Zeichenbefehle
'---------------------------------------------------------
Public Property Get Commands() As Collection

    Set Commands = mCommands

End Property

'---------------------------------------------------------
' Text zeichnen
'---------------------------------------------------------
Public Sub DrawText( _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Text As String)

    Dim Cmd As PdfCmdText

    Set Cmd = New PdfCmdText
    Set Cmd.State = mState.Clone

    Cmd.X = X
    Cmd.Y = Y
    Cmd.Text = Text

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Aktueller Zeichenzustand
'---------------------------------------------------------
Friend Property Get State() As PdfGraphicsState

    Set State = mState

End Property

'---------------------------------------------------------
' Zeichenstift
'---------------------------------------------------------
Public Property Get Pen() As PdfPen

    Set Pen = mState.Pen

End Property

'---------------------------------------------------------
' Aktueller Textzustand
'---------------------------------------------------------
Friend Property Get Text() As PdfTextState

    Set Text = mState.Text

End Property

'---------------------------------------------------------
' Aktuelle Schrift
'---------------------------------------------------------
Public Property Get Font() As PdfFont

    Set Font = mState.Text.Font

End Property

'---------------------------------------------------------
' Aktuelle Schrift
'---------------------------------------------------------
Public Property Set Font(ByVal Value As PdfFont)

    Set mState.Text.Font = Value

End Property

'---------------------------------------------------------
' Füllpinsel
'---------------------------------------------------------
Public Property Get Brush() As PdfBrush

    Set Brush = mState.Brush

End Property

'---------------------------------------------------------
' Rechteck zeichnen
'---------------------------------------------------------
Public Sub DrawRectangle( _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Width As Double, _
        ByVal Height As Double)

    Dim Cmd As PdfCmdRectangle

    Set Cmd = New PdfCmdRectangle

    Set Cmd.State = mState.Clone

    Cmd.X = X
    Cmd.Y = Y

    Cmd.Width = Width
    Cmd.Height = Height

    Cmd.Stroke = True
    Cmd.Fill = False

    mCommands.Add Cmd

End Sub
'---------------------------------------------------------
' Rechteck füllen
'---------------------------------------------------------
Public Sub FillRectangle( _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Width As Double, _
        ByVal Height As Double)

    Dim Cmd As PdfCmdRectangle

    Set Cmd = New PdfCmdRectangle

    Set Cmd.State = mState.Clone

    Cmd.X = X
    Cmd.Y = Y

    Cmd.Width = Width
    Cmd.Height = Height

    Cmd.Stroke = False
    Cmd.Fill = True

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Ellipse zeichnen
'---------------------------------------------------------
Public Sub DrawEllipse( _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Width As Double, _
        ByVal Height As Double)

    Dim Cmd As PdfCmdEllipse

    Set Cmd = New PdfCmdEllipse

    Set Cmd.State = mState.Clone

    Cmd.X = X
    Cmd.Y = Y

    Cmd.Width = Width
    Cmd.Height = Height

    Cmd.Stroke = True
    Cmd.Fill = False

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Ellipse füllen
'---------------------------------------------------------
Public Sub FillEllipse( _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Width As Double, _
        ByVal Height As Double)

    Dim Cmd As PdfCmdEllipse

    Set Cmd = New PdfCmdEllipse

    Set Cmd.State = mState.Clone

    Cmd.X = X
    Cmd.Y = Y

    Cmd.Width = Width
    Cmd.Height = Height

    Cmd.Stroke = False
    Cmd.Fill = True

    mCommands.Add Cmd

End Sub

Public Sub DrawPolyline( _
        ByVal Points As PdfPointCollection)

    Dim Cmd As PdfCmdPolyline

    Set Cmd = New PdfCmdPolyline

    Set Cmd.State = mState.Clone
    Set Cmd.Points = Points.Clone

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Polygon zeichnen
'---------------------------------------------------------
Public Sub DrawPolygon( _
        ByVal Points As PdfPointCollection)

    Dim Cmd As PdfCmdPolygon

    Set Cmd = New PdfCmdPolygon

    Set Cmd.State = mState.Clone
    Set Cmd.Points = Points.Clone

    Cmd.Stroke = True
    Cmd.Fill = False

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Polygon füllen
'---------------------------------------------------------
Public Sub FillPolygon( _
        ByVal Points As PdfPointCollection)

    Dim Cmd As PdfCmdPolygon

    Set Cmd = New PdfCmdPolygon

    Set Cmd.State = mState.Clone
    Set Cmd.Points = Points.Clone

    Cmd.Stroke = False
    Cmd.Fill = True

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Bild in Originalgröße zeichnen
'---------------------------------------------------------
Public Sub DrawImage( _
        ByVal Image As PdfImage, _
        ByVal X As Double, _
        ByVal Y As Double)

    DrawImageEx Image, _
                X, _
                Y, _
                Image.Width, _
                Image.Height

End Sub

'---------------------------------------------------------
' Bild skaliert zeichnen
'---------------------------------------------------------
Public Sub DrawImageEx( _
        ByVal Image As PdfImage, _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal Width As Double, _
        ByVal Height As Double)

    Dim Cmd As PdfCmdImage

    Set Cmd = New PdfCmdImage

    Set Cmd.Image = Image

    Cmd.X = X
    Cmd.Y = Y

    Cmd.Width = Width
    Cmd.Height = Height

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Bild proportional in ein Rechteck einpassen
'---------------------------------------------------------
Public Sub DrawImageFit( _
        ByVal Image As PdfImage, _
        ByVal X As Double, _
        ByVal Y As Double, _
        ByVal MaxWidth As Double, _
        ByVal MaxHeight As Double)

    Dim Factor As Double

    Factor = FitScale( _
                Image.Width, _
                Image.Height, _
                MaxWidth, _
                MaxHeight)

    DrawImageEx _
        Image, _
        X, _
        Y, _
        Image.Width * Factor, _
        Image.Height * Factor

End Sub


'---------------------------------------------------------
' Skalierungsfaktor für proportionales Einpassen
'---------------------------------------------------------
Private Function FitScale( _
        ByVal SourceWidth As Double, _
        ByVal SourceHeight As Double, _
        ByVal TargetWidth As Double, _
        ByVal TargetHeight As Double) As Double

    Dim ScaleX As Double
    Dim ScaleY As Double

    ScaleX = TargetWidth / SourceWidth
    ScaleY = TargetHeight / SourceHeight

    If ScaleX < ScaleY Then
        FitScale = ScaleX
    Else
        FitScale = ScaleY
    End If

End Function

