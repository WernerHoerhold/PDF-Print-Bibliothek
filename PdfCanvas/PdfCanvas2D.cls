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

'===============================================================
' KLASSE......: PdfCanvas2D
' KATEGORIE...: A - HAUPTKLASSE (Zeichenfläche)
' VERSION.....: 1.1 (mit Rotationsunterstützung)
'
' BESCHREIBUNG:
' Die 2D-Zeichenfläche einer PDF-Seite. Sie sammelt alle
' Zeichenbefehle (Linien, Rechtecke, Text, Bilder, etc.)
' und speichert den aktuellen Zeichenzustand (Stift, Pinsel,
' Schrift).
'
' VERWENDUNG:
'   Dim Doc As New PdfDocument
'   Dim Page As PdfPage
'   Set Page = Doc.AddPage()
'   Page.Canvas.DrawLine 50, 50, 200, 200
'   Page.Canvas.DrawText 100, 100, "Hallo!"
'   Page.Canvas.DrawRectangle 50, 50, 100, 50
'
' ABHÄNGIGKEITEN: PdfGraphicsState, alle PdfCmd*-Klassen
' AUTOR: Kapp Niles GmbH & Co, KG
'===============================================================

Private mCommands As Collection
Private mState As PdfGraphicsState

'---------------------------------------------------------------
' EVENT: Class_Initialize
' BESCHREIBUNG: Initialisiert die Canvas mit leerer
'               Befehls-Sammlung und Standard-Zeichenzustand
'---------------------------------------------------------------
Private Sub Class_Initialize()

    Set mCommands = New Collection
    Set mState = New PdfGraphicsState
    
End Sub

'---------------------------------------------------------------
' METHODE: DrawLine
' PARAMETER:
'   - X1, Y1 (Double): Startpunkt
'   - X2, Y2 (Double): Endpunkt
' BESCHREIBUNG: Zeichnet eine Linie vom Start- zum Endpunkt
'               unter Verwendung der aktuellen Stift-Eigenschaften
'---------------------------------------------------------------
Public Sub DrawLine(ByVal X1 As Double, ByVal Y1 As Double, _
                    ByVal X2 As Double, ByVal Y2 As Double)

    Dim Cmd As PdfCmdLine

    Set Cmd = New PdfCmdLine
    Set Cmd.State = mState.Clone
    
    Cmd.X1 = X1
    Cmd.Y1 = Y1

    Cmd.X2 = X2
    Cmd.Y2 = Y2

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------------
' PROPERTY: CommandCount (READ-ONLY)
' RÜCKGABE: Long - Anzahl der Zeichenbefehle
'---------------------------------------------------------------
Public Property Get CommandCount() As Long

    CommandCount = mCommands.Count

End Property

'---------------------------------------------------------------
' PROPERTY: Command(Index)
' PARAMETER: Index (Long) - Index des Befehls (1-basiert)
' RÜCKGABE: Object - Der Befehl an dieser Position
'---------------------------------------------------------------
Public Property Get Command(ByVal Index As Long) As Object

    Set Command = mCommands(Index)

End Property

'---------------------------------------------------------------
' PROPERTY: Commands (READ-ONLY / INTERNAL)
' RÜCKGABE: Collection - Alle Zeichenbefehle
'---------------------------------------------------------------
Public Property Get Commands() As Collection

    Set Commands = mCommands

End Property

'---------------------------------------------------------------
' METHODE: DrawText
' PARAMETER:
'   - X, Y (Double): Position des Texts
'   - Text (String): Der zu zeichnende Text
' BESCHREIBUNG: Zeichnet Text an der angegebenen Position
'               unter Verwendung der aktuellen Font-Eigenschaften
'---------------------------------------------------------------
Public Sub DrawText(ByVal X As Double, ByVal Y As Double, _
                    ByVal Text As String)

    Dim Cmd As PdfCmdText
    
    Set Cmd = New PdfCmdText
    Set Cmd.State = mState.Clone

    Cmd.X = X
    Cmd.Y = Y
    Cmd.Text = Text
    Cmd.Angle = 0

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' METHODE: DrawTextRotated
' PARAMETER:
'   - X, Y (Double): Position des Texts
'   - Text (String): Der zu zeichnende Text
'   - AngleDegrees (Double): Rotationswinkel in Grad
' BESCHREIBUNG: Zeichnet Text an der angegebenen Position
'               mit Rotation um einen beliebigen Winkel.
'               Der Text wird um die Startposition (X, Y) rotiert.
'---------------------------------------------------------
Public Sub DrawTextRotated(ByVal X As Double, ByVal Y As Double, _
                           ByVal Text As String, ByVal AngleDegrees As Double)

    Dim Cmd As PdfCmdText
    
    Set Cmd = New PdfCmdText
    Set Cmd.State = mState.Clone

    Cmd.X = X
    Cmd.Y = Y
    Cmd.Text = Text
    Cmd.Angle = AngleDegrees

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------
' Text rechtsbündig zeichnen
'
' X = rechte Kante des Textes
' Y = Y-Position
'---------------------------------------------------------
Public Sub DrawTextRight( _
                         ByVal X As Double, _
                         ByVal Y As Double, _
                         ByVal Text As String)

    Dim Width As Double

    Width = mState.Text.Font.TextWidth(Text)

    DrawText X - Width, Y, Text

End Sub

'---------------------------------------------------------
' Text zentriert zeichnen
'
' X = Mittelpunkt des Textes
' Y = Baseline
'---------------------------------------------------------
Public Sub DrawTextCenter( _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal Text As String)

    Dim Width As Double

    Width = mState.Text.Font.TextWidth(Text)

    DrawText X - Width / 2#, Y, Text

End Sub

'---------------------------------------------------------------
' PROPERTY: State (READ-ONLY / INTERNAL)
' RÜCKGABE: PdfGraphicsState - Der aktuelle Zeichenzustand
'---------------------------------------------------------------
Friend Property Get State() As PdfGraphicsState

    Set State = mState

End Property

'---------------------------------------------------------------
' PROPERTY: Pen
' RÜCKGABE: PdfPen - Der aktuelle Zeichenstift
' BESCHREIBUNG: Liefert den Stift für Linieneigenschaften
'               (Farbe, Dicke, Linientyp, etc.)
'---------------------------------------------------------------
Public Property Get Pen() As PdfPen

    Set Pen = mState.Pen

End Property

'---------------------------------------------------------------
' PROPERTY: Text (READ-ONLY / INTERNAL)
' RÜCKGABE: PdfTextState - Der aktuelle Text-Zustand
'---------------------------------------------------------------
Public Property Get Text() As PdfTextState

    Set Text = mState.Text

End Property

'---------------------------------------------------------------
' PROPERTY: Font
' RÜCKGABE: PdfFont - Die aktuelle Schrift
' BESCHREIBUNG: Liefert oder setzt die Schrift für
'               Text-Zeichenbefehle
'---------------------------------------------------------------
Public Property Get Font() As PdfFont

    Set Font = mState.Text.Font

End Property

'---------------------------------------------------------
' Aktuelle Schrift
'---------------------------------------------------------
Public Property Set Font(ByVal Value As PdfFont)

    Set mState.Text.Font = Value

End Property

'---------------------------------------------------------------
' PROPERTY: Brush
' RÜCKGABE: PdfBrush - Der aktuelle Füllpinsel
' BESCHREIBUNG: Liefert den Pinsel für Füllfarben
'---------------------------------------------------------------
Public Property Get Brush() As PdfBrush

    Set Brush = mState.Brush

End Property

'---------------------------------------------------------------
' METHODE: DrawRectangle
' PARAMETER: X, Y (Double), Width, Height (Double)
' BESCHREIBUNG: Zeichnet ein leeres Rechteck (nur Umriss)
'---------------------------------------------------------------
Public Sub DrawRectangle(ByVal X As Double, ByVal Y As Double, _
                         ByVal Width As Double, ByVal Height As Double)

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

'---------------------------------------------------------------
' METHODE: FillRectangle
' PARAMETER: X, Y (Double), Width, Height (Double)
' BESCHREIBUNG: Zeichnet ein gefülltes Rechteck
'---------------------------------------------------------------
Public Sub FillRectangle(ByVal X As Double, ByVal Y As Double, _
        ByVal Width As Double, ByVal Height As Double)

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

'---------------------------------------------------------------
' METHODE: DrawEllipse
' PARAMETER: X, Y (Double), Width, Height (Double)
' BESCHREIBUNG: Zeichnet eine leere Ellipse (nur Umriss)
'---------------------------------------------------------------
Public Sub DrawEllipse(ByVal X As Double, ByVal Y As Double, _
                       ByVal Width As Double, ByVal Height As Double)

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

'---------------------------------------------------------------
' METHODE: FillEllipse
' PARAMETER: X, Y (Double), Width, Height (Double)
' BESCHREIBUNG: Zeichnet eine gefüllte Ellipse
'---------------------------------------------------------------
Public Sub FillEllipse(ByVal X As Double, ByVal Y As Double, _
                       ByVal Width As Double, ByVal Height As Double)

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

'---------------------------------------------------------------
' METHODE: DrawPolyline
' PARAMETER: Points (PdfPointCollection) - Punkte der Polylinie
' BESCHREIBUNG: Zeichnet eine Polylinie (offene Kurve)
'---------------------------------------------------------------
Public Sub DrawPolyline(ByVal Points As PdfPointCollection)

    Dim Cmd As PdfCmdPolyline

    Set Cmd = New PdfCmdPolyline

    Set Cmd.State = mState.Clone
    Set Cmd.Points = Points.Clone

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------------
' METHODE: DrawPolygon
' PARAMETER: Points (PdfPointCollection) - Eckpunkte des Polygons
' BESCHREIBUNG: Zeichnet ein leeres Polygon (nur Umriss)
'---------------------------------------------------------------
Public Sub DrawPolygon(ByVal Points As PdfPointCollection)

    Dim Cmd As PdfCmdPolygon

    Set Cmd = New PdfCmdPolygon

    Set Cmd.State = mState.Clone
    Set Cmd.Points = Points.Clone

    Cmd.Stroke = True
    Cmd.Fill = False

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------------
' METHODE: FillPolygon
' PARAMETER: Points (PdfPointCollection) - Eckpunkte des Polygons
' BESCHREIBUNG: Zeichnet ein gefülltes Polygon
'---------------------------------------------------------------
Public Sub FillPolygon(ByVal Points As PdfPointCollection)

    Dim Cmd As PdfCmdPolygon

    Set Cmd = New PdfCmdPolygon

    Set Cmd.State = mState.Clone
    Set Cmd.Points = Points.Clone

    Cmd.Stroke = False
    Cmd.Fill = True

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------------
' METHODE: DrawImage
' PARAMETER:
'   - Image (PdfImage): Das zu zeichnende Bild
'   - X, Y (Double): Position
' BESCHREIBUNG: Zeichnet ein Bild in Originalgröße
'---------------------------------------------------------------
Public Sub DrawImage(ByVal Image As PdfImage, _
                     ByVal X As Double, ByVal Y As Double)

    DrawImageEx Image, _
                X, _
                Y, _
                Image.Width, _
                Image.Height

End Sub

'---------------------------------------------------------------
' METHODE: DrawImageEx
' PARAMETER:
'   - Image (PdfImage): Das zu zeichnende Bild
'   - X, Y (Double): Position
'   - Width, Height (Double): Skalierte Größe
' BESCHREIBUNG: Zeichnet ein Bild mit benutzerdefin. Größe
'---------------------------------------------------------------
Public Sub DrawImageEx(ByVal Image As PdfImage, _
                        ByVal X As Double, ByVal Y As Double, _
                        ByVal Width As Double, ByVal Height As Double)

    Dim Cmd As PdfCmdImage

    Set Cmd = New PdfCmdImage

    Set Cmd.Image = Image

    Cmd.X = X
    Cmd.Y = Y

    Cmd.Width = Width
    Cmd.Height = Height

    mCommands.Add Cmd

End Sub

'---------------------------------------------------------------
' METHODE: DrawImageFit
' PARAMETER:
'   - Image (PdfImage): Das zu zeichnende Bild
'   - X, Y (Double): Position
'   - MaxWidth, MaxHeight (Double): Maximale Größe
' BESCHREIBUNG: Zeichnet ein Bild proportional in ein Rechteck
'---------------------------------------------------------------
Public Sub DrawImageFit(ByVal Image As PdfImage, _
                        ByVal X As Double, ByVal Y As Double, _
                        ByVal MaxWidth As Double, ByVal MaxHeight As Double)

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

'---------------------------------------------------------------
' INTERNER HELPER: FitScale
' BESCHREIBUNG: Berechnet Skalierungsfaktor für proportionales
'               Einpassen eines Bildes in ein Rechteck
'---------------------------------------------------------------
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

