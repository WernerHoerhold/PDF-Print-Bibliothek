VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfEncoding
' Aufgabe....: Zeichencodierung für PDF
'
' Version....: 1.1.0
'
' Beschreibung
' -------------
' Diese Klasse übernimmt sämtliche
' Zeichencodierungen der PDF-Bibliothek.
'
' Zur Zeit:
'
'   - WinAnsiEncoding
'   - Escaping von PDF-Strings
'
' Später:
'
'   - UTF-16
'   - Unicode
'   - TrueType Fonts
'=========================================================



'---------------------------------------------------------
' Wandelt einen VB-String in einen
' PDF-konformen WinAnsi-String um.
'---------------------------------------------------------
Public Function EncodePdfString( _
                ByVal Text As String) As String

    'Text = EncodeWinAnsi(Text)

    Text = EscapePdfString(Text)

    EncodePdfString = Text

End Function



'---------------------------------------------------------
' Unicode -> WinAnsi
'
' VB6 arbeitet intern mit Unicode.
' PDF Standardfonts erwarten WinAnsi.
'---------------------------------------------------------
'Public Function EncodeWinAnsi( _
'                ByVal Text As String) As String
'
'    EncodeWinAnsi = StrConv(Text, vbFromUnicode)
'
'End Function



'---------------------------------------------------------
' Maskiert Sonderzeichen innerhalb
' eines PDF-Strings.
'
' PDF Reference:
'
'     \      ->  \\
'     (      ->  \(
'     )      ->  \)
'
'---------------------------------------------------------
Public Function EscapePdfString( _
                ByVal Text As String) As String

    Text = Replace$(Text, "\", "\\")

    Text = Replace$(Text, "(", "\(")

    Text = Replace$(Text, ")", "\)")

    EscapePdfString = Text

End Function

