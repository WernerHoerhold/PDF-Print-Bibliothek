VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfFont
' Aufgabe....: Beschreibt eine PDF-Schriftart
'
' Version....: 0.8.0
'=========================================================

Private mFamily As PdfStandardFont
Private mAlias As String
Private mSize As Double
Private mObjectNumber As Long
Private mBold As Boolean
Private mItalic As Boolean
Private mUnderline As Boolean
Private mStrikeOut As Boolean

Public Sub Assign(ByVal Source As PdfFont)

    mFamily = Source.Family
    mAlias = Source.Alias
    mSize = Source.Size

    mBold = Source.Bold
    mItalic = Source.Italic

End Sub

'---------------------------------------------------------
' Schriftfamilie
'---------------------------------------------------------
Public Property Get Family() As Long

    Family = CLng(mFamily)

End Property

Public Property Let Family(ByVal Value As Long)

    Select Case Value
        Case psfHelvetica, _
             psfTimesRoman, _
             psfCourier, _
             psfSymbol, _
             psfZapfDingbats, _
             psfArial, _
             psfTimesNewRoman, _
             psfCourierNew

             mFamily = Value
   
        Case Else
             mFamily = psfHelvetica ' Standardfont = Helvetica
   End Select

End Property

'---------------------------------------------------------
' Ressourcenname
' z.B. F1
'---------------------------------------------------------
Public Property Get Alias() As String

    Alias = mAlias

End Property

Public Property Let Alias(ByVal Value As String)

    mAlias = Value

End Property

'---------------------------------------------------------
' Schriftgröße
'---------------------------------------------------------
Public Property Get Size() As Double

    Size = mSize

End Property

Public Property Let Size(ByVal Value As Double)

    mSize = Value

End Property

Public Property Get ObjectNumber() As Long

    ObjectNumber = mObjectNumber

End Property

Public Property Let ObjectNumber(ByVal Value As Long)

    mObjectNumber = Value

End Property

'---------------------------------------------------------
' Fett
'---------------------------------------------------------
Public Property Get Bold() As Boolean

    Bold = mBold

End Property

Public Property Let Bold(ByVal Value As Boolean)

    mBold = Value

End Property


'---------------------------------------------------------
' Kursiv
'---------------------------------------------------------
Public Property Get Italic() As Boolean

    Italic = mItalic

End Property

Public Property Let Italic(ByVal Value As Boolean)

    mItalic = Value

End Property

'---------------------------------------------------------
' Unterstichen
'---------------------------------------------------------
Public Property Get Underline() As Boolean

    Underline = mUnderline

End Property

Public Property Let Underline(ByVal Value As Boolean)

    mUnderline = Value

End Property


'---------------------------------------------------------
' Durchgestrichen
'---------------------------------------------------------
Public Property Get StrikeOut() As Boolean

    StrikeOut = mStrikeOut

End Property

Public Property Let StrikeOut(ByVal Value As Boolean)

    mStrikeOut = Value

End Property

'---------------------------------------------------------
' PDF-Basisfont
'---------------------------------------------------------
Public Property Get BaseFontName() As String

    Select Case mFamily

        Case psfHelvetica, psfArial
            If mBold Then
                If mItalic Then
                    BaseFontName = "Helvetica-BoldOblique"
                Else
                    BaseFontName = "Helvetica-Bold"
                End If
            Else
                If mItalic Then
                    BaseFontName = "Helvetica-Oblique"
                Else
                    BaseFontName = "Helvetica"
                End If
            End If

        Case psfTimesRoman, psfTimesNewRoman
            If mBold Then
                If mItalic Then
                    BaseFontName = "Times-BoldItalic"
                Else
                    BaseFontName = "Times-Bold"
                End If
            Else
                If mItalic Then
                    BaseFontName = "Times-Italic"
                Else
                    BaseFontName = "Times-Roman"
                End If
            End If

        Case psfCourier, psfCourierNew
            If mBold Then
                If mItalic Then
                    BaseFontName = "Courier-BoldOblique"
                Else
                    BaseFontName = "Courier-Bold"
                End If
            Else
                If mItalic Then
                    BaseFontName = "Courier-Oblique"
                Else
                    BaseFontName = "Courier"
                End If
            End If

        Case psfSymbol
            BaseFontName = "Symbol"
'            If mBold Then
'                If mItalic Then
'                    BaseFontName = "Symbol"
'                Else
'                    BaseFontName = "Symbol"
'                End If
'            Else
'                If mItalic Then
'                    BaseFontName = "Symbol"
'                Else
'                    BaseFontName = "Symbol"
'                End If
'            End If

        Case psfZapfDingbats
            BaseFontName = "ZapfDingbats"
'            If mBold Then
'                If mItalic Then
'                    BaseFontName = "ZapfDingbats"
'                Else
'                    BaseFontName = "ZapfDingbats"
'                End If
'            Else
'                If mItalic Then
'                    BaseFontName = "ZapfDingbats"
'                Else
'                    BaseFontName = "ZapfDingbats"
'                End If
'            End If

        Case Else
            BaseFontName = "Helvetica" 'mName
    End Select
 
End Property

'---------------------------------------------------------
' Textbreite in PDF-Punkten
'
' Ermittelt die Breite eines Textes mit der aktuellen
' Schrift und Schriftgröße.
'
' Rückgabe:
'   Breite in PDF-Punkten
'---------------------------------------------------------
Public Function TextWidth(ByVal Text As String) As Double

    Dim I As Long
    Dim Width As Double

    Width = 0#

    For I = 1 To Len(Text)
        Width = Width + CharacterWidth(Mid$(Text, I, 1))
    Next I

    TextWidth = Width * mSize / 1000#

End Function


'---------------------------------------------------------
' Breite eines einzelnen Zeichens in 1/1000 em
'
' Rückgabe in PDF-Font-Einheiten (1/1000)
'
' Aktuell für die unterstützten Standardfonts.
'---------------------------------------------------------
'Private Function CharacterWidth(ByVal Ch As String) As Double
'
'    Select Case mFamily
'
'        Case psfHelvetica, psfArial
'
'            CharacterWidth = HelveticaCharacterWidth(Ch)
'
'        Case psfTimesRoman, psfTimesNewRoman
'            CharacterWidth = TimesCharacterWidth(Ch)
'
'        Case psfCourier, psfCourierNew
'
'            CharacterWidth = 600
'
'        Case Else
'
'            ' Vorläufiger Fallback
'            CharacterWidth = HelveticaCharacterWidth(Ch)
'
'    End Select
'
'End Function

Private Function CharacterWidth(ByVal Ch As String) As Double

    Select Case mFamily

        Case psfHelvetica, psfArial

            If mBold Then
                If mItalic Then
                    CharacterWidth = HelveticaBoldItalicCharacterWidth(Ch)
                Else
                    CharacterWidth = HelveticaBoldCharacterWidth(Ch)
                End If
            Else
                If mItalic Then
                    CharacterWidth = HelveticaItalicCharacterWidth(Ch)
                Else
                    CharacterWidth = HelveticaCharacterWidth(Ch)
                End If
            End If

        Case psfTimesRoman, psfTimesNewRoman

            If mBold Then
                If mItalic Then
                    CharacterWidth = TimesBoldItalicCharacterWidth(Ch)
                Else
                    CharacterWidth = TimesBoldCharacterWidth(Ch)
                End If
            Else
                If mItalic Then
                    CharacterWidth = TimesItalicCharacterWidth(Ch)
                Else
                    CharacterWidth = TimesCharacterWidth(Ch)
                End If
            End If

        Case psfCourier, psfCourierNew

            CharacterWidth = 600

        Case Else

            CharacterWidth = HelveticaCharacterWidth(Ch)

    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Helvetica
'
' Rückgabe in PDF-Font-Einheiten (1/1000)
'---------------------------------------------------------
Private Function HelveticaCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            HelveticaCharacterWidth = 278

        Case "!"
            HelveticaCharacterWidth = 278
        Case """"
            HelveticaCharacterWidth = 355
        Case "#"
            HelveticaCharacterWidth = 556
        Case "$"
            HelveticaCharacterWidth = 556
        Case "%"
            HelveticaCharacterWidth = 889
        Case "&"
            HelveticaCharacterWidth = 667
        Case "'"
            HelveticaCharacterWidth = 222
        Case "("
            HelveticaCharacterWidth = 333
        Case ")"
            HelveticaCharacterWidth = 333
        Case "*"
            HelveticaCharacterWidth = 389
        Case "+"
            HelveticaCharacterWidth = 584
        Case ","
            HelveticaCharacterWidth = 278
        Case "-"
            HelveticaCharacterWidth = 333
        Case "."
            HelveticaCharacterWidth = 278
        Case "/"
            HelveticaCharacterWidth = 278
        
        Case "0"
            HelveticaCharacterWidth = 556
        Case "1"
            HelveticaCharacterWidth = 556
        Case "2"
            HelveticaCharacterWidth = 556
        Case "3"
            HelveticaCharacterWidth = 556
        Case "4"
            HelveticaCharacterWidth = 556
        Case "5"
            HelveticaCharacterWidth = 556
        Case "6"
            HelveticaCharacterWidth = 556
        Case "7"
            HelveticaCharacterWidth = 556
        Case "8"
            HelveticaCharacterWidth = 556
        Case "9"
            HelveticaCharacterWidth = 556
        
        Case ":"
            HelveticaCharacterWidth = 278
        Case ";"
            HelveticaCharacterWidth = 278
        Case "<"
            HelveticaCharacterWidth = 584
        Case "="
            HelveticaCharacterWidth = 584
        Case ">"
            HelveticaCharacterWidth = 584
        Case "?"
            HelveticaCharacterWidth = 556
        Case "@"
            HelveticaCharacterWidth = 1015

        Case "A"
            HelveticaCharacterWidth = 667
        Case "B"
            HelveticaCharacterWidth = 667
        Case "C"
            HelveticaCharacterWidth = 722
        Case "D"
            HelveticaCharacterWidth = 722
        Case "E"
            HelveticaCharacterWidth = 667
        Case "F"
            HelveticaCharacterWidth = 611
        Case "G"
            HelveticaCharacterWidth = 778
        Case "H"
            HelveticaCharacterWidth = 722
        Case "I"
            HelveticaCharacterWidth = 278
        Case "J"
            HelveticaCharacterWidth = 500
        Case "K"
            HelveticaCharacterWidth = 667
        Case "L"
            HelveticaCharacterWidth = 556
        Case "M"
            HelveticaCharacterWidth = 833
        Case "N"
            HelveticaCharacterWidth = 722
        Case "O"
            HelveticaCharacterWidth = 778
        Case "P"
            HelveticaCharacterWidth = 667
        Case "Q"
            HelveticaCharacterWidth = 778
        Case "R"
            HelveticaCharacterWidth = 722
        Case "S"
            HelveticaCharacterWidth = 667
        Case "T"
            HelveticaCharacterWidth = 611
        Case "U"
            HelveticaCharacterWidth = 722
        Case "V"
            HelveticaCharacterWidth = 667
        Case "W"
            HelveticaCharacterWidth = 944
        Case "X"
            HelveticaCharacterWidth = 667
        Case "Y"
            HelveticaCharacterWidth = 667
        Case "Z"
            HelveticaCharacterWidth = 611

        Case "["
            HelveticaCharacterWidth = 278
        Case "\"
            HelveticaCharacterWidth = 278
        Case "]"
            HelveticaCharacterWidth = 278
        Case "^"
            HelveticaCharacterWidth = 469
        Case "_"
            HelveticaCharacterWidth = 556
        Case "`"
            HelveticaCharacterWidth = 222
        
        Case "a"
            HelveticaCharacterWidth = 556
        Case "b"
            HelveticaCharacterWidth = 556
        Case "c"
            HelveticaCharacterWidth = 500
        Case "d"
            HelveticaCharacterWidth = 556
        Case "e"
            HelveticaCharacterWidth = 556
        Case "f"
            HelveticaCharacterWidth = 278
        Case "g"
            HelveticaCharacterWidth = 556
        Case "h"
            HelveticaCharacterWidth = 556
        Case "i"
            HelveticaCharacterWidth = 222
        Case "j"
            HelveticaCharacterWidth = 222
        Case "k"
            HelveticaCharacterWidth = 500
        Case "l"
            HelveticaCharacterWidth = 222
        Case "m"
            HelveticaCharacterWidth = 833
        Case "n"
            HelveticaCharacterWidth = 556
        Case "o"
            HelveticaCharacterWidth = 556
        Case "p"
            HelveticaCharacterWidth = 556
        Case "q"
            HelveticaCharacterWidth = 556
        Case "r"
            HelveticaCharacterWidth = 333
        Case "s"
            HelveticaCharacterWidth = 500
        Case "t"
            HelveticaCharacterWidth = 278
        Case "u"
            HelveticaCharacterWidth = 556
        Case "v"
            HelveticaCharacterWidth = 500
        Case "w"
            HelveticaCharacterWidth = 722
        Case "x"
            HelveticaCharacterWidth = 500
        Case "y"
            HelveticaCharacterWidth = 500
        Case "z"
            HelveticaCharacterWidth = 500
        
        Case Else
            ' Nicht unterstütztes Zeichen
            HelveticaCharacterWidth = 500
    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Helvetica-Bold
'
' Rückgabe in 1/1000 em
'---------------------------------------------------------
Private Function HelveticaBoldCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            HelveticaBoldCharacterWidth = 278

        Case "!"
            HelveticaBoldCharacterWidth = 333
        Case """"
            HelveticaBoldCharacterWidth = 474
        Case "#"
            HelveticaBoldCharacterWidth = 556
        Case "$"
            HelveticaBoldCharacterWidth = 556
        Case "%"
            HelveticaBoldCharacterWidth = 889
        Case "&"
            HelveticaBoldCharacterWidth = 722
        Case "'"
            HelveticaBoldCharacterWidth = 278
        Case "("
            HelveticaBoldCharacterWidth = 333
        Case ")"
            HelveticaBoldCharacterWidth = 333
        Case "*"
            HelveticaBoldCharacterWidth = 500
        Case "+"
            HelveticaBoldCharacterWidth = 584
        Case ","
            HelveticaBoldCharacterWidth = 278
        Case "-"
            HelveticaBoldCharacterWidth = 333
        Case "."
            HelveticaBoldCharacterWidth = 278
        Case "/"
            HelveticaBoldCharacterWidth = 278

        Case "0"
            HelveticaBoldCharacterWidth = 556
        Case "1"
            HelveticaBoldCharacterWidth = 556
        Case "2"
            HelveticaBoldCharacterWidth = 556
        Case "3"
            HelveticaBoldCharacterWidth = 556
        Case "4"
            HelveticaBoldCharacterWidth = 556
        Case "5"
            HelveticaBoldCharacterWidth = 556
        Case "6"
            HelveticaBoldCharacterWidth = 556
        Case "7"
            HelveticaBoldCharacterWidth = 556
        Case "8"
            HelveticaBoldCharacterWidth = 556
        Case "9"
            HelveticaBoldCharacterWidth = 556

        Case ":"
            HelveticaBoldCharacterWidth = 333
        Case ";"
            HelveticaBoldCharacterWidth = 333
        Case "<"
            HelveticaBoldCharacterWidth = 584
        Case "="
            HelveticaBoldCharacterWidth = 584
        Case ">"
            HelveticaBoldCharacterWidth = 584
        Case "?"
            HelveticaBoldCharacterWidth = 611
        Case "@"
            HelveticaBoldCharacterWidth = 975

        Case "A"
            HelveticaBoldCharacterWidth = 722
        Case "B"
            HelveticaBoldCharacterWidth = 722
        Case "C"
            HelveticaBoldCharacterWidth = 722
        Case "D"
            HelveticaBoldCharacterWidth = 722
        Case "E"
            HelveticaBoldCharacterWidth = 667
        Case "F"
            HelveticaBoldCharacterWidth = 611
        Case "G"
            HelveticaBoldCharacterWidth = 778
        Case "H"
            HelveticaBoldCharacterWidth = 722
        Case "I"
            HelveticaBoldCharacterWidth = 278
        Case "J"
            HelveticaBoldCharacterWidth = 556
        Case "K"
            HelveticaBoldCharacterWidth = 722
        Case "L"
            HelveticaBoldCharacterWidth = 611
        Case "M"
            HelveticaBoldCharacterWidth = 833
        Case "N"
            HelveticaBoldCharacterWidth = 722
        Case "O"
            HelveticaBoldCharacterWidth = 778
        Case "P"
            HelveticaBoldCharacterWidth = 667
        Case "Q"
            HelveticaBoldCharacterWidth = 778
        Case "R"
            HelveticaBoldCharacterWidth = 722
        Case "S"
            HelveticaBoldCharacterWidth = 667
        Case "T"
            HelveticaBoldCharacterWidth = 611
        Case "U"
            HelveticaBoldCharacterWidth = 722
        Case "V"
            HelveticaBoldCharacterWidth = 667
        Case "W"
            HelveticaBoldCharacterWidth = 944
        Case "X"
            HelveticaBoldCharacterWidth = 667
        Case "Y"
            HelveticaBoldCharacterWidth = 667
        Case "Z"
            HelveticaBoldCharacterWidth = 611

        Case "["
            HelveticaBoldCharacterWidth = 333
        Case "\"
            HelveticaBoldCharacterWidth = 278
        Case "]"
            HelveticaBoldCharacterWidth = 333
        Case "^"
            HelveticaBoldCharacterWidth = 584
        Case "_"
            HelveticaBoldCharacterWidth = 556
        Case "`"
            HelveticaBoldCharacterWidth = 333

        Case "a"
            HelveticaBoldCharacterWidth = 556
        Case "b"
            HelveticaBoldCharacterWidth = 611
        Case "c"
            HelveticaBoldCharacterWidth = 556
        Case "d"
            HelveticaBoldCharacterWidth = 611
        Case "e"
            HelveticaBoldCharacterWidth = 556
        Case "f"
            HelveticaBoldCharacterWidth = 333
        Case "g"
            HelveticaBoldCharacterWidth = 611
        Case "h"
            HelveticaBoldCharacterWidth = 611
        Case "i"
            HelveticaBoldCharacterWidth = 278
        Case "j"
            HelveticaBoldCharacterWidth = 278
        Case "k"
            HelveticaBoldCharacterWidth = 556
        Case "l"
            HelveticaBoldCharacterWidth = 278
        Case "m"
            HelveticaBoldCharacterWidth = 889
        Case "n"
            HelveticaBoldCharacterWidth = 611
        Case "o"
            HelveticaBoldCharacterWidth = 611
        Case "p"
            HelveticaBoldCharacterWidth = 611
        Case "q"
            HelveticaBoldCharacterWidth = 611
        Case "r"
            HelveticaBoldCharacterWidth = 389
        Case "s"
            HelveticaBoldCharacterWidth = 556
        Case "t"
            HelveticaBoldCharacterWidth = 333
        Case "u"
            HelveticaBoldCharacterWidth = 611
        Case "v"
            HelveticaBoldCharacterWidth = 556
        Case "w"
            HelveticaBoldCharacterWidth = 778
        Case "x"
            HelveticaBoldCharacterWidth = 556
        Case "y"
            HelveticaBoldCharacterWidth = 556
        Case "z"
            HelveticaBoldCharacterWidth = 500

        Case "{"
            HelveticaBoldCharacterWidth = 389
        Case "|"
            HelveticaBoldCharacterWidth = 280
        Case "}"
            HelveticaBoldCharacterWidth = 389
        Case "~"
            HelveticaBoldCharacterWidth = 584

        Case Else
            HelveticaBoldCharacterWidth = 500

    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Helvetica-Oblique
'
' Rückgabe in 1/1000 em
'---------------------------------------------------------
Private Function HelveticaItalicCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            HelveticaItalicCharacterWidth = 278
        Case "!"
            HelveticaItalicCharacterWidth = 278
        Case """"
            HelveticaItalicCharacterWidth = 355
        Case "#"
            HelveticaItalicCharacterWidth = 556
        Case "$"
            HelveticaItalicCharacterWidth = 556
        Case "%"
            HelveticaItalicCharacterWidth = 889
        Case "&"
            HelveticaItalicCharacterWidth = 667
        Case "'"
            HelveticaItalicCharacterWidth = 222
        Case "("
            HelveticaItalicCharacterWidth = 333
        Case ")"
            HelveticaItalicCharacterWidth = 333
        Case "*"
            HelveticaItalicCharacterWidth = 389
        Case "+"
            HelveticaItalicCharacterWidth = 584
        Case ","
            HelveticaItalicCharacterWidth = 278
        Case "-"
            HelveticaItalicCharacterWidth = 333
        Case "."
            HelveticaItalicCharacterWidth = 278
        Case "/"
            HelveticaItalicCharacterWidth = 278

        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            HelveticaItalicCharacterWidth = 556

        Case ":"
            HelveticaItalicCharacterWidth = 278
        Case ";"
            HelveticaItalicCharacterWidth = 278
        Case "<"
            HelveticaItalicCharacterWidth = 584
        Case "="
            HelveticaItalicCharacterWidth = 584
        Case ">"
            HelveticaItalicCharacterWidth = 584
        Case "?"
            HelveticaItalicCharacterWidth = 556
        Case "@"
            HelveticaItalicCharacterWidth = 1015

        Case "A"
            HelveticaItalicCharacterWidth = 667
        Case "B"
            HelveticaItalicCharacterWidth = 667
        Case "C"
            HelveticaItalicCharacterWidth = 667
        Case "D"
            HelveticaItalicCharacterWidth = 722
        Case "E"
            HelveticaItalicCharacterWidth = 667
        Case "F"
            HelveticaItalicCharacterWidth = 611
        Case "G"
            HelveticaItalicCharacterWidth = 722
        Case "H"
            HelveticaItalicCharacterWidth = 722
        Case "I"
            HelveticaItalicCharacterWidth = 278
        Case "J"
            HelveticaItalicCharacterWidth = 500
        Case "K"
            HelveticaItalicCharacterWidth = 667
        Case "L"
            HelveticaItalicCharacterWidth = 556
        Case "M"
            HelveticaItalicCharacterWidth = 833
        Case "N"
            HelveticaItalicCharacterWidth = 722
        Case "O"
            HelveticaItalicCharacterWidth = 722
        Case "P"
            HelveticaItalicCharacterWidth = 667
        Case "Q"
            HelveticaItalicCharacterWidth = 722
        Case "R"
            HelveticaItalicCharacterWidth = 611
        Case "S"
            HelveticaItalicCharacterWidth = 556
        Case "T"
            HelveticaItalicCharacterWidth = 611
        Case "U"
            HelveticaItalicCharacterWidth = 722
        Case "V"
            HelveticaItalicCharacterWidth = 667
        Case "W"
            HelveticaItalicCharacterWidth = 944
        Case "X"
            HelveticaItalicCharacterWidth = 667
        Case "Y"
            HelveticaItalicCharacterWidth = 667
        Case "Z"
            HelveticaItalicCharacterWidth = 611

        Case "["
            HelveticaItalicCharacterWidth = 278
        Case "\"
            HelveticaItalicCharacterWidth = 278
        Case "]"
            HelveticaItalicCharacterWidth = 278
        Case "^"
            HelveticaItalicCharacterWidth = 469
        Case "_"
            HelveticaItalicCharacterWidth = 556
        Case "`"
            HelveticaItalicCharacterWidth = 333

        Case "a"
            HelveticaItalicCharacterWidth = 556
        Case "b"
            HelveticaItalicCharacterWidth = 556
        Case "c"
            HelveticaItalicCharacterWidth = 500
        Case "d"
            HelveticaItalicCharacterWidth = 556
        Case "e"
            HelveticaItalicCharacterWidth = 556
        Case "f"
            HelveticaItalicCharacterWidth = 278
        Case "g"
            HelveticaItalicCharacterWidth = 556
        Case "h"
            HelveticaItalicCharacterWidth = 556
        Case "i"
            HelveticaItalicCharacterWidth = 222
        Case "j"
            HelveticaItalicCharacterWidth = 222
        Case "k"
            HelveticaItalicCharacterWidth = 500
        Case "l"
            HelveticaItalicCharacterWidth = 222
        Case "m"
            HelveticaItalicCharacterWidth = 833
        Case "n"
            HelveticaItalicCharacterWidth = 556
        Case "o"
            HelveticaItalicCharacterWidth = 500
        Case "p"
            HelveticaItalicCharacterWidth = 556
        Case "q"
            HelveticaItalicCharacterWidth = 556
        Case "r"
            HelveticaItalicCharacterWidth = 333
        Case "s"
            HelveticaItalicCharacterWidth = 500
        Case "t"
            HelveticaItalicCharacterWidth = 278
        Case "u"
            HelveticaItalicCharacterWidth = 556
        Case "v"
            HelveticaItalicCharacterWidth = 500
        Case "w"
            HelveticaItalicCharacterWidth = 722
        Case "x"
            HelveticaItalicCharacterWidth = 500
        Case "y"
            HelveticaItalicCharacterWidth = 500
        Case "z"
            HelveticaItalicCharacterWidth = 500

        Case "{"
            HelveticaItalicCharacterWidth = 334
        Case "|"
            HelveticaItalicCharacterWidth = 260
        Case "}"
            HelveticaItalicCharacterWidth = 334
        Case "~"
            HelveticaItalicCharacterWidth = 584

        Case Else
            HelveticaItalicCharacterWidth = 500

    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Helvetica-BoldOblique
'
' Rückgabe in 1/1000 em
'---------------------------------------------------------
Private Function HelveticaBoldItalicCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            HelveticaBoldItalicCharacterWidth = 278
        Case "!"
            HelveticaBoldItalicCharacterWidth = 333
        Case """"
            HelveticaBoldItalicCharacterWidth = 474
        Case "#"
            HelveticaBoldItalicCharacterWidth = 556
        Case "$"
            HelveticaBoldItalicCharacterWidth = 556
        Case "%"
            HelveticaBoldItalicCharacterWidth = 889
        Case "&"
            HelveticaBoldItalicCharacterWidth = 722
        Case "'"
            HelveticaBoldItalicCharacterWidth = 278
        Case "("
            HelveticaBoldItalicCharacterWidth = 333
        Case ")"
            HelveticaBoldItalicCharacterWidth = 333
        Case "*"
            HelveticaBoldItalicCharacterWidth = 500
        Case "+"
            HelveticaBoldItalicCharacterWidth = 584
        Case ","
            HelveticaBoldItalicCharacterWidth = 278
        Case "-"
            HelveticaBoldItalicCharacterWidth = 333
        Case "."
            HelveticaBoldItalicCharacterWidth = 278
        Case "/"
            HelveticaBoldItalicCharacterWidth = 278

        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            HelveticaBoldItalicCharacterWidth = 556

        Case ":"
            HelveticaBoldItalicCharacterWidth = 333
        Case ";"
            HelveticaBoldItalicCharacterWidth = 333
        Case "<"
            HelveticaBoldItalicCharacterWidth = 584
        Case "="
            HelveticaBoldItalicCharacterWidth = 584
        Case ">"
            HelveticaBoldItalicCharacterWidth = 584
        Case "?"
            HelveticaBoldItalicCharacterWidth = 611
        Case "@"
            HelveticaBoldItalicCharacterWidth = 975

        Case "A"
            HelveticaBoldItalicCharacterWidth = 667
        Case "B"
            HelveticaBoldItalicCharacterWidth = 667
        Case "C"
            HelveticaBoldItalicCharacterWidth = 722
        Case "D"
            HelveticaBoldItalicCharacterWidth = 722
        Case "E"
            HelveticaBoldItalicCharacterWidth = 667
        Case "F"
            HelveticaBoldItalicCharacterWidth = 611
        Case "G"
            HelveticaBoldItalicCharacterWidth = 778
        Case "H"
            HelveticaBoldItalicCharacterWidth = 722
        Case "I"
            HelveticaBoldItalicCharacterWidth = 278
        Case "J"
            HelveticaBoldItalicCharacterWidth = 556
        Case "K"
            HelveticaBoldItalicCharacterWidth = 722
        Case "L"
            HelveticaBoldItalicCharacterWidth = 611
        Case "M"
            HelveticaBoldItalicCharacterWidth = 833
        Case "N"
            HelveticaBoldItalicCharacterWidth = 722
        Case "O"
            HelveticaBoldItalicCharacterWidth = 778
        Case "P"
            HelveticaBoldItalicCharacterWidth = 667
        Case "Q"
            HelveticaBoldItalicCharacterWidth = 778
        Case "R"
            HelveticaBoldItalicCharacterWidth = 722
        Case "S"
            HelveticaBoldItalicCharacterWidth = 667
        Case "T"
            HelveticaBoldItalicCharacterWidth = 611
        Case "U"
            HelveticaBoldItalicCharacterWidth = 722
        Case "V"
            HelveticaBoldItalicCharacterWidth = 667
        Case "W"
            HelveticaBoldItalicCharacterWidth = 944
        Case "X"
            HelveticaBoldItalicCharacterWidth = 667
        Case "Y"
            HelveticaBoldItalicCharacterWidth = 667
        Case "Z"
            HelveticaBoldItalicCharacterWidth = 611

        Case "["
            HelveticaBoldItalicCharacterWidth = 333
        Case "\"
            HelveticaBoldItalicCharacterWidth = 278
        Case "]"
            HelveticaBoldItalicCharacterWidth = 333
        Case "^"
            HelveticaBoldItalicCharacterWidth = 584
        Case "_"
            HelveticaBoldItalicCharacterWidth = 556
        Case "`"
            HelveticaBoldItalicCharacterWidth = 333

        Case "a"
            HelveticaBoldItalicCharacterWidth = 556
        Case "b"
            HelveticaBoldItalicCharacterWidth = 556
        Case "c"
            HelveticaBoldItalicCharacterWidth = 500
        Case "d"
            HelveticaBoldItalicCharacterWidth = 556
        Case "e"
            HelveticaBoldItalicCharacterWidth = 556
        Case "f"
            HelveticaBoldItalicCharacterWidth = 333
        Case "g"
            HelveticaBoldItalicCharacterWidth = 556
        Case "h"
            HelveticaBoldItalicCharacterWidth = 556
        Case "i"
            HelveticaBoldItalicCharacterWidth = 278
        Case "j"
            HelveticaBoldItalicCharacterWidth = 278
        Case "k"
            HelveticaBoldItalicCharacterWidth = 556
        Case "l"
            HelveticaBoldItalicCharacterWidth = 278
        Case "m"
            HelveticaBoldItalicCharacterWidth = 833
        Case "n"
            HelveticaBoldItalicCharacterWidth = 556
        Case "o"
            HelveticaBoldItalicCharacterWidth = 556
        Case "p"
            HelveticaBoldItalicCharacterWidth = 556
        Case "q"
            HelveticaBoldItalicCharacterWidth = 556
        Case "r"
            HelveticaBoldItalicCharacterWidth = 389
        Case "s"
            HelveticaBoldItalicCharacterWidth = 500
        Case "t"
            HelveticaBoldItalicCharacterWidth = 333
        Case "u"
            HelveticaBoldItalicCharacterWidth = 556
        Case "v"
            HelveticaBoldItalicCharacterWidth = 556
        Case "w"
            HelveticaBoldItalicCharacterWidth = 778
        Case "x"
            HelveticaBoldItalicCharacterWidth = 556
        Case "y"
            HelveticaBoldItalicCharacterWidth = 556
        Case "z"
            HelveticaBoldItalicCharacterWidth = 500

        Case "{"
            HelveticaBoldItalicCharacterWidth = 389
        Case "|"
            HelveticaBoldItalicCharacterWidth = 280
        Case "}"
            HelveticaBoldItalicCharacterWidth = 389
        Case "~"
            HelveticaBoldItalicCharacterWidth = 584

        Case Else
            HelveticaBoldItalicCharacterWidth = 500

    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Times-Roman
'
' Rückgabe in 1/1000 em
'---------------------------------------------------------
Private Function TimesCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            TimesCharacterWidth = 250

        Case "!"
            TimesCharacterWidth = 333
        Case """"
            TimesCharacterWidth = 408
        Case "#"
            TimesCharacterWidth = 500
        Case "$"
            TimesCharacterWidth = 500
        Case "%"
            TimesCharacterWidth = 833
        Case "&"
            TimesCharacterWidth = 778
        Case "'"
            TimesCharacterWidth = 333
        Case "("
            TimesCharacterWidth = 333
        Case ")"
            TimesCharacterWidth = 333
        Case "*"
            TimesCharacterWidth = 500
        Case "+"
            TimesCharacterWidth = 564
        Case ","
            TimesCharacterWidth = 250
        Case "-"
            TimesCharacterWidth = 333
        Case "."
            TimesCharacterWidth = 250
        Case "/"
            TimesCharacterWidth = 278

        Case "0"
            TimesCharacterWidth = 500
        Case "1"
            TimesCharacterWidth = 500
        Case "2"
            TimesCharacterWidth = 500
        Case "3"
            TimesCharacterWidth = 500
        Case "4"
            TimesCharacterWidth = 500
        Case "5"
            TimesCharacterWidth = 500
        Case "6"
            TimesCharacterWidth = 500
        Case "7"
            TimesCharacterWidth = 500
        Case "8"
            TimesCharacterWidth = 500
        Case "9"
            TimesCharacterWidth = 500

        Case ":"
            TimesCharacterWidth = 278
        Case ";"
            TimesCharacterWidth = 278
        Case "<"
            TimesCharacterWidth = 564
        Case "="
            TimesCharacterWidth = 564
        Case ">"
            TimesCharacterWidth = 564
        Case "?"
            TimesCharacterWidth = 444
        Case "@"
            TimesCharacterWidth = 921

        Case "A"
            TimesCharacterWidth = 722
        Case "B"
            TimesCharacterWidth = 667
        Case "C"
            TimesCharacterWidth = 667
        Case "D"
            TimesCharacterWidth = 722
        Case "E"
            TimesCharacterWidth = 611
        Case "F"
            TimesCharacterWidth = 556
        Case "G"
            TimesCharacterWidth = 722
        Case "H"
            TimesCharacterWidth = 722
        Case "I"
            TimesCharacterWidth = 333
        Case "J"
            TimesCharacterWidth = 389
        Case "K"
            TimesCharacterWidth = 722
        Case "L"
            TimesCharacterWidth = 611
        Case "M"
            TimesCharacterWidth = 889
        Case "N"
            TimesCharacterWidth = 722
        Case "O"
            TimesCharacterWidth = 722
        Case "P"
            TimesCharacterWidth = 556
        Case "Q"
            TimesCharacterWidth = 722
        Case "R"
            TimesCharacterWidth = 667
        Case "S"
            TimesCharacterWidth = 556
        Case "T"
            TimesCharacterWidth = 611
        Case "U"
            TimesCharacterWidth = 722
        Case "V"
            TimesCharacterWidth = 722
        Case "W"
            TimesCharacterWidth = 944
        Case "X"
            TimesCharacterWidth = 722
        Case "Y"
            TimesCharacterWidth = 722
        Case "Z"
            TimesCharacterWidth = 611

        Case "["
            TimesCharacterWidth = 333
        Case "\"
            TimesCharacterWidth = 278
        Case "]"
            TimesCharacterWidth = 333
        Case "^"
            TimesCharacterWidth = 469
        Case "_"
            TimesCharacterWidth = 500
        Case "`"
            TimesCharacterWidth = 333

        Case "a"
            TimesCharacterWidth = 444
        Case "b"
            TimesCharacterWidth = 500
        Case "c"
            TimesCharacterWidth = 444
        Case "d"
            TimesCharacterWidth = 500
        Case "e"
            TimesCharacterWidth = 444
        Case "f"
            TimesCharacterWidth = 333
        Case "g"
            TimesCharacterWidth = 500
        Case "h"
            TimesCharacterWidth = 500
        Case "i"
            TimesCharacterWidth = 278
        Case "j"
            TimesCharacterWidth = 278
        Case "k"
            TimesCharacterWidth = 500
        Case "l"
            TimesCharacterWidth = 278
        Case "m"
            TimesCharacterWidth = 778
        Case "n"
            TimesCharacterWidth = 500
        Case "o"
            TimesCharacterWidth = 500
        Case "p"
            TimesCharacterWidth = 500
        Case "q"
            TimesCharacterWidth = 500
        Case "r"
            TimesCharacterWidth = 333
        Case "s"
            TimesCharacterWidth = 389
        Case "t"
            TimesCharacterWidth = 278
        Case "u"
            TimesCharacterWidth = 500
        Case "v"
            TimesCharacterWidth = 500
        Case "w"
            TimesCharacterWidth = 722
        Case "x"
            TimesCharacterWidth = 500
        Case "y"
            TimesCharacterWidth = 500
        Case "z"
            TimesCharacterWidth = 444

        Case "{"
            TimesCharacterWidth = 480
        Case "|"
            TimesCharacterWidth = 200
        Case "}"
            TimesCharacterWidth = 480
        Case "~"
            TimesCharacterWidth = 541

        Case Else
            ' Nicht unterstütztes Zeichen
            TimesCharacterWidth = 500

    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Times-Bold
'
' Rückgabe in 1/1000 em
'---------------------------------------------------------
Private Function TimesBoldCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            TimesBoldCharacterWidth = 250

        Case "!"
            TimesBoldCharacterWidth = 333
        Case """"
            TimesBoldCharacterWidth = 555
        Case "#"
            TimesBoldCharacterWidth = 500
        Case "$"
            TimesBoldCharacterWidth = 500
        Case "%"
            TimesBoldCharacterWidth = 1000
        Case "&"
            TimesBoldCharacterWidth = 833
        Case "'"
            TimesBoldCharacterWidth = 333
        Case "("
            TimesBoldCharacterWidth = 333
        Case ")"
            TimesBoldCharacterWidth = 333
        Case "*"
            TimesBoldCharacterWidth = 500
        Case "+"
            TimesBoldCharacterWidth = 570
        Case ","
            TimesBoldCharacterWidth = 250
        Case "-"
            TimesBoldCharacterWidth = 333
        Case "."
            TimesBoldCharacterWidth = 250
        Case "/"
            TimesBoldCharacterWidth = 278

        Case "0", "1", "2", "3", "4", _
             "5", "6", "7", "8", "9"
            TimesBoldCharacterWidth = 500

        Case ":"
            TimesBoldCharacterWidth = 333
        Case ";"
            TimesBoldCharacterWidth = 333
        Case "<"
            TimesBoldCharacterWidth = 570
        Case "="
            TimesBoldCharacterWidth = 570
        Case ">"
            TimesBoldCharacterWidth = 570
        Case "?"
            TimesBoldCharacterWidth = 500
        Case "@"
            TimesBoldCharacterWidth = 930

        Case "A"
            TimesBoldCharacterWidth = 722
        Case "B"
            TimesBoldCharacterWidth = 667
        Case "C"
            TimesBoldCharacterWidth = 722
        Case "D"
            TimesBoldCharacterWidth = 722
        Case "E"
            TimesBoldCharacterWidth = 667
        Case "F"
            TimesBoldCharacterWidth = 611
        Case "G"
            TimesBoldCharacterWidth = 778
        Case "H"
            TimesBoldCharacterWidth = 722
        Case "I"
            TimesBoldCharacterWidth = 389
        Case "J"
            TimesBoldCharacterWidth = 500
        Case "K"
            TimesBoldCharacterWidth = 778
        Case "L"
            TimesBoldCharacterWidth = 667
        Case "M"
            TimesBoldCharacterWidth = 944
        Case "N"
            TimesBoldCharacterWidth = 722
        Case "O"
            TimesBoldCharacterWidth = 778
        Case "P"
            TimesBoldCharacterWidth = 611
        Case "Q"
            TimesBoldCharacterWidth = 778
        Case "R"
            TimesBoldCharacterWidth = 722
        Case "S"
            TimesBoldCharacterWidth = 556
        Case "T"
            TimesBoldCharacterWidth = 667
        Case "U"
            TimesBoldCharacterWidth = 722
        Case "V"
            TimesBoldCharacterWidth = 722
        Case "W"
            TimesBoldCharacterWidth = 1000
        Case "X"
            TimesBoldCharacterWidth = 722
        Case "Y"
            TimesBoldCharacterWidth = 722
        Case "Z"
            TimesBoldCharacterWidth = 667

        Case "["
            TimesBoldCharacterWidth = 333
        Case "\"
            TimesBoldCharacterWidth = 278
        Case "]"
            TimesBoldCharacterWidth = 333
        Case "^"
            TimesBoldCharacterWidth = 581
        Case "_"
            TimesBoldCharacterWidth = 500
        Case "`"
            TimesBoldCharacterWidth = 333

        Case "a"
            TimesBoldCharacterWidth = 500
        Case "b"
            TimesBoldCharacterWidth = 556
        Case "c"
            TimesBoldCharacterWidth = 444
        Case "d"
            TimesBoldCharacterWidth = 556
        Case "e"
            TimesBoldCharacterWidth = 444
        Case "f"
            TimesBoldCharacterWidth = 333
        Case "g"
            TimesBoldCharacterWidth = 500
        Case "h"
            TimesBoldCharacterWidth = 556
        Case "i"
            TimesBoldCharacterWidth = 278
        Case "j"
            TimesBoldCharacterWidth = 333
        Case "k"
            TimesBoldCharacterWidth = 556
        Case "l"
            TimesBoldCharacterWidth = 278
        Case "m"
            TimesBoldCharacterWidth = 833
        Case "n"
            TimesBoldCharacterWidth = 556
        Case "o"
            TimesBoldCharacterWidth = 500
        Case "p"
            TimesBoldCharacterWidth = 556
        Case "q"
            TimesBoldCharacterWidth = 556
        Case "r"
            TimesBoldCharacterWidth = 444
        Case "s"
            TimesBoldCharacterWidth = 389
        Case "t"
            TimesBoldCharacterWidth = 333
        Case "u"
            TimesBoldCharacterWidth = 556
        Case "v"
            TimesBoldCharacterWidth = 500
        Case "w"
            TimesBoldCharacterWidth = 722
        Case "x"
            TimesBoldCharacterWidth = 500
        Case "y"
            TimesBoldCharacterWidth = 500
        Case "z"
            TimesBoldCharacterWidth = 444

        Case "{"
            TimesBoldCharacterWidth = 394
        Case "|"
            TimesBoldCharacterWidth = 220
        Case "}"
            TimesBoldCharacterWidth = 394
        Case "~"
            TimesBoldCharacterWidth = 520

        Case Else
            TimesBoldCharacterWidth = 500

    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Times-Italic
'
' Rückgabe in 1/1000 em
'---------------------------------------------------------
Private Function TimesItalicCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            TimesItalicCharacterWidth = 250

        Case "!"
            TimesItalicCharacterWidth = 333
        Case """"
            TimesItalicCharacterWidth = 420
        Case "#"
            TimesItalicCharacterWidth = 500
        Case "$"
            TimesItalicCharacterWidth = 500
        Case "%"
            TimesItalicCharacterWidth = 833
        Case "&"
            TimesItalicCharacterWidth = 778
        Case "'"
            TimesItalicCharacterWidth = 214
        Case "("
            TimesItalicCharacterWidth = 333
        Case ")"
            TimesItalicCharacterWidth = 333
        Case "*"
            TimesItalicCharacterWidth = 500
        Case "+"
            TimesItalicCharacterWidth = 675
        Case ","
            TimesItalicCharacterWidth = 250
        Case "-"
            TimesItalicCharacterWidth = 333
        Case "."
            TimesItalicCharacterWidth = 250
        Case "/"
            TimesItalicCharacterWidth = 278

        Case "0", "1", "2", "3", "4", _
             "5", "6", "7", "8", "9"
            TimesItalicCharacterWidth = 500

        Case ":"
            TimesItalicCharacterWidth = 333
        Case ";"
            TimesItalicCharacterWidth = 333
        Case "<"
            TimesItalicCharacterWidth = 675
        Case "="
            TimesItalicCharacterWidth = 675
        Case ">"
            TimesItalicCharacterWidth = 675
        Case "?"
            TimesItalicCharacterWidth = 500
        Case "@"
            TimesItalicCharacterWidth = 920

        Case "A"
            TimesItalicCharacterWidth = 611
        Case "B"
            TimesItalicCharacterWidth = 611
        Case "C"
            TimesItalicCharacterWidth = 667
        Case "D"
            TimesItalicCharacterWidth = 722
        Case "E"
            TimesItalicCharacterWidth = 611
        Case "F"
            TimesItalicCharacterWidth = 611
        Case "G"
            TimesItalicCharacterWidth = 722
        Case "H"
            TimesItalicCharacterWidth = 722
        Case "I"
            TimesItalicCharacterWidth = 333
        Case "J"
            TimesItalicCharacterWidth = 444
        Case "K"
            TimesItalicCharacterWidth = 667
        Case "L"
            TimesItalicCharacterWidth = 556
        Case "M"
            TimesItalicCharacterWidth = 833
        Case "N"
            TimesItalicCharacterWidth = 667
        Case "O"
            TimesItalicCharacterWidth = 722
        Case "P"
            TimesItalicCharacterWidth = 611
        Case "Q"
            TimesItalicCharacterWidth = 722
        Case "R"
            TimesItalicCharacterWidth = 611
        Case "S"
            TimesItalicCharacterWidth = 500
        Case "T"
            TimesItalicCharacterWidth = 556
        Case "U"
            TimesItalicCharacterWidth = 722
        Case "V"
            TimesItalicCharacterWidth = 667
        Case "W"
            TimesItalicCharacterWidth = 944
        Case "X"
            TimesItalicCharacterWidth = 667
        Case "Y"
            TimesItalicCharacterWidth = 667
        Case "Z"
            TimesItalicCharacterWidth = 556

        Case "["
            TimesItalicCharacterWidth = 389
        Case "\"
            TimesItalicCharacterWidth = 278
        Case "]"
            TimesItalicCharacterWidth = 389
        Case "^"
            TimesItalicCharacterWidth = 422
        Case "_"
            TimesItalicCharacterWidth = 500
        Case "`"
            TimesItalicCharacterWidth = 333

        Case "a"
            TimesItalicCharacterWidth = 500
        Case "b"
            TimesItalicCharacterWidth = 500
        Case "c"
            TimesItalicCharacterWidth = 444
        Case "d"
            TimesItalicCharacterWidth = 500
        Case "e"
            TimesItalicCharacterWidth = 444
        Case "f"
            TimesItalicCharacterWidth = 278
        Case "g"
            TimesItalicCharacterWidth = 500
        Case "h"
            TimesItalicCharacterWidth = 500
        Case "i"
            TimesItalicCharacterWidth = 278
        Case "j"
            TimesItalicCharacterWidth = 278
        Case "k"
            TimesItalicCharacterWidth = 500
        Case "l"
            TimesItalicCharacterWidth = 278
        Case "m"
            TimesItalicCharacterWidth = 722
        Case "n"
            TimesItalicCharacterWidth = 500
        Case "o"
            TimesItalicCharacterWidth = 500
        Case "p"
            TimesItalicCharacterWidth = 500
        Case "q"
            TimesItalicCharacterWidth = 500
        Case "r"
            TimesItalicCharacterWidth = 389
        Case "s"
            TimesItalicCharacterWidth = 389
        Case "t"
            TimesItalicCharacterWidth = 278
        Case "u"
            TimesItalicCharacterWidth = 500
        Case "v"
            TimesItalicCharacterWidth = 500
        Case "w"
            TimesItalicCharacterWidth = 722
        Case "x"
            TimesItalicCharacterWidth = 500
        Case "y"
            TimesItalicCharacterWidth = 500
        Case "z"
            TimesItalicCharacterWidth = 444

        Case "{"
            TimesItalicCharacterWidth = 400
        Case "|"
            TimesItalicCharacterWidth = 275
        Case "}"
            TimesItalicCharacterWidth = 400
        Case "~"
            TimesItalicCharacterWidth = 541

        Case Else
            TimesItalicCharacterWidth = 500

    End Select

End Function

'---------------------------------------------------------
' Zeichenbreite Times-BoldItalic
'
' Rückgabe in 1/1000 em
'---------------------------------------------------------
Private Function TimesBoldItalicCharacterWidth(ByVal Ch As String) As Double

    Select Case Ch

        Case " "
            TimesBoldItalicCharacterWidth = 250

        Case "!"
            TimesBoldItalicCharacterWidth = 389
        Case """"
            TimesBoldItalicCharacterWidth = 555
        Case "#"
            TimesBoldItalicCharacterWidth = 500
        Case "$"
            TimesBoldItalicCharacterWidth = 500
        Case "%"
            TimesBoldItalicCharacterWidth = 833
        Case "&"
            TimesBoldItalicCharacterWidth = 778
        Case "'"
            TimesBoldItalicCharacterWidth = 278
        Case "("
            TimesBoldItalicCharacterWidth = 333
        Case ")"
            TimesBoldItalicCharacterWidth = 333
        Case "*"
            TimesBoldItalicCharacterWidth = 500
        Case "+"
            TimesBoldItalicCharacterWidth = 570
        Case ","
            TimesBoldItalicCharacterWidth = 250
        Case "-"
            TimesBoldItalicCharacterWidth = 333
        Case "."
            TimesBoldItalicCharacterWidth = 250
        Case "/"
            TimesBoldItalicCharacterWidth = 278

        Case "0", "1", "2", "3", "4", _
             "5", "6", "7", "8", "9"
            TimesBoldItalicCharacterWidth = 500

        Case ":"
            TimesBoldItalicCharacterWidth = 333
        Case ";"
            TimesBoldItalicCharacterWidth = 333
        Case "<"
            TimesBoldItalicCharacterWidth = 570
        Case "="
            TimesBoldItalicCharacterWidth = 570
        Case ">"
            TimesBoldItalicCharacterWidth = 570
        Case "?"
            TimesBoldItalicCharacterWidth = 500
        Case "@"
            TimesBoldItalicCharacterWidth = 832

        Case "A"
            TimesBoldItalicCharacterWidth = 667
        Case "B"
            TimesBoldItalicCharacterWidth = 667
        Case "C"
            TimesBoldItalicCharacterWidth = 667
        Case "D"
            TimesBoldItalicCharacterWidth = 722
        Case "E"
            TimesBoldItalicCharacterWidth = 667
        Case "F"
            TimesBoldItalicCharacterWidth = 667
        Case "G"
            TimesBoldItalicCharacterWidth = 722
        Case "H"
            TimesBoldItalicCharacterWidth = 778
        Case "I"
            TimesBoldItalicCharacterWidth = 389
        Case "J"
            TimesBoldItalicCharacterWidth = 500
        Case "K"
            TimesBoldItalicCharacterWidth = 667
        Case "L"
            TimesBoldItalicCharacterWidth = 611
        Case "M"
            TimesBoldItalicCharacterWidth = 944
        Case "N"
            TimesBoldItalicCharacterWidth = 722
        Case "O"
            TimesBoldItalicCharacterWidth = 722
        Case "P"
            TimesBoldItalicCharacterWidth = 611
        Case "Q"
            TimesBoldItalicCharacterWidth = 722
        Case "R"
            TimesBoldItalicCharacterWidth = 667
        Case "S"
            TimesBoldItalicCharacterWidth = 556
        Case "T"
            TimesBoldItalicCharacterWidth = 667
        Case "U"
            TimesBoldItalicCharacterWidth = 722
        Case "V"
            TimesBoldItalicCharacterWidth = 667
        Case "W"
            TimesBoldItalicCharacterWidth = 944
        Case "X"
            TimesBoldItalicCharacterWidth = 667
        Case "Y"
            TimesBoldItalicCharacterWidth = 667
        Case "Z"
            TimesBoldItalicCharacterWidth = 611

        Case "["
            TimesBoldItalicCharacterWidth = 333
        Case "\"
            TimesBoldItalicCharacterWidth = 278
        Case "]"
            TimesBoldItalicCharacterWidth = 333
        Case "^"
            TimesBoldItalicCharacterWidth = 570
        Case "_"
            TimesBoldItalicCharacterWidth = 500
        Case "`"
            TimesBoldItalicCharacterWidth = 333

        Case "a"
            TimesBoldItalicCharacterWidth = 500
        Case "b"
            TimesBoldItalicCharacterWidth = 500
        Case "c"
            TimesBoldItalicCharacterWidth = 444
        Case "d"
            TimesBoldItalicCharacterWidth = 500
        Case "e"
            TimesBoldItalicCharacterWidth = 444
        Case "f"
            TimesBoldItalicCharacterWidth = 333
        Case "g"
            TimesBoldItalicCharacterWidth = 500
        Case "h"
            TimesBoldItalicCharacterWidth = 556
        Case "i"
            TimesBoldItalicCharacterWidth = 278
        Case "j"
            TimesBoldItalicCharacterWidth = 278
        Case "k"
            TimesBoldItalicCharacterWidth = 500
        Case "l"
            TimesBoldItalicCharacterWidth = 278
        Case "m"
            TimesBoldItalicCharacterWidth = 778
        Case "n"
            TimesBoldItalicCharacterWidth = 556
        Case "o"
            TimesBoldItalicCharacterWidth = 500
        Case "p"
            TimesBoldItalicCharacterWidth = 500
        Case "q"
            TimesBoldItalicCharacterWidth = 500
        Case "r"
            TimesBoldItalicCharacterWidth = 389
        Case "s"
            TimesBoldItalicCharacterWidth = 389
        Case "t"
            TimesBoldItalicCharacterWidth = 278
        Case "u"
            TimesBoldItalicCharacterWidth = 556
        Case "v"
            TimesBoldItalicCharacterWidth = 500
        Case "w"
            TimesBoldItalicCharacterWidth = 722
        Case "x"
            TimesBoldItalicCharacterWidth = 500
        Case "y"
            TimesBoldItalicCharacterWidth = 500
        Case "z"
            TimesBoldItalicCharacterWidth = 444

        Case "{"
            TimesBoldItalicCharacterWidth = 348
        Case "|"
            TimesBoldItalicCharacterWidth = 220
        Case "}"
            TimesBoldItalicCharacterWidth = 348
        Case "~"
            TimesBoldItalicCharacterWidth = 570

        Case Else
            TimesBoldItalicCharacterWidth = 500

    End Select

End Function

