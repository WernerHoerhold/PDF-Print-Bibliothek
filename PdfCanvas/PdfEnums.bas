Attribute VB_Name = "PdfEnums"
Option Explicit

'=========================================================
' Modul......: PdfEnums
' Aufgabe....: Gemeinsame Enumerationen der PDF-Bibliothek
'
' Version....: 0.6.0
'=========================================================

'---------------------------------------------------------
' PDF-Objekttypen
'---------------------------------------------------------
Public Enum PdfObjectType
    potUnknown = 0
    potCatalog = 1
    potpages = 2
    potPage = 3
    potContent = 4
    potFont = 5
    potImage = 6
End Enum

Public Enum PdfPageSize
    psA5 = 0
    psA4 = 1
    psA3 = 2
    psLetter = 3
    psLegal = 4
End Enum

Public Enum PdfOrientation
    poPortrait = 0
    poLandscape = 1
End Enum
