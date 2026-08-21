VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfFormat
' Aufgabe....: Formatierung von PDF-Werten
'
' Version....: 0.7.1 (mit verbesserter Exponentialbehandlung)
'=========================================================

'---------------------------------------------------------
' Zahl im PDF-Format
' (immer Dezimalpunkt, keine Exponentialdarstellung)
'---------------------------------------------------------
Public Function Number(ByVal Value As Double) As String

    Dim S As String
    Dim DecimalPos As Long

    'Wenn Wert 0 ist, direkt zurückgeben
    If Value = 0 Then
        Number = "0"
        Exit Function
    End If

    'Mit Format$ arbeiten um Exponentialdarstellung zu vermeiden
    'Format mit fester Dezimalstelle (bis zu 6 Dezimalstellen)
    S = Format$(Value, "0.######")

    'Komma durch Punkt ersetzen (für deutsche Lokalisierung)
    If InStr(S, ",") > 0 Then
        S = Replace$(S, ",", ".")
    End If

    'Führende Null entfernen wenn nötig (z.B. ".5" statt "0.5")
    'Aber nur wenn Wert zwischen -1 und 1 liegt
    If Left$(S, 2) = "0." Then
        'Bleibt "0.xxx"
    ElseIf Left$(S, 3) = "-0." Then
        'Bleibt "-0.xxx"
    End If

    Number = S

End Function

'---------------------------------------------------------
' PDF-Transformationsmatrix
' (garantiert keine Exponentialdarstellung)
'---------------------------------------------------------
Public Function Matrix( _
        ByVal A As Double, _
        ByVal B As Double, _
        ByVal C As Double, _
        ByVal D As Double, _
        ByVal E As Double, _
        ByVal F As Double) As String

    Matrix = _
        Number(A) & " " & _
        Number(B) & " " & _
        Number(C) & " " & _
        Number(D) & " " & _
        Number(E) & " " & _
        Number(F)

End Function

