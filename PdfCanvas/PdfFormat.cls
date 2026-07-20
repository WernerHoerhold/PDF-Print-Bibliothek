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
' Version....: 0.7.0
'=========================================================

'---------------------------------------------------------
' Zahl im PDF-Format
' (immer Dezimalpunkt)
'---------------------------------------------------------
Public Function Number(ByVal Value As Double) As String

    Dim S As String

    S = Trim$(Str$(Value))

    If InStr(S, ",") > 0 Then
        S = Replace$(S, ",", ".")
    End If

    Number = S

End Function

'---------------------------------------------------------
' PDF-Transformationsmatrix
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

