VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfFile
' Aufgabe....: Speichert Textdateien
'=========================================================

Public Sub Save(ByVal FileName As String, _
                ByVal Text As String)

    Dim F As Integer

    F = FreeFile

    Open FileName For Binary Access Write As #F

    Put #F, , Text

    Close #F

End Sub
