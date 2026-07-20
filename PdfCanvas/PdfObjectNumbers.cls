VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfObjectNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfObjectNumbers
' Aufgabe....: Vergibt PDF-Objektnummern
'
' Version....: 0.9.5
'=========================================================

Private mNextNumber As Long

Private Sub Class_Initialize()

    mNextNumber = 1

End Sub

'---------------------------------------------------------
' Nummerierung zurücksetzen
'---------------------------------------------------------
Public Sub Clear()

    mNextNumber = 1

End Sub

'---------------------------------------------------------
' Nächste Objektnummer
'---------------------------------------------------------
Public Function NextNumber() As Long

    NextNumber = mNextNumber

    mNextNumber = mNextNumber + 1

End Function
