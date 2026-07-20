VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfBinaryFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfBinaryFile
' Aufgabe....: Binärdateien laden und speichern
'
' Version....: 0.9.5
'=========================================================



'---------------------------------------------------------
' Binärdatei laden
'---------------------------------------------------------

Public Function Load(ByVal FileName As String) As Byte()

    Dim FF As Integer
    Dim Data() As Byte

    FF = FreeFile

    Open FileName For Binary Access Read As #FF

        If LOF(FF) > 0 Then

            'Data = String$(LOF(FF), Chr$(0))
            ReDim Data(0 To LOF(FF) - 1)

            Get #FF, , Data

        End If

    Close #FF

    Load = Data

End Function

'---------------------------------------------------------
' Binärdatei speichern
'---------------------------------------------------------
Public Sub Save( _
        ByVal FileName As String, _
        ByVal Data As String)

    Dim FF As Integer

    FF = FreeFile

    Open FileName For Binary Access Write As #FF

        Put #FF, , Data

    Close #FF

End Sub

