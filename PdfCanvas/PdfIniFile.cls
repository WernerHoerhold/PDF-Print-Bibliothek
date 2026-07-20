VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfIniFile
' Aufgabe....: Lesen und Schreiben von INI-Dateien
'
' Nur interne Verwendung.
'
' Version....: 1.1.0
'=========================================================


Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpString As String, _
        ByVal lpFileName As String) As Long


'---------------------------------------------------------
' Zeichenkette aus INI-Datei lesen
'---------------------------------------------------------
Public Function ReadString( _
        ByVal FileName As String, _
        ByVal Section As String, _
        ByVal Key As String, _
        Optional ByVal DefaultValue As String = "") As String

    Dim Buffer As String
    Dim Length As Long

    Buffer = String$(512, vbNullChar)

    Length = GetPrivateProfileString( _
                Section, _
                Key, _
                DefaultValue, _
                Buffer, _
                Len(Buffer), _
                FileName)

    ReadString = Left$(Buffer, Length)

End Function

'---------------------------------------------------------
' Zeichenkette in INI-Datei schreiben
'---------------------------------------------------------
Public Sub WriteString( _
        ByVal FileName As String, _
        ByVal Section As String, _
        ByVal Key As String, _
        ByVal Value As String)

    Call WritePrivateProfileString( _
            Section, _
            Key, _
            Value, _
            FileName)

End Sub

'---------------------------------------------------------
' Ganzzahl aus INI-Datei lesen
'---------------------------------------------------------
Public Function ReadInteger( _
        ByVal FileName As String, _
        ByVal Section As String, _
        ByVal Key As String, _
        Optional ByVal DefaultValue As Long = 0) As Long

    ReadInteger = Val(ReadString( _
                        FileName, _
                        Section, _
                        Key, _
                        CStr(DefaultValue)))

End Function

'---------------------------------------------------------
' Ganzzahl in INI-Datei schreiben
'---------------------------------------------------------
Public Sub WriteInteger( _
        ByVal FileName As String, _
        ByVal Section As String, _
        ByVal Key As String, _
        ByVal Value As Long)

    WriteString _
        FileName, _
        Section, _
        Key, _
        CStr(Value)

End Sub
