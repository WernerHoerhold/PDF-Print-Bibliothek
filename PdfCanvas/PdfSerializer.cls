VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mBuffer As PdfBuffer

Friend Sub SetBuffer(ByVal Buffer As PdfBuffer)

    Set mBuffer = Buffer

End Sub

Public Sub BeginObject(ByVal Number As Long)

    mBuffer.AddLine CStr(Number) & " 0 obj"

End Sub

Public Sub EndObject()

    mBuffer.AddLine "endobj"
    mBuffer.AddLine ""

End Sub

Public Sub DictionaryBegin()

    mBuffer.AddLine "<<"

End Sub

Public Sub DictionaryEnd()

    mBuffer.AddLine ">>"

End Sub

'---------------------------------------------------------
' Array beginnen
'---------------------------------------------------------
Public Sub BeginArray()

    mBuffer.Add "["

End Sub

'---------------------------------------------------------
' Array beenden
'---------------------------------------------------------
Public Sub EndArray()

    mBuffer.AddLine "]"

End Sub

'---------------------------------------------------------
' Stream beginnen
'---------------------------------------------------------
'Public Sub BeginStream(ByVal StreamLength As Long)
Public Sub BeginStream()

    'DictionaryBegin

    'AddInteger "Length", StreamLength

    'DictionaryEnd

    AddRaw "stream"

End Sub

'---------------------------------------------------------
' Stream beenden
'---------------------------------------------------------

'---------------------------------------------------------
' Stream innerhalb eines bereits geöffneten Dictionarys
'---------------------------------------------------------
'Public Sub BeginStreamInDictionary(ByVal StreamLength As Long)
'
'    AddInteger "Length", StreamLength
'
'    DictionaryEnd
'
'    AddRaw "stream"
'
'End Sub
Public Sub EndStream()

    AddRaw "endstream"

End Sub

Public Sub AddName(ByVal Name As String, _
                   ByVal Value As String)

    mBuffer.AddLine "/" & Name & " " & Value

End Sub

Public Sub AddInteger(ByVal Name As String, _
                      ByVal Value As Long)

    mBuffer.AddLine "/" & Name & " " & CStr(Value)

End Sub

Public Sub AddReference(ByVal Name As String, _
                        ByVal ObjectNumber As Long)

    mBuffer.AddLine "/" & Name & " " & _
                    CStr(ObjectNumber) & " 0 R"

End Sub

'----------------------------------------------------------
' Beliebigen Eintrag schreiben
'----------------------------------------------------------
Public Sub AddValue(ByVal Name As String, _
                    ByVal Value As String)

    mBuffer.AddLine "/" & Name & " " & Value

End Sub

Public Sub AddRaw(ByVal Text As String)

    mBuffer.AddLine Text

End Sub

