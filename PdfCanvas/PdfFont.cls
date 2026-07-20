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

Private mName As String
Private mAlias As String
Private mSize As Double
Private mObjectNumber As Long

'---------------------------------------------------------
' PDF-Name
' z.B. Helvetica
'---------------------------------------------------------
Public Property Get Name() As String

    Name = mName

End Property

Public Property Let Name(ByVal Value As String)

    mName = Value

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
