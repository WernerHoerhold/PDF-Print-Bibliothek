VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Klasse.....: PdfObject
' Aufgabe....: Beschreibt ein PDF-Objekt
'
' Version....: 0.6.0
'=========================================================

Private mNumber As Long
Private mPosition As Long
Private mKind As Long

'---------------------------------------------------------
' Objektnummer
'---------------------------------------------------------
Public Property Get Number() As Long

    Number = mNumber

End Property

Public Property Let Number(ByVal Value As Long)

    mNumber = Value

End Property

'---------------------------------------------------------
' Byteposition
'---------------------------------------------------------
Public Property Get Position() As Long

    Position = mPosition

End Property

Public Property Let Position(ByVal Value As Long)

    mPosition = Value

End Property

'---------------------------------------------------------
' Objekttyp
'---------------------------------------------------------
Public Property Get Kind() As Long

    Kind = mKind

End Property

Public Property Let Kind(ByVal Value As Long)

    mKind = Value

End Property
