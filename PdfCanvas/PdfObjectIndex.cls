VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfObjectIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================
' KLASSE......: PdfObjectIndex
' KATEGORIE...: B - UNTERSTÜTZUNGSKLASSE
' VERSION.....: 0.6.0
'
' BESCHREIBUNG:
' Verwaltet alle PDF-Objekte und deren Positionen.
' Diese Klasse führt ein Verzeichnis aller Objekte
' für die Cross-Reference-Tabelle (XRef).
'
' VERWENDUNG: Intern durch PdfWriter
' AUTOR: Kapp Niles GmbH & Co, KG
'===============================================================

Private mObjects As Collection

'---------------------------------------------------------
' Initialisierung
'---------------------------------------------------------
Private Sub Class_Initialize()

    Set mObjects = New Collection

End Sub

'---------------------------------------------------------
' Alle Einträge löschen
'---------------------------------------------------------
Public Sub Clear()

    Set mObjects = New Collection

End Sub

'---------------------------------------------------------
' Neues PDF-Objekt hinzufügen
'---------------------------------------------------------
Public Sub Add( _
        ByVal Number As Long, _
        ByVal Position As Long, _
        ByVal Kind As Long)

    Dim Obj As PdfObject

    On Error Resume Next
    mObjects.Remove CStr(Number)
    On Error GoTo 0
    
    Set Obj = New PdfObject

    Obj.Number = Number
    Obj.Position = Position
    Obj.Kind = Kind
    
    On Error Resume Next
    mObjects.Remove CStr(Number)
    On Error GoTo 0
    
    mObjects.Add Obj, CStr(Number)

End Sub

'---------------------------------------------------------
' Anzahl der Objekte
'---------------------------------------------------------
Public Property Get Count() As Long

    Count = mObjects.Count

End Property

'---------------------------------------------------------
' Objekt liefern
'---------------------------------------------------------
Public Property Get Item(ByVal Number As Long) As PdfObject

    Set Item = mObjects(CStr(Number))

End Property

'---------------------------------------------------------
' Position eines Objektes
'---------------------------------------------------------
Public Property Get Position(ByVal Number As Long) As Long

    Position = mObjects(CStr(Number)).Position

End Property

'---------------------------------------------------------
' Objekttyp liefern
'---------------------------------------------------------
Public Property Get Kind(ByVal Number As Long) As Long

    Kind = mObjects(CStr(Number)).Kind

End Property

