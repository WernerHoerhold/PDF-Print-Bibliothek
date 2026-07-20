VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPointCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfPointCollection
' Aufgabe....: Sammlung von PdfPoint-Objekten
'
' Version....: 0.9.4
'=========================================================

Private mPoints As Collection

Private Sub Class_Initialize()

    Set mPoints = New Collection

End Sub

Public Sub Add(ByVal Point As PdfPoint)

    mPoints.Add Point.Clone

End Sub

Public Sub AddXY( _
        ByVal X As Double, _
        ByVal Y As Double)

    Dim P As PdfPoint

    Set P = New PdfPoint

    P.Assign X, Y

    mPoints.Add P

End Sub

Public Property Get Count() As Long

    Count = mPoints.Count

End Property

Public Property Get Item( _
        ByVal Index As Long) As PdfPoint

    Set Item = mPoints(Index)

End Property

Public Sub Clear()

    Set mPoints = New Collection

End Sub

'---------------------------------------------------------
' Punktliste kopieren
'---------------------------------------------------------
Public Function Clone() As PdfPointCollection

    Dim C As PdfPointCollection
    Dim I As Long

    Set C = New PdfPointCollection

    For I = 1 To mPoints.Count

        C.Add mPoints(I).Clone

    Next

    Set Clone = C

End Function

Public Property Get NewEnum() As IUnknown

    Set NewEnum = mPoints.[_NewEnum]


End Property
