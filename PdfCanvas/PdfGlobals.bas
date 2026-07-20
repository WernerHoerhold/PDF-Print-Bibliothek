Attribute VB_Name = "PdfGlobals"
Option Explicit

Private mPageSizes As PdfPageSizes
Private mOrientations As PdfOrientations

Public Function PageSizes() As PdfPageSizes

    If mPageSizes Is Nothing Then
        Set mPageSizes = New PdfPageSizes
    End If

    Set PageSizes = mPageSizes

End Function

Public Function Orientations() As PdfOrientations

    If mOrientations Is Nothing Then
        Set mOrientations = New PdfOrientations
    End If

    Set Orientations = mOrientations

End Function
