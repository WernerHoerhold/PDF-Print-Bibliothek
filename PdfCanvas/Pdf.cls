VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: Pdf
' Aufgabe....: Einstiegspunkt der PDF-Bibliothek
'
' Version....: 1.0
'=========================================================

Public Function CreateDocument() As PdfDocument

    Set CreateDocument = New PdfDocument

End Function

Public Function LoadImage(ByVal FileName As String) As PdfImage

    Dim Img As PdfImage

    Set Img = New PdfImage

    Img.LoadFromFile FileName

    Set LoadImage = Img

End Function


Public Function CreateImage() As PdfImage

End Function

Public Function CreateFont() As PdfFont

End Function

Public Function CreatePointCollection() As PdfPointCollection

    Set CreatePointCollection = New PdfPointCollection

End Function

Public Function CreateBrush() As PdfBrush

End Function

Public Function CreatePen() As PdfPen

End Function

Public Property Get Version() As String

    Version = PDF_VERSION

End Property

Public Property Get Build() As String

    Build = PDF_BUILD

End Property

Public Property Get Name() As String

    Name = PDF_NAME

End Property

Public Property Get Copyright() As String

    Copyright = PDF_COPYRIGHT

End Property

Public Property Get About() As String

    About = PDF_NAME & vbCrLf & _
            " Version " & PDF_VERSION & vbCrLf & _
            " (Build " & PDF_BUILD & ")" & vbCrLf & vbCrLf & _
            PDF_COPYRIGHT

End Property
