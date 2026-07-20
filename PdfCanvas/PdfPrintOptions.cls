VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfPrintOptions
' Aufgabe....: Druckoptionen für PDF-Dokumente
'
' Version....: 1.1.0
'=========================================================

Private mPrinterName As String

Private mCopies As Long

Private mShowDialog As Boolean

Private mDeleteTempFile As Boolean

Private mOpenAfterPrint As Boolean

Private mPageFrom As Long

Private mPageTo As Long

Private Sub Class_Initialize()

    mPrinterName = ""

    mCopies = 1

    mShowDialog = False

    mDeleteTempFile = True

    mOpenAfterPrint = False

    mPageFrom = 0
    mPageTo = 0

End Sub

'---------------------------------------------------------
' Druckername
'---------------------------------------------------------
Public Property Get PrinterName() As String

    PrinterName = mPrinterName

End Property

Public Property Let PrinterName(ByVal Value As String)

    mPrinterName = Trim$(Value)

End Property

'---------------------------------------------------------
' Anzahl der Druckexemplare
'---------------------------------------------------------
Public Property Get Copies() As Long

    Copies = mCopies

End Property

Public Property Let Copies(ByVal Value As Long)

    If Value < 1 Then Value = 1

    mCopies = Value

End Property

'---------------------------------------------------------
' Druckdialog anzeigen
'---------------------------------------------------------
Public Property Get ShowDialog() As Boolean

    ShowDialog = mShowDialog

End Property

Public Property Let ShowDialog(ByVal Value As Boolean)

    mShowDialog = Value

End Property

'---------------------------------------------------------
' Temporäre PDF nach dem Drucken löschen
'---------------------------------------------------------
Public Property Get DeleteTempFile() As Boolean

    DeleteTempFile = mDeleteTempFile

End Property

Public Property Let DeleteTempFile(ByVal Value As Boolean)

    mDeleteTempFile = Value

End Property

'---------------------------------------------------------
' PDF nach dem Drucken öffnen
'---------------------------------------------------------
Public Property Get OpenAfterPrint() As Boolean

    OpenAfterPrint = mOpenAfterPrint

End Property

Public Property Let OpenAfterPrint(ByVal Value As Boolean)

    mOpenAfterPrint = Value

End Property

'---------------------------------------------------------
' Erste zu druckende Seite
'---------------------------------------------------------
Public Property Get PageFrom() As Long

    PageFrom = mPageFrom

End Property

Public Property Let PageFrom(ByVal Value As Long)

    If Value < 0 Then Value = 0

    mPageFrom = Value

End Property

'---------------------------------------------------------
' Letzte zu druckende Seite
'---------------------------------------------------------
Public Property Get PageTo() As Long

    PageTo = mPageTo

End Property

Public Property Let PageTo(ByVal Value As Long)

    If Value < 0 Then Value = 0

    mPageTo = Value

End Property

'---------------------------------------------------------
' Druckoptionen in INI-Datei speichern
'---------------------------------------------------------
Public Sub SaveToIni(ByVal FileName As String)

    Dim Ini As New PdfIniFile

    Ini.WriteString FileName, _
                    "Printer", _
                    "Name", _
                    mPrinterName

    Ini.WriteInteger FileName, _
                     "Printer", _
                     "Copies", _
                     mCopies

    Ini.WriteInteger FileName, _
                     "Printer", _
                     "ShowDialog", _
                     Abs(mShowDialog)

    Ini.WriteInteger FileName, _
                     "Printer", _
                     "DeleteTempFile", _
                     Abs(mDeleteTempFile)

    Ini.WriteInteger FileName, _
                     "Printer", _
                     "OpenAfterPrint", _
                     Abs(mOpenAfterPrint)

    Ini.WriteInteger FileName, _
                     "Printer", _
                     "PageFrom", _
                     mPageFrom

    Ini.WriteInteger FileName, _
                     "Printer", _
                     "PageTo", _
                     mPageTo

End Sub

'---------------------------------------------------------
' Druckoptionen aus INI-Datei laden
'---------------------------------------------------------
Public Sub LoadFromIni(ByVal FileName As String)

    Dim Ini As New PdfIniFile

    mPrinterName = Ini.ReadString( _
                        FileName, _
                        "Printer", _
                        "Name", _
                        "")

    mCopies = Ini.ReadInteger( _
                    FileName, _
                    "Printer", _
                    "Copies", _
                    1)

    mShowDialog = (Ini.ReadInteger( _
                        FileName, _
                        "Printer", _
                        "ShowDialog", _
                        0) <> 0)

    mDeleteTempFile = (Ini.ReadInteger( _
                            FileName, _
                            "Printer", _
                            "DeleteTempFile", _
                            1) <> 0)

    mOpenAfterPrint = (Ini.ReadInteger( _
                            FileName, _
                            "Printer", _
                            "OpenAfterPrint", _
                            0) <> 0)

    mPageFrom = Ini.ReadInteger( _
                    FileName, _
                    "Printer", _
                    "PageFrom", _
                    0)

    mPageTo = Ini.ReadInteger( _
                  FileName, _
                  "Printer", _
                  "PageTo", _
                  0)

End Sub
