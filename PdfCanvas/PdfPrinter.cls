VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfPrinter
' Aufgabe....: Drucken von PDF-Dateien und PdfDocument
'
' Version....: 1.1.0
'=========================================================

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
        
Private Declare Function FindExecutable Lib "shell32.dll" _
    Alias "FindExecutableA" ( _
        ByVal lpFile As String, _
        ByVal lpDirectory As String, _
        ByVal lpResult As String) As Long
        
Private Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" ( _
        ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, _
        ByVal cchBuffer As Long) As Long
        
        
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
        
Private Type STARTUPINFO
    cb As Long

    lpReserved As String
    lpDesktop As String
    lpTitle As String

    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long

    dwXCountChars As Long
    dwYCountChars As Long

    dwFillAttribute As Long

    dwFlags As Long

    wShowWindow As Integer
    cbReserved2 As Integer

    lpReserved2 As Long

    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type


Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long

    dwProcessId As Long
    dwThreadId As Long
End Type


Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long

    hwnd As Long

    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String

    nShow As Long

    hInstApp As Long

    lpIDList As Long
    lpClass As String

    hkeyClass As Long

    dwHotKey As Long

    hIcon As Long
    hProcess As Long
End Type

Private mLastErrorNumber As Long
Private mLastErrorText As String
Private mConfigurationFile As String


Private Sub Class_Initialize()

    mConfigurationFile = ""

End Sub

Public Property Get ConfigurationFile() As String

    ConfigurationFile = mConfigurationFile

End Property

Public Property Let ConfigurationFile( _
        ByVal Value As String)

    mConfigurationFile = Trim$(Value)

End Property

'---------------------------------------------------------
' PDF-Datei drucken
'---------------------------------------------------------
'Public Sub PrintFile(ByVal FileName As String)
Public Function PrintFile(ByVal FileName As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    ClearLastError
    
    Dim Opt As New PdfPrintOptions
    
    If Trim$(mConfigurationFile) <> "" Then
    
        If Len(Dir$(mConfigurationFile)) > 0 Then
    
            Opt.LoadFromIni mConfigurationFile
    
        End If
    
    End If
    
    PrintFile = PrintFileOptions(FileName, Opt)

    Exit Function
    
ErrorHandler:
    SetLastError Err.Number, Err.Description
    PrintFile = False
End Function

'---------------------------------------------------------
' PDF-Datei mit Druckoptionen drucken
'---------------------------------------------------------
Public Function PrintFileOptions( _
        ByVal FileName As String, _
        ByVal Options As PdfPrintOptions) As Boolean
        
    On Error GoTo ErrorHandler

    ClearLastError

    Dim Viewer As String

    'PDF vorhanden?
    If Len(Dir$(FileName)) = 0 Then

        SetLastError vbObjectError + 1104, "PDF-File not found"
        PrintFileOptions = False
        Exit Function

    End If

    'Gewünschten Drucker prüfen
    If Trim$(Options.PrinterName) <> "" Then

        If Not PrinterExists(Options.PrinterName) Then

            SetLastError vbObjectError + 1105, "Printer not found:" & vbCrLf & Options.PrinterName
            PrintFileOptions = False
            Exit Function

        End If

    End If

    'PDF-Viewer ermitteln
    Viewer = GetPdfViewer(FileName)

    If Viewer = "" Then

        SetLastError vbObjectError + 1106, "No Pdf-Viewer (Adobe Reader) found"
        PrintFileOptions = False
        Exit Function

    End If

    Dim Info As PdfPrinterInfo

    If Trim$(Options.PrinterName) <> "" Then
        Set Info = FindPrinter(Options.PrinterName)
    End If
    
    If IsAdobeReader(Viewer) Then
        PrintWithAdobe Viewer, FileName, Info
    Else
        'PrintFile FileName        'oder später PrintWithShell
        'Ret = ShellExecute( _
        '            0, _
        '            "print", _
        '            FileName, _
        '            vbNullString, _
        '            vbNullString, _
        '            SW_HIDE)
        '
        'If Ret <= 32 Then
        '
        '    SetLastError vbObjectError + 1102, "PDF-File could not be printed."
        '
        '    PrintFile = False
        '    Exit Function
        '
        'End If

        SetLastError vbObjectError + 1107, "The installed Viewer is not supported (Only Adobe Reader)"
        PrintFileOptions = False
        Exit Function

    End If

    PrintFileOptions = True
    Exit Function

ErrorHandler:
    SetLastError Err.Number, Err.Description
    PrintFileOptions = False
End Function


'---------------------------------------------------------
' Ermittelt den registrierten PDF-Viewer
'---------------------------------------------------------
Private Function GetPdfViewer(ByVal FileName As String) As String
    
    Dim Buffer As String
    Dim Ret As Long

    Buffer = String$(260, vbNullChar)

    Ret = FindExecutable(FileName, vbNullString, Buffer)

    If Ret > 32 Then
        GetPdfViewer = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    Else
        GetPdfViewer = ""
    End If

End Function

'---------------------------------------------------------
' Prüft, ob Adobe Reader verwendet wird
'---------------------------------------------------------
Private Function IsAdobeReader(ByVal Viewer As String) As Boolean

    IsAdobeReader = _
        (InStr(1, _
               Viewer, _
               "AcroRd32.exe", _
               vbTextCompare) > 0)

End Function

'---------------------------------------------------------
' Drucker anhand seines Namens suchen
'---------------------------------------------------------
Private Function FindPrinter( _
        ByVal PrinterName As String) As PdfPrinterInfo

    Dim Prn As Printer
    Dim Info As PdfPrinterInfo

    For Each Prn In Printers

        If StrComp(Prn.DeviceName, _
                   PrinterName, _
                   vbTextCompare) = 0 Then

            Set Info = New PdfPrinterInfo
            
            If Info.LoadFromWindows(Prn.DeviceName) Then
                Set FindPrinter = Info
            End If
            
            Exit Function

        End If

    Next

End Function

'---------------------------------------------------------
' Prüft, ob ein Drucker vorhanden ist
'---------------------------------------------------------
Public Function PrinterExists( _
        ByVal PrinterName As String) As Boolean

    Dim Info As PdfPrinterInfo

    Set Info = FindPrinter(PrinterName)

    PrinterExists = Not (Info Is Nothing)

End Function

'---------------------------------------------------------
' Ermittelt den DOS-Kurzpfad
'---------------------------------------------------------
Private Function ShortPath(ByVal FileName As String) As String

    Dim Buffer As String
    Dim L As Long

    Buffer = String$(260, vbNullChar)

    L = GetShortPathName( _
            FileName, _
            Buffer, _
            Len(Buffer))

    If L > 0 Then

        ShortPath = Left$(Buffer, L)

    Else

        ShortPath = FileName

    End If

End Function

'---------------------------------------------------------
' String in Anführungszeichen setzen
'---------------------------------------------------------
Private Function Quote(ByVal Text As String) As String

    Quote = """" & Text & """"

End Function
'---------------------------------------------------------
' Erstellt die Befehlszeile für Adobe Reader
'---------------------------------------------------------
Private Function BuildAdobeCommand( _
        ByVal Viewer As String, _
        ByVal FileName As String, _
        ByVal Info As PdfPrinterInfo) As String

    If Info Is Nothing Then
    
        BuildAdobeCommand = _
            Quote(ShortPath(Viewer)) & _
            " /t " & Quote(ShortPath(FileName))
    
    Else
        
        BuildAdobeCommand = Quote(ShortPath(Viewer)) & _
                            " /t " & Quote(ShortPath(FileName)) & _
                            " " & Quote(Info.PrinterName) & _
                            " " & Quote(Info.DriverName) & _
                            " " & Quote(Info.PortName)
                            
    End If

End Function

'---------------------------------------------------------
' Drucken über Adobe Reader
'---------------------------------------------------------
Private Sub PrintWithAdobe( _
        ByVal Viewer As String, _
        ByVal FileName As String, _
        ByVal Info As PdfPrinterInfo)

    RunCommand BuildAdobeCommand(Viewer, FileName, Info)
    
End Sub

'---------------------------------------------------------
' Startet ein externes Programm
'---------------------------------------------------------
Private Sub RunCommand(ByVal Cmd As String)

    Shell Cmd, vbHide

End Sub

Public Property Get LastErrorNumber() As Long
    LastErrorNumber = mLastErrorNumber
End Property

Public Property Get LastErrorText() As String
    LastErrorText = mLastErrorText
End Property

Private Sub ClearLastError()

    mLastErrorNumber = 0
    mLastErrorText = ""

End Sub

Private Sub SetLastError( _
        ByVal Number As Long, _
        ByVal Description As String)

    mLastErrorNumber = Number
    mLastErrorText = Description

End Sub

