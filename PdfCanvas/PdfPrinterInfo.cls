VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPrinterInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' Klasse.....: PdfPrinterInfo
' Aufgabe....: Interne Beschreibung eines Windows-Druckers
'
' Nur interne Verwendung.
'
' Version....: 1.1.0
'=========================================================

Private Declare Function OpenPrinter Lib "winspool.drv" _
    Alias "OpenPrinterA" ( _
        ByVal pPrinterName As String, _
        phPrinter As Long, _
        ByVal pDefault As Long) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" ( _
        ByVal hPrinter As Long) As Long
        
Private Declare Function GetPrinter Lib "winspool.drv" _
    Alias "GetPrinterA" ( _
        ByVal hPrinter As Long, _
        ByVal Level As Long, _
        pPrinter As Any, _
        ByVal cbBuf As Long, _
        pcbNeeded As Long) As Long
        
Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        Destination As Any, _
        Source As Any, _
        ByVal Length As Long)
        
Private Declare Function lstrlen Lib "kernel32" _
    Alias "lstrlenA" ( _
        ByVal lpString As Long) As Long
        
Private Type PRINTER_INFO_2
    pServerName As Long
    pPrinterName As Long
    pShareName As Long
    pPortName As Long
    pDriverName As Long
    pComment As Long
    pLocation As Long
    pDevMode As Long
    pSepFile As Long
    pPrintProcessor As Long
    pDatatype As Long
    pParameters As Long

    pSecurityDescriptor As Long

    Attributes As Long
    Priority As Long
    DefaultPriority As Long

    StartTime As Long
    UntilTime As Long

    Status As Long

    cJobs As Long
    AveragePPM As Long
End Type
        
Private mPrinterName As String
Private mDriverName As String
Private mPortName As String
Private mIsDefault As Boolean
Private mDeviceMode As Long
Private mAttributes As Long

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
' Treibername
'---------------------------------------------------------
Public Property Get DriverName() As String

    DriverName = mDriverName

End Property

Public Property Let DriverName(ByVal Value As String)

    mDriverName = Trim$(Value)

End Property

'---------------------------------------------------------
' Anschluss
'---------------------------------------------------------
Public Property Get PortName() As String

    PortName = mPortName

End Property

Public Property Let PortName(ByVal Value As String)

    mPortName = Trim$(Value)

End Property

'---------------------------------------------------------
' Standarddrucker
'---------------------------------------------------------
Public Property Get IsDefault() As Boolean

    IsDefault = mIsDefault

End Property

Public Property Let IsDefault(ByVal Value As Boolean)

    mIsDefault = Value

End Property

'---------------------------------------------------------
' Druckerinformationen aus Windows laden
'---------------------------------------------------------
'Friend Function LoadFromWindows( _
'        ByVal PrinterName As String) As Boolean
'
'End Function

'Friend Function LoadFromWindows( _
'        ByVal PrinterName As String) As Boolean
'
'    Dim hPrinter As Long
'
'    If OpenPrinter(PrinterName, hPrinter, 0) = 0 Then
'
'        Exit Function
'
'    End If
'
'    ClosePrinter hPrinter
'
'    LoadFromWindows = True
'
'End Function
Friend Function LoadFromWindows( _
        ByVal PrinterName As String) As Boolean

    Dim hPrinter As Long
    Dim Needed As Long
    Dim Buffer() As Byte
    Dim PI2 As PRINTER_INFO_2
    
    If OpenPrinter(PrinterName, hPrinter, 0) = 0 Then

        Exit Function

    End If

    'Benötigte Speichergröße ermitteln
    GetPrinter hPrinter, _
               2, _
               ByVal 0&, _
               0, _
               Needed

    ReDim Buffer(0 To Needed - 1)
    
    If GetPrinter( _
        hPrinter, _
        2, _
        Buffer(0), _
        Needed, _
        Needed) = 0 Then

        ClosePrinter hPrinter
    
        Exit Function
    
    End If
    
    CopyMemory PI2, Buffer(0), Len(PI2)
    
    mPrinterName = PtrToStringA(PI2.pPrinterName)
    mDriverName = PtrToStringA(PI2.pDriverName)
    mPortName = PtrToStringA(PI2.pPortName)
    
    mAttributes = PI2.Attributes
    mDeviceMode = PI2.pDevMode
    
    ClosePrinter hPrinter


    LoadFromWindows = True

End Function

'---------------------------------------------------------
' Wandelt einen ANSI-Stringpointer in einen VB-String um
'---------------------------------------------------------
Private Function PtrToStringA( _
        ByVal lpString As Long) As String

    Dim L As Long

    If lpString = 0 Then Exit Function

    L = lstrlen(lpString)

    If L = 0 Then Exit Function

    PtrToStringA = String$(L, vbNullChar)

    CopyMemory ByVal PtrToStringA, _
               ByVal lpString, _
               L

End Function

