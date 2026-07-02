VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "setsystemdateformat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LOCALE_SSHORTDATE = &H1F
Private Const HWND_BROADCAST = &HFFFF
Private Const WM_WININICHANGE = &H1A
Private Const PRINTER_ATTRIBUTE_SHARED As Long = &H8
Private Const PRINTER_ATTRIBUTE_LOCAL As Long = &H40
Private Const PRINTER_ATTRIBUTE_NETWORK = &H10
Private Const PRINTER_LEVEL2 As Long = &H2
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER As Long = &H4
Private Const PRINTER_ACCESS_USE As Long = &H8
Private Const PRINTER_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Type PRINTER_DEFAULTS
pDatatype As String
pDevMode As Long 'DEVMODE
DesiredAccess As Long
End Type
Private Type PRINTER_INFO_2
pServerName As String
pPrinterName As String
pShareName As String
pPortName As String
pDriverName As String
pComment As String
pLocation As String
pDevMode As Long 'DEVMODE
pSepFile As String
pPrintProcessor As String
pDatatype As String
pParameters As String
pSecurityDescriptor As Long 'SECURITY_DESCRIPTOR
Attributes As Long
Priority As Long
DefaultPriority As Long
StartTime As Long
UntilTime As Long
Status As Long
cJobs As Long
AveragePPM As Long
End Type
Private Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pServerName As String, ByVal Level As Long, pPrinter As Any) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
'Private Declare Function DeletePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long


Private Function InstallPrinter()
Dim hPrinter As Long
Dim sServer As String
Dim pi2 As PRINTER_INFO_2
With pi2
    .pServerName = vbNullString 'local machine
    .pPrinterName = "Generic / Text Only - DAS" 'display name for new printer
    .pShareName = vbNullString 'name for share - must be upper-case
    .pPortName = "LPT1:" 'printer port (NT or later can have multiple separated with commas, e.g. "LPT1:,LPT2:")
    .pDriverName = "Generic / Text Only" 'the name of an existing driver
    .pComment = "Default Generic Printer" 'a comment, if desired
    .pLocation = vbNullString 'the location, if desired
    .pPrintProcessor = "WinPrint" 'the print processor name - manditory
    .pDatatype = "RAW" 'type of data supported, e.g. RAW or EMF this must match type used by the driver - manditory
    .Attributes = PRINTER_ATTRIBUTE_LOCAL  'flags - local printer, with share created (Or PRINTER_ATTRIBUTE_SHARED)
End With
'install on the local machine by passing vbNullString as the server name. If installation on another machine is
'to be made, pass the machine name in the usual 'server name format', eg: sServer = "\laptopxp"
sServer = vbNullString
hPrinter = AddPrinter(sServer, PRINTER_LEVEL2, pi2)
'Debug.Print hPrinter, Err.LastDllError
If hPrinter <> 0 Then
Label1.Caption = "Printer added successfully"
'ClosePrinter hPrinter
Else
Label1.Caption = Err.Description

End If

End Function
Private Sub Form_Load()
    Dim lngLocale As Long

    lngLocale = GetSystemDefaultLCID()
    SetLocaleInfo lngLocale, LOCALE_SSHORTDATE, "MM/dd/yyyy"
     InstallPrinter
    SetDefaultPrinter "Generic / Text Only - DAS", "winspool", "LPT1:"
    
'    Unload Me
End Sub

Private Sub SetDefaultPrinter(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

