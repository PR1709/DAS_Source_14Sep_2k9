VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{09E60688-07A2-11CF-BD27-10005AE6AFFA}#1.0#0"; "RRW32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm M_MDI 
   BackColor       =   &H8000000F&
   Caption         =   "DAS Version 4.01.01 - Distributor Automation System, Utility Modules - Ver.Rel.04-Sep-2009"
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   -810
   ClientWidth     =   12735
   Icon            =   "M_MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   6120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar STS_BAR 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8265
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16351
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1499
            TextSave        =   "09/04/2009"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1589
            MinWidth        =   1589
            TextSave        =   "6:50 PM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   10875
      Left            =   0
      ScaleHeight     =   10815
      ScaleWidth      =   12675
      TabIndex        =   1
      Top             =   0
      Width           =   12735
      Begin VB.Image Image1 
         Height          =   1425
         Left            =   10440
         Picture         =   "M_MDI.frx":08CA
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consoft Systems Pvt Ltd"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   10005
         TabIndex        =   2
         Top             =   7560
         Width           =   2235
      End
   End
   Begin RrocxLib.RRReport RRReport1 
      Left            =   1920
      Top             =   1560
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      WindowControlBox=   0   'False
      WindowBorderStyle=   0
      WindowTitle     =   ""
      PrintFileName   =   ""
      ReportName      =   ""
      ReportLibrary   =   ""
      MemoFileName    =   ""
      Printer         =   ""
      Port            =   ""
      MasterTable     =   ""
      Filter          =   ""
      RelatedTables   =   ""
      WindowParentHandle=   0
      DataDirectory   =   ""
      ReportDirectory =   ""
      ImageDirectory  =   ""
      LastErrorCode   =   0
      LastErrorString =   ""
      LastErrorPage   =   0
      SuppressTitle   =   0   'False
      Parameters      =   ""
      DBContainer     =   ""
      SortFields      =   ""
      GroupFields     =   ""
      ReportSortFields=   ""
      ReportGroupFields=   ""
      ControlFile     =   ""
      SuitableGroupFields=   ""
      SuitableSortFields=   ""
      ReportsRTString =   ""
      RelatedTablesString=   ""
      ParametersString=   ""
      ReportParametersString=   ""
      StatusFileName  =   ""
      MasterIndex     =   ""
      LowScope        =   ""
      HighScope       =   ""
      IndexExtension  =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   1560
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuFastBackup 
         Caption         =   "&Fast Backup"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "&Indexing"
      End
      Begin VB.Menu mnuRepair 
         Caption         =   "&Repair"
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMIBS 
         Caption         =   "MIBS"
         Begin VB.Menu mnuMobileIdentification 
            Caption         =   "&Mobile - Identification"
         End
         Begin VB.Menu mnumobileRetailers 
            Caption         =   "Mobile - &Retailers Information"
         End
         Begin VB.Menu mnuMobileProduts 
            Caption         =   "Mobile - &Products Information"
         End
         Begin VB.Menu mnuDataExchange 
            Caption         =   "Mobile - &Data Exchange"
         End
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSMS 
         Caption         =   "&GSM Modem SMS Send Interface"
      End
      Begin VB.Menu mnuETL 
         Caption         =   "&Accounts and Audit Interface"
      End
      Begin VB.Menu mnuBarCode 
         Caption         =   "&BarCode LabelPrint Interface"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSelectCompany 
      Caption         =   "&Select Company"
   End
   Begin VB.Menu mnuExit1 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "M_MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private REC_POS As Integer

Private Sub MDIForm_Load()
    If App.PrevInstance = True Then
        MsgBox "Already Opened"
        Unload Me
        Exit Sub
    End If
    
GetDBPath
ReSetConnectionObjects EFF_LIB1.CompanyName
End Sub

Private Sub MDIForm_Resize()
Picture1.BorderStyle = 0
'Picture1.BackColor = vbRed
If (Me.Height - STS_BAR.Height - 750) > 0 Then Picture1.Height = Me.Height - STS_BAR.Height - 750
If (Me.Width - Label1.Width - 200) > 0 Then Label1.Left = Me.Width - Label1.Width - 200
If (Label1.Left + ((Label1.Width - Image1.Width) / 2)) > 0 Then Image1.Left = Label1.Left + ((Label1.Width - Image1.Width) / 2)
If (Me.Height - STS_BAR.Height - 1550) > 0 Then Label1.Top = Me.Height - STS_BAR.Height - 1550
If (Label1.Top - Label1.Height - Image1.Height - 50) > 0 Then Image1.Top = Label1.Top - Label1.Height - Image1.Height - 50
'Picture1.Width = Me.Width
End Sub

Private Sub mnuActivationTargets_Click()
If I = 0 Then
M_FRM_RTG.Show 1
ElseIf I = 1 Then
M_FRM_MBL.Show 1
End If
End Sub

Private Sub mnuBarCode_Click()
    frmBarCode.Show 1
End Sub

Private Sub mnuDataExchange_Click()
M_FRM_MIBS.Show 1
End Sub

Private Sub mnuETL_Click()
        ETLForm.Show 1
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuExit1_Click()
End
End Sub

Private Sub mnuFastBackup_Click()
ReSetConnectionObjects EFF_LIB1.CompanyName
    modGeneral.FastBackup
End Sub

Private Sub mnuIndex_Click()
    modGeneral.IndexData
End Sub

Private Sub mnuMobileIdentification_Click()
M_FRM_HHT.Show 1
End Sub

Private Sub mnuMobileProduts_Click()
M_FRM_EVP.Show 1
End Sub

Private Sub mnumobileRetailers_Click()
m_FRM_HUP.Show 1
End Sub

Private Sub mnuRepair_Click()
    modGeneral.RepairDatabase
End Sub

Private Sub mnuSelectCompany_Click()
    CMP_SEL.Show 1
End Sub

Private Sub mnuSMS_Click()
    frmSMS.Show 1
End Sub

