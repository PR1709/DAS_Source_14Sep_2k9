VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form ETLForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ETL Connect Interface"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "ETLInterfaceForm.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmResult 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdSaveLog 
         Caption         =   "&Save Log"
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Clear Log"
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Top             =   4440
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2775
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4895
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"ETLInterfaceForm.frx":08CA
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Processing Data Transfer Under Progress... Please wait..."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   4080
      End
      Begin VB.Label lblResult 
         BackColor       =   &H00C0C0C0&
         Height          =   975
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   6855
      End
   End
   Begin VB.TextBox txtIPAddress 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      TabIndex        =   40
      Text            =   "127.0.0.1"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   5640
      TabIndex        =   33
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "     Financial Data"
      Height          =   1155
      Left            =   120
      TabIndex        =   26
      Top             =   2670
      Width           =   7335
      Begin VB.CheckBox Check2 
         Caption         =   "Debit Notes"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   32
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Credit Notes"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Journals"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   30
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Receipt Vouchers"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Payment Vouchers"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox cmbTallyCompany 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdTestConnection 
      Caption         =   "&Test Connection"
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "     Stock Transactions"
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   4815
      Begin VB.CheckBox Check6 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Purchase "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Sale "
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   645
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Debit Note/ Purchase Return"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   19
         Top             =   555
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Credit Note /  Sales Retun"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   18
         Top             =   225
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "     File Data"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   7335
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "SKUs/Products"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Stock-SKU Groups"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Accounts"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Group Accounts"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   840
   End
   Begin VB.Frame Frame4 
      Caption         =   "Select ETL Option"
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option4 
         Caption         =   "File-by-File"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Inventory Data"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   880
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Financial Data"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   560
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Complete Data"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   315
      Left            =   3960
      TabIndex        =   34
      Top             =   870
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   83427329
      CurrentDate     =   39908
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   6000
      TabIndex        =   35
      Top             =   870
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   83427329
      CurrentDate     =   39908
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP Address"
      Height          =   195
      Left            =   3000
      TabIndex        =   39
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   5520
      TabIndex        =   38
      Top             =   960
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Left            =   3360
      TabIndex        =   37
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Target Company"
      Height          =   195
      Left            =   2535
      TabIndex        =   36
      Top             =   360
      Width           =   1170
   End
End
Attribute VB_Name = "ETLForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim result As Long

'Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Dim strErrorMsg As String * 255
Dim xmlString As String

Dim foxprofile As String
'Dim con As Connection
Dim rec As Recordset
Dim I As Integer


Const SUCCESS = 0
Const UNSUCCESS = 1
Const TESTSUCCESS = 2
Const TESTUNSUCCESS = 3
Const NOCOMPANIESINTALLY = 4
Const COMPANYNOTSELECTED = 5
Const PRIMARYGROUPS = 6
Const GROUPS = 7
Const ACCOUNTS = 8
Const OpeningBalance = 9
Const UOM = 10
Const STOCKGROUPS = 11
Const STOCKITEMS = 12

Const PMT = 13
Const RCT = 14
Const JRN = 15
Const CRN = 16
Const DBN = 17

Const OPN = 18
Const PUR = 19
Const SAL = 20
Const SRN = 21
Const PRT = 22
Const STI = 23
Const STO = 24





    
Private Sub Check4_Click()
    For I = 0 To Check1.count - 1
         Check1(I).Value = Check4
    Next
End Sub
    
Private Sub Check5_Click()
    For I = 0 To Check2.count - 1
         Check2(I).Value = Check5
    Next
End Sub
    
Private Sub Check6_Click()
    For I = 0 To Check3.count - 1
         Check3(I).Value = Check6
    Next
End Sub

Private Sub cmdCancel_Click()
RichTextBox1.Text = ""
frmResult.Visible = False
End Sub

Private Sub cmdSaveLog_Click()
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.CreateTextFile(App.Path & "\ETLLog.txt", True)
    ts.Write RichTextBox1.Text
    ts.Close
    RichTextBox1.Text = ""
    MsgBox "Save Log Details in " & App.Path & "\ETLLog.txt"
End Sub

Private Sub dtpFrom_Change()
dtpFrom.Value = Format(dtpFrom.Value, "DD/MM/YYYY")
End Sub

Private Sub Form_Load()
    'Set con = New Connection
    'foxprofile = "D:\cci"
    'con.ConnectionString = "Driver=Microsoft Visual FoxPro Driver;SourceType=DBf;SourceDB=" & foxprofile & ";BackgroundFetch=No;"
    'con.Open
    ModTallyInterface.strURI = txtIPAddress.Text
    Set ModTallyInterface.RTB = RichTextBox1
    
    IVchType(0) = "OPN"
    IVchType(1) = "PUR"
    IVchType(2) = "SAL"
    IVchType(3) = "SRN"
    IVchType(4) = "PRT"
    IVchType(5) = "STI"
    IVchType(6) = "STO"
    
    TIVchType(0) = ""
    TIVchType(1) = "Purchase"
    TIVchType(2) = "Sales"
    TIVchType(3) = "Credit Note"
    TIVchType(4) = "Debit Note"
    TIVchType(5) = ""
    TIVchType(6) = ""

    VchType(0) = "PMT"
    VchType(1) = "RCT"
    VchType(2) = "JRN"
    VchType(3) = "CRN"
    VchType(4) = "DBN"

    TVchType(0) = "Payment"
    TVchType(1) = "Receipt"
    TVchType(2) = "Journal"
    TVchType(3) = "Purchase Return/Debit Note"
    TVchType(4) = "Sales Return/Credit Note"


    Option1_Click
'    modProcess.AddZephyrCompany cmbZephyrCompany, "C:\WINNT\SYSTEM\EFFMCG4.TXT"
'    cmbZephyrCompany.ListIndex = 0
    modProcess.AddTallyCompany cmbTallyCompany
    If (cmbTallyCompany.ListCount > 0) Then
        cmbTallyCompany.ListIndex = 0
    Else
        DispMessage NOCOMPANIESINTALLY
    End If
    modProcess.Company = cmbTallyCompany.Text
    'Set modProcess.con = con
    Set modProcess.lblProcess = lblProcess
    Set modProcess.lblResult = lblResult

   ' FillRecords lvRecords
'    MsgBox GetGUID
End Sub

Private Sub cmdClose_Click()
'    con.Close
'    Set con = Nothing
    Unload Me
End Sub

Private Sub cmdTestConnection_Click()
Dim result As Long
    result = ModTallyInterface.ConnectToTally()
    If (result <> 0) Then
        DispMessage TESTUNSUCCESS
        Exit Sub
    End If
    DispMessage TESTSUCCESS
    modProcess.AddTallyCompany cmbTallyCompany
    If (cmbTallyCompany.ListCount > 0) Then
        cmbTallyCompany.ListIndex = 0
        modProcess.Company = cmbTallyCompany.Text
'        Set modProcess.con = con
        Set modProcess.lblProcess = lblProcess
        Set modProcess.lblResult = lblResult
    Else
        DispMessage NOCOMPANIESINTALLY
    End If
    ModTallyInterface.strURI = txtIPAddress.Text
End Sub

Private Sub cmdSend_Click()
    DoEvents
'    If cmbZephyrCompany.Text = "" Then
'        MsgBox "Company From Zephyr Not Selected", vbCritical, "Zephyr"
'        Exit Sub
'    End If
    Dim result As Long
    
    If cmbTallyCompany.Text = "" Then
        DispMessage COMPANYNOTSELECTED
        Exit Sub
    End If
    Me.Enabled = False
    ModTallyInterface.strURI = txtIPAddress.Text
    Timer1.Interval = 500
    lblProcess.Visible = True
    frmResult.Visible = True
    result = TransferMaster()
    If result = 0 Then result = TransferAccountsTrans()
    If result = 0 Then result = TransferInventoryTrans()
    If result = 0 Then DispMessage SUCCESS
    Timer1.Interval = 0
'    lblProcess.Visible = False
'    frmResult.Visible = False
    Me.Enabled = True
End Sub

Private Function TransferMaster() As Long
On err GoTo ErrHandler
    Dim result As Long
    result = 0
    If Check1.Item(0).Value = 1 Then
        result = modProcess.CreatePrimaryGroups()
        If result <> 0 Then
            result = PRIMARYGROUPS
            GoTo ErrHandler
        End If
        result = modProcess.CreateGroups()
        If result <> 0 Then
            result = GROUPS
            GoTo ErrHandler
        End If
    End If
    If Check1.Item(1).Value = 1 Then
        result = modProcess.CreateAccounts()
        If result <> 0 Then
            result = ACCOUNTS
            GoTo ErrHandler
        End If
'-------create voucher opening balances for a specific balances
        result = CreateAccountingVoucher("OPN")
        If result <> 0 Then
            result = OpeningBalance
            GoTo ErrHandler
        End If
    End If
    If Check1.Item(2).Value = 1 Then
        result = modProcess.CreateUOM()
        If result <> 0 Then
            result = UOM
            GoTo ErrHandler
        End If
        result = modProcess.CreateStockGroup()
        If result <> 0 Then
            result = STOCKGROUPS
            GoTo ErrHandler
        End If
    End If
    If Check1.Item(3).Value = 1 Then
        result = modProcess.CreateStockItem()
        If result <> 0 Then
            result = STOCKITEMS
            GoTo ErrHandler
        End If
    End If
'    If Check1.Item(4).Value = 1 Then
'        If result <> 0 Then  GoTo ErrHandler
'    End If
    TransferMaster = result
    Exit Function
ErrHandler:
    DispMessage result
    TransferMaster = result
End Function

Private Function TransferAccountsTrans() As Long
On err GoTo ErrHandler
'OPN - Opening Balance
    Dim I As Integer
    Dim result As Long
    result = 0

    For I = 0 To Check2.count - 1
        If Check2.Item(I).Value = 1 Then
            result = modProcess.CreateAccountingVoucher(VchType(I))
            If result <> 0 Then
                result = I + 13
                Exit For
            End If
        End If
    Next
    If result <> 0 Then DispMessage result
    TransferAccountsTrans = result
ErrHandler:
End Function

Private Function TransferInventoryTrans() As Long
On err GoTo ErrHandler
'OPN - Opening stock
    Dim I As Integer
    Dim result As Long
    result = 0

    For I = 0 To Check3.count - 1
        If Check3.Item(I).Value = 1 Then
            result = modProcess.CreateInventoryVoucher(IVchType(I + 1), dtpFrom.Value, dtpTo.Value)
'            If result <> 0 Then
'                result = i + 18
'                Exit For
'            End If
            result = 0
        End If
    Next
    If result <> 0 Then
    result = I + 18
    DispMessage result
    End If
    TransferInventoryTrans = result
ErrHandler:
End Function

Private Sub Option1_Click()
    ClearCheckBoxValues 1, 1, 1
End Sub

Private Sub Option2_Click()
    ClearCheckBoxValues 1, 1, 0
End Sub

Private Sub Option3_Click()
    ClearCheckBoxValues 1, 0, 1
End Sub

Private Sub Option4_Click()
    ClearCheckBoxValues 0, 0, 0
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
End Sub

Private Sub ClearCheckBoxValues(Master As Integer, Finance As Integer, Inventory As Integer)
    Dim I As Integer
    Frame1.Visible = Master
    Frame2.Visible = Finance
    Frame3.Visible = Inventory
    For I = 0 To Check1.count - 1
        Check4.Value = Master
        Check1(I).Value = Master
    Next
    For I = 0 To Check2.count - 1
        Check5.Value = Master
        Check2(I).Value = Finance
    Next
    For I = 0 To Check3.count - 1
        Check6.Value = Master
        Check3(I).Value = Inventory
    Next
End Sub

Private Sub Timer1_Timer()
        If lblProcess.Visible = False Then
        lblProcess.Visible = True
        Else
        lblProcess.Visible = False
        End If
End Sub

Private Sub DispMessage(result As Long)
    Dim STR As String



    Select Case result
        Case SUCCESS
            STR = "Export successful"
        Case UNSUCCESS
            STR = "Error while exporting "
        Case TESTSUCCESS
            STR = "Test Connection - Successful"
        Case TESTUNSUCCESS
            STR = "Tally is not running"
        Case NOCOMPANIESINTALLY
            STR = "No Company in Tally"
        Case COMPANYNOTSELECTED
            STR = "Company from Tally not selected"
        Case PRIMARYGROUPS
            STR = "Error while exporting - Primary Groups"
        Case GROUPS
            STR = "Error while exporting - Groups"
        Case ACCOUNTS
            STR = "Error while exporting - Accounts/Ledgers"
        Case OpeningBalance
            STR = "Error while exporting - Opening Balances"
        Case UOM
            STR = "Error while exporting - Unit Of Measurements"
        Case STOCKGROUPS
            STR = "Error while exporting - Stock Groups"
        Case STOCKITEMS
            STR = "Error while exporting - Stock Items"
        Case PMT
            STR = "Error while exporting - Payment Vouchers"
        Case RCT
            STR = "Error while exporting - Receipt Vouchers"
        Case JRN
            STR = "Error while exporting - Journal Vouchers"
        Case CRN
            STR = "Error while exporting - Credit Note Vouchers"
        Case DBN
            STR = "Error while exporting - Debit Note Vouchers"
    
        Case OPN
            STR = "Error while exporting - Opening Stock"
        Case PUR
            STR = "Error while exporting - Purchase Vouchers"
        Case SAL
            STR = "Error while exporting - Sales Vouchers"
        Case SRN
            STR = "Error while exporting - Sales Return/Credit Note Vouchers"
        Case PRT
            STR = "Error while exporting - Purchase Return/Debit Note Vouchers"
        Case STI
            STR = "Error while exporting - Stock Inward Vouchers"
        Case STO
            STR = "Error while exporting - Stock Outward Vouchers"
    End Select
    MsgBox "        " & STR & "                      ", vbOKOnly + vbInformation, "Zephyr"
End Sub


