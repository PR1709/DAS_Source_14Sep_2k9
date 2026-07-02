VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form USR_FRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Users - Passwords"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "USR_FRM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      Sorted          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   580
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   3420
      TabIndex        =   6
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "User"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Administrator"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1620
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1050
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1620
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1485
      Width           =   1455
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5741
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5213
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_FRM.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl+S)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Enter Password"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   795
      TabIndex        =   4
      Top             =   1118
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   690
      TabIndex        =   3
      Top             =   645
      Width           =   795
   End
End
Attribute VB_Name = "USR_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC_SET As New ADODB.Recordset
Dim MOD_REC As Boolean

Public Sub PRC_ADD_USR()
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "Select distinct(usr_nme) from usr group by usr_nme", con, adOpenKeyset, adLockPessimistic
Combo1.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Combo1.AddItem Trim(REC_SET(0))
        REC_SET.MoveNext
    Next
End If
End Sub

Public Sub PRC_DIS_DET()
If Len(Trim(Combo1)) <= 4 Then
    MsgBox "User Name and Password Should be more than 4 characters long!", vbExclamation, "DAS Version 4.0"
    Combo1.SetFocus
    Exit Sub
End If
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from usr where usr_nme like '" & Trim(Combo1) & "'", con, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then
    MOD_REC = True
    If Trim(REC_SET(3)) = "Admin" Then
        Text2 = DPWD.decPwd(Trim(REC_SET(1)))
        Text3 = DPWD.decPwd(Trim(REC_SET(1)))
        TreeView1.Enabled = False
        Option1.Value = True
        Option2.Value = False
        Text2.SetFocus
        Exit Sub
    Else
        Text2 = DPWD.decPwd(Trim(REC_SET(1)))
        Text3 = DPWD.decPwd(Trim(REC_SET(1)))
        Option2.Value = True
        Option1.Value = False
        
        TreeView1.Enabled = True
        For A = 1 To REC_SET.RecordCount
            For b = 1 To TreeView1.Nodes.count
                If TreeView1.Nodes.Item(b).KEY = Trim(REC_SET(2)) Then
                    TreeView1.Nodes.Item(b).Checked = True
                End If
            Next
            REC_SET.MoveNext
        Next
    End If
    Text2.SetFocus
Else
    If MsgBox("Create New User ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        MOD_REC = False
        Text2.SetFocus
        TreeView1.Enabled = False
        Option1.Value = False
        Option2.Value = False
    Else
        Combo1.SetFocus
        SendKeys "{home}+{end}"
    End If
End If
End Sub


Private Sub Combo1_keypress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If Len(Combo1) = 30 Then K = 0
If K = 13 Then PRC_DIS_DET
End Sub

Private Sub PRC_SAV_REC()
If MsgBox("Confirm New User Creation ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    PRC_CAN_REC
    Exit Sub
End If
If Len(Trim(Combo1)) = 0 Then
MsgBox "User Name Should Not Be Blank, Enter Username !", vbExclamation, "DAS Version 4.0"
Combo1.SetFocus
Exit Sub
End If

If Len(Trim(Text2)) < 4 Then
MsgBox "Password Should be Minimum Four Characters Long!", vbExclamation, "DAS Version 4.0"
Text2 = ""
Text3 = ""
Text2.SetFocus
Exit Sub
End If

If Len(Trim(Text2)) > 8 Then
    MsgBox "Password should not be greater than 8 characters", vbExclamation, "DAS Version 4.0"
    Text2 = ""
    Text3 = ""
    Text2.SetFocus
End If

If Trim(Text3) <> Trim(Text2) Then
MsgBox "Password Mismatch, Re-enter Password !", vbCritical, "DAS Version 4.0"
Text3.SetFocus
Exit Sub
End If

If REC_SET.State = 1 Then REC_SET.Close
If MOD_REC = True Then
    con.Execute "DELETE  FROM USR WHERE USR_NME LIKE '" & Trim(Combo1) & "'"
End If
REC_SET.Open "select * from usr", con, adOpenKeyset, adLockPessimistic
'If TreeView1.Nodes.Item(9).Checked = True Then MsgBox "child1 is cheked"
 If Option1.Value = True Then
     REC_SET.AddNew
     REC_SET(0) = Trim(Combo1)
     REC_SET(1) = EPWD.encPwd((Text2))
     REC_SET(3) = "Admin"
     REC_SET.Update
     PRC_CAN_REC
     PRC_ADD_USR
     Exit Sub
End If
If Option2.Value = True Then
    For A = 1 To TreeView1.Nodes.count
        If TreeView1.Nodes.Item(A).Checked = True Then
            REC_SET.AddNew
            REC_SET(0) = Trim(Combo1)
            REC_SET(1) = EPWD.encPwd((Text2))
            REC_SET(2) = Trim(TreeView1.Nodes.Item(A).KEY)
            REC_SET(3) = "User"
            REC_SET.Update
        End If
    Next
End If
PRC_CAN_REC
PRC_ADD_USR
End Sub

Private Sub PRC_CAN_REC()
Combo1 = ""
Text2 = ""
Text3 = ""
Option1.Value = False
Option2.Value = False
For A = 1 To TreeView1.Nodes.count
 TreeView1.Nodes.Item(A).Checked = False
Next
TreeView1.Enabled = False
Combo1.SetFocus
End Sub

Private Sub Form_Activate()
TreeView1.Enabled = False
Combo1.SetFocus
MOD_REC = False
PRC_ADD_USR
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
    If K = 78 And s = 2 Then
        PRC_CAN_REC
    ElseIf K = 83 And s = 2 Then
        PRC_SAV_REC
    ElseIf K = 81 And s = 2 Then
        PRC_CAN_REC
    ElseIf K = 88 And s = 2 Then
        Unload Me
    ElseIf K = 112 Then
'        Shell "c:\windows\hh.exe " & App.Path & "\conthelp.chm", vbMaximizedFocus
    ElseIf K = 80 And s = 2 Then
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If USR_FRM.ActiveControl.name = "Combo1" Then
        Unload Me
    Else
        PRC_SAV_REC
    End If
End If
End Sub

Private Sub Form_Load()

   TreeView1.LineStyle = tvwRootLines  ' Linestyle 1
   Dim nodX As Node  ' Declare Node variable.
   'FILE MENU AND ITS SUB MENU
   Set nodX = TreeView1.Nodes.Add(, , "File", "File")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Categories", "Categories")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Locations", "Locations")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Salesman", "Salesman")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Markets", "Markets")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Group Accounts", "Group Accounts")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Accounts & Parties", "Accounts & Parties")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Product Categories", "Product Categories")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Products (SKU)", "Products (SKU)")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Key Accounts SKU Rates", "Key Accounts SKU Rates")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Schemes - Trade", "Schemes - Trade")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Sales Calender", "Sales Calender")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Targets - Salesmen", "Targets - Salesmen")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Route Sequence", "Route Sequence")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Commissions", "Commissions")
   Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Day ClosUre n Data XChange", "Day ClosUre n Data XChange")
   'Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Breakage Limits", "Breakage Limits")
   'Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Incentives", "Incentives")
   'Set nodX = TreeView1.Nodes.Add("File", tvwChild, "Retailer Fixed Schemes", "Retailer Fixed Schemes")

   
   'CORE MENU AND ITS SUB MENU
   Set nodX = TreeView1.Nodes.Add(, , "Core", "Core")
   Set nodX = TreeView1.Nodes.Add("Core", tvwChild, "Purchases", "Purchases")
   Set nodX = TreeView1.Nodes.Add("Core", tvwChild, "Sales", "Sales")
   'Sales MENU AND ITS SUB MENU
   Set nodX = TreeView1.Nodes.Add("Sales", tvwChild, "Credit Sales", "Sales Invoices")
   Set nodX = TreeView1.Nodes.Add("Sales", tvwChild, "Ready Stock Bills", "Van Stock Issue")
   Set nodX = TreeView1.Nodes.Add("Sales", tvwChild, "Van Variance Check", "Van Variance Check")
   Set nodX = TreeView1.Nodes.Add("Sales", tvwChild, "Invoices vs Receipts", "Collection Entry")
   Set nodX = TreeView1.Nodes.Add("Sales", tvwChild, "Load Sheet", "Load Sheet")
   Set nodX = TreeView1.Nodes.Add("Sales", tvwChild, "Push Cart Sales", "Push Cart Sales")
   
   Set nodX = TreeView1.Nodes.Add("Core", tvwChild, "Stocks", "Stocks")
   Set nodX = TreeView1.Nodes.Add("Stocks", tvwChild, "Opening Stock", "Opening Stock")
   Set nodX = TreeView1.Nodes.Add("Stocks", tvwChild, "Credit Notes", "Credit Notes")
   Set nodX = TreeView1.Nodes.Add("Stocks", tvwChild, "Debit Notes", "Debit Notes")
   Set nodX = TreeView1.Nodes.Add("Stocks", tvwChild, "Transfer Inward", "Transfer Inward")
   Set nodX = TreeView1.Nodes.Add("Stocks", tvwChild, "Transfer Outward", "Transfer Outward")
   Set nodX = TreeView1.Nodes.Add("Stocks", tvwChild, "Stock Adjustment", "Stock Adj. Journal")
   Set nodX = TreeView1.Nodes.Add("Stocks", tvwChild, "Stock Conversion", "Stock Conversion")
   
   Set nodX = TreeView1.Nodes.Add("Core", tvwChild, "snp", "Two Way Packs")
   Set nodX = TreeView1.Nodes.Add("snp", tvwChild, "Empties Returnables & Actual", "Empties Returnables & Actual")
   Set nodX = TreeView1.Nodes.Add("snp", tvwChild, "Empties Additional Entries", "Empties Additional Entries")
   Set nodX = TreeView1.Nodes.Add("snp", tvwChild, "Shell Description", "Shell Description")
   Set nodX = TreeView1.Nodes.Add("snp", tvwChild, "Empties Reports", "Empties Reports")
   
   Set nodX = TreeView1.Nodes.Add("Core", tvwChild, "twp", "Telecom")
   Set nodX = TreeView1.Nodes.Add("twp", tvwChild, "tent", "Activation Target")
   Set nodX = TreeView1.Nodes.Add("twp", tvwChild, "ti", "Telecom Interface")
   
   
   Set nodX = TreeView1.Nodes.Add("Core", tvwChild, "Supplier Order", "Supplier Order")
   Set nodX = TreeView1.Nodes.Add("Core", tvwChild, "Update Stocks", "Update Stocks")
   
   'ACCOUNTS MENU AND ITS SUB MENU
   Set nodX = TreeView1.Nodes.Add(, , "Accounts", "Accounts")
   Set nodX = TreeView1.Nodes.Add("Accounts", tvwChild, "Receipts", "Receipts")
   Set nodX = TreeView1.Nodes.Add("Accounts", tvwChild, "Payments", "Payments")
   Set nodX = TreeView1.Nodes.Add("Accounts", tvwChild, "Journals", "Journals")
   Set nodX = TreeView1.Nodes.Add("Journals", tvwChild, "Journal Entries", "A/c Journals")
   Set nodX = TreeView1.Nodes.Add("Journals", tvwChild, "Credit Note", "Credit Notes")
   Set nodX = TreeView1.Nodes.Add("Journals", tvwChild, "Debit Note", "Debit Notes")
   Set nodX = TreeView1.Nodes.Add("Journals", tvwChild, "Opening Balance", "Opening Balance")
   
   Set nodX = TreeView1.Nodes.Add(, , "Assets", "Assets")
   Set nodX = TreeView1.Nodes.Add("Assets", tvwChild, "Asset Type", "Asset Type")
   Set nodX = TreeView1.Nodes.Add("Assets", tvwChild, "Asset Details", "Asset Details")
   Set nodX = TreeView1.Nodes.Add("Assets", tvwChild, "Inward Entry", "Inward Entry")
   Set nodX = TreeView1.Nodes.Add("Assets", tvwChild, "Outward Entry", "Outward Entry")
   
   'ANALYSIS MENU AND ITS SUBMENU
   Set nodX = TreeView1.Nodes.Add(, , "Analysis", "Analysis")
   Set nodX = TreeView1.Nodes.Add("Analysis", tvwChild, "Retailers", "Retailers")
   
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Ranking", "Ranking")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Last Billing Date", "Last Billing Date")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Retailers Route NB", "Retailers Route NB")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Retailers Salesmen NB", "Retailers Salesmen NB")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Non Operational", "Non Operational")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Categories Count", "Categories Count")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Sales Pareto", "Sales Pareto")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Coverage Frequency", "Coverage Frequency")
   Set nodX = TreeView1.Nodes.Add("Retailers", tvwChild, "Yearly Sales", "Yearly Sales")
   
   Set nodX = TreeView1.Nodes.Add("Analysis", tvwChild, "Products(SKU)", "Products (SKU)")
   
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "SKU Pareto", "SKU Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "SKU Segment Pareto", "SKU Segment Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "SKU Category Pareto", "SKU Category Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Sale Value Pareto", "Sale Value Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Retailers Pareto", "Retailers Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Categories Pareto", "Categories Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Locations Pareto", "Locations Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Markets Pareto", "Markets Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Routes Pareto", "Routes Pareto")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Penetration Count", "Penetration Count")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Sales Report - Splr", "Sales Report - Splr")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Sales Report - Stock", "Sales Report - Stock")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Product Sales", "Product Sales")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Yearly Sales - Units", "Yearly Sales - Units")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Yearly Sales - Cases", "Yearly Sales - Cases")
   Set nodX = TreeView1.Nodes.Add("Products(SKU)", tvwChild, "Actual Sales Report", "Actual Sales Report")
   
   Set nodX = TreeView1.Nodes.Add("Analysis", tvwChild, "Others", "Others")
   Set nodX = TreeView1.Nodes.Add("Others", tvwChild, "Month Wise Business Record", "Month Wisee Bussiness Record")
   Set nodX = TreeView1.Nodes.Add("Others", tvwChild, "Salesman Turnover Report", "Salesman Turnover Report")
   Set nodX = TreeView1.Nodes.Add("Others", tvwChild, "Market Coverage Frequency", "Market Coverage Frequency")
   
   Set nodX = TreeView1.Nodes.Add("Analysis", tvwChild, "Specific Reports", "Specific Reports")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Van Variance", "Van Variance")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Shell Variance", "Shell Variance")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Van Shell Variance", "Van Shell Variance")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Smn Cumulative Balance", "Smn Cumulative Balance")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Daily Productivity Summary", "Daily Productivity Summary")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Discount Scheme CPC", "Discount Scheme CPC")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Vehicle Utilization", "Vehicle Utilization")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Van Challans Status", "Van Challans Status")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Distributors Stock Report", "Distributors Stock Report")
   Set nodX = TreeView1.Nodes.Add("Specific Reports", tvwChild, "Net Load Stock Movement", "Net Load Stock Movement")
   
   Set nodX = TreeView1.Nodes.Add("Analysis", tvwChild, "Progress Reports", "Progress Reports")
   Set nodX = TreeView1.Nodes.Add("Progress Reports", tvwChild, "Category Sales Progress", "Category Sales Progress")
   Set nodX = TreeView1.Nodes.Add("Progress Reports", tvwChild, "Location Sales Progress", "Location Sales Progress")
   Set nodX = TreeView1.Nodes.Add("Progress Reports", tvwChild, "Market Sales Progress", "Market Sales Progress")
   Set nodX = TreeView1.Nodes.Add("Progress Reports", tvwChild, "Salesman Sales Progress", "Salesman Sales Progress")
   Set nodX = TreeView1.Nodes.Add("Progress Reports", tvwChild, "Route Sales Progress", "Route Sales Progress")
   Set nodX = TreeView1.Nodes.Add("Progress Reports", tvwChild, "Retailer Sales Progress", "Retailer Sales Progress")
   
   Set nodX = TreeView1.Nodes.Add("Analysis", tvwChild, "Combination Reports", "Combination Reports")
   'Set nodX = TreeView1.Nodes.Add("Analysis", tvwChild, "Analysis Reports", "Anlaysis Reports")
   
   'REPORTS MENU AND ITS SUBMENU
   Set nodX = TreeView1.Nodes.Add(, , "Reports", "Reports")
   
   Set nodX = TreeView1.Nodes.Add("Reports", tvwChild, "Core(R)", "Core")
   'REPORT CORE SUBMENU
   Set nodX = TreeView1.Nodes.Add("Core(R)", tvwChild, "Sales(R)", "Sales")
   Set nodX = TreeView1.Nodes.Add("Core(R)", tvwChild, "Purchases(R)", "Purchases")
   Set nodX = TreeView1.Nodes.Add("Core(R)", tvwChild, "Stocks(R)", "Stocks")
   
   'REPORT ACCOUNTS MENU
   Set nodX = TreeView1.Nodes.Add("Reports", tvwChild, "Accounts(R)", "Accounts")
    
   'REPORT REGULAR MENU
   Set nodX = TreeView1.Nodes.Add("Reports", tvwChild, "Regular", "Regular")
'   Set nodX = TreeView1.Nodes.Add("Regular", tvwChild, "MIS", "Additional Reports - MIS")
   'REPORT RANGEPRINT MENU
   Set nodX = TreeView1.Nodes.Add("Reports", tvwChild, "Range Print", "Range Print")
   'RANGE PRINT SUBMENU
   Set nodX = TreeView1.Nodes.Add("Range Print", tvwChild, "Core(RP)", "Core")
   Set nodX = TreeView1.Nodes.Add("Range Print", tvwChild, "Accounts(RP)", "Accounts")
   Set nodX = TreeView1.Nodes.Add("Reports", tvwChild, "Direct Print", "Direct Print")
   ' SYSTEM MENU
   Set nodX = TreeView1.Nodes.Add(, , "System", "System")
   Set nodX = TreeView1.Nodes.Add("System", tvwChild, "Settings", "Settings")
   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "CSW", "Company Switchover")
   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "CLI", "Change Log in")
   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "DF", "Distributor Profile")
   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "AFY", "Financial Year")
   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "SD", "Specify Defaults")
   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "RS", "Report Sets")
   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "UAP", "User and Password")
'   Set nodX = TreeView1.Nodes.Add("Settings", tvwChild, "MC", "Module Check")
   
   Set nodX = TreeView1.Nodes.Add("System", tvwChild, "Database", "Database")
   Set nodX = TreeView1.Nodes.Add("Database", tvwChild, "IED", "Import DAS Version 4.0 Data")
   Set nodX = TreeView1.Nodes.Add("Database", tvwChild, "EFD", "External File Data")
   Set nodX = TreeView1.Nodes.Add("Database", tvwChild, "VDB", "Valedate Database")
   
   Set nodX = TreeView1.Nodes.Add("System", tvwChild, "Others", "Others")
'   Set nodX = TreeView1.Nodes.Add("MIBS and Others", tvwChild, "DAS Mobile - Identification", "DAS Mobile - Identification")
'   Set nodX = TreeView1.Nodes.Add("MIBS and Others", tvwChild, "DAS Mobile - Retailers Information", "DAS Mobile - Retailers Information")
'   Set nodX = TreeView1.Nodes.Add("MIBS and Others", tvwChild, "DAS Mobile - Products Information", "DAS Mobile - Products Information")
'   Set nodX = TreeView1.Nodes.Add("MIBS and Others", tvwChild, "DAS Mobile - Data Exchange", "DAS Mobile - Data Exchange")
   
End Sub


Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PRC_SAV_REC
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TreeView1.Enabled = True
    TreeView1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PRC_DIS_DET
End Sub

Private Sub TEXT2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If Len(Trim(Text2)) >= 4 Then
        Text3.SetFocus
    Else
        MsgBox "Password Should Not be Less Than Four Characters !", vbExclamation, "DAS Version 4.0"
        SendKeys "{home}+{end}"
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    If Trim(Text3) = Trim(Text2) Then
        If MOD_REC = True And Option1.Value = True Then
            Option1.SetFocus
            Exit Sub
        End If
        Option2.SetFocus
    Else
        MsgBox "Password Mismatch, Re-enter Password !", vbCritical, "DAS Version 4.0"
        SendKeys "{home}+{end}"
        Text3.SetFocus
    End If
ElseIf K = 38 Then
    Text2.SetFocus
End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    PRC_CAN_REC
ElseIf Button.KEY = "s" Then
    PRC_SAV_REC
ElseIf Button.KEY = "c" Then
    PRC_CAN_REC
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
'    Shell "c:\windows\hh.exe " & App.Path & "\conthelp.chm", vbMaximizedFocus
ElseIf Button.KEY = "p" Then
    M_FRM_DSN.Tag = "LOGREP"
    M_FRM_DSN.Show 1
'    MDIForm1.RRReport1.ReportName = App.Path & "\Users list.rrw"
    'ON ERROR GoTo er
'    MDIForm1.RRReport1.Destination = 2
'    MDIForm1.RRReport1.Printer = "?"
'    MDIForm1.RRReport1.RunReport 1
'    Exit Sub
'er:
'    MsgBox "Printer Unavailable, Set Printer !", vbExclamation, "DAS Version 4.0"
'ElseIf Button.Key = "pr" Then
'    'ON ERROR GoTo er1
'    MDIForm1.RRReport1.ReportName = App.Path & "\Users list.rrw"
'    MDIForm1.RRReport1.Destination = 6
'    MDIForm1.RRReport1.RunReport 1
'    Exit Sub
'er1:
'    MsgBox "Users Not Found, Define User...!", vbExclamation, "DAS Version 4.0"
End If
End Sub


