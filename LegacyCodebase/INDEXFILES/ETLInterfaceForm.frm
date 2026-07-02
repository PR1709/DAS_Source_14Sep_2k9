VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ETLForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ETL(Accounts) Connect Interface"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   ControlBox      =   0   'False
   Icon            =   "ETLInterfaceForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmResult 
      Height          =   3855
      Left            =   1440
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   3240
      End
      Begin VB.CommandButton cmdSaveLog 
         Caption         =   "&Save Log"
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Clear Log"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   3360
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"ETLInterfaceForm.frx":058A
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Processing Data Transfer Under Progress... Please wait..."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   4080
      End
      Begin VB.Label lblResult 
         BackColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   6975
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      ItemData        =   "ETLInterfaceForm.frx":060C
      Left            =   1440
      List            =   "ETLInterfaceForm.frx":0619
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      ItemData        =   "ETLInterfaceForm.frx":0638
      Left            =   1440
      List            =   "ETLInterfaceForm.frx":0645
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "ETLInterfaceForm.frx":0669
      Left            =   1440
      List            =   "ETLInterfaceForm.frx":0673
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtFilter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox cmbDescription 
      Height          =   315
      ItemData        =   "ETLInterfaceForm.frx":0681
      Left            =   1440
      List            =   "ETLInterfaceForm.frx":06B2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtIPAddress 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   540
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   5040
      Width           =   855
   End
   Begin VB.ComboBox cmbTallyCompany 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   3495
   End
   Begin VB.CommandButton cmdTestConnection 
      Caption         =   "&Test Connection"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   510
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvValues 
      Height          =   2415
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Batch Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pur Rate"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sell Rate"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MRP"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   1740
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   75169793
      CurrentDate     =   39908
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   1740
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   75169793
      CurrentDate     =   39908
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Rec.Selected"
      Height          =   195
      Left            =   6930
      TabIndex        =   29
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Recs"
      Height          =   195
      Left            =   7080
      TabIndex        =   28
      Top             =   1860
      Width           =   780
   End
   Begin VB.Label lblSelected 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8040
      TabIndex        =   27
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblRecords 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8040
      TabIndex        =   26
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   3120
      TabIndex        =   25
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "By"
      Height          =   195
      Left            =   1005
      TabIndex        =   22
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Left            =   840
      TabIndex        =   21
      Top             =   1800
      Width           =   345
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   8760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   390
      TabIndex        =   19
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP Address"
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Target Company"
      Height          =   195
      Left            =   75
      TabIndex        =   17
      Top             =   120
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

Private Sub Form_Load()

    ReSetConnectionObjects EFF_LIB1.CompanyName
    ModTallyInterface.strURI = txtIPAddress.Text
    Set ModTallyInterface.rtb = RichTextBox1
    
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
    TIVchType(5) = "Contra"
    TIVchType(6) = "Contra"

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


'    modProcess.AddZephyrCompany cmbZephyrCompany, "C:\WINNT\SYSTEM\EFFMCG4.TXT"
'    cmbZephyrCompany.ListIndex = 0
    modProcess.AddTallyCompany cmbTallyCompany
    If (cmbTallyCompany.ListCount > 0) Then
        cmbTallyCompany.ListIndex = 0
    Else
        DispMessage NOCOMPANIESINTALLY
    End If
    modProcess.company = cmbTallyCompany.Text
    'Set modProcess.con = con
    Set modProcess.lblProcess = lblProcess
    Set modProcess.lblResult = lblResult
    dtpFrom.Value = Format(Now, "MM/dd/yyyy")
    dtpTo.Value = Format(Now, "MM/dd/yyyy")
    Exit Sub
   ' FillRecords lvRecords
'    MsgBox GetGUID
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub


Private Sub cmbDescription_Click()
    If cmbDescription.ListIndex = 2 Then cmbDescription.ListIndex = 3
    If cmbDescription.ListIndex = 8 Then cmbDescription.ListIndex = 9
    Combo1(0).ListIndex = 0
    Combo1(1).ListIndex = 0
    Combo1(2).ListIndex = 0
    Combo1(0).Visible = False
    Combo1(1).Visible = False
    Combo1(2).Visible = False
    
    If cmbDescription.ListIndex <= 2 Then
        lblDate.Visible = False
        dtpFrom.Visible = False
        dtpTo.Visible = False
        Label4.Visible = False
        Combo1(0).Visible = True
    ElseIf cmbDescription.ListIndex >= 3 And cmbDescription.ListIndex < 8 Then
        lblDate.Visible = True
        dtpFrom.Visible = True
        dtpTo.Visible = True
        Label4.Visible = True
        Combo1(1).Visible = True
    ElseIf cmbDescription.ListIndex >= 8 And cmbDescription.ListIndex < cmbDescription.ListCount Then
        lblDate.Visible = True
        dtpFrom.Visible = True
        dtpTo.Visible = True
        Label4.Visible = True
        Combo1(2).Visible = True
    End If
    query

End Sub

Private Sub cmbDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dtpFrom.Visible = True Then
            dtpFrom.SetFocus
        Else
            dtpTo_KeyDown 13, 0
        End If
    End If
End Sub


Private Sub cmbTallyCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtIPAddress.SetFocus
End Sub

Private Sub Combo1_Change(Index As Integer)
    query
End Sub

Private Sub Combo1_Click(Index As Integer)
    query
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    txtFilter.SetFocus
End Sub

Private Sub dtpFrom_Change()
    query
End Sub

Private Sub dtpFrom_Click()
    query
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTo.SetFocus
End Sub

Private Sub dtpTo_Change()
    query
End Sub

Private Sub dtpTo_Click()
    query
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Combo1(0).Visible = True Then Combo1(0).SetFocus
        If Combo1(1).Visible = True Then Combo1(1).SetFocus
        If Combo1(2).Visible = True Then Combo1(2).SetFocus
    End If
End Sub

Private Sub lvValues_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvValues.SortKey = ColumnHeader.Index - 1 And lvValues.SortOrder = lvwDescending Then
        lvValues.SortKey = ColumnHeader.Index - 1
        lvValues.SortOrder = lvwAscending
    ElseIf lvValues.SortKey = ColumnHeader.Index - 1 And lvValues.SortOrder = lvwAscending Then
        lvValues.SortKey = ColumnHeader.Index - 1
        lvValues.SortOrder = lvwDescending
    Else
         lvValues.SortKey = ColumnHeader.Index - 1
        lvValues.SortOrder = lvwAscending
    End If
    lvValues.Sorted = True
End Sub

Private Sub lvValues_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim COUNT As Integer
COUNT = 0
For I = 1 To lvValues.ListItems.COUNT
    If lvValues.ListItems(I).Checked = True Then COUNT = COUNT + 1
Next
lblSelected.Caption = COUNT
End Sub

Private Sub lvValues_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdSend.SetFocus
End If
End Sub

Private Sub txtFilter_Change()
query
End Sub


Private Sub cmdCancel_Click()
RichTextBox1.Text = ""
frmResult.Visible = False
cmbTallyCompany.SetFocus
End Sub

Private Sub cmdSaveLog_Click()
    Dim FSO As New FileSystemObject
    Dim ts As TextStream
    Set ts = FSO.CreateTextFile(SetupFolder & "\ETLLog.txt", True)
    ts.Write RichTextBox1.Text
    ts.Close
    RichTextBox1.Text = ""
    MsgBox "Save Log Details in " & SetupFolder & "\ETLLog.txt"
    frmResult.Visible = False
    cmbTallyCompany.SetFocus
End Sub

Private Sub cmdSelectAll_Click()
    Dim COUNT As Integer
    If (cmdSelectAll.Caption = "Deselect &All") Then
        For I = 1 To lvValues.ListItems.COUNT
            lvValues.ListItems(I).Checked = False
        Next
        cmdSelectAll.Caption = "Select &All"
    Else
        For I = 1 To lvValues.ListItems.COUNT
            lvValues.ListItems(I).Checked = True
        Next
        cmdSelectAll.Caption = "Deselect &All"
    End If
    COUNT = 0
    For I = 1 To lvValues.ListItems.COUNT
        If (lvValues.ListItems(I).Checked = True) Then
        COUNT = COUNT + 1
        End If
    Next
    lblSelected.Caption = COUNT
End Sub

Private Sub cmdSend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmbTallyCompany.SetFocus
    End If
End Sub

Private Sub cmdClose_Click()
'    con.Close
'    Set con = Nothing
    Unload Me
End Sub

Private Sub cmdTestConnection_Click()
Dim RESULT As Long
    ModTallyInterface.strURI = txtIPAddress.Text
    RESULT = ModTallyInterface.ConnectToTally()
    If (RESULT <> 0) Then
        DispMessage TESTUNSUCCESS
        Exit Sub
    End If
    DispMessage TESTSUCCESS
    modProcess.AddTallyCompany cmbTallyCompany
    If (cmbTallyCompany.ListCount > 0) Then
        cmbTallyCompany.ListIndex = 0
        modProcess.company = cmbTallyCompany.Text
'        Set modProcess.con = con
        Set modProcess.lblProcess = lblProcess
        Set modProcess.lblResult = lblResult
    Else
        DispMessage NOCOMPANIESINTALLY
    End If
    ModTallyInterface.strURI = txtIPAddress.Text
End Sub

Private Sub cmdSend_Click()
    Me.Enabled = False
    Dim RESULT As Long
    lblProcess.Visible = True
    If cmbTallyCompany.Text = "" Then
        DispMessage COMPANYNOTSELECTED
        Me.Enabled = True
        Exit Sub
    End If
    ModTallyInterface.strURI = txtIPAddress.Text
    Timer1.Interval = 500
    lblProcess.Visible = True
    frmResult.Visible = True
    TransferMaster
    TransferAccountsTrans
    TransferInventoryTrans
    Timer1.Interval = 0
    lblResult.Caption = ""
    lblProcess.Visible = False
    If Trim(RichTextBox1.Text) = "" Then
        frmResult.Visible = False
    End If
    MsgBox "ETL Transfer completed", vbInformation, "DAS Version 4.0"
    Me.Enabled = True
End Sub

Private Sub Timer1_Timer()
        If lblProcess.Visible = False Then
        lblProcess.Visible = True
        Else
        lblProcess.Visible = False
        End If
End Sub



Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If lvValues.ListItems.COUNT > 0 Then lvValues.SetFocus
End If
End Sub

Private Sub txtIPAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmbDescription.SetFocus
End Sub


Private Function FillListView(lv As ListView, SqlStatement As String)
    lv.ListItems.CLEAR
    If Trim(SqlStatement) = "" Then Exit Function
    Dim li As ListItem
    Dim STR As String
    
    Set rec = modGeneral.GetRecords(SqlStatement)
    lv.ColumnHeaders.CLEAR
    If rec.EOF = True Then Exit Function
    For I = 0 To rec.Fields.COUNT - 1
        STR = Trim(rec.Fields(I).name)
        STR = ProperCase(STR)
        'str = UCase(Mid(Trim(rec.Fields(I).name), 1, 1)) & Mid(Trim(rec.Fields(I).name), 2)
        If InStr(1, UCase(rec(I).name), "ROU OFF") > 0 Or InStr(1, UCase(rec(I).name), "CST") > 0 Or InStr(1, UCase(rec(I).name), "PCG") > 0 Or InStr(1, UCase(rec(I).name), "AMT") > 0 Or InStr(1, UCase(rec(I).name), "AMOUNT") > 0 Or InStr(1, UCase(rec(I).name), "DIS") > 0 Or InStr(1, UCase(rec(I).name), "VAL") > 0 Or InStr(1, UCase(rec(I).name), "TAX") > 0 Then
            lv.ColumnHeaders.Add , , STR, , lvwColumnRight
        ElseIf InStr(1, UCase(rec(I).name), "DATE") > 0 Or InStr(1, UCase(rec(I).name), "DT") > 0 Or InStr(1, UCase(rec(I).name), "ID") > 0 Or InStr(1, UCase(rec(I).name), "DAY") > 0 Or InStr(1, UCase(rec(I).name), "MOBILE") > 0 Then
            If lv.ColumnHeaders.COUNT > 0 Then
                lv.ColumnHeaders.Add , , STR, , lvwColumnCenter
            Else
                lv.ColumnHeaders.Add , , STR, , lvwColumnLeft
            End If
        Else
            lv.ColumnHeaders.Add , , STR, , lvwColumnLeft
        End If
    Next
    While Not rec.EOF
        Set li = lv.ListItems.Add(, , Trim(rec(0)))
        For I = 1 To rec.Fields.COUNT - 1
        If InStr(1, UCase(rec(I).name), "ROU") > 0 Or InStr(1, UCase(rec(I).name), "CST") > 0 Or InStr(1, UCase(rec(I).name), "PCG") > 0 Or InStr(1, UCase(rec(I).name), "AMT") > 0 Or InStr(1, UCase(rec(I).name), "AMOUNT") > 0 Or InStr(1, UCase(rec(I).name), "DIS") > 0 Or InStr(1, UCase(rec(I).name), "VAL") > 0 Or InStr(1, UCase(rec(I).name), "TAX") > 0 Then
                li.ListSubItems.Add , , IIf(IsNull(rec(I)) = True, "", Format(Trim(rec(I)), "########0.00"))
            Else
                li.ListSubItems.Add , , IIf(IsNull(rec(I)) = True, "", Trim(rec(I)))
            End If
        Next
        rec.MoveNext
    Wend
    rec.Close
    lblRecords.Caption = lv.ListItems.COUNT
End Function

Private Sub query()
lblRecords.Caption = "0"
lblSelected.Caption = "0"
Dim SqlStatement As String
Dim ToDate As String
Dim FromDate As String
Dim STR As String
If Combo1(0).Visible = True And cmbDescription.ListIndex < 2 Then
    Select Case cmbDescription.ListIndex
        Case 0
            STR = IIf(Combo1(0).ListIndex = 0, "grp_idy", "grp_nme")
        Case 1
            STR = IIf(Combo1(0).ListIndex = 0, "act_idy", "act_nme")
        Case 2
            STR = IIf(Combo1(0).ListIndex = 0, "grp_cat", "grp_cls")
        Case 3
            STR = IIf(Combo1(0).ListIndex = 0, "prd_idy", "prd_nme")
    End Select
ElseIf Combo1(1).Visible = True And cmbDescription.ListIndex >= 3 And cmbDescription.ListIndex < 8 Then
    If Combo1(1).ListIndex = 0 Then STR = "txn_idy"
    If Combo1(1).ListIndex = 1 Then STR = "a1.act_nme"
    If Combo1(1).ListIndex = 2 Then STR = "a2.act_nme"
ElseIf Combo1(2).Visible = True And cmbDescription.ListIndex >= 8 And cmbDescription.ListIndex < cmbDescription.ListCount Then
    If Combo1(2).ListIndex = 0 Then STR = "INL.inv_idy"
    If Combo1(2).ListIndex = 1 Then STR = "a2.act_nme"
    If Combo1(2).ListIndex = 2 Then STR = "a1.act_nme"
End If

FromDate = Month(dtpFrom.Value) & "/" & Day(dtpFrom.Value) & "/" & Year(dtpFrom.Value)
ToDate = Month(dtpTo.Value) & "/" & Day(dtpTo.Value) & "/" & Year(dtpTo.Value)

lvValues.ListItems.CLEAR
'lvValues.ColumnHeaders.CLEAR
    If cmbDescription.ListIndex = 0 Then
        SqlStatement = "select  DISTINCT grp_idy AS ID, grp_nme AS Group_Name, grp_clf AS Group_Type from grp"
            If Trim(txtFilter.Text) <> "" Then
            If STR <> "" Then SqlStatement = SqlStatement & " WHERE " & STR & " like '" & txtFilter.Text & "%' "
            End If
        FillListView lvValues, SqlStatement
    ElseIf cmbDescription.ListIndex = 1 Then
        SqlStatement = "select  DISTINCT act_idy as Id, act_nme as Account_Name, grp_nme as Group_nAME from act,grp where TRIM(grp.grp_idy) = TRIM(act.grp_idy) "
            If Trim(txtFilter.Text) <> "" Then
            If STR <> "" Then SqlStatement = SqlStatement & " and " & STR & " like '" & txtFilter.Text & "%' "
            End If
        FillListView lvValues, SqlStatement
'    ElseIf cmbDescription.ListIndex = 2 Then
'        SqlStatement = "select  DISTINCT grp_idy, grp_cat, grp_cls from pgr  "
'        If STR <> "" Then SqlStatement = SqlStatement & " where " & STR & " like '" & txtFilter.Text & "%'  "
'        FillListView lvValues, SqlStatement
'    ElseIf cmbDescription.ListIndex = 3 Then
'        SqlStatement = "select   DISTINCT prd_idy, prd_nme, PGR.GRP_CLS, prd_uom, pur_rat, sal_rat, mrp_rat, prd.cur_lev, prd.reo_lev, prd.saf_lev from pgr, prd where pgr.grp_idy = prd.grp_idy  "
'        If STR <> "" Then SqlStatement = SqlStatement & " AND " & STR & " like '" & txtFilter.Text & "%' "
'        FillListView lvValues, SqlStatement
    ElseIf cmbDescription.ListIndex >= 3 And cmbDescription.ListIndex <= 7 Then
        SqlStatement = "select txn_idy AS ID, txn_dat AS TXN_DATE, txn_typ AS TXN_TYPE, txn_amt AS TXN_AMT, a1.act_nme as dr_acT, a2.act_nme as cr_ACT  " & _
                " from vch, act  a1, act  a2 " & _
                " where vch.dbt_idy = a1.act_idy And vch.crt_idy = a2.act_idy  " & _
                " and vch.txn_typ='" & VchType(cmbDescription.ListIndex - 3) & "'" & _
                " AND txn_dat  >= CTOD('" & FromDate & "') " & _
                " AND txn_dat  <= CTOD('" & ToDate & "') "
            If Trim(txtFilter.Text) <> "" Then
            If STR <> "" Then SqlStatement = SqlStatement & " and " & STR & " like '" & txtFilter.Text & "%' "
            End If
         SqlStatement = SqlStatement & " order by txn_idy"
        FillListView lvValues, SqlStatement
    ElseIf cmbDescription.ListIndex >= 9 And cmbDescription.ListIndex <= 14 Then
        SqlStatement = "select INL.inv_idy, INL.INV_DAT, INL.txn_typ, sum(inv.col_amt) as prd_val, sum(INV.DIS_AMT) as prd_dis," & _
            " sum(inv.tax_amt) as prd_tax,  INL.grs_amt, INL.spd_pcg, INL.spd_amt, INL.csd_pcg, INL.csd_amt, INL.exd_pcg, " & _
            " INL.exd_amt, INL.tax_pcg, INL.tax_amt as inl_tax_amt, INL.net_amt, INL.frt_amt, INL.pkg_cst, " & _
            " INL.rou_off , INL.cst_idy, INL.act_idy, a1.act_nme as CUST_aCT, a2.act_nme as aCCOUNT " & _
            " from inl, inv, act  a1, act  a2 " & _
            " where INL.inv_idy = inv.inv_idy and INL.cst_idy = a1.act_idy and INL.act_idy = a2.act_idy " & _
            " and INL.txn_typ='" & IVchType(cmbDescription.ListIndex - 8) & "' " & _
            " and INL.INV_DAT   >= CTOD('" & FromDate & "')  " & _
            " and INL.INV_DAT   <= CTOD('" & ToDate & "')  "
            If Trim(txtFilter.Text) <> "" Then
            If STR <> "" Then SqlStatement = SqlStatement & " and " & STR & " like '" & txtFilter.Text & "%' "
            End If
            SqlStatement = SqlStatement & " group by INL.inv_idy, INL.INV_DAT, INL.txn_typ, INL.cst_idy, INL.act_idy, INL.grs_amt, INL.spd_pcg,  INL.spd_amt, INL.csd_pcg, INL.csd_amt,  INL.exd_pcg, INL.exd_amt, INL.tax_pcg, " & _
            " INL.tax_amt,  INL.net_amt, INL.frt_amt, INL.pkg_cst,INL.rou_off,a1.act_nme, a2.act_nme, txn_rm1, txn_rm2"
        FillListView lvValues, SqlStatement
    End If
    
    cmdSelectAll.Caption = "Select &All"
    cmdSelectAll_Click
    cmdSelectAll.Caption = "Deselect &All"
End Sub

Private Function TransferMaster() As Long
On err GoTo ErrHandler
    Dim RESULT As Long
    Dim STR As String
    Dim Counter As Integer
    
    I = cmbDescription.ListIndex
    STR = ""
    If cmbDescription.ListIndex = 0 Then
        modProcess.CreatePrimaryGroups
    End If
    If I >= 0 And I <= 2 Then
            For I = 1 To lvValues.ListItems.COUNT
                If lvValues.ListItems(I).Checked = True Then
                    If InStr(1, STR, lvValues.ListItems(I).Text) <= 0 Then
                        STR = STR & "'" & lvValues.ListItems(I).Text & "'" & ","
                        
                            If Counter = 10 Then
                                If STR <> "" Then
                                    If InStr(1, STR, ",") > 0 Then STR = Mid(STR, 1, Len(STR) - 1)
                                    If cmbDescription.ListIndex = 0 Then
                                        modProcess.CreateGroups STR
                                    End If
                                    If cmbDescription.ListIndex = 1 Then
                                        modProcess.CreateAccounts STR
                                '-------create voucher opening balances for a specific balances
                                        CreateAccountingVoucher ("OPN"), STR
                                    End If
                                End If
                                Counter = 0
                                STR = ""
                            End If
                            Counter = Counter + 1
                    End If
                End If
            Next
    End If
    If InStr(1, STR, ",") > 0 Then STR = Mid(STR, 1, Len(STR) - 1)
    If cmbDescription.ListIndex = 0 Then
        modProcess.CreateGroups STR
    End If
    If cmbDescription.ListIndex = 1 Then
        modProcess.CreateAccounts STR
'-------create voucher opening balances for a specific balances
        CreateAccountingVoucher ("OPN"), STR
    End If
                                    
    TransferMaster = RESULT
    Exit Function
ErrHandler:
    DispMessage RESULT
    TransferMaster = RESULT
End Function

Private Function TransferAccountsTrans()
'On err GoTo ErrHandler
'OPN - Opening Balance
    Dim I As Integer
    Dim STR As String
        Dim Counter As Integer

    I = cmbDescription.ListIndex - 3
    STR = ""
    If I >= 0 And I <= 4 Then
        For I = 1 To lvValues.ListItems.COUNT
            If lvValues.ListItems(I).Checked = True Then
                If InStr(1, STR, lvValues.ListItems(I).Text) <= 0 Then
                    STR = STR & "'" & lvValues.ListItems(I).Text & "'" & ","
                    If Counter = 20 Then
                        STR = Mid(STR, 1, Len(STR) - 1)
                        modProcess.CreateAccountingVoucher VchType(cmbDescription.ListIndex - 3), STR
                        Counter = 0
                        STR = ""
                    End If
                End If
            End If
        Next
        If STR = "" Then
            MsgBox "No Value From the List Is Selected"
            Exit Function
        End If
        STR = Mid(STR, 1, Len(STR) - 1)
        modProcess.CreateAccountingVoucher VchType(cmbDescription.ListIndex - 3), STR
    End If
'ErrHandler:
End Function

Private Function TransferInventoryTrans()
'On err GoTo ErrHandler
'OPN - Opening stock
    Dim I As Integer
    Dim STR As String
    Dim Counter As Integer
    I = cmbDescription.ListIndex - 9
    STR = ""
    If I >= 0 And I <= 6 Then
        For I = 1 To lvValues.ListItems.COUNT
            If lvValues.ListItems(I).Checked = True Then
                If InStr(1, STR, lvValues.ListItems(I).Text) <= 0 Then
                    STR = STR & "'" & lvValues.ListItems(I).Text & "'" & ","
                    If Counter = 20 Then
                        STR = Mid(STR, 1, Len(STR) - 1)
                        modProcess.CreateInventoryVoucher IVchType(cmbDescription.ListIndex - 8), STR
                        Counter = 0
                        STR = ""
                    End If
                    Counter = Counter + 1
                End If
            End If
        Next
'        If STR = "" Then
'            MsgBox "No Value From the List Is Selected"
'            Exit Function
'        End If
                         STR = Mid(STR, 1, Len(STR) - 1)
                        modProcess.CreateInventoryVoucher IVchType(cmbDescription.ListIndex - 8), STR
   End If
'ErrHandler:
End Function


Private Sub DispMessage(RESULT As Long)
    Dim STR As String



    Select Case RESULT
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
    MsgBox "        " & STR & "                      ", vbOKOnly + vbInformation, "DAS Version 4.0"
End Sub

