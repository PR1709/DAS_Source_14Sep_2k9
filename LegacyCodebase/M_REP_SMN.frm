VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_SMN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "M_REP_SMN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Report Filters"
      TabPicture(0)   =   "M_REP_SMN.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_SMN.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Select Salesman"
         Height          =   1215
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   3975
         Begin VB.CheckBox Check6 
            Caption         =   "All &Salesman"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   360
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Salesman"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   795
            Width           =   690
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Dates"
         Height          =   700
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   3975
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   2520
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16580611
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   600
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16580611
            CurrentDate     =   37258
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   308
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   2160
            TabIndex        =   11
            Top             =   308
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   4320
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Salesman"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         Width           =   5505
      End
      Begin VB.Label Label9 
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4440
         TabIndex        =   2
         Top             =   2715
         Visible         =   0   'False
         Width           =   45
      End
   End
End
Attribute VB_Name = "M_REP_SMN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check6_Click()
If Check6.Value = 0 Then
    Check6.Enabled = False
    Text6.Enabled = True
    Text6.SetFocus
Else
    Command1.SetFocus
End If
End Sub
Private Sub Check6_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    If Check6.Value = 1 Then
        Command1.SetFocus
    Else
        Text6.Enabled = True
        Text6.SetFocus
    End If
End If
End Sub
Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Command1.Enabled = False
If rs.State = 1 Then rs.Close
'dates
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
    SAL_QRY = "inv->inv_dat>=ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "').and.inv->inv_dat<=ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "')"
End If
con.Execute "UPDATE DIR SET DIR_MSG = '" & Trim(Text6) & "' WHERE DIR_IDY = 'SMN' AND DIR_IDX = 'IDY'"
If Me.Tag = "CUM" Then
    'UPDATEOB
    Smn_OpgBalUpdate        'New Update Route written by PNR
    If Check6.Value = 1 Then
        SAL_QRY = "INV_DAT >= CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') .AND. INV_DAT <= CTOD('" & Format(DTPicker2, "mm/DD/yyyy") & "')"
        OpenReport CURDIR & "\REP_LIB.RP1", "Van Cumulative Balance", 2, SAL_QRY, 6
        Clear_Balance
    Else
        SAL_QRY = "INV_DAT >= CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') .AND. INV_DAT <= CTOD('" & Format(DTPicker2, "mm/DD/yyyy") & "').AND. VAN->SMN_IDY = '" & Trim(Label9) & "'"
        OpenReport CURDIR & "\REP_LIB.RP1", "Van Cumulative Balance", 2, SAL_QRY, 6
        Clear_Balance
    End If
ElseIf Me.Tag = "SHT_CSH" Then
    Smn_OpgBalUpdate        'New Update Route written by PNR
    If Check6.Value = 1 Then
        SAL_QRY = "INV_DAT >= CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') .AND. INV_DAT <= CTOD('" & Format(DTPicker2, "mm/DD/yyyy") & "')"
        OpenReport CURDIR & "\REP_LIB.RP1", "Salesman Shortage Report", 2, SAL_QRY, 6
        Clear_Balance
    Else
        SAL_QRY = "INV_DAT >= CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') .AND. INV_DAT <= CTOD('" & Format(DTPicker2, "mm/DD/yyyy") & "').AND. VAN->SMN_IDY = '" & Trim(Label9) & "'"
        OpenReport CURDIR & "\REP_LIB.RP1", "Salesman Shortage Report", 2, SAL_QRY, 6
        Clear_Balance
    End If
ElseIf Me.Tag = "EXP_RPT" Then
    Smn_OpgBalUpdate        'New Update Route written by PNR
    If Check6.Value = 1 Then
        SAL_QRY = "INV_DAT >= CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') .AND. INV_DAT <= CTOD('" & Format(DTPicker2, "mm/DD/yyyy") & "')"
        OpenReport CURDIR & "\REP_LIB.RP1", "Route Expense Report", 2, SAL_QRY, 6
        Clear_Balance
    Else
        SAL_QRY = "INV_DAT >= CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') .AND. INV_DAT <= CTOD('" & Format(DTPicker2, "mm/DD/yyyy") & "').AND. VAN->SMN_IDY = '" & Trim(Label9) & "'"
        OpenReport CURDIR & "\REP_LIB.RP1", "Route Expense Report", 2, SAL_QRY, 6
        Clear_Balance
    End If
ElseIf Me.Tag = "VAN" Then
    updatevan
ElseIf Me.Tag = "SHL" Then
    SAL_QRY = "SNV->INV_DAT >= ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "') .and. SNV->INV_DAT <=ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "') .and. VAN->SMN_IDY = '" & Trim(Label9) & "'"
    OpenReport CURDIR & "\REP_LIB.RP1", "Salesman Reconcile Shells", 2, SAL_QRY, 6
    'updateshells - commented by PNR
ElseIf Me.Tag = "VSV" Then
    SAL_QRY = "SNV->INV_DAT >= ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "') .and. SNV->INV_DAT <=ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "') .and. SNV->SMN_IDY = '" & Trim(Label9) & "'"
    OpenReport CURDIR & "\REP_LIB.RP1", "Salesman Settlement Shells", 2, SAL_QRY, 6
ElseIf Me.Tag = "COMREP" Or Me.Tag = "VAN_SCR" Then
    SAL_QRY = "VAN->INV_DAT >= ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "').and. VAN->INV_DAT <= ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "') .and. (VAN->VAN_CLS = 'S' .OR. VAN->VAN_CLS = 'C')"
    If Check6.Value = 1 Then
    Else
    SAL_QRY = SAL_QRY & ".AND. VAN->SMN_IDY = '" & Trim(Label9) & "'"
    End If
    OpenReport CURDIR & "\REP_LIB.RP1", "Commission Report", 2, SAL_QRY, 6
    OpenReport CURDIR & "\REP_LIB.RP1", "Commission Summary", 2, SAL_QRY, 6
End If
Unload Me
End Sub
Private Sub updatevan()
Dim RSINV As New ADODB.Recordset
Dim rsVan As New ADODB.Recordset
con.Execute "DELETE FROM CLD"
If rsVan.State = 1 Then rsVan.Close
rsVan.Open "SELECT SUM(((tot_qty- int(tot_qty))*1000)+ (int(tot_qty)*sub_unt)) p_tot_qty,SUM(((ret_qty- int(ret_qty))*1000)+ (int(ret_qty)*sub_unt)+((brk_qty- int(brk_qty))*1000)+ (int(brk_qty)*sub_unt)+((lkg_qty- int(lkg_qty))*1000)+ (int(lkg_qty)*sub_unt)) p_ret_qty,*  FROM VAN,PRD WHERE PRD.PRD_IDY LIKE VAN.PRD_IDY AND INV_Dat >= ctod('" & DTPicker1 & "') and inv_dat<=ctod('" & DTPicker2 & "') and smn_idy like '" & Trim(Label9) & "%' group by VAN.prd_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rsVan.RecordCount
'    CON.Execute "INSERT INTO CLD VALUES('" & Trim(rsVan!PRD_IDY) & "'," & Val(rsVan!p_tot_qty) & "," & Val(rsVan!p_ret_qty) & ",0,0,0,0,0,0)"
    con.Execute "INSERT INTO CLD VALUES('" & Trim(rsVan!BAT_IDY) & "'," & Val(rsVan!p_tot_qty) & "," & Val(rsVan!p_ret_qty) & ",0,0,0,0,0,0)" 'Added by PNR
    rsVan.MoveNext
Next
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT SUM(TOT_PCS),BAT_IDY,STK_TYP FROM INV WHERE ORD_IDY in(SELECT INV_IDY FROM VAN WHERE INV_Dat >= ctod('" & DTPicker1 & "') and inv_dat<=ctod('" & DTPicker2 & "') and smn_idy like '" & Trim(Label9) & "%' group by inv_idy)AND STK_TYP IN('Salable','Free','Damage','TakeBack','Breakage','Leakage') GROUP BY BAT_IDY,STK_TYP", con, adOpenKeyset, adLockPessimistic
For A = 1 To RSINV.RecordCount
    If Trim(RSINV(2)) = "Salable" Or Trim(RSINV(2)) = "Free" Then
        con.Execute "UPDATE CLD SET PRT_STK= PRT_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!BAT_IDY) & "'", d
        If d = 0 Then con.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!BAT_IDY) & "',0,0," & Val(RSINV(0)) & ",0,0,0,0,0)"
    Else
        con.Execute "UPDATE CLD SET SAL_STK= SAL_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!BAT_IDY) & "'", d
        If d = 0 Then con.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!BAT_IDY) & "',0,0,0," & Val(RSINV(0)) & ",0,0,0,0)"
    End If
    RSINV.MoveNext
Next
OpenReport CURDIR & "\REP_LIB.RP1", "Van Variance Report", 0, SAL_QRY, 6
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub
Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Check6.Visible = True Then
    Check6.Enabled = True
    Check6.SetFocus
Else
    Text6.SetFocus
End If
ElseIf KeyCode = 38 And Shift = 1 Then
DTPicker1.SetFocus
End If
End Sub
Private Sub DTPicker2_LostFocus()
If DTPicker1 > DTPicker2 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker1.SetFocus
    Exit Sub
End If
End Sub
Private Sub Form_Activate()
'If Me.Tag = "CUM" Then         'Code Changed by PNR as per Prince Chopra Request
'    Check6.Visible = False
'    Text6.Enabled = True
'End If
SETYEAROPEN M_REP_SMN
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
Check6.Value = 1
DTPicker1.SetFocus
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        If Label20.Caption = "Select Salesman" Then Text6.SetFocus
    Else
    
        Unload Me
    End If
End If
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If list1.ListItems.count > 0 Then
        If Label20 = "Select Salesman" Then
            Label9 = list1.SelectedItem.ListSubItems.Item(1)
            Text6 = Trim(list1.SelectedItem)
            Check6.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text6_Change()
If Text6 = "" Then Label9 = ""
End Sub

Private Sub TEXT6_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("smn", "smn_idy", Trim(Label9)) = True Then
    If Check6.Enabled = True Then Check6.SetFocus
Else
    rs.Open "select * from smn where smn_nme like '" & Trim(Text6) & "%'", con, adOpenStatic
    list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
        rs.MoveNext
    Next
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    SSTab1.Tab = 1
    Label20 = "Select Salesman"
    list1.ColumnHeaders(1).Text = "Salesman Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text6.Enabled = True
        Text6.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub Text6_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub
Public Sub Smn_OpgBalUpdate()
    
    Dim RS_SMN As New ADODB.Recordset
    If RS_SMN.State = 1 Then RS_SMN.Close
    RS_SMN.Open "SELECT * FROM SMN", con, adOpenKeyset, adLockPessimistic
    
    Dim rs As New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT SUM(((((tot_qty- int(tot_qty))*1000)+ (int(tot_qty)*sub_unt)) - (((ret_qty- int(ret_qty))*1000)+ (int(ret_qty)*sub_unt)+((brk_qty- int(brk_qty))*1000)+ (int(brk_qty)*sub_unt)+((lkg_qty- int(lkg_qty))*1000)+ (int(lkg_qty)*sub_unt) + ((fre_qty- int(fre_qty))*1000)+ (int(fre_qty)*sub_unt)))*(PRD_RAT/SUB_UNT)), SMN_IDY FROM VAN,PRD WHERE VAN.PRD_IDY LIKE PRD.PRD_IDY  AND INV_DAT < CTOD('" & DTPicker1 & "') GROUP BY SMN_IDY ", con, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount > 0 Then rs.MoveFirst
    For A = 1 To rs.RecordCount
        RS_SMN.MoveFirst
        RS_SMN.Find "SMN_IDY = '" & rs("SMN_IDY") & "'"
        If RS_SMN.EOF = False Then
            RS_SMN("SMN_AMT") = Round(Val(rs(0)), 2) & ""
            RS_SMN.Update
        End If
        rs.MoveNext
    Next

    If RS_SMN.State = 1 Then RS_SMN.Close
    RS_SMN.Open "SELECT * FROM SMN", con, adOpenKeyset, adLockPessimistic
    If RS_SMN.RecordCount > 0 Then RS_SMN.MoveFirst
    
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT CAS_AMT+CHQ_AMT,COM_AMT,ADD_AMT,SMN_IDY FROM VAN WHERE INV_DAT < CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    
    For A = 1 To rs.RecordCount
        RS_SMN.Find "SMN_IDY = '" & rs("SMN_IDY") & "'"
        If RS_SMN.EOF = False Then
            RS_SMN("SMN_AMT") = Val(rs(0) & "") - RS_SMN("SMN_AMT")
            RS_SMN("SMN_AMT") = Val(rs(1) & "") - RS_SMN("SMN_AMT")
            RS_SMN("SMN_AMT") = Val(rs(2) & "") + RS_SMN("SMN_AMT")
        End If
        rs.MoveNext
    Next
    rs.Close

End Sub
Public Sub UPDATEOB()
'If Check6.Value = 1 Then
'    Dim RS_SMN As New ADODB.Recordset
'    If RS_SMN.State = 1 Then RS_SMN.Close
'    RS_SMN.Open "SELECT * FROM SMN", CON, adOpenKeyset, adlockpessimistic
'
'    Dim rs As New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "SELECT SUM(((((tot_qty- int(tot_qty))*1000)+ (int(tot_qty)*sub_unt)) - (((ret_qty- int(ret_qty))*1000)+ (int(ret_qty)*sub_unt)+((brk_qty- int(brk_qty))*1000)+ (int(brk_qty)*sub_unt)+((lkg_qty- int(lkg_qty))*1000)+ (int(lkg_qty)*sub_unt) + ((fre_qty- int(fre_qty))*1000)+ (int(fre_qty)*sub_unt)))*(PRD_RAT/SUB_UNT)), SMN_IDY FROM VAN,PRD WHERE VAN.PRD_IDY LIKE PRD.PRD_IDY  AND INV_DAT < CTOD('" & DTPicker1 & "') GROUP BY SMN_IDY ", CON, adOpenKeyset, adlockpessimistic
'
'    'AND SMN_IDY LIKE '" & Trim(Label9) & "'
'
'    If rs.RecordCount > 0 Then rs.MoveFirst 'SALAMT = rs(0) & ""
'    For A = 1 To rs.RecordCount
'        RS_SMN.MoveFirst
'        RS_SMN.Find "SMN_IDY = '" & rs("SMN_IDY") & "'"
'        If RS_SMN.EOF = False Then
'            RS_SMN("SMN_AMT") = Val(rs(0)) & ""
'            RS_SMN.Update
'        End If
'        rs.MoveNext
'    Next
'
'    If RS_SMN.State = 1 Then RS_SMN.Close
'    RS_SMN.Open "SELECT * FROM SMN", CON, adOpenKeyset, adlockpessimistic
'
'    If rs.State = 1 Then rs.Close
'    rs.Open "SELECT CAS_AMT+CHQ_AMT,COM_AMT,ADD_AMT,SMN_IDY FROM VAN WHERE INV_DAT < CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') GROUP BY INV_IDY", CON, adOpenKeyset, adlockpessimistic
'    ' SMN_IDY LIKE '" & Trim(Label9) & "' AND
'
'    For A = 1 To rs.RecordCount
'        RS_SMN.MoveFirst
'        RS_SMN.Find "SMN_IDY = '" & rs("SMN_IDY") & "'"
'        If RS_SMN.EOF = False Then
'            RS_SMN("SMN_AMT") = Val(rs(0) & "") - RS_SMN("SMN_AMT")
'            RS_SMN("SMN_AMT") = Val(rs(1) & "") - RS_SMN("SMN_AMT")
'            RS_SMN("SMN_AMT") = Val(rs(2) & "") + RS_SMN("SMN_AMT")
'        End If
'        rs.MoveNext
'    Next
'''    RS_SMN.Close
'    rs.Close
'Else
'    'Dim rs As New ADODB.Recordset
'    rs.Open "SELECT SUM(((((tot_qty- int(tot_qty))*1000)+ (int(tot_qty)*sub_unt)) - (((ret_qty- int(ret_qty))*1000)+ (int(ret_qty)*sub_unt)+((brk_qty- int(brk_qty))*1000)+ (int(brk_qty)*sub_unt)+((lkg_qty- int(lkg_qty))*1000)+ (int(lkg_qty)*sub_unt) + ((fre_qty- int(fre_qty))*1000)+ (int(fre_qty)*sub_unt)))*(PRD_RAT/SUB_UNT)) FROM VAN,PRD WHERE VAN.PRD_IDY LIKE PRD.PRD_IDY AND SMN_IDY LIKE '" & Trim(Label9) & "' AND INV_DAT < CTOD('" & DTPicker1 & "') ", CON, adOpenKeyset, adlockpessimistic
'    If rs.RecordCount > 0 Then SALAMT = rs(0) & ""
'    If rs.State = 1 Then rs.Close
'    rs.Open "SELECT CAS_AMT+CHQ_AMT,COM_AMT,ADD_AMT FROM VAN WHERE SMN_IDY LIKE '" & Trim(Label9) & "' AND INV_DAT < CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "') GROUP BY INV_IDY ", CON, adOpenKeyset, adlockpessimistic
'    For A = 1 To rs.RecordCount
'        CASAMT = Val(rs(0) & "") + CASAMT
'        COMAMT = Val(rs(1) & "") + COMAMT
'        ADDAMT = Val(rs(2) & "") + ADDAMT
'        rs.MoveNext
'    Next
'    CON.Execute "UPDATE DIR SET DIR_MSG = '" & (Val(SALAMT) + Val(ADDAMT)) - (Val(COMAMT) + Val(CASAMT)) & "' WHERE DIR_IDY = 'PRO' AND DIR_IDX = 'LOS'"
'End If
End Sub
Public Sub updateshells()
Dim RSINV As New ADODB.Recordset
Dim rsVan As New ADODB.Recordset
con.Execute "DELETE FROM CLD"
If rsVan.State = 1 Then rsVan.Close
rsVan.Open "SELECT PRD_IDY,SUM(I_SHL_NOS),SUM(I_PCK_NOS),SUM(R_SHL_NOS),SUM(R_PCK_NOS) FROM SNV WHERE INV_IDY IN(SELECT INV_IDY FROM VAN WHERE INV_Dat >= ctod('" & DTPicker1 & "') and inv_dat<=ctod('" & DTPicker2 & "') and smn_idy like '" & Trim(Label9) & "%') GROUP BY PRD_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rsVan.RecordCount
    con.Execute "INSERT INTO CLD VALUES('" & Trim(rsVan!prd_idy) & "'," & Val(rsVan(1)) & "," & Val(rsVan(2)) & "," & Val(rsVan(3)) & "," & Val(rsVan(4)) & ",0,0,0,0)"
    rsVan.MoveNext
Next
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT INV_IDY FROM INL WHERE ORD_IDY IN(SELECT INV_IDY FROM VAN WHERE INV_Dat >= ctod('" & DTPicker1 & "') and inv_dat<=ctod('" & DTPicker2 & "') and smn_idy like '" & Trim(Label9) & "%' group by inv_idy)", con, adOpenKeyset, adLockPessimistic
con.Execute "DELETE FROM NPO"
For A = 1 To RSINV.RecordCount
    con.Execute "INSERT INTO NPO VALUES('" & Trim(RSINV(0)) & "')"
    RSINV.MoveNext
Next
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT PRD_IDY,SUM(R_SHL_NOS),SUM(R_PCK_NOS) FROM SNV WHERE INV_IDY IN(SELECT CST_IDY FROM NPO) GROUP BY PRD_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To RSINV.RecordCount
    con.Execute "UPDATE CLD SET SRN_STK =" & Val(RSINV(1)) & ",STI_STK=" & Val(RSINV(2)) & " WHERE PRD_IDY LIKE '" & Trim(RSINV(0)) & "'", d
    If d = 0 Then con.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV(0)) & "',0,0,0,0," & Val(RSINV(1)) & "," & Val(RSINV(2)) & ",0,0)"
    RSINV.MoveNext
Next
If Me.Tag = "SHL" Then
    OpenReport CURDIR & "\REP_LIB.RP1", "Shells Variance", 0, SAL_QRY, 6
ElseIf Me.Tag = "VSV" Then
    OpenReport CURDIR & "\REP_LIB.RP1", "Van Shells Variance", 0, SAL_QRY, 6
End If
End Sub
Public Sub Clear_Balance()
Dim RS_SMN As New ADODB.Recordset
If RS_SMN.State = 1 Then RS_SMN.Close
RS_SMN.Open ("SELECT * FROM SMN"), con, adOpenKeyset, adLockPessimistic
While RS_SMN.EOF = False
    RS_SMN("SMN_AMT") = 0#
    RS_SMN.Update
    RS_SMN.MoveNext
Wend
End Sub
