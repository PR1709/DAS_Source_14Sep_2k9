VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_VAN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "M_REP_VAN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Report Filters"
      TabPicture(0)   =   "M_REP_VAN.frx":0442
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
      TabPicture(1)   =   "M_REP_VAN.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "lv"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Select Vehicle"
         Height          =   725
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   3975
         Begin VB.TextBox vanid 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   960
            MaxLength       =   40
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   247
            Width           =   2895
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Vehicle"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   315
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Dates"
         Height          =   725
         Left            =   240
         TabIndex        =   8
         Top             =   960
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
            Format          =   74317827
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   840
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74317827
            CurrentDate     =   37258
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   315
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   2200
            TabIndex        =   11
            Top             =   315
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   4320
         TabIndex        =   4
         Top             =   960
         Width           =   1215
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView lv 
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
         Caption         =   "Select Vehicle"
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
Attribute VB_Name = "M_REP_VAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Check6_Click()
'If Check6.Value = 0 Then
'    Check6.Enabled = False
'    Text6.Enabled = True
'    Text6.SetFocus
'Else
'    Command1.SetFocus
'End If
'End Sub
'
'Private Sub Check6_KeyDown(K As Integer, S As Integer)
'If K = 13 Then
'    If Check6.Value = 1 Then
'        Command1.SetFocus
'    Else
'        Text6.Enabled = True
'        Text6.SetFocus
'    End If
'End If
'End Sub

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
    SAL_QRY = "van->inv_dat>=ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "').and.van->inv_dat<=ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "')"
End If
con.Execute "UPDATE DIR SET DIR_MSG = '" & Trim(vanid) & "' WHERE DIR_IDY = 'VAN' AND DIR_IDX = 'IDY'"
If Me.Tag = "CHL_STS" Then
    OpenReport CURDIR & "\VVR_RPT.RP1", "Van Challans Status", 2, SAL_QRY, 6
Else
    SAL_QRY = SAL_QRY + ".and. ACT->ACT_NME = '" & Trim(vanid) & "'" 'Report to be Redone by PNR & TSR
    OpenReport CURDIR & "\SAL_PRG.RP1", "Vehicle Utilization", 2, SAL_QRY, 6
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
    con.Execute "INSERT INTO CLD VALUES('" & Trim(rsVan!prd_idy) & "'," & Val(rsVan!p_tot_qty) & "," & Val(rsVan!p_ret_qty) & ",0,0,0,0,0,0)"
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
OpenReport CURDIR & "\REP_LIB.RP1", "Van Variance", 0, SAL_QRY, 6
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    vanid.Enabled = True
    vanid.SetFocus
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
SETYEAROPEN M_REP_VAN
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
DTPicker1.SetFocus
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If lv.Visible = True Then
        lv.Visible = False
        SSTab1.Tab = 0
        If Label20.Caption = "Select Salesman" Then Text6.SetFocus
    Else
    
        Unload Me
    End If
End If
End Sub

Private Sub LV_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If lv.SortOrder = lvwAscending Then
lv.SortOrder = lvwDescending
ElseIf lv.SortOrder = lvwDescending Then
lv.SortOrder = lvwAscending
End If
lv.SortKey = CH.Index - 1
End Sub

Private Sub LV_DblClick()
lv_KeyDown 13, 0
End Sub

Private Sub lV_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If lv.ListItems.count > 0 Then
        vanid = lv.SelectedItem
        SSTab1.Tab = 0
        lv.Visible = False
        Command1.SetFocus
    End If
End If
End Sub
Public Sub UPDATEOB()
Dim rs As New ADODB.Recordset
rs.Open "SELECT SUM(((((tot_qty- int(tot_qty))*1000)+ (int(tot_qty)*sub_unt)) - (((ret_qty- int(ret_qty))*1000)+ (int(ret_qty)*sub_unt)+((brk_qty- int(brk_qty))*1000)+ (int(brk_qty)*sub_unt)+((lkg_qty- int(lkg_qty))*1000)+ (int(lkg_qty)*sub_unt)))*(PRD_RAT/SUB_UNT)) FROM VAN,PRD WHERE VAN.PRD_IDY LIKE PRD.PRD_IDY AND SMN_IDY LIKE '" & Trim(Label9) & "' AND INV_DAT < CTOD('" & DTPicker1 & "') ", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then SALAMT = rs(0) & ""
If rs.State = 1 Then rs.Close
rs.Open "SELECT CAS_AMT+CHQ_AMT,COM_AMT FROM VAN WHERE SMN_IDY LIKE '" & Trim(Label9) & "' AND INV_DAT <CTOD('" & DTPicker1 & "') GROUP BY INV_IDY ", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    CASAMT = Val(rs(0) & "") + CASAMT
    COMAMT = Val(rs(1) & "") + COMAMT
    rs.MoveNext
Next
con.Execute "UPDATE DIR SET DIR_MSG = '" & Val(SALAMT) - (Val(COMAMT) + Val(CASAMT)) & "' WHERE DIR_IDY = 'PRO' AND DIR_IDX = 'LOS'"
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
OpenReport CURDIR & "\REP_LIB.RP1", "Van Shells Variance", 0, SAL_QRY, 6
End Sub
Private Sub vanidGotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Vehicle ID"
End Sub

Private Sub VANID_KeyDown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If CHECKTB("act.dbf,smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_CST_IDY)) & "' AND (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') AND DBT_TYP='Y' and cst_typ like 'Van'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        vanid = Sincrement(vanid)
        If Len(Trim(vanid)) = 0 Then Exit Sub
        M_CST_NME = Trim(REC_SET("act_nme") & "")
        If Len(Trim(REC_SET("ROU_IDY"))) <> 0 Then M_ROU_IDY = Trim(REC_SET("ROU_IDY") & "")
'        SMN = Split(GetSalesman(vanid, M_INV_DAT), "|")
'        If UBound(SMN) > 0 Then
'            M_SMN_IDY = Trim(SMN(0))
'            M_SMN_NME = Trim(SMN(1))
'        End If
        If REC_SET1.State = 1 Then REC_SET1.Close
        M_DBT_RAP = Trim(REC_SET("dbt_rap"))
        SSTab1.Tab = 0
'        If K = 38 Then
'            M_INV_DAT.SetFocus
'        Else
'            If TRP_NOS.Enabled = True Then
'                TRP_NOS.SetFocus
'            Else
'                M_SMN_IDY.SetFocus
'            End If
'        End If
    Else
        V_DLR_LST
    End If
End If
End Sub
Private Sub V_DLR_LST()
If CHECKTB("act.dbf,mar.dbf,smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
'If Check1.Value = 1 Then
    REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,RDY,smn where (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and  act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.ACT_IDY LIKE RDY.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND DBT_TYP='Y' and act.cst_typ like 'Van' GROUP BY ACT.ACT_IDY", con, adOpenKeyset, adLockPessimistic
'Else
'    REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,RDY,smn where (act.ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.ACT_IDY LIKE RDY.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND DBT_TYP='Y' and act.cst_typ like 'Van' GROUP BY ACT.ACT_IDY", CON, adOpenKeyset, adlockpessimistic
'End If
lv.ListItems.CLEAR
lv.ColumnHeaders.CLEAR
lv.ColumnHeaders.Add , , "Retailer", 2000
lv.ColumnHeaders.Add , , "ID", 1000
lv.ColumnHeaders.Add , , "Address", 2000
lv.ColumnHeaders.Add , , "Salesman", 2000
lv.ColumnHeaders.Add , , "Market", 2000
lv.ColumnHeaders.Add , , "Route", 1000
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_nme") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("MAR_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("rou_idy") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    lv.Visible = True
    lv.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Vehicle Not Found, Define Vehicle !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        vanid.SetFocus
    End If
End If
End Sub


Private Sub CANCELRECORD()
vanid = ""
SSTab1.Tab = 0
DTPicker1.Enabled = True
DTPicker1.SetFocus
End Sub
