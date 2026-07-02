VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form M_FRM_ASD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Date Select"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   Icon            =   "M_FRM_ASD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   360
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   255
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16515075
         CurrentDate     =   37315
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   375
         Width           =   45
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "M_FRM_ASD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_val As Integer
Private Sub Command1_Click()
Dim d As Date
Command1.Enabled = False
Dim rec_agd As New ADODB.Recordset
Dim rec_opa As New ADODB.Recordset
Dim rec_inl As New ADODB.Recordset
Dim rec_vch As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
con.Execute "update dir set DIR_MSG = '" & Format(DTPicker1, "DD/MM/YYYY") & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
'CON.Execute "update dir set DIR_DAT = '" & Format(DTPicker1, "DD/MM/YYYY") & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
If Me.Tag = "AGD" Or Me.Tag = "AGD_CAS" Or Me.Tag = "AGD_VAL" Or Me.Tag = "AGD_MFG" Then con.Execute "update dir set dir_MSG='" & DTPicker1 & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    'rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_DAT") = Format(DTPicker1, "MM/DD/YYYY")
    rs.Update
End If

'If Me.Tag = "AGD_6WB" Then CON.Execute "update dir set DIR_DAT = '" & DTPicker1 & "' where dir_idy like 'HDR' and dir_idx like 'DAT'"
If M_FRM_ASD.Tag = "Balance Sheet" Then
    Me.MousePointer = vbHourglass
    d = DateAdd("D", 1, DTPicker1)
    TRAIL_BALANCE Format(DTPicker1, "MM/DD/YYYY")
    PAL
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Finance.RP1", "Balance Sheet", 0, "", 6
ElseIf M_FRM_ASD.Tag = "Ageing" Then
    d = DateAdd("D", 1, DTPicker1)
    con.Execute "delete from agd"
    If rec_opa.State = 1 Then rec_opa.Close
    rec_opa.Open "select * from opa", con, adOpenDynamic, adLockPessimistic
    m_val = 0
    m_val = rec_opa("opa_dbt") - rec_opa("opa_crt")
    If m_val < 0 Then
     Do While rec_opa.EOF = False
        If rec_opa.State = 1 Then rec_opa.Close
        rec_opa.Open "select * from opa", con, adOpenDynamic, adLockPessimistic
        If rec_agd.State = 1 Then rec_agd.Close
        rec_agd.Open "select * from agd", con, adOpenDynamic, adLockPessimistic
        rec_agd.AddNew
        rec_agd("inv_idy") = "Excess"
        rec_agd("inv_dat") = Format(d, "mm/dd/yyyy")
        rec_agd("due_dat") = Format(d, "mm/dd/yyyy")
        rec_agd("cst_idy") = Trim(rec_opa("act_idy"))
        rec_agd("net_amt") = m_val
        rec_agd("day_ovr") = (Format(d, "mm/dd/yyyy")) - (Format(d, "mm/dd/yyyy"))
        rec_agd("txn_typ") = "Exc"
     Loop
    End If
    If rec_inl.State = 1 Then rec_inl.Close
    rec_inl.Open "select * from inl where txn_typ in ('LSL','SAL')", con, adOpenDynamic, adLockPessimistic
    If rec_inl("inv_dat") = Format(d, "mm/dd/yyyy") Then
        Do While rec_inl.EOF = False
            If rs.State = 1 Then rs.Close
            rs.Open "select * from act", con, adOpenDynamic, adLockPessimistic
            If rec_inl.State = 1 Then rec_inl.Close
            rec_inl.Open "select * from inl where act_idy like '" & Trim(rs("act_idy")) & "'", con, adOpenDynamic, adLockPessimistic
            If rec_agd.State = 1 Then rec_agd.Close
            rec_agd.Open "select * from agd", con, adOpenDynamic, adLockPessimistic
            rec_agd.AddNew
            rec_agd("inv_idy") = Trim(rec_inl("inv_idy"))
            rec_agd("inv_dat") = Trim(rec_inl("inv_dat"))
            rec_agd("due_dat") = Trim(rec_inl("due_dat"))
            rec_agd("cst_idy") = Trim(rec_inl("act_idy"))
            rec_agd("net_amt") = Trim(rec_inl("net_amt"))
            rec_agd("day_ovr") = Trim(rec_inl("due_dat")) - Trim(rec_inl("inv_dat"))
            rec_agd("txn_typ") = Trim(rec_inl("txn_typ"))
        Loop
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "select * from vch where txn_typ in ('OPN','PMT','DBN')", con, adOpenDynamic, adLockPessimistic
    If rs("TXN_dat") <= Format(d, "mm/dd/yyyy") Then
        Do While rec_inl.EOF = False
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from act", con, adOpenDynamic, adLockPessimistic
            If rec_vch.State = 1 Then rec_vch.Close
            rec_vch.Open "select * from vch where crt_idy like '" & Trim(RS1("act_idy")) & "'", con, adOpenDynamic, adLockPessimistic
            If rec_agd.State = 1 Then rec_agd.Close
            rec_agd.Open "select * from agd", con, adOpenKeyset, adLockPessimistic
            rec_agd.AddNew
            rec_agd("inv_idy") = Trim(rec_vch("txn_idy"))
            rec_agd("inv_dat") = Trim(rec_vch("txn_dat"))
            rec_agd("due_dat") = Trim(rec_vch("TXN_dat"))
            rec_agd("cst_idy") = Trim(rec_vch("crt_idy"))
            rec_agd("net_amt") = Trim(rec_vch("txn_amt"))
            rec_agd("day_ovr") = DateDiff("D", Format(d, "mm/dd/yyyy"), rec_vch("txn_dat"))
            rec_agd("txn_typ") = Trim(rec_vch("txn_typ"))
        Loop
    End If
    If rec_agd.State = 1 Then rec_agd.Close
    rec_agd.Open "select * from agd", con, adOpenDynamic, adLockPessimistic
    rec_agd.MoveFirst
    If rec_opa.State = 1 Then rec_opa.Close
    rec_opa.Open "select * from opa", con, adOpenDynamic, adLockPessimistic
    If rec_agd.State = 1 Then rec_agd.Close
    rec_agd.Open "select * from agd where cst_idy like '" & Trim(rec_opa("act_idy")) & "'", con, adOpenDynamic, adLockPessimistic
    m_val = 0
    m_val = Val(Trim(rec_agd("net_amt"))) - Val(Trim(rec_opa("opa_crt")))
    If Val(Trim(rec_agd("net_amt"))) > 0 Then
        If (m_val >= 0) Then
            con.Execute "update agd set net_amt = '" & Trim(m_val) & "'"
            con.Execute "update opa set opa_crt = 0"
        Else
            con.Execute "update agd set net_amt = 0"
            con.Execute "update opa set opa_crt = '" & Abs(m_val) & "'"
        End If
    End If
    con.Execute "delete from agd where net_amt = 0 or inv_dat > DTPicker1"
    SAL_QRY = "agd->inv_dat='" & Format(d, "mm/dd/yyyy") & "'"
    'Openreport CURDIR & "\Aging.RP1", "Ageing Detail (80 column)", 2, "", 6
'end of ageing code
ElseIf M_FRM_ASD.Tag = "Group Balance Sheet" Then
    Me.MousePointer = vbHourglass
    d = DateAdd("D", 1, DTPicker1)
    TRAIL_BALANCE Format(DTPicker1, "MM/DD/YYYY")
    PAL
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Finance.RP1", "Group Balance Sheet", 0, "", 6
ElseIf M_FRM_ASD.Tag = "ROI" Then
    Me.MousePointer = vbHourglass
    d = DateAdd("D", 1, DTPicker1)
    TRAIL_BALANCE Format(d, "MM/DD/YYYY")
    PAL
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Finance.RP1", "ROI Report", 2, "ACT->GRP_CLF='Asset'.or.ACT->GRP_CLF='Capital Account'", 6
ElseIf M_FRM_ASD.Tag = "Trial Balance" Then
    Me.MousePointer = vbHourglass
    d = DateAdd("D", 1, DTPicker1)
    TRAIL_BALANCE Format(DTPicker1, "MM/DD/YYYY")
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Finance.RP1", "Trial Balance", 0, Trim(SAL_QRY), 6
ElseIf M_FRM_ASD.Tag = "Group Trial Balance" Then
    Me.MousePointer = vbHourglass
    d = DateAdd("D", 1, DTPicker1)
    TRAIL_BALANCE Format(DTPicker1, "MM/DD/YYYY")
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Finance.RP1", "Group Trial Balance", 0, Trim(SAL_QRY), 6
ElseIf M_FRM_ASD.Tag = "Income Statement" Then
    Me.MousePointer = vbHourglass
    d = DateAdd("D", 1, DTPicker1)
    TRAIL_BALANCE Format(DTPicker1, "MM/DD/YYYY")
    PAL
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Finance.RP1", "Income Statement", 0, Trim(SAL_QRY), 6
ElseIf M_FRM_ASD.Tag = "CRL" Then
    C = Format(DTPicker1, "DD/MM/YYYY")
    d = DateAdd("D", 1, DTPicker1)
    TRAIL_BALANCE Format(d, "MM/DD/YYYY")
    con.Execute "update dir set dir_MSG='" & C & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
    SAL_QRY = "OPA->DB_CR='Cr'.AND.(ACT->GRP_CLF='Creditor'.OR.ACT->GRP_CLF='Debtor')"
    OpenReport CURDIR & "\Finance.RP1", "Creditors List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "OPN" Then
    SAL_QRY = "VCH->TXN_TYP='OPN'.AND.VCH->TXN_DAT<=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\Finance.RP1", "Opening Balance Vouchers List", 2, Trim(SAL_QRY), 6
ElseIf Me.Caption = "Invoice Date" Then
    INVOICE_DATE = DTPicker1
    Unload Me
ElseIf Me.Tag = "PCS" Then
    M_FRM_PCS_PCS.Tag = "PCS"
    INVOICE_DATE = DTPicker1
ElseIf Me.Tag = "UNS" Then
    M_FRM_PCS.Tag = "UNS"
    INVOICE_DATE = DTPicker1
ElseIf Me.Tag = "Van" Then
    M_FRM_RSI.Tag = "Van"
    INVOICE_DATE = DTPicker1
ElseIf Me.Tag = "DAS_CRS" Then
    Label1.Caption = "Start From "
    Label1.Refresh
    DAS_Cola_Extract DTPicker1, Date
    
ElseIf Me.Tag = "MyPage_SQA" Then
    OpenReport CURDIR & "\MYP.RP1", "Sales Quantity Analysis", 2, "VAN->INV_DAT = CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "')", 6
ElseIf Me.Tag = "MyPage_SVA" Then
    OpenReport CURDIR & "\MYP.RP1", "Sales Value Analysis", 2, "VAN->INV_DAT = CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "')", 6
ElseIf Me.Tag = "MyPage_FUA" Then
    OpenReport CURDIR & "\MYP.RP1", "Funds Analysis", 2, "VAN->INV_DAT = CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "')", 6
ElseIf Me.Tag = "MyPage_VUT" Then
    OpenReport CURDIR & "\MYP.RP1", "Vehicle Utilization", 2, "VAN->INV_DAT = CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "')", 6
ElseIf Me.Tag = "MyPage_VUV" Then
    OpenReport CURDIR & "\MYP.RP1", "Vehicle Utilization - Vehicle", 2, "VAN->INV_DAT = CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "')", 6
ElseIf Me.Tag = "MyPage_STS" Then
    StatusSheet DTPicker1
    OpenReport CURDIR & "\MYP.RP1", "My Page Status Sheet", 0, "", 6
ElseIf M_FRM_ASD.Tag = "SCB_SMY" Then
    OpenReport CURDIR & "\REP_LIB.RP1", "Van Cumulative Smry", 2, "INV_DAT <= CTOD('" & Format(DTPicker1, "mm/DD/yyyy") & "')", 6
ElseIf M_FRM_ASD.Tag = "CLG_CRT" Then
    C = Format(DTPicker1, "DD/MM/YYYY")
    d = DateAdd("D", 0, DTPicker1)
    TRAIL_BALANCECLG Format(d, "MM/DD/YYYY")
    con.Execute "update dir set dir_MSG='" & C & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
    SAL_QRY = "OPA->DB_CR='Cr'.AND.(ACT->GRP_CLF='Creditor'.OR.ACT->GRP_CLF='Debtor')"
    OpenReport CURDIR & "\Finance.RP1", "Creditors List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "AGD" Or Me.Tag = "AGD_CAS" Or Me.Tag = "AGD_VAL" Or Me.Tag = "AGD_MFG" Or Me.Tag = "AGD_6WB" Then
    TO_DAT = DTPicker1
    M_FRM_UPD.Tag = Me.Tag
    Unload Me
    M_FRM_UPD.Show 1
ElseIf Me.Tag = "invl" Then
    TO_DAT = DTPicker1.Value
    M_FRM_SAL_PSH.Tag = "invl"
    Unload Me
End If
Unload Me
Set Form = Nothing
End Sub
Private Sub Command2_Click()
Unload Me
Set Form = Nothing
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
SETYEAROPEN M_FRM_ASD
If DTPicker1.Enabled = True Then DTPicker1.SetFocus
If Me.Tag = "DAS_CRS" Then
    Me.Caption = "Purchase Invoice Date"
    Me.Refresh
    Label1.Caption = "Start Date"
    Label1.Refresh
    Dim INL_JMD As New ADODB.Recordset
    If INL_JMD.State = 1 Then INL_JMD.Close
    INL_JMD.Open "SELECT * FROM INL WHERE TXN_TYP = 'PUR'", con, adOpenKeyset, adLockPessimistic
    If INL_JMD.RecordCount > 0 Then
        INL_JMD.MoveLast
        DTPicker1 = Trim(INL_JMD("INV_DAT"))
    End If
    INL_JMD.Close
Else
    Label1.Caption = "As On Date"
    Label1.Refresh
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub StatusSheet(TODAY As Date)

Me.MousePointer = vbHourglass
d = DateAdd("D", 1, TODAY)
TRAIL_BALANCE Format(TODAY, "MM/DD/YYYY")
Me.MousePointer = vbDefault

Me.MousePointer = vbHourglass
con.Execute "DELETE FROM STS"
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select SUM(NET_AMT),TXN_TYP from inl where INV_DAT >= CTOD('" & Format(TODAY, "MM/DD/YYYY") & "') AND inv_dat <= ctod('" & Format(TODAY, "MM/DD/YYYY") & "') group by txn_typ", con, adOpenKeyset, adLockPessimistic
RS1.Open "SELECT * FROM STS", con, adOpenKeyset, adLockPessimistic
RS1.AddNew
RS1(0) = 0
    RS1.Update
While rs.EOF = False
    If Trim(rs(1)) = "SAL" Then
'            RS1.Update "crd_sal", Val(rs(0))
        RS1.Update "tot_sal", Val(RS1("tot_sal") & "") + Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "LSL" Then
        RS1.Update "cas_sal", Val(rs(0))
        RS1.Update "tot_sal", Val(RS1("tot_sal") & "") + Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "PUR" Then
        RS1.Update "tot_pur", Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "PRT" Then
        RS1.Update "tot_prt", Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "SRN" Then
        RS1.Update "tot_srn", Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "STI" Then
        RS1.Update "tot_sti", Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "STO" Then
        RS1.Update "tot_sto", Val(rs(0) & "")
    End If
    rs.MoveNext
Wend
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select sum(net_amt),txn_typ from inl where inv_dat = ctod('" & TODAY & "') group by txn_typ", con, adOpenKeyset, adLockPessimistic
RS1.Open "select * from sts", con, adOpenKeyset, adLockPessimistic
If rs.EOF = False Then rs.MoveFirst
While rs.EOF = False
    If Trim(rs(1)) = "SAL" Then
'            RS1.Update "D_crd_sal", Val(rs(0))
        RS1.Update "day_sal", Val(RS1("day_sal") & "") + Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "LSL" Then
'        RS1.Update "D_cas_sal", Val(rs(0))
        RS1.Update "day_sal", Val(RS1("day_sal") & "") + Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "PUR" Then
        RS1.Update "day_pur", Val(rs(0))
    ElseIf Trim(rs(1)) = "PRT" Then
        RS1.Update "day_prt", Val(rs(0))
    ElseIf Trim(rs(1)) = "SRN" Then
        RS1.Update "day_srn", Val(rs(0))
    ElseIf Trim(rs(1)) = "STI" Then
        RS1.Update "day_sti", Val(rs(0))
    ElseIf Trim(rs(1)) = "STO" Then
        RS1.Update "day_sto", Val(rs(0))
    End If
    rs.MoveNext
Wend

If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select sum(prd_qty),txn_typ from inv where inv_dat = ctod('" & TODAY & "') group by txn_typ", con, adOpenKeyset, adLockPessimistic
RS1.Open "select * from sts", con, adOpenKeyset, adLockPessimistic
If rs.EOF = False Then rs.MoveFirst
While rs.EOF = False
    If Trim(rs(1)) = "SAL" Then
        RS1.Update "CAS_SAL", Val(RS1("CAS_SAL") & "") + Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "LSL" Then
'        RS1.Update "D_cas_sal", Val(rs(0))
        RS1.Update "CAS_SAL", Val(RS1("CAS_SAL") & "") + Val(rs(0) & "")
    ElseIf Trim(rs(1)) = "PUR" Then
        RS1.Update "CAS_PUR", Val(rs(0))
    ElseIf Trim(rs(1)) = "PRT" Then
        RS1.Update "CAS_PRT", Val(rs(0))
    ElseIf Trim(rs(1)) = "SRN" Then
        RS1.Update "CAS_SRN", Val(rs(0))
    ElseIf Trim(rs(1)) = "STI" Then
        RS1.Update "CAS_STI", Val(rs(0))
    ElseIf Trim(rs(1)) = "STO" Then
        RS1.Update "CAS_STO", Val(rs(0))
    End If
    rs.MoveNext
Wend

If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
While rs.EOF = False
    STK_VAL = Val(STK_VAL) + Val(RATCAL(QTY(rs("cur_lev") & ""), Val(rs("PUR_RAT") & ""), Val(rs("SUB_UNT") & "")))
    STK_QTY = Val(STK_QTY) + Val(QTY(rs("CUR_LEV")))
    rs.MoveNext
Wend
RS1.Update "STK_VAL", Val(STK_VAL)
RS1.Update "CAS_STK", Val(STK_QTY)
NET_BAL = 0
CRT_TTL = 0
DBT_TTL = 0
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM OPA", con, adOpenKeyset, adLockPessimistic
While rs.EOF = False
    If GetValue("GRP_CLF", "ACT WHERE ACT_IDY like '" & Trim(rs("ACT_IDY")) & "'") = "Creditor" Then
        NET_BAL = (Val(rs(1)) - Val(rs(2)))
        CRT_TTL = Val(CRT_TTL) + NET_BAL
    End If
    rs.MoveNext
Wend
RS1.Update "CRD_AMT", Val(CRT_TTL)
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM OPA", con, adOpenKeyset, adLockPessimistic
While rs.EOF = False
    If GetValue("GRP_CLF", "ACT WHERE ACT_IDY like '" & Trim(rs("ACT_IDY")) & "'") = "Debtor" Then
        DBT_TTL = Val(DBT_TTL) + (Val(rs(1)) - Val(rs(2)))
    End If
    rs.MoveNext
Wend
RS1.Update "DBT_AMT", Val(DBT_TTL)
RS1.Update "NET_AMT", Val(CRT_TTL) - Val(DBT_TTL)

RS1.Close
rs.Close
Me.MousePointer = vbDefault
End Sub
Private Sub DAS_Cola_Extract(B_DATE As Date, E_DATE As Date)
' * Distributor ID Identification
        Dim DIR_RS As New ADODB.Recordset
        Dim M_WDS_IDY As String
        Dim M_WDS_SRL As String
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                M_WDS_IDY = Trim(DIR_RS("DIR_ALS"))
            End If
            If Len(M_WDS_IDY) < 8 Then
                MsgBox ("Distributor ID Required for this Process! Obtain Distributor ID and Start Again"), vbCritical, "DAS Version 4.0"
'                cancel
                Exit Sub
            End If
            M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
            If Len(Trim(M_WDS_SRL)) < 3 Then
                 While Len(Trim(M_WDS_SRL)) < 3
                      M_WDS_SRL = "0" & M_WDS_SRL
                 Wend
            End If
        End If
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'SYS' and DIR_IDX = 'VER'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                DIR_RS("DIR_MSG") = Trim("01.01.01")
                DIR_RS.Update
            End If
        End If
' * Distributor ID Identification

If MsgBox("Create Data File for Data Transfer From " & Format(B_DATE, "DD/MM/YYYY"), vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
        
    ' * : Beginning of DASCRS_PURCHASE.DAT Creation Code for DAS_CRS Data Transfer
    
    Dim F_INL
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_INL = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    If FS_INL.FileExists(CURDIR & "\HHT\DASCRS_PURCHASE.DAT") = True Then
        FS_INL.DeleteFile (CURDIR & "\HHT\DASCRS_PURCHASE.DAT")
    End If
    Set F_INL = FS_INL.OpenTextFile(CURDIR & "\HHT\DASCRS_PURCHASE.DAT", 2, True)
    
    Dim K As Double
    Dim r As Double
    Dim s As Double
    Dim T As Double
    K = 0
    r = 0
    s = 0
    T = 0
    Dim INL_RS As New ADODB.Recordset
    Dim INV_RS As New ADODB.Recordset
    Dim LNV_RS As New ADODB.Recordset
    Dim PRD_RS As New ADODB.Recordset
    Dim BAT_RS As New ADODB.Recordset
    Dim SCH_RS As New ADODB.Recordset
    Dim VAN_RS As New ADODB.Recordset
    Dim SNV_RS As New ADODB.Recordset
    Dim ACT_RS As New ADODB.Recordset
    
    Dim L_CST_IDY As String
    Dim L_ORD_IDY As String
    Dim L_BAT_NME As String
    Dim L_PRD_COD As String
    Dim L_MER_NUM As String
    
    If INL_RS.State = 1 Then INL_RS.Close
    If INV_RS.State = 1 Then INV_RS.Close
    If LNV_RS.State = 1 Then LNV_RS.Close
    If PRD_RS.State = 1 Then PRD_RS.Close
    If BAT_RS.State = 1 Then BAT_RS.Close
    If SCH_RS.State = 1 Then SCH_RS.Close
       
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & B_DATE & "') .AND. INV_DAT <= CTOD('" & E_DATE & "')", con, adOpenKeyset, adLockPessimistic
    PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    SCH_RS.Open "SELECT * FROM SCH", con, adOpenKeyset, adLockPessimistic
    
    If INL_RS.State = 1 Then INL_RS.Close
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & B_DATE & "') .AND. INV_DAT <= CTOD('" & E_DATE & "')", con, adOpenKeyset, adLockPessimistic
    If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
    While Not INL_RS.EOF
        M_ADD_AMT = 0
        M_DIS_AMT = 0
        M_TAX_AMT = 0
        T_SKU_QTY = 0
        If INV_RS.State = 1 Then INV_RS.Close
        INV_RS.Open "SELECT * FROM INV WHERE INV.INV_IDY = '" & Trim(INL_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
        If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
        While Not INV_RS.EOF
            If IsNull(INV_RS("ADD_AMT")) = False Then
                M_ADD_AMT = M_ADD_AMT + INV_RS("ADD_AMT")
            End If
            If IsNull(INV_RS("DIS_AMT")) = False Then
                M_DIS_AMT = M_DIS_AMT + INV_RS("DIS_AMT")
            End If
            If IsNull(INV_RS("TAX_AMT")) = False Then
                M_TAX_AMT = M_TAX_AMT + INV_RS("TAX_AMT")
            End If
            If IsNull(INV_RS("TOT_PCS")) = False Then
                If Val(INV_RS("SUB_UNT")) > 0 Then
                    T_SKU_QTY = Val(T_SKU_QTY) + Round(Val(INV_RS("TOT_PCS")) / Val(INV_RS("SUB_UNT")), 2)
                End If
            End If
            INV_RS.MoveNext
        Wend
        If Trim(INL_RS("TXN_TYP")) <> "OPS" Then
            If INV_RS.RecordCount > 0 Then
                If Trim(INL_RS("TXN_TYP")) = "PUR" Then
                    F_INL.WriteLine Format(INL_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(M_WDS_IDY) & "|" & Trim(INL_RS("INV_IDY")) & "|" & Trim(INL_RS("ORD_IDY")) & "|" & Format(INL_RS("DCH_DAT"), "YYYYMMDD") & "|" & Val(T_SKU_QTY) & "|" & Val(INL_RS("NET_AMT"))
                    r = r + 1
                End If
            End If
        End If
        INL_RS.MoveNext
    Wend
    F_INL.Close
    
'    cancel

    ' * : Beginning of DASCRS_STOCK.DAT Creation Code for Data Transfer

    Dim F_STK
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DASCRS_STOCK.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DASCRS_STOCK.DAT")
    End If
    
    Set F_STK = FS_SOU.OpenTextFile(CURDIR & "\HHT\DASCRS_STOCK.DAT", 2, True)
    Dim l As Integer
    l = 0
            
            If PRD_RS.State = 1 Then PRD_RS.Close
            PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
            While Not PRD_RS.EOF
                If PRD_RS("SUB_UNT") = 0 Then
                    PRD_RS("SUB_UNT") = 1
                    PRD_RS.Update
                End If
                PRD_RS.MoveNext
            Wend
            If PRD_RS.State = 1 Then PRD_RS.Close
'            PRD_RS.Open "SELECT ROUND(sum(PRD_PCS)/(SUB_UNT),2), PRD_IDY, ALT_IDY, PRD_PRN FROM PRD GROUP BY PRD_PRN", CON, adOpenKeyset, adLockPessimistic
            PRD_RS.Open "SELECT sum(PRD_PCS), PRD_IDY, ALT_IDY, PRD_PRN FROM PRD GROUP BY ALT_IDY WHERE CUR_LEV > 0", con, adOpenKeyset, adLockPessimistic
            If PRD_RS.RecordCount > 0 Then
                While Not PRD_RS.EOF
                    F_STK.WriteLine Format(E_DATE, "YYYYMMDD") & "|" & Format(E_DATE, "YYYYMMDD") & "|" & Trim(PRD_RS!ALT_IDY) & "|" & Trim(M_WDS_IDY) & "|" & Trim(str(Val(PRD_RS(0))))
                    PRD_RS.MoveNext
                    l = l + 1
                Wend
            End If
    F_STK.Close
    
    ' * : End of DASCRS_STOCK.DAT Creation Code for Data Transfer
    
    ' * : Log Text File Creation Begin Process
    
    Dim F_VER
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FileExists(CURDIR & "\HHT\DASCRS_VERSION.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DASCRS_VERSION.DAT")
    End If
    
    Dim Q As Integer
    Dim LDC_DAT As Date
    Q = 0
    Set F_VER = FS_SOU.OpenTextFile(CURDIR & "\HHT\DASCRS_VERSION.DAT", 2, True)
    
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) like 'DAY' AND alltrim(DIR_MSG) like 'Y'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveLast
        LDC_DAT = DIR_RS("DIR_DAT")
    Else
        LDC_DAT = Date
    End If
    M_EXE_VER = Format(FileDateTime(App.Path & "\EFFMCG.EXE"), "dd/mmm/yyyy hh:mm")
    F_VER.WriteLine Trim(M_WDS_IDY) & "|" & Trim("01.01.01") & "|" & Format(E_DATE, "dd/MM/yyyy") & "|" & Format(E_DATE, "dd/MM/yyyy") & "|" & Format(LDC_DAT, "dd/MM/yyyy") & "|" & Trim(M_EXE_VER)
    Q = Q + 1
    
    F_VER.Close
    
    Dim F_LOG
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DASCRS_LOG.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DASCRS_LOG.DAT")
    End If
    
    Set F_LOG = FS_SOU.OpenTextFile(CURDIR & "\HHT\DASCRS_LOG.DAT", 2, True)
    
    F_LOG.WriteLine "DASCRS_PURCHASE.DAT" & "|" & str(r)
    F_LOG.WriteLine "DASCRS_STOCK.DAT" & "|" & str(l)
    F_LOG.WriteLine "DASCRS_VERSION.DAT" & "|" & str(Q)
    
    Dim MDD As String
    Dim MMM As String
    Dim MYY As String
    MDD = Trim(str(Day(E_DATE)))
    MMM = Trim(str(Month(E_DATE)))
    MYY = Trim(str(Year(E_DATE)))
    F_LOG.Close
    ' * : Log Text File Creation End Process
    Me.MousePointer = vbDefault
    MsgBox ("File Creation Process Successful! Log File Generated!"), vbInformation, "DAS Version 4.0"
    
' Beginning of Distributor Information for MAPI Data Xfer
    
    Dim M_NME, M_ADR, M_TFX, M_WDS As String
    Dim RST_DIR As New ADODB.Recordset
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT1'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_NME = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT2'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_ADR = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT3'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_TFX = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'WDS' .AND. DIR_IDX LIKE 'IDY'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_WDS = Trim(RST_DIR("DIR_ALS"))
    End If

    If Dir(CURDIR & "\DST_INF.DAT") = "" Then
        Set F_VER = FS_SOU.OpenTextFile(CURDIR & "\DST_INF.DAT", 2, True)
        F_VER.WriteLine "Distributor : " & M_NME
        F_VER.WriteLine "Address     : " & M_ADR
        F_VER.WriteLine "Telephone   : " & M_TFX
        F_VER.WriteLine "DAS Code-ID : " & M_WDS
        F_VER.Close
    End If

' End of Distributor Information for MAPI Data Xfer

    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then DIR_RS.MoveFirst
       DIR_RS("DIR_MSG") = Trim(M_WDS_SRL)
       DIR_RS.Update
    If DIR_RS.State = 1 Then DIR_RS.Close
    'Zip File Creation
            Dim DEX As String
            Dim CDO As String
            DEX = "C:\DEX"
            CDO = "C:\DEX\OutBox"
            If Dir(DEX, vbDirectory) = "" Then
                MkDir DEX
            End If
            If Dir(CDO, vbDirectory) = "" Then
                MkDir CDO
            End If
            
'            CopyFile CURDIR & "\HHT" & "\DASCRS_PURCHASE.DAT", CDO & "\DASCRS_PURCHASE.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DASCRS_STOCK.DAT", CDO & "\DASCRS_STOCK.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DASCRS_VERSION.DAT", CDO & "\DASCRS_VERSION.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DASCRS_LOG.DAT", CDO & "\DASCRS_LOG.DAT", 0
'            CopyFile CURDIR & "\DST_INF.DAT", CDO & "\DST_INF.DAT", 0

            CopyFile CURDIR & "\HHT" & "\DASCRS_PURCHASE.DAT", DEX & "\DASCRS_PURCHASE.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DASCRS_STOCK.DAT", DEX & "\DASCRS_STOCK.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DASCRS_VERSION.DAT", DEX & "\DASCRS_VERSION.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DASCRS_LOG.DAT", DEX & "\DASCRS_LOG.DAT", 0
            CopyFile CURDIR & "\DST_INF.DAT", DEX & "\DST_INF.DAT", 0
                        
                        
            
            CopyFile App.Path & "\" & "Zip32.dll", CURDIR & "\" & "Zip32.dll", 0
            M_EML_IDY = GetValue("DIR_MSG", "DIR where DIR_IDY like 'EML' AND DIR_IDX LIKE 'IDY'")
            
            M_FIL_NME = "\" & "DASCRS" & Trim(M_WDS_IDY) & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".ZIP"
            Dim Myzip As ColaZip.GSColaZip32
            Set Myzip = New ColaZip.GSColaZip32
            Myzip.RootDirectory = "\"
            Myzip.ZipFileName = "C:\DEX" & M_FIL_NME
            Myzip.AddFile ("C:\DEX" & "\DASCRS_PURCHASE.DAT")
            Myzip.AddFile ("C:\DEX" & "\DASCRS_STOCK.DAT")
            Myzip.AddFile ("C:\DEX" & "\DASCRS_VERSION.DAT")
            Myzip.AddFile ("C:\DEX" & "\DASCRS_LOG.DAT")
            Myzip.AddFile ("C:\DEX" & "\DST_INF.DAT")
            
            Myzip.MakeColaZip ("DASCRS")
            
            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
                      
            M_MSG_STR = "DAS CRS Extract Zipped Data From : " & Trim(M_NME) & " , " & Trim(M_ADR) & " , " & Trim(M_TFX) & " , " & Trim(M_WDS) & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
            M_MSG_EML = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'")
            M_MSG_SUB = "Zip Data Xfer DAS CRS Data Extract : " & M_WDS & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
            
            M_MAIL_ATTMNT = "C:\DEX" & M_FIL_NME
            
            ' Disabled Email Send of DAS CRS Extract as per Prince Chopras Request
'            If MsgBox("Send Zip File Data to DAS CRS Desk Now ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
'                If Trim(M_EML_IDY) = "" Then
'                    MailSend "effmcg@consoftsystems.com", Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
'                    MsgBox ("Data Send via MAPI Executed! Email ID: effmcg@consoftsystems.com" & Trim(M_EML_IDY)), vbInformation, "DAS Version 4.0"
'                Else
'                    MailSend Trim(M_EML_IDY), Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
'                    MsgBox ("Data Send via MAPI Executed! Email ID: " & Trim(M_EML_IDY)), vbInformation, "DAS Version 4.0"
'                End If
'            Else
'                'cancel
'            End If
            'File gets copied into DEX OutBox Folder
            
            CopyFile DEX & M_FIL_NME, CDO & M_FIL_NME, 0
    ' Zip File Creation
            
    'cancel
Else
    DTPicker1.SetFocus
End If
End Sub
Function MailSend(sSendTo As String, sSubject As String, sText As String, Optional sAttachFile As String) As Boolean
'Dim MAPISession1 As MAPISession
'Dim MAPIMessages1 As MAPIMessages
    On Error GoTo ErrHandler
    With MAPISession1
        If .SessionID = 0 Then
            .DownLoadMail = False
            .LogonUI = True
            .SignOn
            .NewSession = True
            MAPIMessages1.SessionID = .SessionID
        End If
    End With
    With MAPIMessages1
        .Compose
        .RecipAddress = sSendTo
        .AddressResolveUI = True
        .ResolveName
        .MsgSubject = sSubject
        .MsgNoteText = sText
        If Len(sAttachFile) > 0 And Len(Dir$(sAttachFile)) > 0 Then
            .AttachmentPathName = sAttachFile
        Else
            .AttachmentCount = 0
        End If
        .Send False
    End With
    MailSend = True
    Exit Function
ErrHandler:
    Debug.Print err.Description
    MailSend = False
End Function

