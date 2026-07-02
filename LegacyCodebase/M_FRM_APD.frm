VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_APD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   ControlBox      =   0   'False
   Icon            =   "M_FRM_APD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&Details  >>"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   900
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&No"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   -15
      TabIndex        =   1
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   1215
      TabIndex        =   8
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processed"
      Height          =   195
      Left            =   825
      TabIndex        =   7
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   855
      Left            =   360
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Update Accounts...?"
      Height          =   195
      Left            =   1185
      TabIndex        =   0
      Top             =   0
      Width           =   1485
   End
End
Attribute VB_Name = "M_FRM_APD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CTR As Double
Dim COM_CTR As Double
Private Sub Command1_Click()
If Me.Tag = "FA" Then
    FA_UPDATE
ElseIf Me.Tag = "TB" Then
    TRAILBALANCE TO_DAT
ElseIf Me.Tag = "TB_CLG" Then
    TRAILBALANCECLG TO_DAT
ElseIf Me.Tag = "FA_CLG" Then
    FA_UPDATECLG
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub inc_pgb()
On Error Resume Next
DoEvents
Label4 = COM_CTR
Label4.Refresh
ProgressBar1.value = (100 / CTR) * COM_CTR
End Sub

Public Sub FA_UPDATE()
ProgressBar1.value = 0
CTR = 0
COM_CTR = 0
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
CON.Execute "delete from led"
TRAILBALANCE DateAdd("D", -1, FRM_DAT)
Label1.Caption = "Generating Ledger"
M_FRM_APD.Refresh
rs.Open "select * from vch WHERE txn_dat >=ctod('" & FRM_DAT & "') AND TXN_DAT <= CTOD('" & TO_DATE & "')", CON, adOpenKeyset, adLockPessimistic
rs2.Open "select * from INL WHERE INV_dat >=ctod('" & FRM_DAT & "') AND INV_DAT <= CTOD('" & TO_DATE & "')", CON, adOpenKeyset, adLockPessimistic
RS1.Open "select * from led", CON, adOpenKeyset, adLockPessimistic
CTR = rs.RecordCount + rs2.RecordCount
Label5 = CTR
For A = 1 To rs.RecordCount
    RS1.AddNew
    RS1(0) = Trim(rs("txn_idy") & "")
    RS1(1) = rs("txn_dat")
    RS1(2) = rs("crt_idy")
    RS1(4) = rs("txn_amt")
    RS1(3) = 0
    RS1(5) = rs("txn_typ")
    RS1(6) = Trim(rs("txn_rmk") & "")
    RS1(7) = Trim(rs("dbt_idy") & "")
    RS1.Update
    
    RS1.AddNew
    RS1(0) = Trim(rs("txn_idy") & "")
    RS1(1) = rs("txn_dat")
    RS1(2) = rs("dbt_idy")
    RS1(4) = 0
    RS1(3) = Val(Trim(rs("txn_amt") & ""))
    RS1(5) = rs("txn_typ")
    RS1(6) = Trim(rs("txn_rmk") & "")
    RS1(7) = Trim(rs("crt_idy") & "")
    RS1.Update
    rs.MoveNext
    inc_pgb
    COM_CTR = COM_CTR + 1
Next

For A = 1 To rs2.RecordCount
    If Trim(rs2("TXN_TYP")) = "SAL" Or Trim(rs2("TXN_TYP")) = "STO" Or Trim(rs2("TXN_TYP")) = "RSI" Or Trim(rs2("TXN_TYP")) = "PRT" Then
        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("act_idy"))
        RS1(3) = 0
        RS1(4) = Val(Trim(rs2("net_amt") & ""))
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("cst_idy"))
        RS1.Update
        
        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("cst_idy"))
        RS1(3) = rs2("net_amt")
        RS1(4) = 0
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("act_idy"))
        RS1.Update
    Else
        
        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("cst_idy"))
        RS1(3) = 0
        RS1(4) = Val(Trim(rs2("net_amt") & ""))
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("act_idy"))
        RS1.Update

        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("act_idy"))
        RS1(3) = rs2("net_amt")
        RS1(4) = 0
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("cst_idy"))
        RS1.Update
    End If
    rs2.MoveNext
    inc_pgb
    COM_CTR = COM_CTR + 1
Next
End Sub
Public Sub FA_UPDATECLG()
ProgressBar1.value = 0
CTR = 0
COM_CTR = 0
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
CON.Execute "delete from led"
TRAILBALANCECLG DateAdd("D", -1, FRM_DAT)
Label1.Caption = "Generating Ledger"
M_FRM_APD.Refresh
rs.Open "select * from vch WHERE CLG_dat >=ctod('" & FRM_DAT & "') AND CLG_DAT <= CTOD('" & TO_DATE & "')", CON, adOpenKeyset, adLockPessimistic
rs2.Open "select * from INL WHERE INV_dat >=ctod('" & FRM_DAT & "') AND INV_DAT <= CTOD('" & TO_DATE & "')", CON, adOpenKeyset, adLockPessimistic
RS1.Open "select * from led", CON, adOpenKeyset, adLockPessimistic
CTR = rs.RecordCount + rs2.RecordCount
Label5 = CTR
For A = 1 To rs.RecordCount
    RS1.AddNew
    RS1(0) = Trim(rs("txn_idy") & "")
    RS1(1) = rs("txn_dat")
    RS1(2) = rs("crt_idy")
    RS1(4) = rs("txn_amt")
    RS1(3) = 0
    RS1(5) = rs("txn_typ")
    RS1(6) = Trim(rs("txn_rmk") & "")
    RS1(7) = Trim(rs("dbt_idy") & "")
    RS1.Update
    
    RS1.AddNew
    RS1(0) = Trim(rs("txn_idy") & "")
    RS1(1) = rs("txn_dat")
    RS1(2) = rs("dbt_idy")
    RS1(4) = 0
    RS1(3) = Val(Trim(rs("txn_amt") & ""))
    RS1(5) = rs("txn_typ")
    RS1(6) = Trim(rs("txn_rmk") & "")
    RS1(7) = Trim(rs("crt_idy") & "")
    RS1.Update
    rs.MoveNext
    inc_pgb
    COM_CTR = COM_CTR + 1
Next

For A = 1 To rs2.RecordCount
    If Trim(rs2("TXN_TYP")) = "SAL" Or Trim(rs2("TXN_TYP")) = "STO" Or Trim(rs2("TXN_TYP")) = "RSI" Or Trim(rs2("TXN_TYP")) = "PRT" Then
        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("act_idy"))
        RS1(3) = 0
        RS1(4) = Val(Trim(rs2("net_amt") & ""))
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("cst_idy"))
        RS1.Update
        
        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("cst_idy"))
        RS1(3) = rs2("net_amt")
        RS1(4) = 0
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("act_idy"))
        RS1.Update
    Else
        
        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("cst_idy"))
        RS1(3) = 0
        RS1(4) = Val(Trim(rs2("net_amt") & ""))
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("act_idy"))
        RS1.Update

        RS1.AddNew
        RS1(0) = Trim(rs2("inv_idy") & "")
        RS1(1) = rs2("inv_dat")
        RS1(2) = Trim(rs2("act_idy"))
        RS1(3) = rs2("net_amt")
        RS1(4) = 0
        RS1(5) = Trim(rs2("txn_typ"))
        RS1(6) = Trim(rs2("txn_rmk") & "")
        RS1(7) = Trim(rs2("cst_idy"))
        RS1.Update
    End If
    rs2.MoveNext
    inc_pgb
    COM_CTR = COM_CTR + 1
Next
End Sub
Public Sub TRAILBALANCE(FRMDAT As Date)
On Error Resume Next
M_FRM_APD.MousePointer = vbHourglass
Command1.Enabled = False
Command2.Enabled = False
ProgressBar1.Visible = True
CTR = 0
COM_CTR = 0
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rs6 As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
CON.Execute "update opa set opa_dbt=0", d
CON.Execute "update opa set opa_crt=0", d
CON.Execute "DELETE FROM OPA", d
CON.Execute "UPDATE DIR SET DIR_MSG='" & Format(FRMDAT, "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX <> 'DAT'"

rs.Open "select sum(txn_amt),dbt_idy from vch where txn_dat <=ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') group by dbt_idy", CON, adOpenKeyset, adLockPessimistic
rs2.Open "select sum(txn_amt),crt_idy from vch where txn_dat <=ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') group by crt_idy", CON, adOpenKeyset, adLockPessimistic

rs3.Open "select sum(NET_amt),CST_idy from INL where INV_dat <=ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') AND  txn_typ in('LSL','SAL','PRT','RSI','STO') group by CST_idy", CON, adOpenKeyset, adLockPessimistic
rs4.Open "select sum(NET_amt),ACT_idy from INL where INV_dat <= ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') AND txn_typ in('LSL','SAL','PRT','RSI','STO') group by ACT_idy", CON, adOpenKeyset, adLockPessimistic

rs5.Open "select sum(net_amt),cst_idy from inl where inv_dat <= ctod('" & FRMDAT & "') and txn_typ in('PUR','SRN','RSR','STI','OPS') group by cst_idy", CON, adOpenKeyset, adLockPessimistic
rs6.Open "select sum(net_amt),act_idy from inl where inv_dat <= ctod('" & FRMDAT & "') and txn_typ in('PUR','SRN','RSR','STI','OPS') group by act_idy", CON, adOpenKeyset, adLockPessimistic

CTR = rs.RecordCount + rs2.RecordCount + rs3.RecordCount + rs4.RecordCount + rs5.RecordCount + rs6.RecordCount
Label5 = CTR
For A = 1 To rs.RecordCount
    CON.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs(0)) & " where act_idy like '" & Trim(rs!dbt_idy) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs!dbt_idy) & "'," & Val(rs(0)) & ",0,ctod('" & TO_DAT & "'))"
    rs.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next

For A = 1 To rs2.RecordCount
    CON.Execute "update opa set opa_crt = opa_crt+ " & Val(rs2(0)) & " where act_idy like '" & Trim(rs2!crt_idy) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs2!crt_idy) & "',0," & Val(rs2(0)) & ",ctod('" & TO_DAT & "'))"
    rs2.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next
For A = 1 To rs3.RecordCount
    CON.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs3(0)) & " where act_idy like '" & Trim(rs3!CST_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs3!CST_IDY) & "'," & Val(rs3(0)) & ",0,ctod('" & TO_DAT & "'))"
    rs3.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next
For A = 1 To rs4.RecordCount
    CON.Execute "update opa set opa_crt = opa_crt+ " & Val(rs4(0)) & " where act_idy like '" & Trim(rs4!ACT_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs4!ACT_IDY) & "',0," & Val(rs4(0)) & ",ctod('" & TO_DAT & "'))"
    COM_CTR = COM_CTR + 1
    inc_pgb
    rs4.MoveNext
Next
For A = 1 To rs6.RecordCount
    CON.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs6(0)) & " where act_idy like '" & Trim(rs6!ACT_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs6!ACT_IDY) & "'," & Val(rs6(0)) & ",0,ctod('" & TO_DAT & "'))"
    COM_CTR = COM_CTR + 1
    inc_pgb
    rs6.MoveNext
Next
For A = 1 To rs5.RecordCount
    CON.Execute "update opa set opa_crt = opa_crt+ " & Val(rs5(0)) & " where act_idy like '" & Trim(rs5!CST_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs5!CST_IDY) & "',0," & Val(rs5(0)) & ",ctod('" & TO_DAT & "'))"
    rs5.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next
ProgressBar1.value = 100
End Sub

Private Sub Command3_Click()
If Me.Height = 1725 Then
Command3.Caption = "&Details  <<"
Me.Height = 2790
Else
Me.Height = 1725
Command3.Caption = "&Details  >>"
End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If ProgressBar1.value > 1 Then
    Me.MousePointer = vbHourglass
End If
End Sub

Public Sub TRAILBALANCECLG(FRMDAT As Date)
M_FRM_APD.MousePointer = vbHourglass
Command1.Enabled = False
Command2.Enabled = False
ProgressBar1.Visible = True
CTR = 0
COM_CTR = 0
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rs6 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
CON.Execute "update opa set opa_dbt=0", d
CON.Execute "update opa set opa_crt=0", d
CON.Execute "DELETE FROM OPA", d
CON.Execute "UPDATE DIR SET DIR_MSG='" & Format(FRMDAT, "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'HDR'"

rs.Open "select sum(txn_amt),dbt_idy from vch where CLG_dat <=ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') group by dbt_idy", CON, adOpenKeyset, adLockPessimistic
rs2.Open "select sum(txn_amt),crt_idy from vch where CLG_dat <=ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') group by crt_idy", CON, adOpenKeyset, adLockPessimistic

rs3.Open "select sum(NET_amt),CST_idy from INL where INV_dat <=ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') AND  txn_typ in('LSL','SAL','PRT','RSI','STO') group by CST_idy", CON, adOpenKeyset, adLockPessimistic
rs4.Open "select sum(NET_amt),ACT_idy from INL where INV_dat <= ctod('" & Format(FRMDAT, "MM/DD/YYYY") & "') AND txn_typ in('LSL','SAL','PRT','RSI','STO') group by ACT_idy", CON, adOpenKeyset, adLockPessimistic

rs5.Open "select sum(net_amt),cst_idy from inl where inv_dat <= ctod('" & FRMDAT & "') and txn_typ in('PUR','SRN','RSR','STI','OPS') group by cst_idy", CON, adOpenKeyset, adLockPessimistic
rs6.Open "select sum(net_amt),act_idy from inl where inv_dat <= ctod('" & FRMDAT & "') and txn_typ in('PUR','SRN','RSR','STI','OPS') group by act_idy", CON, adOpenKeyset, adLockPessimistic

'RS1.Open "SELECT * FROM opa", CON, adOpenKeyset, adlockpessimistic
CTR = rs.RecordCount + rs2.RecordCount + rs3.RecordCount + rs4.RecordCount + rs5.RecordCount + rs6.RecordCount
Label5 = CTR
For A = 1 To rs.RecordCount
'    RS1.Find "act_idy like '" & Trim(rs("dbt_idy")) & "'"
'    If RS1.EOF = False Then
'        RS1("opa_dbt") = Val(RS1("opa_dbt")) + Val(rs(0) & "")
'        RS1.Update
'        RS1.MoveFirst
'    Else
    CON.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs(0)) & " where act_idy like '" & Trim(rs!dbt_idy) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs!dbt_idy) & "'," & Val(rs(0)) & ",0,ctod('" & TO_DAT & "'))"
'        RS1.Requery
'    End If
    rs.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next

For A = 1 To rs2.RecordCount
'    RS1.Find "act_idy like '" & Trim(rs2("CRT_idy")) & "'"
'    If RS1.EOF = False Then
'        RS1("opa_crt") = Val(RS1("opa_crt")) + Val(rs2(0) & "")
'        RS1.Update
'        RS1.MoveFirst
'    Else
'        RS1.AddNew
'            RS1("act_idy") = Trim(rs2("crt_idy") & "")
'            RS1("opa_crt") = Val(rs2(0) & "")
'            RS1("opa_dbt") = 0
'            RS1("fyr_dat") = TO_DAT
'        RS1.Update
'        RS1.Requery
'    End If
    CON.Execute "update opa set opa_crt = opa_crt+ " & Val(rs2(0)) & " where act_idy like '" & Trim(rs2!crt_idy) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs2!crt_idy) & "',0," & Val(rs2(0)) & ",ctod('" & TO_DAT & "'))"
    rs2.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next
'sales customer
For A = 1 To rs3.RecordCount
'    RS1.Find "act_idy like '" & Trim(rs3("cst_idy")) & "'"
'    If RS1.EOF = False Then
'        RS1("opa_dbt") = Val(RS1("opa_dbt")) + Val(rs3(0) & "")
'        RS1.Update
'        RS1.MoveFirst
'    Else
'        RS1.AddNew
'        RS1("act_idy") = Trim(rs3("cst_idy") & "")
'        RS1("opa_dbt") = Val(rs3(0) & "")
'        RS1("opa_crt") = 0
'        RS1("fyr_dat") = TO_DAT
'        RS1.Update
'        RS1.Requery
'    End If
    CON.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs3(0)) & " where act_idy like '" & Trim(rs3!CST_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs3!CST_IDY) & "'," & Val(rs3(0)) & ",0,ctod('" & TO_DAT & "'))"
    rs3.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next
'sales accounts
'if RS1.State = 1 Then RS1.Close
'RS1.Open "SELECT * FROM OPA", CON, adOpenKeyset, adlockpessimistic
For A = 1 To rs4.RecordCount
'    RS1.Find "act_idy like '" & Trim(RS4("act_idy")) & "'"
'    If RS1.EOF = False Then
'        RS1("opa_crt") = Val(RS1("opa_crt")) + Val(RS4(0) & "")
'        RS1.Update
'        RS1.MoveFirst
'    Else
'        RS1.AddNew
'            RS1("act_idy") = Trim(RS4("act_idy") & "")
'            RS1("opa_crt") = Val(RS4(0) & "")
'            RS1("opa_dbt") = 0
'            RS1("fyr_dat") = TO_DAT
'        RS1.Update
'        RS1.Requery
'    End If
    CON.Execute "update opa set opa_crt = opa_crt+ " & Val(rs4(0)) & " where act_idy like '" & Trim(rs4!ACT_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs4!ACT_IDY) & "',0," & Val(rs4(0)) & ",ctod('" & TO_DAT & "'))"
    COM_CTR = COM_CTR + 1
    inc_pgb
    rs4.MoveNext
Next
'purchases customer
'If RS1.State = 1 Then RS1.Close
'RS1.Open "SELECT * FROM OPA", CON, adOpenKeyset, adlockpessimistic
For A = 1 To rs6.RecordCount
'    RS1.Find "act_idy like '" & Trim(rs6("act_idy")) & "'"
'    If RS1.EOF = False Then
'        RS1("opa_dbt") = Val(RS1("opa_dbt")) + Val(rs6(0) & "")
'        RS1.Update
'        RS1.MoveFirst
'    Else
'        RS1.AddNew
'            RS1("act_idy") = Trim(rs6("act_idy") & "")
'            RS1("opa_dbt") = Val(rs6(0) & "")
'            RS1("opa_crt") = 0
'            RS1("fyr_dat") = TO_DAT
'        RS1.Update
'        RS1.Requery
'    End If
    CON.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs6(0)) & " where act_idy like '" & Trim(rs6!ACT_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs6!ACT_IDY) & "'," & Val(rs6(0)) & ",0,ctod('" & TO_DAT & "'))"
    COM_CTR = COM_CTR + 1
    inc_pgb
    rs6.MoveNext
Next
'purchases customer
For A = 1 To rs5.RecordCount
'    RS1.Find "act_idy like '" & Trim(rs5("cst_idy")) & "'"
'    If RS1.EOF = False Then
'        RS1("opa_crt") = Val(RS1("opa_crt")) + Val(rs5(0) & "")
'        RS1.Update
'        RS1.MoveFirst
'    Else
'        RS1.AddNew
'        RS1("act_idy") = Trim(rs5("cst_idy") & "")
'        RS1("opa_crt") = Val(rs5(0) & "")
'        RS1("opa_dbt") = 0
'        RS1("fyr_dat") = TO_DAT
'        RS1.Update
'        RS1.Requery
'    End If
    CON.Execute "update opa set opa_crt = opa_crt+ " & Val(rs5(0)) & " where act_idy like '" & Trim(rs5!CST_IDY) & "'", d
    If d = 0 Then CON.Execute "insert into opa values('" & Trim(rs5!CST_IDY) & "',0," & Val(rs5(0)) & ",ctod('" & TO_DAT & "'))"
    rs5.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next
ProgressBar1.value = 100
End Sub
