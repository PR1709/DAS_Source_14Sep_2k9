VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_UPD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Version 4.0"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H80000006&
   Icon            =   "M_FRM_UPD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3945
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
      Caption         =   "&Details  >>"
      Height          =   375
      Left            =   2745
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   -240
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "&No"
      Height          =   375
      Left            =   1425
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   45
      TabIndex        =   8
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   1215
      TabIndex        =   7
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processed"
      Height          =   195
      Left            =   825
      TabIndex        =   6
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   855
      Left            =   405
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Update Stock...?"
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   50
      Width           =   3855
   End
End
Attribute VB_Name = "M_FRM_UPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COM_CTR As Long
Dim CTR As Long
Private Sub Command1_Click()
On Error Resume Next
If CHECKTB("prd.dbf,brd.dbf,cld.dbf,inv.dbf,van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
CTR = 0
COM_CTR = 0
Me.MousePointer = vbHourglass
If Me.Tag = "AGD" Or Me.Tag = "AGD_CAS" Or Me.Tag = "AGD_VAL" Or Me.Tag = "AGD_MFG" Or Me.Tag = "AGD_6WB" Then
    Ageing_Stock
Else
Label1.Caption = "Updation In Progress..."
Me.Refresh
M_FRM_UPD.MousePointer = vbHourglass
Command1.Enabled = False
Command2.Enabled = False
Dim REC_SET_P As New ADODB.Recordset
Dim REC_SET_I As New ADODB.Recordset
Dim REC_SET_I1 As New ADODB.Recordset
Dim REC_SET_I2 As New ADODB.Recordset
Dim REC_SET_I3 As New ADODB.Recordset
Dim REC_SET_I4 As New ADODB.Recordset
Dim REC_SET_I5 As New ADODB.Recordset
Dim REC_SET_I6 As New ADODB.Recordset
Dim REC_SET_L As New ADODB.Recordset
'CON.BeginTrans
con.Execute "Update BRD SET PRD_QTY = 0"
con.Execute "Update PRD SET PRD.CUR_LEV = 0"
con.Execute "Update PRD SET PRD.PRD_PCS = 0"
con.Execute "delete from cld"

If REC_SET_I.State = 1 Then REC_SET_I.Close
    REC_SET_L.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
    REC_SET_P.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY = PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
    REC_SET_I.Open "select sum(tot_pcs),bat_idy,SUB_UNT from inv where stk_upd = 'Y' and txn_typ in('PUR','OPS','SRN','STI') and stk_typ in('Salable','Free') group by bat_idy", con, adOpenKeyset, adLockPessimistic
    REC_SET_I1.Open "select sum(tot_pcs),bat_idy,SUB_UNT,stk_typ from inv where stk_upd = 'Y' and  txn_typ in('SAL','PRT') group by bat_idy,stk_typ", con, adOpenKeyset, adLockPessimistic
    REC_SET_I4.Open "select sum(tot_pcs),bat_idy,SUB_UNT,stk_typ from inv where stk_upd = 'Y' and  txn_typ in('STO') group by bat_idy,stk_typ", con, adOpenKeyset, adLockPessimistic
    REC_SET_I5.Open "select sum(tot_pcs),bat_idy,SUB_UNT,stk_typ from inv where stk_upd = 'Y' and  txn_typ in('ADJ') group by bat_idy,stk_typ", con, adOpenKeyset, adLockPessimistic
' Code modified by PNR - Stock Update Going Wrong
    REC_SET_I2.Open "SELECT (sum(int(TOT_QTY))* SUB_UNT)+((sum(tot_qty)-sum(int(TOT_QTY)))*1000),BAT_IDY FROM VAN,PRD WHERE VAN.PRD_IDY = PRD.PRD_IDY AND VAN_CHK='V' AND VAN_CLS in('S','C','N') GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
    REC_SET_I3.Open "SELECT (INT(tot_QTY)* SUB_UNT)+((tot_qty-int(TOT_QTY))*1000),BAT_IDY FROM VAN,PRD WHERE VAN.PRD_IDY = PRD.PRD_IDY AND VAN_CHK='P' AND VAN_CLS='N' GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
    REC_SET_I6.Open "SELECT (sum(int(RET_QTY))* SUB_UNT)+((sum(RET_QTY)-sum(int(RET_QTY)))*1000),BAT_IDY FROM VAN,PRD WHERE VAN.PRD_IDY = PRD.PRD_IDY AND VAN_CHK='V' AND VAN_CLS in('S','C','N') GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic

    CTR = REC_SET_P.RecordCount + REC_SET_I.RecordCount + REC_SET_I1.RecordCount + REC_SET_I2.RecordCount + REC_SET_I1.RecordCount + REC_SET_P.RecordCount + REC_SET_P.RecordCount + REC_SET_I4.RecordCount + REC_SET_I5.RecordCount + REC_SET_I6.RecordCount
    Label5 = CTR
    With REC_SET_L
    For A = 1 To REC_SET_P.RecordCount
        .AddNew
        .Fields("PRD_IDY") = REC_SET_P!BAT_IDY
        .Fields("sal_STK") = 0
        .Fields("PUR_STK") = 0
        .Fields("srn_STK") = 0
        .Fields("CLS_STK") = REC_SET_P!SUB_UNT
        .Update
        COM_CTR = COM_CTR + 1
        inc_pgb
        REC_SET_P.MoveNext
    Next
    .Close
    End With
    
    For A = 1 To REC_SET_I.RecordCount
        con.Execute "update cld set pur_stk=" & Val(REC_SET_I(0)) & " where prd_idy = '" & Trim(REC_SET_I!BAT_IDY) & "'", d
        REC_SET_I.MoveNext
        inc_pgb
        COM_CTR = COM_CTR + 1
    Next
    
    For A = 1 To REC_SET_I1.RecordCount
        If Trim(REC_SET_I1!stk_typ) = "Salable" Or Trim(REC_SET_I1!stk_typ) = "Free" Then
            con.Execute "update cld set sal_stk= sal_stk +" & Val(REC_SET_I1(0)) & " where prd_idy = '" & Trim(REC_SET_I1!BAT_IDY) & "'", d
        'ElseIf (Trim(REC_SET_I1!stk_typ) = "Damage" Or Trim(REC_SET_I1!stk_typ) = "TakeBack" Or Trim(REC_SET_I1!stk_typ) = "Breakage" Or Trim(REC_SET_I1!stk_typ) = "Leakage") Then
        ElseIf Trim(REC_SET_I1!stk_typ) = "TakeBack" Then
            con.Execute "update cld set srn_stk = srn_stk+" & Val(REC_SET_I1(0)) & " where prd_idy = '" & Trim(REC_SET_I1!BAT_IDY) & "'", d
        End If
        REC_SET_I1.MoveNext
        inc_pgb
        COM_CTR = COM_CTR + 1
    Next
    
    For A = 1 To REC_SET_I2.RecordCount
        'CON.Execute "update cld set sal_stk= SAL_STK + " & Val(REC_SET_I2(0)) & " where prd_idy = '" & Trim(REC_SET_I2!BAT_IDY) & "'", d
        con.Execute "update cld set sal_stk= SAL_STK + " & Val(REC_SET_I2(0)) & " where prd_idy = '" & Trim(REC_SET_I2!BAT_IDY) & "'", d
        REC_SET_I2.MoveNext
        inc_pgb
        COM_CTR = COM_CTR + 1
    Next
    
    For A = 1 To REC_SET_I6.RecordCount
        'CON.Execute "update cld set pur_stk=" & Val(REC_SET_I(0)) & " where prd_idy = '" & Trim(REC_SET_I!BAT_IDY) & "'", d
        con.Execute "update cld set pur_stk = PUR_STK + " & Val(REC_SET_I6(0)) & " where prd_idy = '" & Trim(REC_SET_I6!BAT_IDY) & "'", d
        REC_SET_I6.MoveNext
        inc_pgb
        COM_CTR = COM_CTR + 1
    Next
    
    For A = 1 To REC_SET_I4.RecordCount
        con.Execute "update cld set sal_stk= SAL_STK + " & Val(REC_SET_I4(0)) & " where prd_idy = '" & Trim(REC_SET_I4!BAT_IDY) & "'", d
        REC_SET_I4.MoveNext
        inc_pgb
        COM_CTR = COM_CTR + 1
    Next
    
    For A = 1 To REC_SET_I5.RecordCount
        con.Execute "update cld set sal_stk= SAL_STK + " & Val(REC_SET_I5(0)) & " where prd_idy = '" & Trim(REC_SET_I5!BAT_IDY) & "'", d
        REC_SET_I5.MoveNext
        inc_pgb
        COM_CTR = COM_CTR + 1
    Next
    
    For A = 1 To REC_SET_I3.RecordCount
        con.Execute "update cld set sal_stk = SAL_STK + " & Val(REC_SET_I3(0)) & " where prd_idy = '" & Trim(REC_SET_I3!BAT_IDY) & "'", d
        REC_SET_I3.MoveNext
        inc_pgb
        COM_CTR = COM_CTR + 1
    Next
    
    If REC_SET_L.State = 1 Then REC_SET_L.Close
    REC_SET_L.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
    For A = 1 To REC_SET_L.RecordCount
        con.Execute "UPDATE BRD WHERE BAT_IDY = '" & Trim(REC_SET_L!prd_idy) & "' SET PRD_QTY=" & STM((Val(REC_SET_L!PUR_STK) + Val(REC_SET_L!srn_stk & "")) - Val(REC_SET_L!sal_stk & ""), Val(REC_SET_L!cls_stk)), d
        REC_SET_L.MoveNext
        COM_CTR = COM_CTR + 1
        inc_pgb
    Next
    
    REC_SET_L.Close
    REC_SET_P.Close
    If REC_SET_L.State = 1 Then REC_SET_L.Close
    REC_SET_L.Open "SELECT sum(PUR_STK),sum(sal_stk),sum(srn_stk),CLS_STK,BRD.PRD_IDY FROM CLD,BRD WHERE CLD.PRD_IDY = BRD.BAT_IDY group by BRD.PRD_IDY", con, adOpenKeyset
    For A = 1 To REC_SET_L.RecordCount
        b = (Val(REC_SET_L(0)) + Val(REC_SET_L(2) & "") - Val(REC_SET_L(1)))
        con.Execute "UPDATE PRD SET PRD_PCS= PRD_PCS + " & Val(b) & "  WHERE PRD_IDY = '" & Trim(REC_SET_L!prd_idy) & "'", d
        con.Execute "UPDATE PRD SET CUR_LEV=" & STM(Val(b), Val(REC_SET_L(3) & "")) & "  WHERE PRD_IDY = '" & Trim(REC_SET_L!prd_idy) & "'", d
        REC_SET_L.MoveNext
        COM_CTR = COM_CTR + 1
        inc_pgb
    Next
    'CON.CommitTrans
End If
If Me.Tag = "DIS_STK_REP" Then
    OpenReport CURDIR & "\lists.RP1", "Products List (80) Column", 0, "", 6
End If
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Public Sub inc_pgb()
On Error Resume Next
DoEvents
Label4.Caption = COM_CTR
Label4.Refresh
ProgressBar1.Value = (100 / CTR) * COM_CTR
End Sub

Private Sub Command3_Click()
If Me.Height = 1725 Then
Command3.Caption = "&Details  <<"
Me.Height = 2850
Else
Me.Height = 1725
Command3.Caption = "&Details  >>"
End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
If Me.Tag = "LDSORD" Then
    Command2.Enabled = False
End If
End Sub

Private Sub Form_Load()
Me.Height = 1725
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If ProgressBar1.Value > 1 Then
    Me.MousePointer = vbHourglass
End If
End Sub
Private Sub Ageing_Stock()
On Error GoTo lab
If CHECKTB("inv.dbf,prd.dbf,brd.dbf,sta.dbf,tpd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim REC_INV As New ADODB.Recordset
Dim REC_INV1 As New ADODB.Recordset
Dim REC_INV2 As New ADODB.Recordset
Dim REC_INV3 As New ADODB.Recordset
Dim REC_STA As New ADODB.Recordset
Dim REC_TPD As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
con.Execute "DELETE FROM TPD"
con.Execute "DELETE FROM STA"
REC_INV.Open "SELECT SUM(TOT_PCS) AS TOTAL,INV.PRD_IDY,BAT_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND TXN_TYP IN('SAL','STO','PRT') AND INV_DAT <= CTOD('" & TO_DAT & "') AND INV.STK_TYP IN ('Free','Salable')  AND STK_UPD = 'Y' GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
REC_INV2.Open "SELECT SUM(TOT_PCS) AS TOTAL,INV.PRD_IDY,BAT_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND TXN_TYP IN('SAL','PRT') AND INV_DAT <= CTOD('" & TO_DAT & "') AND INV.STK_TYP IN ('TakeBack','Damage','Breakage','Leakage')  AND STK_UPD = 'Y' GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
REC_INV3.Open "SELECT SUM(TOT_PCS) AS TOTAL,INV.PRD_IDY,BAT_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND TXN_TYP IN('STO') AND INV_DAT <= CTOD('" & TO_DAT & "') AND INV.STK_TYP IN ('Damage') AND STK_UPD = 'Y' GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic

' This part of the code is modified by Radhika and Ramanesh
'REC_INV.Open "SELECT SUM(TOT_PCS) AS TOTAL,INV.PRD_IDY,BAT_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND TXN_TYP IN('SAL','STO','PRT') AND INV_DAT <= CTOD('" & TO_DAT & "') AND INV.STK_TYP IN ('Free','Salable')  GROUP BY BAT_IDY", CON, adOpenKeyset, adlockpessimistic
'REC_INV2.Open "SELECT SUM(TOT_PCS) AS TOTAL,INV.PRD_IDY,BAT_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND TXN_TYP IN('SAL','PRT') AND INV_DAT <= CTOD('" & TO_DAT & "') AND INV.STK_TYP IN ('TakeBack','Damage','Breakage','Leakage')  GROUP BY BAT_IDY", CON, adOpenKeyset, adlockpessimistic
'REC_INV3.Open "SELECT SUM(TOT_PCS) AS TOTAL,INV.PRD_IDY,BAT_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND TXN_TYP IN('STO') AND INV_DAT <= CTOD('" & TO_DAT & "') AND INV.STK_TYP IN ('Damage') GROUP BY BAT_IDY", CON, adOpenKeyset, adlockpessimistic
' End Modify

REC_TPD.Open "SELECT * FROM TPD", con, adOpenKeyset, adLockPessimistic
REC_INV1.Open "SELECT TOT_PCS AS TOTAL,* FROM INV WHERE TXN_TYP IN ('OPS','PUR','SRN','STI') AND STK_UPD = 'Y'", con, adOpenKeyset, adLockPessimistic
REC_STA.Open "SELECT * FROM STA", con, adOpenKeyset, adLockPessimistic
REC_PRD.Open "SELECT * FROM PRD,BRD WHERE PRD.PRD_IDY = BRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
CTR = ((REC_INV.RecordCount + REC_INV1.RecordCount) * 2) + Val(REC_INV3.RecordCount)
Label5 = CTR
For CTR1 = 1 To REC_INV.RecordCount
    REC_TPD.AddNew
    REC_TPD!prd_idy = Trim(REC_INV!BAT_IDY)
    REC_TPD!prd_qty = Val(REC_INV!Total)
    REC_TPD!PRD_VAL = 0
    REC_TPD!PRD_UNR = 0
    REC_TPD.Update
    REC_INV.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next

For CTR1 = 1 To REC_INV3.RecordCount
    REC_TPD.AddNew
    REC_TPD!prd_idy = Trim(REC_INV3!BAT_IDY)
    REC_TPD!prd_qty = Val(REC_INV3!Total)
    REC_TPD!PRD_VAL = 0
    REC_TPD!PRD_UNR = 0
    REC_TPD.Update
    REC_INV3.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next

For CTR1 = 1 To REC_INV2.RecordCount
    con.Execute "update tpd set prd_qty = prd_qty - " & Val(REC_INV2!Total) & " WHERE prd_idy = '" & Trim(REC_INV2!BAT_IDY) & "'", d
    If d = 0 Then
        REC_TPD.AddNew
        REC_TPD!prd_idy = Trim(REC_INV2!BAT_IDY)
        REC_TPD!prd_qty = Val(REC_INV2!Total) * -1
        REC_TPD!PRD_VAL = 0
        REC_TPD!PRD_UNR = 0
        REC_TPD.Update
    End If
    REC_INV2.MoveNext
Next

For CTR1 = 1 To REC_INV1.RecordCount
    REC_STA.AddNew
    REC_STA!prd_idy = Trim(REC_INV1!BAT_IDY)
    REC_STA!TXN_IDY = Trim(REC_INV1!INV_IDY)
    REC_STA!txn_dat = Trim(REC_INV1!INV_DAT)
    REC_STA!prd_qty = Trim(REC_INV1!Total)
    REC_STA!COL_AMT = 0
    REC_STA.Update
    REC_INV1.MoveNext
    COM_CTR = COM_CTR + 1
    inc_pgb
Next
REC_TPD.Requery
REC_STA.Requery
For A = 1 To REC_STA.RecordCount
    REC_TPD.Find "prd_idy = '" & Trim(REC_STA("prd_idy")) & "'"
    If REC_TPD.EOF = False Then
        If Val(REC_TPD!prd_qty & "") > Val(REC_STA!prd_qty & "") Then
            TPDQTY = Val(REC_TPD!prd_qty)
            STAQTY = Val(REC_STA!prd_qty)
            REC_TPD!prd_qty = Val(TPDQTY) - Val(STAQTY)
            REC_STA!prd_qty = 0
            REC_STA!COL_AMT = 0
            REC_TPD.Update
            REC_STA.Update
        Else
            TPDQTY = Val(REC_TPD!prd_qty)
            STAQTY = Val(REC_STA!prd_qty)
            REC_STA!prd_qty = Val(STAQTY) - Val(TPDQTY)
            REC_TPD!prd_qty = 0
            REC_STA!COL_AMT = 0
            REC_TPD.Update
            REC_STA.Update
        End If
    End If
REC_TPD.Requery
REC_STA.MoveNext
COM_CTR = COM_CTR + 1
inc_pgb
Next
con.Execute "DELETE FROM STA WHERE PRD_QTY = 0"
Me.MousePointer = vbDefault
If CHECKTB("stock.RP1") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If Me.Tag = "AGD" Then
    OpenReport CURDIR & "\STOCK.RP1", "Stocks Ageing Batch", 2, "BRD->PRD_QTY > 0", 6
ElseIf Me.Tag = "AGD_6WB" Then
    OpenReport CURDIR & "\KOR.RP1", "Stocks Ageing 6 Weeks", 2, "BRD->PRD_QTY > 0", 6
ElseIf Me.Tag = "AGD_CAS" Then
    OpenReport CURDIR & "\STOCK.RP1", "Stocks Ageing Batch Cases", 2, "BRD->PRD_QTY > 0", 6
ElseIf Me.Tag = "AGD_VAL" Then
    OpenReport CURDIR & "\STOCK.RP1", "Stocks Ageing Batch Value", 2, "BRD->PRD_QTY > 0", 6
ElseIf Me.Tag = "AGD_MFG" Then
    OpenReport CURDIR & "\STOCK.RP1", "Stocks Ageing Batch Mfg", 2, "BRD->PRD_QTY > 0", 6
End If
Exit Sub
lab:
Resume Next
End Sub
