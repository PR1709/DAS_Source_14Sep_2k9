VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form M_FRM_DSN1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Report Dates"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "M_FRM_DSN1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3855
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   1710
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16515075
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1710
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16515075
         CurrentDate     =   37258
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   195
         Left            =   750
         TabIndex        =   6
         Top             =   915
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "M_FRM_DSN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim REC_SET As New ADODB.Recordset
Command1.Enabled = False
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
End If
con.Execute "update dir set dir_als= '" & Format(DTPicker1, "DD/MM/YYYY") & "', dir_msg='" & Format(DTPicker2, "DD/MM/YYYY") & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
con.Execute "update dir set dir_als= '" & DateDiff("D", DTPicker1, DTPicker2) & "' where  dir_idy like 'AUT' and dir_idx like 'IND'"
SAL_QRY = "inV->inv_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inv->inv_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
If M_FRM_DSN1.Tag = "Sales Return Report" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='SRN'"
    OpenReport CURDIR & "\report.RP1", "Sales Return Report", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "Adjustment" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='ADJ'"
    OpenReport CURDIR & "\Vouch.RP1", "Adjustment Report", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "Salable CNV" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='CNV'.AND.INV->STK_TYP='Damage'"
    OpenReport CURDIR & "\Vouch.RP1", "Stock Conversion Report", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "Damage CNV" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='CNV'.AND.INV->STK_TYP='Salable'"
    OpenReport CURDIR & "\Vouch.RP1", "Stock Conversion Report", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "SAL_PRD_TAX_TTL" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='SAL'"
    OpenReport CURDIR & "\REPORT.RP1", "Sales Product Tax Total", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "PUR_PRD_TAX_TTL" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PUR'"
    OpenReport CURDIR & "\REPORT.RP1", "Purchase Product Tax Total", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "Sales Return Product" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='SRN'"
    OpenReport CURDIR & "\report.RP1", "Sales Return Product", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "Tax Report" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL').AND.(INL->TAX_AMT>0.OR.INL->FRT_AMT>0.OR.INV->TAX_AMT>0)"
    OpenReport CURDIR & "\report.RP1", "Sales Tax Register", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "VAT" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='PUR'.OR.INV->TXN_TYP='STO'.OR.INV->TXN_TYP='PRT'.OR.INV->TXN_TYP='STI'.OR.INV->TXN_TYP='SRN').AND.(INL->TAX_AMT>0.OR.INL->FRT_AMT>0.OR.INV->TAX_AMT>0)"
    OpenReport CURDIR & "\report.RP1", "Tax Report Vat", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "Purchase Register" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PUR'"
    OpenReport CURDIR & "\report.RP1", "Purchase Register", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "Purchase Return Report" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PUR'"
    OpenReport CURDIR & "\report.RP1", "Purchase Return Report", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "Purchase Book" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PUR'"
    OpenReport CURDIR & "\report.RP1", "Purchase Book", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_DSN1.Tag = "Stock Issues" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='STI'"
    OpenReport CURDIR & "\report.RP1", "Stock Receipts Report", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Stock Receipts" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='STO'"
    OpenReport CURDIR & "\report.RP1", "Stock Issues Report", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Inward Product" Then
    SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP='STI'"
    OpenReport CURDIR & "\report.RP1", "Inward Product", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Outward Product" Then
    SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP='STO'"
    OpenReport CURDIR & "\report.RP1", "Outward Product", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Sales Tax Product" Then
    SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP='SAL' .and. INV->TRD_TAX > 0"
    OpenReport CURDIR & "\report.RP1", "Sales Product Tax", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Purchase Product Tax" Then
    SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP='PUR' .and. INV->TRD_TAX > 0"
    OpenReport CURDIR & "\report.RP1", "Purchase Product Tax", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Opening Stocks" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='OPS'"
    OpenReport CURDIR & "\report.RP1", "Opening Stocks Register", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "A9" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PUR'"
    OpenReport CURDIR & "\report.RP1", "A-9 Statement", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Daily(Stock)" Then
    SAL_QRY = "inl->inv_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inl->inv_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\report.RP1", "Daily Transactions", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Daily(Finance)" Then
    SAL_QRY = "VCH->TXN_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.VCH->TXN_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\report.RP1", "Daily Transactions (Fin)", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Ready Stocks Report (Issues)" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='RSI'"
    OpenReport CURDIR & "\report.RP1", "Ready Stocks Report (Issues)", 2, SAL_QRY, 6
ElseIf M_FRM_DSN1.Tag = "Van Challan Returns" Then
    VAN_RET_DAT
    OpenReport CURDIR & "\vouch.RP1", "Van Returns", 2, "", 6
ElseIf Me.Tag = "PMT" Then
    SAL_QRY = "VCH->TXN_TYP='PMT'.AND.VCH->TXN_DAT>=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "').AND.VCH->TXN_DAT<=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\Finance.RP1", "Payments List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "RCT" Then
    SAL_QRY = "VCH->TXN_TYP='RCT'.AND.VCH->TXN_DAT>=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "').AND.VCH->TXN_DAT<=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\Finance.RP1", "Receipts List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "JVL" Then
    SAL_QRY = "VCH->TXN_TYP='JRN'.AND.VCH->TXN_DAT>=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "').AND.VCH->TXN_DAT<=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\Finance.RP1", "Journals List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "CRN" Then
    SAL_QRY = "VCH->TXN_TYP='CRN'.AND.VCH->TXN_DAT>=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "').AND.VCH->TXN_DAT<=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\Finance.RP1", "Credit Notes List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "DRN" Then
    SAL_QRY = "VCH->TXN_TYP='DBN'.AND.VCH->TXN_DAT>=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "').AND.VCH->TXN_DAT<=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\Finance.RP1", "Debit Notes List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "Schemes List" Then
    SAL_QRY = "SCH->SCH_BDT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.SCH->SCH_BDT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\Lists.RP1", "Schemes List", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "Sales Register" Then
    SAL_QRY = "INV->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.INV->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    SAL_QRY = SAL_QRY & ".AND.INV->TXN_TYP='SAL'.AND.(INV->STK_TYP='Salable'.or.inv->stk_typ='Free')"
    OpenReport "", CURDIR & "\Sales Register.RRW", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "Claims Report" Or Me.Tag = "Claims Report Brand" Or Me.Tag = "Claims Report Segment" Or Me.Tag = "Claims Report SKU" Then
    CLAIM
ElseIf Me.Tag = "SSM" Or Me.Tag = "SUN" Or Me.Tag = "PRD" Or Me.Tag = "Inventory" Or Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Or Me.Tag = "DIS_STK_REP" Or Me.Tag = "SNV_STT" Then
    If rs.State = 1 Then rs.Close
    con.Execute "DELETE FROM CLD"
    rs.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
    If Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Then
        If RS1.State = 1 Then RS1.Close
        'RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat < ctod('" & DTPicker1 & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & DTPicker1 & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika and Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where (STK_UPD ='Y' or ORD_IDY # ' ') AND inv_dat < ctod('" & DTPicker1 & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
        'If Me.Tag = "DIS_STK_REP" Then
            RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & DTPicker1 & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic
        'Else
        '    RS1.Open "SELECT TOT_QTY,RET_QTY FROM VAN WHERE "
        'End If
    End If
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0)) * -1
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "OPS" Then
            con.Execute "UPDATE CLD SET OPN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!opn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PUR" Then
        con.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PUR_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PRT" Then
        con.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PRT_STK = Val(RS1(0))
                rs!PUR_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STI" Then
        con.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!STI_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!PUR_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SRN" Then
        con.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        con.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs(7) = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs!PUR_STK = 0
                rs!cls_stk = 0
                rs.Update
            End If
        End If
        RS1.MoveNext
    Next
    rs.Close
    con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
    con.Execute "UPDATE CLD SET OPN_STK=CLS_STK"
    con.Execute "UPDATE CLD SET PUR_STK=0"
    con.Execute "UPDATE CLD SET PRT_STK=0"
    con.Execute "UPDATE CLD SET SAL_STK=0"
    con.Execute "UPDATE CLD SET STO_STK=0"
    con.Execute "UPDATE CLD SET SRN_STK=0"
    con.Execute "UPDATE CLD SET STI_STK=0"
    con.Execute "UPDATE CLD SET CLS_STK=0"
    
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
    If Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Then
        If RS1.State = 1 Then RS1.Close
        'RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    End If
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0)) * -1
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "OPS" Then
            con.Execute "UPDATE CLD SET OPN_STK = OPN_STK +" & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!opn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PUR" Then
        con.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PUR_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PRT" Then
        con.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PRT_STK = Val(RS1(0))
                rs!PUR_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STI" Then
        con.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!STI_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!PUR_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SRN" Then
        con.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        con.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs(7) = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs!PUR_STK = 0
                rs!cls_stk = 0
                rs.Update
            End If
        End If
        RS1.MoveNext
    Next
    rs.Close
    con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
    If Me.Tag = "SSM" Then OpenReport CURDIR & "\STOCK.RP1", "Stock Summary (80) Column", 0, "", 6
    If Me.Tag = "SUN" Or Me.Tag = "DIS_STK_REP" Then OpenReport CURDIR & "\STOCK.RP1", "Stock Summary Cases", 0, "", 6
    If Me.Tag = "PRD" Then OpenReport CURDIR & "\STOCK.RP1", "Stock Summary Product", 0, "", 6
    If Me.Tag = "Inventory" Then OpenReport CURDIR & "\STOCK.RP1", "Inventory Days", 0, "", 6
    If Me.Tag = "COK_STK_BAT" Then OpenReport CURDIR & "\CR.RP1", "Stock Summary Cases", 0, "", 6
    If Me.Tag = "COK_STK_BRD" Then OpenReport CURDIR & "\CR.RP1", "Stock Summary Brand wise", 0, "", 6
    If Me.Tag = "COK_STK_SEG" Then OpenReport CURDIR & "\CR.RP1", "Stock Summary Segment Wise", 0, "", 6
    If Me.Tag = "SNV_STT" Then OpenReport CURDIR & "\STOCK.RP1", "Stock and Value Statement", 0, "", 6
ElseIf Me.Tag = "STS" Then
    Me.MousePointer = vbHourglass
    con.Execute "delete from sts"
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    rs.Open "select SUM(net_amt),TXN_TYP from inl where INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND inv_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "') group by txn_typ", con, adOpenKeyset, adLockPessimistic
    RS1.Open "select * from sts", con, adOpenKeyset, adLockPessimistic
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
    rs.Open "select sum(net_amt),txn_typ from inl where inv_dat = ctod('" & DTPicker2 & "') group by txn_typ", con, adOpenKeyset, adLockPessimistic
    RS1.Open "select * from sts", con, adOpenKeyset, adLockPessimistic
    If rs.EOF = False Then rs.MoveFirst
    While rs.EOF = False
        If Trim(rs(1)) = "SAL" Then
'            RS1.Update "D_crd_sal", Val(rs(0))
            RS1.Update "day_sal", Val(RS1("day_sal") & "") + Val(rs(0) & "")
        ElseIf Trim(rs(1)) = "LSL" Then
            RS1.Update "D_cas_sal", Val(rs(0))
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
    rs.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    While rs.EOF = False
        STK_VAL = Val(STK_VAL) + Val(RATCAL(QTY(rs("cur_lev") & ""), Val(rs("PUR_RAT") & ""), Val(rs("SUB_UNT") & "")))
        rs.MoveNext
    Wend
    RS1.Update "STK_VAL", Val(STK_VAL)
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Report.RP1", "Status Sheet", 0, "", 6
ElseIf Me.Tag = "Day Summary" Then
    con.Execute "update opa set opa_dbt=0"
    con.Execute "update opa set opa_crt=0"
    con.Execute "update dsu set rct_amt=0"
    con.Execute "update dsu set pmt_amt=0"
    Me.MousePointer = vbHourglass
    If rs.State = 1 Then rs.Close
    rs.Open "select * from vch where txn_dat < ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "')", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM opa where act_idy like '" & Trim(rs("dbt_idy")) & "'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then
            RS1("opa_dbt") = Val(RS1("opa_dbt")) + Val(rs("txn_amt") & "")
            RS1.Update
        Else
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from opa", con, adOpenKeyset, adLockPessimistic
            RS1.AddNew
                RS1("act_idy") = Trim(rs("dbt_idy") & "")
                RS1("opa_dbt") = Val(rs("txn_amt") & "")
                RS1("opa_crt") = 0
                RS1("fyr_dat") = TO_DAT
            RS1.Update
        End If
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM opa where act_idy like '" & Trim(rs("CRt_idy")) & "'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then
            RS1("opa_crt") = Val(RS1("opa_crt")) + Val(rs("txn_amt") & "")
            RS1.Update
        Else
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from opa", con, adOpenKeyset, adLockPessimistic
            RS1.AddNew
                RS1("act_idy") = Trim(rs("crt_idy") & "")
                RS1("opa_crt") = Val(rs("txn_amt") & "")
                RS1("opa_dbt") = 0
                RS1("fyr_dat") = TO_DAT
            RS1.Update
        End If
        rs.MoveNext
    Next
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "select * from vch where txn_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND TXN_DAT <=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "')", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM DSU where act_idy like '" & Trim(rs("dbt_idy")) & "'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then
            RS1("RCT_AMT") = Val(RS1("RCT_AMT")) + Val(rs("txn_amt") & "")
            RS1.Update
        Else
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from DSU", con, adOpenKeyset, adLockPessimistic
            RS1.AddNew
                RS1("act_idy") = Trim(rs("dbt_idy") & "")
                RS1("RCT_AMT") = Val(rs("txn_amt") & "")
                RS1("PMT_AMT") = 0
            RS1.Update
        End If
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM DSU where act_idy like '" & Trim(rs("CRt_idy")) & "'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then
            RS1("PMT_AMT") = Val(RS1("PMT_AMT")) + Val(rs("txn_amt") & "")
            RS1.Update
        Else
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from DSU", con, adOpenKeyset, adLockPessimistic
            RS1.AddNew
                RS1("act_idy") = Trim(rs("crt_idy") & "")
                RS1("PMT_AMT") = Val(rs("txn_amt") & "")
                RS1("RCT_AMT") = 0
            RS1.Update
        End If
        rs.MoveNext
    Next
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\Finance.RP1", "Day Summary", 2, "act->grp_clf='Bank Book'.or.act->grp_clf='Cash Book'", 6
ElseIf Me.Tag = "AVG" Then
    FRM_DAT = DTPicker1
    TO_DATE = DTPicker2
ElseIf Me.Tag = "CLG_CHQ_RCT" Then
    SAL_QRY = "Vch->clg_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.vch->clg_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    SAL_QRY = SAL_QRY & ".and.(vch->txn_typ='RCT'.AND.ACT1->GRP_CLF='Bank Book')"
    OpenReport CURDIR & "\FINANCE.RP1", "Outstanding Cheques", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "CLG_CHQ_PMT" Then
    SAL_QRY = "Vch->clg_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.vch->clg_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    SAL_QRY = SAL_QRY & ".and.(vch->txn_typ='PMT'.AND.ACT->GRP_CLF='Bank Book')"
    OpenReport CURDIR & "\FINANCE.RP1", "Outstanding Cheques", 2, Trim(SAL_QRY), 6
ElseIf Me.Tag = "AST" Then
    'SAL_QRY = "txn_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.txn_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    'SAL_QRY = "txn_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.ast->txn_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\ASSETS.RP1", "Assets", 0, SAL_QRY, 6
ElseIf Me.Tag = "DPS" Then
    UPDATEDPS
ElseIf Me.Tag = "RRN" Then
    AddScheduledCalls DTPicker1, DTPicker2, "ROU"
    doOutletsnotBilled DTPicker1, DTPicker2, "ROU"
    OpenReport CURDIR & "\PNT_MIS.RP1", "Retailers Route NB", 0, "", 6
ElseIf Me.Tag = "RSN" Then
    AddScheduledCalls DTPicker1, DTPicker2, "SMN"
    doOutletsnotBilled DTPicker1, DTPicker2, "SMN"
    OpenReport CURDIR & "\PNT_MIS.RP1", "Retailers Salesman NB", 0, "", 6
ElseIf Me.Tag = "LOGREP" Then
    OpenReport CURDIR & "\vvr_rpt.RP1", "Log Report", 0, SAL_QRY, 6
    'M_REP_IQY.Tag = M_FRM_DSN1.Tag
    'M_REP_IQY.Show 1
End If
Unload Me
If M_FRM_DSN1.Caption = "Stock Ledger " Then
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(K As Integer, s As Integer)
If K = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    Command1.SetFocus
ElseIf K = 37 And s = 1 Then
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
SETYEAR M_FRM_DSN1
DTPicker1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then Shell "c:\windows\hh.exe " & App.Path & "\conthelp.chm", vbMaximizedFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Public Sub VAN_RET_DAT()
Dim INV_DAT As New ADODB.Recordset
Dim VAN_DAT As New ADODB.Recordset
Dim RSI_DAT As New ADODB.Recordset
Dim RSR_DAT As New ADODB.Recordset
Me.MousePointer = vbHourglass
con.Execute "DELETE FROM VAN"
VAN_DAT.Open "SELECT * FROM VAN", con, adOpenKeyset, adLockPessimistic
INV_DAT.Open "SELECT * FROM INV WHERE INV_DAT >= CTOD('" & DTPicker1 & "') AND INV_DAT <= CTOD('" & DTPicker2 & "') AND TXN_TYP LIKE 'RSI'", con, adOpenKeyset, adLockPessimistic
For A = 1 To INV_DAT.RecordCount
    If RSI_DAT.State = 1 Then RSI_DAT.Close
        RSI_DAT.Open "SELECT * FROM INL WHERE INV_IDY LIKE '" & INV_DAT("INV_IDY") & "' AND TXN_TYP LIKE 'RSI'", con, adOpenDynamic, adLockPessimistic
    If RSR_DAT.State = 1 Then RSR_DAT.Close
        RSR_DAT.Open "SELECT * FROM INL WHERE ORD_IDY LIKE '" & RSI_DAT("INV_IDY") & "' AND TXN_TYP LIKE 'RSR'", con, adOpenDynamic, adLockPessimistic
        If RSR_DAT.EOF = False Then
            VAN_DAT.AddNew
            VAN_DAT("INV_IDY") = RSR_DAT("INV_IDY")
            VAN_DAT("INV_DAT") = INV_DAT("INV_DAT")
            VAN_DAT("ORD_IDY") = INV_DAT("INV_IDY")
            VAN_DAT("PRD_IDY") = INV_DAT("PRD_IDY")
            VAN_DAT("PRD_QTY") = INV_DAT("PRD_QTY")
            VAN_DAT("RET_QTY") = INV_DAT("RET_QTY") & ""
            VAN_DAT("BIL_QTY") = INV_DAT("PRD_QTY") - INV_DAT("RET_QTY")
            VAN_DAT.Update
            INV_DAT.MoveNext
        Else
            INV_DAT.MoveNext
        End If
Next
Me.MousePointer = vbNormal
End Sub

Public Sub CLAIM()
Me.MousePointer = vbHourglass
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
con.Execute "DELETE FROM CLD"
rs.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
If RS1.State = 1 Then RS1.Close
'RS1.Open "select sum(tot_pcs),prd_idy,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat < ctod('" & DTPicker1 & "') group by prd_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
RS1.Open "select sum(tot_pcs),prd_idy,txn_typ,stk_typ from inv where inv_dat < ctod('" & DTPicker1 & "') group by prd_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic ' Modified by Radhika & Ramanesh
For A = 1 To RS1.RecordCount
    If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0)) * -1
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "OPS" Then
        con.Execute "UPDATE CLD SET OPN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!opn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PUR" Then
    con.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PUR_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PRT" Then
    con.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PRT_STK = Val(RS1(0))
            rs!PUR_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STI" Then
    con.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!STI_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!PUR_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SRN" Then
    con.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STO" Then
    con.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs(7) = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs!PUR_STK = 0
            rs!cls_stk = 0
            rs.Update
        End If
    End If
    RS1.MoveNext
Next
rs.Close

con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
con.Execute "UPDATE CLD SET OPN_STK=CLS_STK"
con.Execute "UPDATE CLD SET PUR_STK=0"
con.Execute "UPDATE CLD SET PRT_STK=0"
con.Execute "UPDATE CLD SET SAL_STK=0"
con.Execute "UPDATE CLD SET STO_STK=0"
con.Execute "UPDATE CLD SET SRN_STK=0"
con.Execute "UPDATE CLD SET STI_STK=0"
con.Execute "UPDATE CLD SET CLS_STK=0"
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
If RS1.State = 1 Then RS1.Close
'RS1.Open "select sum(tot_pcs),prd_idy,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') group by prd_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
RS1.Open "select sum(tot_pcs),prd_idy,txn_typ,stk_typ from inv WHERE inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') group by prd_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh & Radhika
For A = 1 To RS1.RecordCount
    If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0)) * -1
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "OPS" Then
        con.Execute "UPDATE CLD SET OPN_STK = OPN_STK +" & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!opn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PUR" Then
    con.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PUR_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PRT" Then
    con.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PRT_STK = Val(RS1(0))
            rs!PUR_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STI" Then
    con.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!STI_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!PUR_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SRN" Then
    con.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STO" Then
    con.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!prd_idy)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs(7) = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs!PUR_STK = 0
            rs!cls_stk = 0
            rs.Update
        End If
    End If
    RS1.MoveNext
Next
rs.Close
con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
con.Execute "DELETE FROM SCL"
RS1.Open "SELECT * FROM SCL", con, adOpenKeyset, adLockPessimistic
rs.Open "SELECT * FROM SCH WHERE SCH_IDY IN(SELECT SCH_IDY FROM INV WHERE inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') GROUP BY SCH_IDY )GROUP BY SCH_IDY,B_PRD_IDY,TXN_TYP", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    RS1.AddNew
    RS1(0) = Date
    RS1(2) = Trim(rs!B_PRD_IDY & "")
    RS1(1) = Trim(rs!Sch_Idy)
    RS1(3) = 0
    RS1(4) = 0
    RS1(5) = 0
    RS1(6) = 0
    RS1(7) = 0
    RS1(8) = 0
    RS1(9) = 0
    RS1(10) = 0
    RS1(11) = Trim(rs!TXN_TYP)
    RS1.Update
    rs.MoveNext
Next
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        con.Execute "UPDATE SCL SET TOT_SAL =  " & Val(rs!sal_stk & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!prd_idy) & "'", d
        con.Execute "UPDATE SCL SET CLS_STK =  " & Val(rs!cls_stk & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!prd_idy) & "'", d
        con.Execute "UPDATE SCL SET OPN_STK =  " & Val(rs!opn_stk & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!prd_idy) & "'", d
        con.Execute "UPDATE SCL SET RCT_STK =  " & Val(rs!PUR_STK & "") + Val(rs!srn_stk & "") + Val(rs!STI_STK & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!prd_idy) & "'", d
        con.Execute "UPDATE SCL SET ISS_STK =  " & Val(rs!PRT_STK & "") + Val(rs!STO_STK & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!prd_idy) & "'", d
        rs.MoveNext
    Next
    
    If rs.State = 1 Then rs.Close
    'RS.Open "select sum(tot_pcs),prd_idy,SCH_IDY,SUM((PRD_RAT/SUB_UNT) * TOT_PCS) SCH_AMT from inv where STK_UPD ='Y' AND inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') AND TXN_TYP ='SAL' AND STK_TYP LIKE 'Free' group by prd_idy,SCH_IDY", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select sum(tot_pcs),prd_idy,SCH_IDY,SUM((PRD_RAT/SUB_UNT) * TOT_PCS) SCH_AMT from inv where inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') AND TXN_TYP ='SAL' AND STK_TYP LIKE 'Free' group by prd_idy,SCH_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        con.Execute "UPDATE SCL SET fre_STK =  " & Val(rs(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!prd_idy) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'", d
        con.Execute "UPDATE SCL SET SCH_AMT =  " & Val(rs(3) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!prd_idy) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'", d
        rs.MoveNext
    Next
    con.Execute "DELETE FROM SCL WHERE SCH_IDY NOT IN(SELECT SCH_IDY FROM SCL WHERE FRE_STK >0)"
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM SCL", con, adOpenKeyset, adLockPessimistic
    While rs.EOF = False
        If RS1.State = 1 Then RS1.Close
        'RS1.Open "select sum(tot_pcs),prd_idy from inv where inv_idy in(SELECT inv_idy from inv where STK_UPD ='Y' AND inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') AND TXN_TYP ='SAL' AND SCH_IDY LIKE '" & Trim(RS!Sch_Idy & "") & "' AND STK_TYP LIKE 'Free') and stk_typ like 'Salable' group by prd_idy", CON, adOpenKeyset, adlockpessimistic
        RS1.Open "select sum(tot_pcs),prd_idy from inv where inv_idy in(SELECT inv_idy from inv where inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') AND TXN_TYP ='SAL' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy & "") & "' AND STK_TYP LIKE 'Free') and stk_typ like 'Salable' group by prd_idy", con, adOpenKeyset, adLockPessimistic
        For A = 1 To RS1.RecordCount
            con.Execute "UPDATE SCL SET SAL_STK =" & Val(RS1(0) & "") & " WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND PRD_IDY LIKE '" & Trim(RS1!prd_idy & "") & "'", d
            RS1.MoveNext
        Next
    rs.MoveNext
    Wend
    Me.MousePointer = vbDefault
    If Me.Tag = "Claims Report" Then
        OpenReport CURDIR & "\CR.RP1", "Schemes Report Cases", 0, "", 6
    ElseIf Me.Tag = "Claims Report Brand" Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM SCL", con, adOpenKeyset, adLockPessimistic
        For A = 1 To RS1.RecordCount
            SUBUNT = GetValue("CON_FAC/SUB_UNT", "PRD WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy & "'"))
            RS1!fre_stk = Val(RS1!fre_stk & "") * (Val(SUBUNT))
            RS1!sal_stk = Val(RS1!sal_stk & "") * Val(SUBUNT)
            RS1.Update
            RS1.MoveNext
        Next
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT SCH_IDY,PRD.PRD_IDY,SUM(FRE_STK),SUM(SAL_STK),BAS_TYP,SUM(SCH_AMT) FROM SCL,PRD WHERE SCL.PRD_IDY LIKE PRD.PRD_IDY GROUP BY SCH_IDY,BRD_NME,BAS_TYP", con, adOpenStatic, adLockReadOnly
        con.Execute "DELETE FROM SCL"
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM SCL", con, adOpenKeyset, adLockPessimistic
        For A = 1 To rs.RecordCount
            RS1.AddNew
            For b = 0 To RS1.Fields.count - 1
                If RS1(b).Type = 133 Then
                    RS1(b) = Date
                ElseIf RS1(b).Type = adChar Then
                    RS1(b) = ""
                Else
                    RS1(b) = 0
                End If
            Next
            RS1!Sch_Idy = Trim(rs(0))
            RS1!prd_idy = Trim(rs(1))
            RS1!fre_stk = Val(rs(2) & "")
            RS1!sal_stk = Trim(rs(3) & "")
            RS1!bas_typ = Trim(rs(4))
            RS1!SCH_AMT = Val(rs(5) & "")
            RS1.Update
            rs.MoveNext
        Next
        OpenReport CURDIR & "\CR.RP1", "Schemes Report Brand", 0, "", 6
    ElseIf Me.Tag = "Claims Report Segment" Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM SCL", con, adOpenKeyset, adLockPessimistic
        For A = 1 To RS1.RecordCount
            SUBUNT = GetValue("SUB_UNT", "PRD WHERE PRD_IDY LIKE '" & Trim(RS1!prd_idy & "'"))
            RS1!fre_stk = STM(Val(RS1!fre_stk & ""), Val(SUBUNT))
            RS1!sal_stk = STM(Val(RS1!sal_stk & ""), Val(SUBUNT))
            RS1.Update
            RS1.MoveNext
        Next
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT SCH_IDY,PGR.GRP_IDY,SUM(FRE_STK),SUM(SAL_STK),BAS_TYP,SUM(SCH_AMT) FROM SCL,PRD,PGR WHERE PRD.GRP_IDY LIKE PGR.GRP_IDY AND SCL.PRD_IDY LIKE PRD.PRD_IDY GROUP BY SCH_IDY,GRP_CLS,BAS_TYP", con, adOpenStatic, adLockReadOnly
        con.Execute "DELETE FROM SCL"
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM SCL", con, adOpenKeyset, adLockPessimistic
        For A = 1 To rs.RecordCount
            RS1.AddNew
            For b = 0 To RS1.Fields.count - 1
                If RS1(b).Type = 133 Then
                    RS1(b) = Date
                ElseIf RS1(b).Type = adChar Then
                    RS1(b) = ""
                Else
                    RS1(b) = 0
                End If
            Next
            RS1!Sch_Idy = Trim(rs(0))
            RS1!prd_idy = Trim(rs(1))
            RS1!fre_stk = Val(rs(2) & "")
            RS1!sal_stk = Trim(rs(3) & "")
            RS1!bas_typ = Trim(rs(4))
            RS1!SCH_AMT = Val(rs(5) & "")
            RS1.Update
            rs.MoveNext
        Next
        OpenReport CURDIR & "\CR.RP1", "Schemes Report Segment", 0, "", 6
    ElseIf Me.Tag = "Claims Report SKU" Then
        OpenReport CURDIR & "\CR.RP1", "Schemes n Free Report ", 0, "", 6
    End If
End Sub
Public Sub UPDATEDPS()
Dim RSSMN As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
Dim DT1 As Date
Dim DT2 As Date
Dim DAYCOUNT(6) As Integer
DT1 = DTPicker1
DT2 = DTPicker2
con.Execute "DELETE FROM RP1"
RSSMN.Open "SELECT * FROM SMN", con, adOpenKeyset
Do While DT1 <= DT2
    If Format(DT1, "DDDD") = "Monday" Then
        DAYCOUNT(0) = Val(DAYCOUNT(0)) + 1
    ElseIf Format(DT1, "DDDD") = "Tuesday" Then
        DAYCOUNT(1) = Val(DAYCOUNT(1)) + 1
    ElseIf Format(DT1, "DDDD") = "Wednesday" Then
        DAYCOUNT(2) = Val(DAYCOUNT(2)) + 1
    ElseIf Format(DT1, "DDDD") = "Thursday" Then
        DAYCOUNT(3) = Val(DAYCOUNT(3)) + 1
    ElseIf Format(DT1, "DDDD") = "Friday" Then
        DAYCOUNT(4) = Val(DAYCOUNT(4)) + 1
    ElseIf Format(DT1, "DDDD") = "Saturday" Then
        DAYCOUNT(5) = Val(DAYCOUNT(5)) + 1
    ElseIf Format(DT1, "DDDD") = "Sunday" Then
        DAYCOUNT(6) = Val(DAYCOUNT(6)) + 1
    End If
    DT1 = DateAdd("D", 1, DT1)
Loop
If rs.State = 1 Then rs.Close
rs.Open "SELECT COUNT(*),SMN_IDY,DAY_IDY FROM RDY GROUP BY DAY_IDY,SMN_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To RSSMN.RecordCount
    CTR = 0
    For dayid = 0 To 6
        rs.filter = "SMN_IDY LIKE '" & Trim(RSSMN!SMN_IDY) & "' AND DAY_IDY =" & dayid + 1
        If rs.RecordCount > 0 Then CTR = CTR + (Val(rs(0)) * DAYCOUNT(dayid))
    Next
    ADDRP1 RSSMN!SMN_IDY, Val(CTR)
    RSSMN.MoveNext
Next
If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
For A = 1 To RSSMN.RecordCount
    If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "SELECT sum(tot_pcs),sub_unt FROM INV WHERE INV_IDY IN(SELECT INV_IDY FROM INL WHERE inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "')AND SMN_IDY LIKE '" & Trim(RSSMN!SMN_IDY) & "' AND TXN_TYP LIKE 'SAL') AND STK_TYP IN('Salable','Free') GROUP BY SUB_UNT", con, adOpenKeyset, adLockPessimistic
    TOTCASES = 0
    For b = 1 To RSINV.RecordCount
        M = STM(RSINV(0), Val(RSINV(1)))
        TOTCASES = Val(TOTCASES) + Round(Int(M) + Val(((Val(M) - Int(M)) * 1000) / Val(RSINV!SUB_UNT)), 2)
        RSINV.MoveNext
    Next
    con.Execute "UPDATE RP1 SET NUM_ONE= " & Round(TOTCASES, 0) & " WHERE CHR_TEN LIKE '" & Trim(RSSMN!SMN_IDY) & "'"
    RSSMN.MoveNext
Next

If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT COUNT(*),SMN_IDY FROM INL WHERE inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') AND TXN_TYP LIKE 'SAL' GROUP BY SMN_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To RSINV.RecordCount
    con.Execute "UPDATE RP1 SET NUM_THR= " & Val(RSINV(0) & "") & " WHERE CHR_TEN LIKE '" & Trim(RSINV!SMN_IDY) & "'"
    RSINV.MoveNext
Next

If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
For A = 1 To RSSMN.RecordCount
    For b = 1 To 7
        If b = 1 Then
            Weekname = "Monday"
        ElseIf b = 2 Then
            Weekname = "Tuesday"
        ElseIf b = 3 Then
            Weekname = "Wednesday"
        ElseIf b = 4 Then
            Weekname = "Thursday"
        ElseIf b = 5 Then
            Weekname = "Friday"
        ElseIf b = 6 Then
            Weekname = "Saturday"
        ElseIf b = 7 Then
            Weekname = "Sunday"
        End If
        If RSINV.State = 1 Then RSINV.Close
        RSINV.Open "SELECT * FROM INL WHERE inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') AND TXN_TYP LIKE 'SAL' and cdow(inv_dat)='" & Weekname & "' AND CST_IDY NOT IN(SELECT ACT_IDY FROM RDY WHERE DAY_IDY =" & Val(b) & " AND SMN_IDY LIKE '" & Trim(RSSMN!SMN_IDY) & "') AND SMN_IDY LIKE '" & Trim(RSSMN!SMN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        con.Execute "UPDATE RP1 SET NUM_FOR= NUM_FOR+" & Val(RSINV.RecordCount & "") & " WHERE CHR_TEN LIKE '" & Trim(RSSMN!SMN_IDY) & "'"
    Next
    RSSMN.MoveNext
Next

If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
For A = 1 To RSSMN.RecordCount
    If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "SELECT PRD_IDY FROM INV WHERE INV_IDY IN(SELECT INV_IDY FROM INL WHERE inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "')AND SMN_IDY LIKE '" & Trim(RSSMN!SMN_IDY) & "' AND TXN_TYP LIKE 'SAL') AND STK_TYP IN('Salable','Free') GROUP BY PRD_IDY", con, adOpenKeyset, adLockPessimistic
'    MsgBox Round(RSINV.RecordCount / GetValue("COUNT(*)", "PRD"), 2)
    con.Execute "UPDATE RP1 SET NUM_FIV= " & Round(RSINV.RecordCount / GetValue("COUNT(*)", "PRD WHERE STK_TYP = 'Salable'"), 2) & " WHERE CHR_TEN LIKE '" & Trim(RSSMN!SMN_IDY) & "'"
    RSSMN.MoveNext
Next
OpenReport CURDIR & "\SAL_PRG.RP1", "Daily Productivity Summary", 0, "", 6
End Sub

Private Sub ADDRP1(smnid As String, CTR As Long)
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic
rs.AddNew
rs(0) = ""
rs(1) = ""
rs(2) = ""
rs(3) = ""
rs(4) = ""
rs(5) = ""
rs(6) = ""
rs(7) = ""
rs(8) = ""
rs(9) = smnid
rs(10) = 0
rs(11) = CTR
rs(12) = 0
rs(13) = 0
rs(14) = 0
rs(15) = 0
rs(16) = 0
rs(17) = 0
rs(18) = 0
rs(19) = 0
rs.Update
End Sub
