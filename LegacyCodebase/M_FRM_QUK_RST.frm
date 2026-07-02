VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_QUK_RST 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DB Quick Restore"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   Icon            =   "M_FRM_QUK_RST.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   113
      TabIndex        =   1
      Top             =   2320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Quick DB Restore Date Range"
      Height          =   1575
      Left            =   113
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1275
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   75104259
         CurrentDate     =   39392
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1275
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   75104259
         CurrentDate     =   39392
         MinDate         =   36526
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   195
         Left            =   510
         TabIndex        =   5
         Top             =   1050
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   450
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1448
      TabIndex        =   6
      Top             =   1880
      Width           =   1215
   End
End
Attribute VB_Name = "M_FRM_QUK_RST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COM_CTR As Double
Dim CTR As Double
Dim P2B_CTR As Double
Dim REC_CTR As Double
Private Sub Command1_Click()
If CHECKTB("DIR.DBF,ACT.DBF,GRP.DBF,LOC.DBF,CAT.DBF,SMN.DBF,MAR.DBF,SCH.DBF,WEK.DBF,YER.DBF,LOG.DBF,PRD.DBF,PGR.DBF,BRD.DBF,TGT.DBF,USR.DBF,IDY.DBF,SHL.DBF,AST.DBF,TYP.DBF,RDY.DBF,SID.DBF,SIC.DBF") = 1 Then
    MsgBox UCase(TNAME) & " missing in " & Me.Caption
    Exit Sub
End If
If CHECKTB("INV.DBF,INL.DBF,PNV.DBF,PNL.DBF,VAN.DBF,SNV.DBF,CUR.DBF,IRT.DBF,LNV.DBF,VCH.DBF,ORD.DBF") = 1 Then
    MsgBox UCase(TNAME) & " missing in " & Me.Caption
    Exit Sub
End If
If CHECKTB("BNV.DBF,BNL.DBF,XNV.DBF,XNL.DBF,VBK.DBF,CNV.DBF,LDS.DBF,SBK.DBF,EFF.DBC,EFF.DCT") = 1 Then
    MsgBox UCase(TNAME) & " missing in " & Me.Caption
    Exit Sub
End If
' Verify Important Tables Availability in EFF Database
Command1.Enabled = False
COM_CTR = 0
CTR = 15
Me.MousePointer = vbHourglass
ProgressBar1.Visible = True

If Me.Tag = "XL_SAL" Then
    XLSalesRegister
ElseIf Me.Tag = "XL_VAN" Then
    XLVanRegister
Else
'inv_dat >= ctod('" & DTPicker1 & "') AND INV_DAT <=CTOD('" & DTPicker2 & "') GROUP BY SCH_IDY) and sch_typ not in('Pack Size','Pack Type','Brand','Segment','Category') GROUP BY SCH_IDY,B_PRD_IDY,TXN_TYP ", CON, adOpenKeyset, adLockPessimistic
CheckShutDownStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'")
If Trim(CheckShutDownStatus) = "IMPROPER SHUTDOWN" Then
    Dim RST_RSET As New ADODB.Recordset
    If RST_RSET.State = 1 Then RST_RSET.Close
    RST_RSET.Open ("SELECT * FROM BNV WHERE INV_DAT >= ctod('" & DTPicker1 & "') AND INV_DAT <= ctod('" & DTPicker2 & "') GROUP BY INV_IDY"), con, adOpenKeyset, adLockPessimistic
    If RST_RSET.RecordCount > 0 Then
        While Not RST_RSET.EOF
'            TxnMoveBackup "BNV", "INV", Trim(RST_RSET!INV_IDY), Trim(RST_RSET!TXN_TYP)
'            TxnMoveBackup "CNV", "LNV", Trim(RST_RSET!INV_IDY), Trim(RST_RSET!TXN_TYP)
            RST_RSET.MoveNext
        Wend
    End If
    INC_PCB
    If RST_RSET.State = 1 Then RST_RSET.Close
    RST_RSET.Open ("SELECT * FROM XNV WHERE INV_DAT >= ctod('" & DTPicker1 & "') AND INV_DAT <= ctod('" & DTPicker2 & "') GROUP BY INV_IDY"), con, adOpenKeyset, adLockPessimistic
    If RST_RSET.RecordCount > 0 Then
        While Not RST_RSET.EOF
'            TxnMoveBackup "XNV", "PNV", Trim(RST_RSET!INV_IDY), Trim(RST_RSET!TXN_TYP)
            RST_RSET.MoveNext
        Wend
    End If
    INC_PCB
    If RST_RSET.State = 1 Then RST_RSET.Close
    RST_RSET.Open ("SELECT * FROM BNL WHERE INV_DAT >= ctod('" & DTPicker1 & "') AND INV_DAT <= ctod('" & DTPicker2 & "') GROUP BY INV_IDY"), con, adOpenKeyset, adLockPessimistic
    If RST_RSET.RecordCount > 0 Then
        While Not RST_RSET.EOF
'            TxnMoveBackup "BNL", "INL", Trim(RST_RSET!INV_IDY), Trim(RST_RSET!TXN_TYP)
            RST_RSET.MoveNext
        Wend
    End If
    INC_PCB
    If RST_RSET.State = 1 Then RST_RSET.Close
    RST_RSET.Open ("SELECT * FROM XNL WHERE INV_DAT >= ctod('" & DTPicker1 & "') AND INV_DAT <= ctod('" & DTPicker2 & "') GROUP BY INV_IDY"), con, adOpenKeyset, adLockPessimistic
    If RST_RSET.RecordCount > 0 Then
        While Not RST_RSET.EOF
'            TxnMoveBackup "XNL", "PNL", Trim(RST_RSET!INV_IDY), Trim(RST_RSET!TXN_TYP)
            RST_RSET.MoveNext
        Wend
    End If
    INC_PCB
    If RST_RSET.State = 1 Then RST_RSET.Close
    RST_RSET.Open ("SELECT * FROM VBK WHERE INV_DAT >= ctod('" & DTPicker1 & "') AND INV_DAT <= ctod('" & DTPicker2 & "') GROUP BY INV_IDY"), con, adOpenKeyset, adLockPessimistic
    If RST_RSET.RecordCount > 0 Then
        While Not RST_RSET.EOF
'            TxnMoveBackup "VBK", "VAN", Trim(RST_RSET!INV_IDY), ""
            RST_RSET.MoveNext
        Wend
    End If
    INC_PCB
    If RST_RSET.State = 1 Then RST_RSET.Close
    RST_RSET.Open ("SELECT * FROM SBK WHERE INV_DAT >= ctod('" & DTPicker1 & "') AND INV_DAT <= ctod('" & DTPicker2 & "') GROUP BY INV_IDY"), con, adOpenKeyset, adLockPessimistic
    If RST_RSET.RecordCount > 0 Then
        While Not RST_RSET.EOF
'            TxnMoveBackup "SBK", "SNV", Trim(RST_RSET!INV_IDY), Trim(RST_RSET!TXN_TYP)
            RST_RSET.MoveNext
        Wend
    End If
    RST_RSET.Close
End If
'*
con.Execute "DELETE FROM INV WHERE INV_DAT = ctod('" & "  /  /    " & "')", "D"
INC_PCB
con.Execute "DELETE FROM PNV WHERE INV_DAT = ctod('" & "  /  /    " & "')", "D"
INC_PCB
con.Execute "DELETE FROM INL WHERE INV_DAT = ctod('" & "  /  /    " & "')", "D"
INC_PCB
con.Execute "DELETE FROM PNL WHERE INV_DAT = ctod('" & "  /  /    " & "')", "D"
INC_PCB
con.Execute "DELETE FROM VAN WHERE INV_DAT = ctod('" & "  /  /    " & "')", "D"
INC_PCB
con.Execute "DELETE FROM SNV WHERE INV_DAT = ctod('" & "  /  /    " & "')", "D"
INC_PCB
con.Execute "DELETE FROM SNV WHERE TXN_DAT = ctod('" & "  /  /    " & "')", "D"
INC_PCB
con.Execute "DELETE FROM LNV WHERE LEN(TRIM(INV_IDY)) = 0", "D"
INC_PCB
' * Initialize Intermediate Data Hold Files
con.Execute "DELETE FROM BNV", "D"
con.Execute "DELETE FROM BNL", "D"
con.Execute "DELETE FROM XNV", "D"
con.Execute "DELETE FROM XNL", "D"
con.Execute "DELETE FROM CNV", "D"
con.Execute "DELETE FROM VBK", "D"
con.Execute "DELETE FROM CBK", "D"
con.Execute "DELETE FROM SBK", "D"
' * End of Intermediate Hold Files Deletion
INC_PCB
'* Routine to Verify and Fill INL DBF from INV DBF for partial record write loss
REC_CTR = 0
ADD_CTR = 0
Dim INV_RSET As New ADODB.Recordset
Dim INL_RSET As New ADODB.Recordset
Dim PNL_RSET As New ADODB.Recordset
If INV_RSET.State = 1 Then INV_RSET.Close
INV_RSET.Open ("SELECT DISTINCT * FROM INV GROUP BY INV_IDY, TXN_TYP ORDER BY TXN_TYP, INV_DAT"), con, adOpenKeyset, adLockPessimistic
P2B_CTR = INV_RSET.RecordCount
If INV_RSET.RecordCount > 0 Then
    For A = 1 To INV_RSET.RecordCount
        If INL_RSET.State = 1 Then INL_RSET.Close
        INL_RSET.Open "SELECT * FROM INL WHERE INV_IDY LIKE '" & Trim(INV_RSET!INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If INL_RSET.RecordCount = 0 Then
            INL_RSET.AddNew
            INL_RSET("INV_IDY") = Trim(INV_RSET!INV_IDY)
            INL_RSET("INV_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("TXN_TYP") = Trim(INV_RSET!TXN_TYP)
            INL_RSET("DUE_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("WBL_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("LRN_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("DCH_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("VCH_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("SYS_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("ORD_IDY") = Trim(INV_RSET!ORD_IDY)
            INL_RSET("CST_IDY") = "0000000001"
            INL_RSET("SCH_IDY") = ""
            INL_RSET("DCH_IDY") = ""
            INL_RSET("VCH_IDY") = ""
            INL_RSET("NET_AMT") = 0#
            INL_RSET.Update
            ADD_CTR = ADD_CTR + 1
        End If
        INC_P2B
        INV_RSET.MoveNext
    Next
End If
Me.MousePointer = vbDefault
If ADD_CTR > 0 Then
    MsgBox "Added " & Trim(str(ADD_CTR)) & " missing records ", vbInformation, "DAS Version 4.0"
End If
Me.Refresh
'* End of Routine to Verify and Fill INL DBF from INV DBF for partial record write loss

'* Routine to Verify and Fill INV DBF from INL DBF for partial record write loss
REC_CTR = 0
ADD_CTR = 0

If INL_RSET.State = 1 Then INL_RSET.Close
INL_RSET.Open ("SELECT DISTINCT * FROM INL GROUP BY INV_IDY, TXN_TYP ORDER BY TXN_TYP, INV_DAT"), con, adOpenKeyset, adLockPessimistic
P2B_CTR = INL_RSET.RecordCount
If INL_RSET.RecordCount > 0 Then
    For A = 1 To INL_RSET.RecordCount
        If INV_RSET.State = 1 Then INV_RSET.Close
        INV_RSET.Open "SELECT * FROM INV WHERE INV_IDY LIKE '" & Trim(INL_RSET!INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If INV_RSET.RecordCount = 0 Then
            INL_RSET("NET_AMT") = 0#
            INL_RSET.Update
            INV_RSET.AddNew
            INV_RSET("INV_IDY") = Trim(INL_RSET!INV_IDY)
            INV_RSET("INV_DAT") = Trim(INL_RSET!INV_DAT)
            INV_RSET("TXN_TYP") = Trim(INL_RSET!TXN_TYP)
            INV_RSET("PRD_IDY") = Trim("0000000001")
            INV_RSET("BAT_IDY") = Trim("BAT0000001")
            INV_RSET("LST_RAT") = 0#
            INV_RSET("LST_DIS") = 0#
            INV_RSET("PRD_QTY") = 0#
            INV_RSET("PRD_RAT") = 0#
            INV_RSET("RET_QTY") = 0#
            INV_RSET("RAT_APP") = ""
            INV_RSET("CPN_RAT") = 0#
            INV_RSET("CPN_AMT") = 0#
            INV_RSET("ADD_DIS") = 0#
            INV_RSET("ADD_AMT") = 0#
            INV_RSET("TRD_DIS") = 0#
            INV_RSET("DIS_AMT") = 0#
            INV_RSET("TRD_TAX") = 0#
            INV_RSET("TAX_AMT") = 0#
            INV_RSET("COL_AMT") = 0#
            INV_RSET("STK_UPD") = "Y"
            INV_RSET("FST_FLG") = ""
            INV_RSET("OPN_QTY") = 0#
            INV_RSET("ISS_QTY") = 0#
            INV_RSET("TOT_QTY") = 0#
            INV_RSET("SCH_AMT") = 0#
            INV_RSET("ORD_IDY") = Trim(INL_RSET!ORD_IDY)
            INV_RSET("PRD_UOM") = 0#
            INV_RSET("SUB_UNT") = 0#
            INV_RSET("PRD_UNT") = 0#
            INV_RSET("PRD_PCS") = 0#
            INV_RSET("TOT_PCS") = 0#
            INV_RSET("STK_TYP") = "Salable"
            INV_RSET("SCH_IDY") = ""
            INV_RSET("INC_TTL") = 0#
            INV_RSET.Update
            ADD_CTR = ADD_CTR + 1
        End If
        INC_P2B

        INL_RSET.MoveNext
    Next
End If
Me.MousePointer = vbDefault
If ADD_CTR > 0 Then
    MsgBox "Added " & Trim(str(ADD_CTR)) & " missing records ", vbInformation, "DAS Version 4.0"
End If
Me.Refresh
'* End of Routine to Verify and Fill INV DBF from INL DBF for partial record write loss


'* Routine to Verify and Fill PNL DBF from PNV DBF for partial record write loss
REC_CTR = 0
ADD_CTR = 0
If INV_RSET.State = 1 Then INV_RSET.Close
INV_RSET.Open ("SELECT DISTINCT * FROM PNV GROUP BY INV_IDY, TXN_TYP ORDER BY TXN_TYP, INV_DAT"), con, adOpenKeyset, adLockPessimistic
P2B_CTR = INV_RSET.RecordCount
If INV_RSET.RecordCount > 0 Then
    For A = 1 To INV_RSET.RecordCount
        If INL_RSET.State = 1 Then INL_RSET.Close
        INL_RSET.Open "SELECT * FROM PNL WHERE INV_IDY LIKE '" & Trim(INV_RSET!INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If INL_RSET.RecordCount = 0 Then
            INL_RSET.AddNew
            INL_RSET("INV_IDY") = Trim(INV_RSET!INV_IDY)
            INL_RSET("INV_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("TXN_TYP") = Trim(INV_RSET!TXN_TYP)
            INL_RSET("DUE_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("WBL_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("LRN_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("DCH_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("VCH_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("SYS_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("ORD_IDY") = Trim(INV_RSET!ORD_IDY)
            INL_RSET("CST_NME") = ""
            INL_RSET("CST_AD1") = ""
            INL_RSET("CST_AD2") = ""
            INL_RSET("cst_ad3") = ""
            INL_RSET("cst_ad4") = ""
            INL_RSET("cst_zip") = ""
            INL_RSET("Act_idy") = ""
            INL_RSET("dch_idy") = ""
            INL_RSET("csd_pcg") = 0#
            INL_RSET("csd_amt") = 0#
            INL_RSET("spd_pcg") = 0#
            INL_RSET("spd_amt") = 0#
            INL_RSET("tax_pcg") = 0#
            INL_RSET("tax_amt") = 0#
            INL_RSET("exd_pcg") = 0#               'SURCHARGE PCG
            INL_RSET("EXD_AMT") = 0#               'SURCHARGE AMT
            INL_RSET("pkg_cst") = 0#               'AMOUNT (+)
            INL_RSET("FRT_AMT") = 0#          'AMOUNT (-)
            INL_RSET("LRN_IDY") = ""
            INL_RSET("LRN_AMT") = 0#
            INL_RSET("LRN_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("WBL_IDY") = ""
            INL_RSET("WBL_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("TXN_RMK") = ""
            INL_RSET("TTL_CAS") = 0#
            INL_RSET("txn_typ") = "SAL"
            INL_RSET("grs_amt") = 0#
            INL_RSET("net_amt") = 0#
            INL_RSET("crt_day") = 0#
            INL_RSET("due_dat") = Trim(INV_RSET!INV_DAT)
            INL_RSET("rou_off") = 0#
            INL_RSET("ROU_IDY") = ""
            INL_RSET("TXN_TPT") = ""
            INL_RSET("TXN_RM1") = ""
            INL_RSET("TXN_RM2") = ""
            INL_RSET("USR_NME") = ""
            INL_RSET("SYS_DAT") = Trim(INV_RSET!INV_DAT)
            INL_RSET("SYS_TIM") = Format(Time, "HH:MM:SS")
            INL_RSET("SCH_CHK") = "Y"
            INL_RSET("SRN_AMT") = 0#
            INL_RSET!vch_idy = ""
            INL_RSET!Sch_Idy = ""
            INL_RSET!cst_gst = ""
            INL_RSET!cst_cst = ""
            INL_RSET!MAR_IDY = ""
            INL_RSET!mar_nme = ""
            INL_RSET!SMN_NME = ""
            INL_RSET!loc_IDY = ""
            INL_RSET!loc_nme = ""
            INL_RSET!cat_IDY = ""
            INL_RSET!cat_nme = ""
            INL_RSET!exd_cal = ""
            INL_RSET!SMN_IDY = ""
            INL_RSET!frt_typ = ""
            INL_RSET!tax_act = ""
            INL_RSET!frt_act = ""
            INL_RSET!des_mod = ""
            INL_RSET!sto_rmk = ""
            INL_RSET!vch_dat = Trim(INV_RSET!INV_DAT)
            INL_RSET!ins_nme = ""
            INL_RSET!ins_adr = ""
            INL_RSET!sys_rm1 = ""
            INL_RSET!sys_rm2 = ""
            INL_RSET!sys_rm3 = ""
            INL_RSET!sys_rm4 = ""
            INL_RSET!ins_gst = ""
            INL_RSET!ins_cst = ""
            INL_RSET!FRM_DAT = Trim(INV_RSET!INV_DAT)
            INL_RSET!TO_DAT = Trim(INV_RSET!INV_DAT)
            INL_RSET!rou_act = ""
            INL_RSET!dch_dat = Trim(INV_RSET!INV_DAT)
            INL_RSET!DIS_CHK = ""
            INL_RSET("CST_IDY") = "0000000001"
            INL_RSET("SCH_IDY") = ""
            INL_RSET("DCH_IDY") = ""
            INL_RSET("VCH_IDY") = ""
            INL_RSET.Update
            ADD_CTR = ADD_CTR + 1
        End If
        INC_P2B
        INV_RSET.MoveNext
    Next
End If
Me.MousePointer = vbDefault
If ADD_CTR > 0 Then
    MsgBox "Added " & Trim(str(ADD_CTR)) & " missing records ", vbInformation, "DAS Version 4.0"
End If
'* End of Routine to Verify and Fill PNL DBF from PNV DBF for partial record write loss
Me.Refresh
' CON.Execute "DELETE FROM INL WHERE INV_IDY = '101TEMP'", "D" ' Command for verification ie, test case to check below code written by PNR

'* Routine to Verify Records Addition into INL DBF from INV DBF for partial record write loss
REC_CTR = 0
ADD_CTR = 0
If INV_RSET.State = 1 Then INV_RSET.Close
INV_RSET.Open ("SELECT DISTINCT * FROM INV GROUP BY INV_IDY, TXN_TYP ORDER BY TXN_TYP, INV_DAT"), con, adOpenKeyset, adLockPessimistic
P2B_CTR = INV_RSET.RecordCount
If INV_RSET.RecordCount > 0 Then
    For A = 1 To INV_RSET.RecordCount
        If INL_RSET.State = 1 Then INL_RSET.Close
        INL_RSET.Open "SELECT * FROM INL WHERE INV_IDY LIKE '" & Trim(INV_RSET!INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If INL_RSET.RecordCount = 0 Then
            ADD_CTR = ADD_CTR + 1
        End If
        INC_P2B
        INV_RSET.MoveNext
    Next
End If
Me.MousePointer = vbDefault
If ADD_CTR > 0 Then
    Dim DIR_RS As New ADODB.Recordset
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'"), con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveFirst
        DIR_RS("DIR_MSG") = "IMPROPER SHUTDOWN"
        DIR_RS("DIR_DAT") = Date
        DIR_RS("DIR_ALS") = ""
        DIR_RS.Update
        DIR_RS.Close
    End If
    MsgBox "Failed " & Trim(str(ADD_CTR)) & " records addition! Improper shutdown was very severe! " & Chr(13) & "Check your IT Infrastructure and run 'Repair' Utility", vbInformation, "DAS Version 4.0"
    Unload Me
    Set INV_RSET = Nothing
    Set INL_RSET = Nothing
    Set PNL_RSET = Nothing
    Set Form = Nothing
    con.Close
    PNL.Close
    Set con = Nothing
    Set PNL = Nothing
    Shell App.Path & "\IndexData.EXE", vbNormalFocus
    MsgBox "Restart DAS Application now! ", vbInformation, "DAS Version 4.0"
    End
    End If
Me.Refresh
'* End of Routine to Verify Records Addition into INL DBF from INV DBF for partial record write loss
Unload Me
Set INV_RSET = Nothing
Set INL_RSET = Nothing
Set PNL_RSET = Nothing
Set Form = Nothing
End If
End Sub
Private Sub DTPicker1_GotFocus()
On Error Resume Next
If Me.Tag = "XL_SAL" Or Me.Tag = "XL_VAN" Then
'    DTPicker1.value = Date
Else
    CheckShutDownDate = GetValue("DIR_DAT ", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'")
    DTPicker1.Value = CheckShutDownDate
End If
End Sub
Private Sub DTPicker2_GotFocus()
If Me.Tag = "XL_SAL" Or Me.Tag = "XL_VAN" Then
Else
    CheckShutDownDate = GetValue("DIR_DAT ", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'")
    DTPicker2.Value = CheckShutDownDate
End If
End Sub
Private Sub Form_Activate()
If Me.Tag = "XL_SAL" Then
    Me.Caption = "Sales Register SoftCopy"
    Frame1.Caption = "Select SalesRegister Date Range"
    Me.Refresh
ElseIf Me.Tag = "XL_VAN" Then
    Me.Caption = "Van Register SoftCopy"
    Frame1.Caption = "Select VanRegister Date Range"
    Me.Refresh
End If
SETYEAROPEN M_FRM_QUK_RST
DTPicker1.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
Set Form = Nothing
End Sub
Private Sub DTPicker1_KeyDown(K As Integer, s As Integer)
If K = 13 Then DTPicker2.SetFocus
End Sub
Private Sub DTPicker2_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    If Me.Tag = "XL_SAL" Or Me.Tag = "XL_VAN" Then
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        If Pending_RS.RecordCount = 0 Then
            MsgBox "There are no Day Closure Dates BETWEEN THE SELECTED DATES for report generation!", vbInformation, "DAS Version 4.0"
            Set Pending_RS = Nothing
            Unload Me
            Exit Sub
        End If
    End If
    Command1.SetFocus
ElseIf K = 37 And s = 1 Then
    DTPicker1.SetFocus
End If
End Sub
Private Sub DTPicker2_LostFocus()
If DTPicker1 > DTPicker2 Then
        MsgBox "'From' Date Should Not Be Greater Than 'To' Date!", vbInformation, "DAS Version 4.0"
        DTPicker1.SetFocus
        Exit Sub
    End If
End Sub
Private Sub XLVanRegister()

On Error GoTo errh
REC_CTR = 0
ADD_CTR = 0
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
' * Start of Distributor ID Identification
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
            Exit Sub
        End If
        M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
        If Len(Trim(M_WDS_SRL)) < 3 Then
             While Len(Trim(M_WDS_SRL)) < 3
                  M_WDS_SRL = "0" & M_WDS_SRL
             Wend
        End If
    End If
        
' * End of Distributor ID Identification

        
    ' * : Beginning of _VanReg.XLS - XLSalesRegister Creation Code for XL Data Export
    
    Me.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    M_FIL_NME = "C:\DEX\" & Trim(M_WDS_IDY) & "_VanReg.XLS"
       
    ' Title Line Record for Excel Data Export File
    xlSheet.Cells(1, 1).Value = "Salesman Name"
    xlSheet.Cells(1, 2).Value = "Van Number"
    xlSheet.Cells(1, 3).Value = "Van Location"
    xlSheet.Cells(1, 4).Value = "Capacity"
    xlSheet.Cells(1, 5).Value = "Route"
    xlSheet.Cells(1, 6).Value = "Date"
    xlSheet.Cells(1, 7).Value = "Day"
    xlSheet.Cells(1, 8).Value = "Challan Number"
    xlSheet.Cells(1, 9).Value = "Challan Refrce"
    xlSheet.Cells(1, 10).Value = "Total Trips"
    xlSheet.Cells(1, 11).Value = "Load-Out Cs"
    xlSheet.Cells(1, 12).Value = "Load-Out Rs"
    xlSheet.Cells(1, 13).Value = "Load-In Cs"
    xlSheet.Cells(1, 14).Value = "Load-In Rs"
    xlSheet.Cells(1, 15).Value = "Gross Sales Rs"
    xlSheet.Cells(1, 16).Value = "Free-Qty Cs"
    xlSheet.Cells(1, 17).Value = "Free-Qty Rs"
    xlSheet.Cells(1, 18).Value = "Net-Sales Cs"
    xlSheet.Cells(1, 19).Value = "Net-Sales Rs"
    xlSheet.Cells(1, 20).Value = "Discount"
    xlSheet.Cells(1, 21).Value = "Commission"
    xlSheet.Cells(1, 22).Value = "Expenses"
    xlSheet.Cells(1, 23).Value = "SalesTax"
    xlSheet.Cells(1, 24).Value = "Dep Refunds"
    xlSheet.Cells(1, 25).Value = "Dep Receipts"
    xlSheet.Cells(1, 26).Value = "Old Dues"
    xlSheet.Cells(1, 27).Value = "CreditSales"
    xlSheet.Cells(1, 28).Value = "Cash"
    xlSheet.Cells(1, 29).Value = "Cheques"
    xlSheet.Cells(1, 30).Value = "Short-Excess"
    xlSheet.Cells(1, 31).Value = "Cml Short-Excess"
    xlSheet.Cells(1, 32).Value = "ChallanStatus"
    xlSheet.Cells(1, 33).Value = "Total Variance"
    xlSheet.Cells(1, 34).Value = "1000s"
    xlSheet.Cells(1, 35).Value = "500s"
    xlSheet.Cells(1, 36).Value = "100s"
    xlSheet.Cells(1, 37).Value = "50s"
    xlSheet.Cells(1, 38).Value = "20s"
    xlSheet.Cells(1, 39).Value = "10s"
    xlSheet.Cells(1, 40).Value = "5s"
    xlSheet.Cells(1, 41).Value = "2s"
    xlSheet.Cells(1, 42).Value = "1s"
    xlSheet.Cells(1, 43).Value = "Change"
    xlSheet.Cells(1, 44).Value = "Glass Net-Load"
    xlSheet.Cells(1, 45).Value = "ShellsNet-Load"
    xlSheet.Cells(1, 46).Value = "Glass COL"
    xlSheet.Cells(1, 47).Value = "ShellsCOL"
    xlSheet.Cells(1, 48).Value = "GlassROL"
    xlSheet.Cells(1, 49).Value = "ShellsROL"
    xlSheet.Cells(1, 50).Value = "GlassGOD"
    xlSheet.Cells(1, 51).Value = "ShellsGOD"
    xlSheet.Cells(1, 52).Value = "GlassROD"
    xlSheet.Cells(1, 53).Value = "ShellsROD"
    xlSheet.Cells(1, 54).Value = "Glass Collected"
    xlSheet.Cells(1, 55).Value = "ShellsCollected"
    xlSheet.Cells(1, 56).Value = "Glass Variance"
    xlSheet.Cells(1, 57).Value = "ShellsVariance"
    xlSheet.Cells(1, 58).Value = "SKUs Count"
    xlSheet.Cells(1, 59).Value = "Sch-Calls"
    xlSheet.Cells(1, 60).Value = "UnSch-Calls"
    xlSheet.Cells(1, 61).Value = "ProdctveCalls"
    xlSheet.Cells(1, 62).Value = "Sales Vs LoadOut"
    xlSheet.Cells(1, 63).Value = "GrossValue Vs LoadOut"
    xlSheet.Cells(1, 64).Value = "CreditSale+Shortage Vs NetSale"
    xlSheet.Cells(1, 65).Value = "Veh Utilz%"
    xlSheet.Cells(1, 66).Value = "OneWaySale%"
    xlSheet.Cells(1, 67).Value = "RGB Sale%"
    xlSheet.Cells(1, 68).Value = "Empts Retrn% of RGB Sale"
    xlSheet.Cells(1, 69).Value = "Strike Rate%"
    xlSheet.Cells(1, 70).Value = "Index"
    xlSheet.Cells(1, 71).Value = "Avg Drop"
    
    Dim J As Integer
    Dim CCnt As Integer
    J = 0
    CCnt = 2
    ADD_CTR = 0
    Dim Rrs As New ADODB.Recordset
    Dim Category_RS As New ADODB.Recordset
    Dim Location_RS As New ADODB.Recordset
    Dim Market_RS As New ADODB.Recordset
    Dim Salesman_RS As New ADODB.Recordset
    Dim RDY_RS As New ADODB.Recordset
    Dim INV_RS As New ADODB.Recordset
    Dim SNV_RS As New ADODB.Recordset
    Dim CUR_RS As New ADODB.Recordset
    Dim DET_RS As New ADODB.Recordset
    Dim PRC_RS As New ADODB.Recordset
    
    If INV_RS.State = 1 Then INV_RS.Close
'    INV_RS.Open "SELECT INL.CST_IDY, INL.INV_DAT, PRD.PRD_PRN, PRD.BRD_NME, PRD.PRD_PCK, PRD.PRD_PKT, PGR.GRP_CAT, PRD.SUB_UNT, BRD.PUR_LPP, BRD.SAL_LPP, BRD.PRD_MRP, PRD.PRD_NME, SUM(INV.TOT_PCS), INV.BAT_IDY, INV.STK_TYP, INV.SCH_IDY, INL.FRT_AMT, INV.TXN_TYP, INL.ROU_IDY, INL.SMN_IDY FROM INV, BRD, PRD, PGR, INL  WHERE INV.TXN_TYP = 'SAL' AND INV.BAT_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY AND PRD.GRP_IDY = PGR.GRP_IDY AND INV.INV_IDY = INL.INV_IDY AND INV.INV_DAT >= CTOD('" & DTPicker1 & "') AND INV.INV_DAT <= CTOD('" & DTPicker2 & "') GROUP BY INL.CST_IDY, INV.BAT_IDY, INV.STK_TYP, INV.SCH_IDY, INV.TXN_TYP HAVING INV.TXN_TYP = 'SAL' ORDER BY INL.CST_IDY, INL.INV_DAT, PRD.PRD_PRN", CON, adOpenKeyset, adLockPessimistic
    INV_RS.Open "SELECT *  FROM VAN WHERE VAN.INV_DAT >= CTOD('" & DTPicker1 & "') AND VAN.INV_DAT <= CTOD('" & DTPicker2 & "') GROUP BY VAN.INV_IDY  ORDER BY VAN.INV_IDY", con, adOpenKeyset, adLockPessimistic
'    INV_RS.Open "SELECT *  FROM VAN WHERE VAN.INV_DAT >= CTOD('" & DTPicker1 & "') AND VAN.INV_DAT <= CTOD('" & DTPicker2 & "') ORDER BY VAN.INV_IDY", CON, adOpenKeyset, adLockPessimistic
    P2B_CTR = INV_RS.RecordCount
    While Not INV_RS.EOF
        If Rrs.State = 1 Then Rrs.Close
        Rrs.Open "select * from ACT where ACT_TYP = 'SAL' AND DBT_TYP = 'Y' AND ACT_IDY = '" & Trim(INV_RS("CST_IDY")) & "'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
        If Rrs.RecordCount > 0 Then
                Dim M_NET_SAL_INR, M_SHT_EXC, C_SHT_EXC, M_LOD_OUT, M_LOD_IN, M_FRE_QT As Double
                
                Dim L_CAT_NME As String
                Dim L_LOC_NME As String
                Dim L_MAR_NME As String
                Dim L_SMN_NME As String
                Dim L_ROU_IDY As String
                Dim L_DST_NME As String
                Dim L_MAR_GRD As String
                Dim R_SMN_NME As String
                Dim XL_ACT_IDY As String
                Dim D_MON, D_TUE, D_WED, D_THU, D_FRI, D_SAT, D_SUN As String
                
                
                M_LOD_OUT = 0#
                M_LOD_IN = 0#
                M_FRE_QT = 0#
                
                If Category_RS.State = 1 Then Category_RS.Close
                Category_RS.Open "SELECT * FROM CAT WHERE CAT_IDY = '" & Trim(Rrs("CAT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Category_RS.RecordCount > 0 Then
                      L_CAT_NME = Category_RS("CAT_NME")
                Else
                      L_CAT_NME = ""
                End If
                
                If Location_RS.State = 1 Then Location_RS.Close
                Location_RS.Open "SELECT * FROM LOC WHERE LOC_IDY = '" & Trim(Rrs("LOC_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Location_RS.RecordCount > 0 Then
                      L_LOC_NME = Location_RS("LOC_NME")
                Else
                      L_LOC_NME = ""
                End If
                
                If Market_RS.State = 1 Then Market_RS.Close
                Market_RS.Open "SELECT * FROM MAR WHERE MAR_IDY = '" & Trim(Rrs("MAR_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Market_RS.RecordCount > 0 Then
                      L_MAR_NME = Market_RS("MAR_NME")
                      L_DST_NME = Market_RS("MAR_PGR")
                      L_MAR_GRD = Market_RS("MAR_GRD")
                         ' Added for Margin Minder Data Extract by PNR
                      L_MAR_CSG = Market_RS("MAR_CSG")
                      L_MAR_CLU = Market_RS("MAR_CLU")
                      L_MAR_ROC = Market_RS("MAR_ROC")
                      L_MAR_TWC = Market_RS("MAR_TWC")
                      L_MAR_TWN = Market_RS("MAR_TWN")
                      L_MAR_DST = Market_RS("MAR_DST")
                      L_MAR_STA = Market_RS("MAR_STA")
                      L_MAR_COU = Market_RS("MAR_COU")
                Else
                      L_MAR_NME = ""
                      L_DST_NME = ""
                      L_MAR_GRD = ""
                         ' Added for Margin Minder Data Extract by PNR
                      L_MAR_CSG = ""
                      L_MAR_CLU = ""
                      L_MAR_ROC = ""
                      L_MAR_TWC = ""
                      L_MAR_TWN = ""
                      L_MAR_DST = ""
                      L_MAR_STA = ""
                      L_MAR_COU = ""
                End If
                
                   
                If RDY_RS.State = 1 Then RDY_RS.Close
                RDY_RS.Open "SELECT * FROM RDY ORDER BY DAY_IDY WHERE ACT_IDY = '" & Trim(Rrs("ACT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If RDY_RS.RecordCount > 0 Then RDY_RS.MoveFirst
                   
                   If RDY_RS.RecordCount > 0 Then
                      L_SMN_IDY = Trim(RDY_RS("SMN_IDY"))
                      L_ROU_IDY = Trim(RDY_RS("ROU_IDY"))
                   Else
                      L_SMN_IDY = ""
                      L_ROU_IDY = ""
                   End If
                   
                   While Not RDY_RS.EOF
                   If RDY_RS("ACT_IDY") = Trim(Rrs("ACT_IDY")) Then
                         If RDY_RS("DAY_IDY") = 1 Then
                             D_MON = "MONDAY"
                         ElseIf RDY_RS("DAY_IDY") = 2 Then
                             D_TUE = "TUESDAY"
                         ElseIf RDY_RS("DAY_IDY") = 3 Then
                             D_WED = "WEDNESDAY"
                         ElseIf RDY_RS("DAY_IDY") = 4 Then
                             D_THU = "THURSDAY"
                         ElseIf RDY_RS("DAY_IDY") = 5 Then
                             D_FRI = "FRIDAY"
                         ElseIf RDY_RS("DAY_IDY") = 6 Then
                             D_SAT = "SATURDAY"
                         ElseIf RDY_RS("DAY_IDY") = 7 Then
                             D_SUN = "SUNDAY"
                         End If
                   End If
                   RDY_RS.MoveNext
                   Wend
                   
                     If Salesman_RS.State = 1 Then Salesman_RS.Close
                     Salesman_RS.Open "SELECT * FROM SMN WHERE SMN_IDY = '" & Trim(L_SMN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
                     If Salesman_RS.RecordCount > 0 Then
                         L_SMN_NME = Salesman_RS("SMN_NME")
                     Else
                         L_SMN_NME = ""
                     End If
                     If Salesman_RS.State = 1 Then Salesman_RS.Close
                     Salesman_RS.Open "SELECT * FROM SMN WHERE SMN_IDY = '" & Trim(INV_RS(19)) & "'", con, adOpenKeyset, adLockPessimistic
                     If Salesman_RS.RecordCount > 0 Then
                         R_SMN_NME = Salesman_RS("SMN_NME")
                     Else
                         R_SMN_NME = ""
                     End If
                       
                     If CUR_RS.State = 1 Then CUR_RS.Close
                     CUR_RS.Open "SELECT * FROM CUR WHERE INV_IDY = '" & Trim(INV_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                       
                     If SNV_RS.State = 1 Then SNV_RS.Close
                     SNV_RS.Open "SELECT * FROM SNV WHERE INV_IDY = '" & Trim(INV_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                       
                    xlSheet.Cells(CCnt, 1).Value = Trim(L_SMN_NME)                          'Salesman Name
                    xlSheet.Cells(CCnt, 2).Value = Trim(Rrs("ACT_NME"))                     'Van Name - Number
                    xlSheet.Cells(CCnt, 3).Value = Trim(L_LOC_NME)                          'Van Location
                    xlSheet.Cells(CCnt, 4).Value = Trim(str(Rrs("VAN_CAP")))                'Van Capacity
                    xlSheet.Cells(CCnt, 5).Value = Trim(INV_RS("ROU_IDY"))                  'Van Route
                    xlSheet.Cells(CCnt, 6).Value = Format(INV_RS("INV_DAT"), "DD/MM/YYYY")  'Challan Date
                    If Trim(Weekday(INV_RS("INV_DAT"), vbUseSystemDayOfWeek)) = 1 Then
                        xlSheet.Cells(CCnt, 7).Value = "Sun"
                    ElseIf Trim(Weekday(INV_RS("INV_DAT"), vbUseSystemDayOfWeek)) = 2 Then
                        xlSheet.Cells(CCnt, 7).Value = "Mon"
                    ElseIf Trim(Weekday(INV_RS("INV_DAT"), vbUseSystemDayOfWeek)) = 3 Then
                        xlSheet.Cells(CCnt, 7).Value = "Tue"
                    ElseIf Trim(Weekday(INV_RS("INV_DAT"), vbUseSystemDayOfWeek)) = 4 Then
                        xlSheet.Cells(CCnt, 7).Value = "Wed"
                    ElseIf Trim(Weekday(INV_RS("INV_DAT"), vbUseSystemDayOfWeek)) = 5 Then
                        xlSheet.Cells(CCnt, 7).Value = "Thu"
                    ElseIf Trim(Weekday(INV_RS("INV_DAT"), vbUseSystemDayOfWeek)) = 6 Then
                        xlSheet.Cells(CCnt, 7).Value = "Fri"
                    ElseIf Trim(Weekday(INV_RS("INV_DAT"), vbUseSystemDayOfWeek)) = 7 Then
                        xlSheet.Cells(CCnt, 7).Value = "Sat"
                    End If
'                    xlSheet.Cells(CCnt, 7).value = Trim(WeekDay(INV_RS("INV_DAT")))         'Day of Date
                    xlSheet.Cells(CCnt, 8).Value = Trim(INV_RS("INV_IDY"))                  'Challan Number
                    xlSheet.Cells(CCnt, 9).Value = Trim(INV_RS("MNL_REF"))                  'Manual Reference
'                    xlSheet.Cells(CCnt, 10).value = Trim(STR(INV_RS("TRP_IDY")))            'Trips Total
                    
                    'Write another subquery select with sigma values to fill the following numeric data
                    If DET_RS.State = 1 Then DET_RS.Close
                    DET_RS.Open "SELECT SUM(TOT_QTY), SUM(RET_QTY), SUM(FRE_QTY), SUM(LKG_QTY), SUM(BRK_QTY), SUM(COL_AMT) FROM VAN WHERE INV_IDY = '" & Trim(INV_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                    If DET_RS.RecordCount > 0 Then
                        xlSheet.Cells(CCnt, 11).Value = Val(DET_RS(0))                          'Load Out Cases
                        xlSheet.Cells(CCnt, 13).Value = Val(DET_RS(1))                          'Load In Cases
                        xlSheet.Cells(CCnt, 16).Value = Val(DET_RS(2))                          'Free Qty Cases
                        xlSheet.Cells(CCnt, 18).Value = Val(DET_RS(0)) - (Val(DET_RS(1)) + Val(DET_RS(2)) + Val(DET_RS(3)) + Val(DET_RS(4)))
                        xlSheet.Cells(CCnt, 15).Value = Val(DET_RS(5))                          'Gross Sales INR
                    End If
                    If PRC_RS.State = 1 Then PRC_RS.Close
                    PRC_RS.Open "SELECT * FROM VAN, BRD WHERE VAN.BAT_IDY LIKE BRD.BAT_IDY AND INV_IDY = '" & Trim(INV_RS("INV_IDY")) & "' GROUP BY BRD.BAT_IDY", con, adOpenKeyset, adLockPessimistic
                    If PRC_RS.RecordCount > 0 Then
                        M_LOD_OUT = 0#
                        M_LOD_IN = 0#
                        M_FRE_QT = 0#
                        M_NET_SAL_INR = 0#
                        While Not PRC_RS.EOF
                            S_TAX = (Val(PRC_RS("S_SAL_VAT")) / Val(PRC_RS("SKU_SBU")))
                            M_LOD_OUT = Val(M_LOD_OUT) + ((MTS(Val(PRC_RS("TOT_QTY")), Val(PRC_RS("SKU_SBU")))) * (Val(PRC_RS("U_SAL_LPP"))))
                            M_LOD_IN = Val(M_LOD_IN) + ((MTS(Val(PRC_RS("RET_QTY")), Val(PRC_RS("SKU_SBU")))) * (Val(PRC_RS("U_SAL_LPP"))))
                            M_FRE_QT = Val(M_FRE_QT) + ((MTS(Val(PRC_RS("FRE_QTY")), Val(PRC_RS("SKU_SBU")))) * (Val(PRC_RS("U_SAL_LPP"))))
                            M_NET_SAL_INR = M_NET_SAL_INR + Val(PRC_RS("COL_AMT"))
                            xlSheet.Cells(CCnt, 10).Value = Trim(str(Val(PRC_RS("TRP_IDY"))))                   'Trips Total
                            PRC_RS.MoveNext
                        Wend
                        xlSheet.Cells(CCnt, 12).Value = Val(M_LOD_OUT)                                      'Load Out INR
                        xlSheet.Cells(CCnt, 14).Value = Val(M_LOD_IN)                                       'Load In INR
                        xlSheet.Cells(CCnt, 17).Value = Val(M_FRE_QT)                                       'Free Qty INR
                        M_NET_SAL_INR = Val(DET_RS(5)) - Val(M_FRE_QT)
                        'xlSheet.Cells(CCnt, 18).value = Val(DET_RS(0)) - Val(DET_RS(1))                     'Net Sales Cases
                        xlSheet.Cells(CCnt, 19).Value = convert(Val(M_NET_SAL_INR))
'                        convert (Val(M_LOD_OUT) - (Val(M_LOD_IN) + Val(M_FRE_QT)))   'Net Sales INR
                    End If
                    If PRC_RS.State = 1 Then PRC_RS.Close
                    PRC_RS.Open "SELECT DISTINCT COUNT(PRD_IDY) FROM VAN WHERE INV_IDY = '" & Trim(INV_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                    If PRC_RS.RecordCount > 0 Then
                        xlSheet.Cells(CCnt, 58).Value = Trim(Val(PRC_RS(0)))
                    End If
                    PRC_RS.Close
                   
                    xlSheet.Cells(CCnt, 20).Value = convert(Trim(Val(INV_RS("DIS_AMT"))))            'Discount Amount
                    xlSheet.Cells(CCnt, 21).Value = convert(Trim(Val(INV_RS("COM_AMT"))))            'Commission Amount
                    xlSheet.Cells(CCnt, 22).Value = convert(Trim(Val(INV_RS("EXP_AMT"))))            'Expense Amount
                    xlSheet.Cells(CCnt, 23).Value = convert(0#)                                      'Sales Tax
                    xlSheet.Cells(CCnt, 24).Value = convert(Trim(Val(INV_RS("DEP_REF"))))            'Deposit Refunds
                    xlSheet.Cells(CCnt, 25).Value = convert(Trim(Val(INV_RS("DEP_RCD"))))            'Deposit Received
                    xlSheet.Cells(CCnt, 26).Value = convert(Trim(Val(INV_RS("OLD_CRD"))))            'Old Credits
                    xlSheet.Cells(CCnt, 27).Value = convert(Trim(Val(INV_RS("CUR_CRD"))))            'Credit Sales
                    xlSheet.Cells(CCnt, 28).Value = convert(Trim(Val(INV_RS("CAS_AMT"))))            'Cash
                    xlSheet.Cells(CCnt, 29).Value = convert(Trim(Val(INV_RS("CHQ_AMT"))))            'Cheques
                    M_SHT_EXC = 0#
                    M_SHT_EXC = convert((Val(M_NET_SAL_INR) + Val(INV_RS("DEP_RCD")) + Val(INV_RS("OLD_CRD"))) - (Val(INV_RS("CAS_AMT")) + Val(INV_RS("DEP_REF")) + Val(INV_RS("CHQ_AMT")) + Val(INV_RS("DIS_AMT")) + Val(INV_RS("EXP_AMT")) + Val(INV_RS("CUR_CRD"))))
                    C_SHT_EXC = Val(C_SHT_EXC) + Val(M_SHT_EXC)
                    xlSheet.Cells(CCnt, 30).Value = convert(Val(M_SHT_EXC))                          'Short/Excess
                    
                    'xlSheet.Cells(CCnt, 31).value = convert(Val(C_SHT_EXC))                          'Cumulative Short/Excess
                    xlSheet.Cells(CCnt, 31).Value = 0#                                                'Cumulative Short/Excess
                    If Trim(INV_RS("VAN_CLS")) = "N" Then                                         'Challan Status, N for New, S for Returns Posted, C Closed
                        xlSheet.Cells(CCnt, 32).Value = "Open"
                    ElseIf Trim(INV_RS("VAN_CLS")) = "S" Then
                        xlSheet.Cells(CCnt, 32).Value = "Returns Posted"
                    ElseIf Trim(INV_RS("VAN_CLS")) = "C" Then
                        xlSheet.Cells(CCnt, 32).Value = "Closed"
                    End If
                    
                    xlSheet.Cells(CCnt, 33).Value = 0                                       'Total Variance
                    
                    If Not CUR_RS.EOF Then
                        xlSheet.Cells(CCnt, 34).Value = Val(CUR_RS("ONE_IDY"))                  '1000s
                        xlSheet.Cells(CCnt, 35).Value = Val(CUR_RS("TWO_IDY"))
                        xlSheet.Cells(CCnt, 36).Value = Val(CUR_RS("THR_IDY"))
                        xlSheet.Cells(CCnt, 37).Value = Val(CUR_RS("FOR_IDY"))
                        xlSheet.Cells(CCnt, 38).Value = Val(CUR_RS("FIV_IDY"))
                        xlSheet.Cells(CCnt, 39).Value = Val(CUR_RS("SIX_IDY"))
                        xlSheet.Cells(CCnt, 40).Value = Val(CUR_RS("SEV_IDY"))
                        xlSheet.Cells(CCnt, 41).Value = Val(CUR_RS("EIG_IDY"))
                        xlSheet.Cells(CCnt, 42).Value = Val(CUR_RS("NIN_IDY"))
                        xlSheet.Cells(CCnt, 43).Value = Val(CUR_RS("TEN_IDY"))
                        If IsNull(CUR_RS("CHG_IDY")) = False Then
                            xlSheet.Cells(CCnt, 44).Value = Val(CUR_RS("CHG_IDY"))
                        Else
                            xlSheet.Cells(CCnt, 44).Value = 0#
                        End If
                    End If
                    
                    
                    J = J + 1
                    CCnt = CCnt + 1
                   
                   ADD_CTR = ADD_CTR + 1
                   INC_P2B
                   L_SMN_IDY = ""
                   L_ROU_IDY = ""
                   L_SMN_NME = ""
                   L_CAT_NME = ""
                   L_LOC_NME = ""
                   L_MAR_NME = ""
                   L_MAR_CSG = ""
                   L_MAR_CLU = ""
                   L_MAR_ROC = ""
                   L_MAR_TWC = ""
                   L_MAR_TWN = ""
                   L_MAR_DST = ""
                   L_MAR_STA = ""
                   L_MAR_COU = ""
                   D_MON = ""
                   D_TUE = ""
                   D_WED = ""
                   D_THU = ""
                   D_FRI = ""
                   D_SAT = ""
                   D_SUN = ""
              End If
          INV_RS.MoveNext

       Wend
        If Dir(M_FIL_NME) <> "" Then
            Kill M_FIL_NME
        Else
            xlSheet.SaveAs M_FIL_NME
        End If
        MsgBox ("Van Register WorkSheet File " & M_FIL_NME & " with " & CCnt & " Records Created in the Folder C:\DEX"), vbInformation, "DAS Version 4.0"
        Me.MousePointer = Normal
        xlBook.Close
        xlApp.Quit
        
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        Set Rrs = Nothing
        Set Category_RS = Nothing
        Set Location_RS = Nothing
        Set Market_RS = Nothing
        Set Salesman_RS = Nothing
        Set RDY_RS = Nothing
        Set INV_RS = Nothing
        Set SNV_RS = Nothing
        Set CUR_RS = Nothing
        Set DET_RS = Nothing
        Set PRC_RS = Nothing
        Unload M_FRM_QUK_RST
    ' * : End of XL Van Register Creation Code for XL/Text Data Export
errh:
    If err.Number = 429 Then
        MsgBox "Required External File Data Component Not Available!", vbCritical, "DAS Version 4.0"
        Unload M_FRM_QUK_RST
        Exit Sub
    ElseIf err.Number = 364 Then
        MsgBox "Component Error!", vbCritical, "DAS Version 4.0"
        Me.MousePointer = Normal
        xlBook.Close
        xlApp.Quit
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        Unload M_FRM_QUK_RST
        Exit Sub
    ElseIf err.Number > 0 Then
        MsgBox "Error No " & str(err.Number), vbCritical, "DAS Version 4.0"
        Me.MousePointer = Normal
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        Unload M_FRM_QUK_RST
        Exit Sub
    End If
End Sub
Public Sub INC_PCB()
COM_CTR = COM_CTR + 1
ProgressBar1.Value = (100 / CTR) * COM_CTR
End Sub
Public Sub INC_P2B()
REC_CTR = REC_CTR + 1
ProgressBar1.Value = (100 / P2B_CTR) * REC_CTR
End Sub
Private Sub XLSalesRegister()
On Error GoTo errh
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
  
' * Start of Distributor ID Identification
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
            Exit Sub
        End If
        M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
        If Len(Trim(M_WDS_SRL)) < 3 Then
             While Len(Trim(M_WDS_SRL)) < 3
                  M_WDS_SRL = "0" & M_WDS_SRL
             Wend
        End If
    End If
    
' * End of Distributor ID Identification

        
    ' * : Beginning of _SaleReg.XLS - XLSalesRegister Creation Code for XL Data Export
    
    Me.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    M_FIL_NME = "C:\DEX\" & Trim(M_WDS_IDY) & "_SaleReg.XLS"
       
    ' Title Line Record for Excel Data Export File
    xlSheet.Cells(1, 1).Value = "DAS_ACT_IDY"
    xlSheet.Cells(1, 2).Value = "NAME"
    xlSheet.Cells(1, 3).Value = "TELEPHONE"
    xlSheet.Cells(1, 4).Value = "CONTACT"
    xlSheet.Cells(1, 5).Value = "LOCATION"
    xlSheet.Cells(1, 6).Value = "MARKET"
    xlSheet.Cells(1, 7).Value = "SALESMAN"
    xlSheet.Cells(1, 8).Value = "BILL SALESMAN"
    xlSheet.Cells(1, 9).Value = "CATEGORY"
    xlSheet.Cells(1, 10).Value = "ROUTE"
    xlSheet.Cells(1, 11).Value = "BILL ROUTE"
    xlSheet.Cells(1, 12).Value = "GRADE"
    xlSheet.Cells(1, 13).Value = "ACTIVE"
    xlSheet.Cells(1, 14).Value = "CAPACITY"
    xlSheet.Cells(1, 15).Value = "RED CODE"
    xlSheet.Cells(1, 16).Value = "HVO CODE"
    xlSheet.Cells(1, 17).Value = "TYPE"
    xlSheet.Cells(1, 18).Value = "CASH-CREDIT"
    xlSheet.Cells(1, 19).Value = "MONDAY"
    xlSheet.Cells(1, 20).Value = "TUESDAY"
    xlSheet.Cells(1, 21).Value = "WEDNESDAY"
    xlSheet.Cells(1, 22).Value = "THURSDAY"
    xlSheet.Cells(1, 23).Value = "FRIDAY"
    xlSheet.Cells(1, 24).Value = "SATURDAY"
    xlSheet.Cells(1, 25).Value = "SUNDAY"
    xlSheet.Cells(1, 26).Value = "CUSTOMER GROUP"
    xlSheet.Cells(1, 27).Value = "RSA CLUSTER"
    xlSheet.Cells(1, 28).Value = "RED OUTLET CLASS"
    xlSheet.Cells(1, 29).Value = "TOWN CLASS"
    xlSheet.Cells(1, 30).Value = "TOWN"
    xlSheet.Cells(1, 31).Value = "DISTRICT"
    xlSheet.Cells(1, 32).Value = "STATE"
    xlSheet.Cells(1, 33).Value = "SUPPLIER CODE"
    xlSheet.Cells(1, 34).Value = "Brand"
    xlSheet.Cells(1, 35).Value = "Pack Size"
    xlSheet.Cells(1, 36).Value = "Pack Type"
    xlSheet.Cells(1, 37).Value = "Product Category"
    xlSheet.Cells(1, 38).Value = "Subunits"
    xlSheet.Cells(1, 39).Value = "PurLanding Rate"
    xlSheet.Cells(1, 40).Value = "SaleLanding Rate"
    xlSheet.Cells(1, 41).Value = "MRP"
    xlSheet.Cells(1, 42).Value = "Date From"
    xlSheet.Cells(1, 43).Value = "Date To"
    xlSheet.Cells(1, 44).Value = "SKU NAME"
    xlSheet.Cells(1, 45).Value = "SalesQty"
    xlSheet.Cells(1, 46).Value = "FreeQty"
    xlSheet.Cells(1, 47).Value = "TotalQty"
    xlSheet.Cells(1, 48).Value = "CashDiscount"
    xlSheet.Cells(1, 49).Value = "MER Ref No"
 '   F_OMS.WriteLine "DAS_ACT_IDY" & "|" & "NAME" & "|" & "GST" & "|" & "CST" & "|" & "ADDRESS1" & "|" & "ADDRESS2" & "|" & "ADDRESS3" & "|" & "ADDRESS4" & "|" & "ZIPCODE" & "|" & "TELEPHONE" & "|" & "CONTACT" & "|" & "LOCATION" & "|" & "MARKET" & "|" & "SALESMAN" & "|" & "CATEGORY" & "|" & "ROUTE" & "|" & "GRADE" & "|" & "ACTIVE" & "|" & "CAPACITY" & "|" & "RED CODE" & "|" & "HVO CODE" & "|" & "TYPE" & "|" & "CASH-CREDIT" & "|" & "MONDAY" & "|" & "TUESDAY" & "|" & "WEDNESDAY" & "|" & "THURSDAY" & "|" & "FRIDAY" & "|" & "SATURDAY" & "|" & "SUNDAY" & "|" & "CUSTOMER GROUP" & "|" & "RSA CLUSTER" & "|" & "RED OUTLET CLASS" & "|" & "TOWN CLASS" & "|" & "TOWN" & "|" & "DISTRICT" & "|" & "STATE" & "|" & "COUNTRY"
    
    Dim J As Integer
    Dim CCnt As Integer
    J = 0
    CCnt = 2
    ADD_CTR = 0
    Dim Rrs As New ADODB.Recordset
    Dim Category_RS As New ADODB.Recordset
    Dim Location_RS As New ADODB.Recordset
    Dim Market_RS As New ADODB.Recordset
    Dim Salesman_RS As New ADODB.Recordset
    Dim RDY_RS As New ADODB.Recordset
    Dim INV_RS As New ADODB.Recordset
    
    If INV_RS.State = 1 Then INV_RS.Close
    INV_RS.Open "SELECT INL.CST_IDY, INL.INV_DAT, PRD.PRD_PRN, PRD.BRD_NME, PRD.PRD_PCK, PRD.PRD_PKT, PGR.GRP_CAT, PRD.SUB_UNT, BRD.PUR_LPP, BRD.SAL_LPP, BRD.PRD_MRP, PRD.PRD_NME, SUM(INV.TOT_PCS), INV.BAT_IDY, INV.STK_TYP, INV.SCH_IDY, INL.FRT_AMT, INV.TXN_TYP, INL.ROU_IDY, INL.SMN_IDY FROM INV, BRD, PRD, PGR, INL  WHERE INV.TXN_TYP = 'SAL' AND INV.BAT_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY AND PRD.GRP_IDY = PGR.GRP_IDY AND INV.INV_IDY = INL.INV_IDY AND INV.INV_DAT >= CTOD('" & DTPicker1 & "') AND INV.INV_DAT <= CTOD('" & DTPicker2 & "') GROUP BY INL.CST_IDY, INV.BAT_IDY, INV.STK_TYP, INV.SCH_IDY, INV.TXN_TYP HAVING INV.TXN_TYP = 'SAL' ORDER BY INL.CST_IDY, INL.INV_DAT, PRD.PRD_PRN", con, adOpenKeyset, adLockPessimistic
    P2B_CTR = INV_RS.RecordCount
    While Not INV_RS.EOF
        If Rrs.State = 1 Then Rrs.Close
        Rrs.Open "select * from ACT where ACT_TYP = 'SAL' AND DBT_TYP = 'Y' AND ACT_IDY = '" & Trim(INV_RS(0)) & "'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
        
        If Rrs.RecordCount > 0 Then
           
                Dim L_CAT_NME As String
                Dim L_LOC_NME As String
                Dim L_MAR_NME As String
                Dim L_SMN_NME As String
                Dim L_ROU_IDY As String
                Dim L_DST_NME As String
                Dim L_MAR_GRD As String
                Dim R_SMN_NME As String
                Dim XL_ACT_IDY As String
                Dim D_MON, D_TUE, D_WED, D_THU, D_FRI, D_SAT, D_SUN As String
                
                If Category_RS.State = 1 Then Category_RS.Close
                Category_RS.Open "SELECT * FROM CAT WHERE CAT_IDY = '" & Trim(Rrs("CAT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Category_RS.RecordCount > 0 Then
                      L_CAT_NME = Category_RS("CAT_NME")
                Else
                      L_CAT_NME = ""
                End If
                
                If Location_RS.State = 1 Then Location_RS.Close
                Location_RS.Open "SELECT * FROM LOC WHERE LOC_IDY = '" & Trim(Rrs("LOC_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Location_RS.RecordCount > 0 Then
                      L_LOC_NME = Location_RS("LOC_NME")
                Else
                      L_LOC_NME = ""
                End If
                
                If Market_RS.State = 1 Then Market_RS.Close
                Market_RS.Open "SELECT * FROM MAR WHERE MAR_IDY = '" & Trim(Rrs("MAR_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Market_RS.RecordCount > 0 Then
                      L_MAR_NME = Market_RS("MAR_NME")
                      L_DST_NME = Market_RS("MAR_PGR")
                      L_MAR_GRD = Market_RS("MAR_GRD")
                         ' Added for Margin Minder Data Extract by PNR
                      L_MAR_CSG = Market_RS("MAR_CSG")
                      L_MAR_CLU = Market_RS("MAR_CLU")
                      L_MAR_ROC = Market_RS("MAR_ROC")
                      L_MAR_TWC = Market_RS("MAR_TWC")
                      L_MAR_TWN = Market_RS("MAR_TWN")
                      L_MAR_DST = Market_RS("MAR_DST")
                      L_MAR_STA = Market_RS("MAR_STA")
                      L_MAR_COU = Market_RS("MAR_COU")
                Else
                      L_MAR_NME = ""
                      L_DST_NME = ""
                      L_MAR_GRD = ""
                         ' Added for Margin Minder Data Extract by PNR
                      L_MAR_CSG = ""
                      L_MAR_CLU = ""
                      L_MAR_ROC = ""
                      L_MAR_TWC = ""
                      L_MAR_TWN = ""
                      L_MAR_DST = ""
                      L_MAR_STA = ""
                      L_MAR_COU = ""
                End If
                
                   
                If RDY_RS.State = 1 Then RDY_RS.Close
                RDY_RS.Open "SELECT * FROM RDY ORDER BY DAY_IDY WHERE ACT_IDY = '" & Trim(Rrs("ACT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If RDY_RS.RecordCount > 0 Then RDY_RS.MoveFirst
                   
                   If RDY_RS.RecordCount > 0 Then
                      L_SMN_IDY = Trim(RDY_RS("SMN_IDY"))
                      L_ROU_IDY = Trim(RDY_RS("ROU_IDY"))
                   Else
                      L_SMN_IDY = ""
                      L_ROU_IDY = ""
                   End If
                   
                   While Not RDY_RS.EOF
                   If RDY_RS("ACT_IDY") = Trim(Rrs("ACT_IDY")) Then
                         If RDY_RS("DAY_IDY") = 1 Then
                             D_MON = "MONDAY"
                         ElseIf RDY_RS("DAY_IDY") = 2 Then
                             D_TUE = "TUESDAY"
                         ElseIf RDY_RS("DAY_IDY") = 3 Then
                             D_WED = "WEDNESDAY"
                         ElseIf RDY_RS("DAY_IDY") = 4 Then
                             D_THU = "THURSDAY"
                         ElseIf RDY_RS("DAY_IDY") = 5 Then
                             D_FRI = "FRIDAY"
                         ElseIf RDY_RS("DAY_IDY") = 6 Then
                             D_SAT = "SATURDAY"
                         ElseIf RDY_RS("DAY_IDY") = 7 Then
                             D_SUN = "SUNDAY"
                         End If
                   End If
                   RDY_RS.MoveNext
                   Wend
                   
                     If Salesman_RS.State = 1 Then Salesman_RS.Close
                     Salesman_RS.Open "SELECT * FROM SMN WHERE SMN_IDY = '" & Trim(L_SMN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
                     If Salesman_RS.RecordCount > 0 Then
                         L_SMN_NME = Salesman_RS("SMN_NME")
                     Else
                         L_SMN_NME = ""
                     End If
                     If Salesman_RS.State = 1 Then Salesman_RS.Close
                     Salesman_RS.Open "SELECT * FROM SMN WHERE SMN_IDY = '" & Trim(INV_RS(19)) & "'", con, adOpenKeyset, adLockPessimistic
                     If Salesman_RS.RecordCount > 0 Then
                         R_SMN_NME = Salesman_RS("SMN_NME")
                     Else
                         R_SMN_NME = ""
                     End If
                       
                    xlSheet.Cells(CCnt, 1).Value = "ID" & Trim(Sincrement(INV_RS(0))) & " "
                    xlSheet.Cells(CCnt, 2).Value = Trim(Rrs("ACT_NME"))
                    xlSheet.Cells(CCnt, 3).Value = Trim(Rrs("ACT_TFX"))
                    xlSheet.Cells(CCnt, 4).Value = Trim(Rrs("ACT_CPN"))
                    xlSheet.Cells(CCnt, 5).Value = Trim(L_LOC_NME)
                    xlSheet.Cells(CCnt, 6).Value = Trim(L_MAR_NME)
                    xlSheet.Cells(CCnt, 7).Value = Trim(L_SMN_NME)
                    xlSheet.Cells(CCnt, 8).Value = Trim(R_SMN_NME)
                    xlSheet.Cells(CCnt, 9).Value = Trim(L_CAT_NME)
                    xlSheet.Cells(CCnt, 10).Value = Trim(L_ROU_IDY)
                    xlSheet.Cells(CCnt, 11).Value = Trim(INV_RS(18))
                    xlSheet.Cells(CCnt, 12).Value = Trim(Rrs("RET_TYP"))
                    xlSheet.Cells(CCnt, 13).Value = Trim(Rrs("DBT_TYP"))
                    xlSheet.Cells(CCnt, 14).Value = Trim(str(Rrs("VAN_CAP")))
                    xlSheet.Cells(CCnt, 15).Value = Trim(Rrs("ACN_TWO"))
                    xlSheet.Cells(CCnt, 16).Value = Trim(Rrs("BNM_TWO"))
                    xlSheet.Cells(CCnt, 17).Value = Trim(Rrs("CST_TYP"))
                    xlSheet.Cells(CCnt, 18).Value = Trim(Rrs("MCL_REQ"))
                    xlSheet.Cells(CCnt, 19).Value = Trim(D_MON)
                    xlSheet.Cells(CCnt, 20).Value = Trim(D_TUE)
                    xlSheet.Cells(CCnt, 21).Value = Trim(D_WED)
                    xlSheet.Cells(CCnt, 22).Value = Trim(D_THU)
                    xlSheet.Cells(CCnt, 23).Value = Trim(D_FRI)
                    xlSheet.Cells(CCnt, 24).Value = Trim(D_SAT)
                    xlSheet.Cells(CCnt, 25).Value = Trim(D_SUN)
                    xlSheet.Cells(CCnt, 26).Value = Trim(L_MAR_CSG)
                    xlSheet.Cells(CCnt, 27).Value = Trim(L_MAR_CLU)
                    xlSheet.Cells(CCnt, 28).Value = Trim(L_MAR_ROC)
                    xlSheet.Cells(CCnt, 29).Value = Trim(L_MAR_TWC)
                    xlSheet.Cells(CCnt, 30).Value = Trim(L_MAR_TWN)
                    xlSheet.Cells(CCnt, 31).Value = Trim(L_MAR_DST)
                    xlSheet.Cells(CCnt, 32).Value = Trim(L_MAR_STA)
                    xlSheet.Cells(CCnt, 33).Value = "ID" & Trim(INV_RS(2)) & " "
                    xlSheet.Cells(CCnt, 34).Value = Trim(INV_RS(3))
                    xlSheet.Cells(CCnt, 35).Value = Trim(INV_RS(4))
                    xlSheet.Cells(CCnt, 36).Value = Trim(INV_RS(5))
                    xlSheet.Cells(CCnt, 37).Value = Trim(INV_RS(6))
                    xlSheet.Cells(CCnt, 38).Value = Trim(INV_RS(7))
                    xlSheet.Cells(CCnt, 39).Value = Trim(str(INV_RS(8)))
                    xlSheet.Cells(CCnt, 40).Value = Trim(str(INV_RS(9)))
                    xlSheet.Cells(CCnt, 41).Value = Trim(str(INV_RS(10))) 'mrp
                    xlSheet.Cells(CCnt, 42).Value = Format(DTPicker1, "DD/MM/YYYY")
                    xlSheet.Cells(CCnt, 43).Value = Format(DTPicker2, "DD/MM/YYYY")
                    xlSheet.Cells(CCnt, 44).Value = Trim(INV_RS(11))
                    
                    If Trim(INV_RS(14)) = "Salable" Then
                        xlSheet.Cells(CCnt, 45).Value = Trim(str(Round((Val(INV_RS(12)) / Val(INV_RS(7))), 2)))
                    End If
                    If Trim(INV_RS(14)) = "Free" Then
                        xlSheet.Cells(CCnt, 46).Value = Trim(str(Round((Val(INV_RS(12)) / Val(INV_RS(7))), 2))) 'Trim(STR(INV_RS(12)))
                    End If
                    xlSheet.Cells(CCnt, 47).Value = Trim(str(Round((Val(INV_RS(12)) / Val(INV_RS(7))), 2))) 'Trim(STR(INV_RS(12)))                      'calculated field in xl
                    xlSheet.Cells(CCnt, 48).Value = Trim(str(INV_RS(16)))
                    xlSheet.Cells(CCnt, 49).Value = Trim(INV_RS(15))
                    J = J + 1
                    CCnt = CCnt + 1
                   
                   ADD_CTR = ADD_CTR + 1
                   INC_P2B
                   L_SMN_IDY = ""
                   L_ROU_IDY = ""
                   L_SMN_NME = ""
                   L_CAT_NME = ""
                   L_LOC_NME = ""
                   L_MAR_NME = ""
                   L_MAR_CSG = ""
                   L_MAR_CLU = ""
                   L_MAR_ROC = ""
                   L_MAR_TWC = ""
                   L_MAR_TWN = ""
                   L_MAR_DST = ""
                   L_MAR_STA = ""
                   L_MAR_COU = ""
                   D_MON = ""
                   D_TUE = ""
                   D_WED = ""
                   D_THU = ""
                   D_FRI = ""
                   D_SAT = ""
                   D_SUN = ""
              End If
          INV_RS.MoveNext

       Wend
        If Dir(M_FIL_NME) <> "" Then
            Kill M_FIL_NME
        Else
            xlSheet.SaveAs M_FIL_NME
        End If
        MsgBox ("Sales Register WorkSheet File " & M_FIL_NME & " with " & CCnt & " Records Created in the Folder C:\DEX"), vbInformation, "DAS Version 4.0"
        Me.MousePointer = Normal
        xlBook.Close
        xlApp.Quit
        
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        Set Rrs = Nothing
        Set Category_RS = Nothing
        Set Location_RS = Nothing
        Set Market_RS = Nothing
        Set Salesman_RS = Nothing
        Set RDY_RS = Nothing
        Set INV_RS = Nothing
        
        Unload M_FRM_QUK_RST
    ' * : End of XL Sales Register Creation Code for XL/`Text Data Export
errh:
    If err.Number = 429 Then
        MsgBox "Required External File Data Component Not Available!", vbCritical, "DAS Version 4.0"
        Unload M_FRM_QUK_RST
        Exit Sub
    ElseIf err.Number = 364 Then
        MsgBox "Component Error!", vbCritical, "DAS Version 4.0"
        Me.MousePointer = Normal
        xlBook.Close
        xlApp.Quit
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        Unload M_FRM_QUK_RST
        Exit Sub
    ElseIf err.Number > 0 Then
        MsgBox "Error No " & str(err.Number), vbCritical, "DAS Version 4.0"
        Me.MousePointer = Normal
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        Unload M_FRM_QUK_RST
        Exit Sub
    End If
End Sub

'            CON.Execute "DELETE FROM INV WHERE INV_IDY LIKE '" & Trim(RST_RSET!INV_IDY) & "' AND TXN_TYP LIKE '" & Trim(RST_RSET!TXN_TYP) & "'", "D"
'            CON.Execute "DELETE FROM LNV WHERE INV_IDY LIKE '" & Trim(RST_RSET!INV_IDY) & "'", "D"
'            CON.Execute "DELETE FROM PNV WHERE INV_IDY LIKE '" & Trim(RST_RSET!INV_IDY) & "' AND TXN_TYP LIKE '" & Trim(RST_RSET!TXN_TYP) & "'", "D"
'            CON.Execute "DELETE FROM INL WHERE INV_IDY LIKE '" & Trim(RST_RSET!INV_IDY) & "' AND TXN_TYP LIKE '" & Trim(RST_RSET!TXN_TYP) & "'", "D"
'            CON.Execute "DELETE FROM PNL WHERE INV_IDY LIKE '" & Trim(RST_RSET!INV_IDY) & "' AND TXN_TYP LIKE '" & Trim(RST_RSET!TXN_TYP) & "'", "D"
'            CON.Execute "DELETE FROM VAN WHERE INV_IDY LIKE '" & Trim(RST_RSET!INV_IDY) & "'", "D"
'            CON.Execute "DELETE FROM SNV WHERE INV_IDY LIKE '" & Trim(RST_RSET!INV_IDY) & "' AND TXN_TYP LIKE '" & Trim(RST_RSET!TXN_TYP) & "'", "D"

'            If Trim(INV_RS(14)) = "Salable" Then
'                xlSheet.Cells(CCnt, 45).value = Trim(STR(Round((Val(INV_RS(12)) / Val(INV_RS(7))), 2)))
'            End If
'            If Trim(INV_RS(14)) = "Free" Then
'                xlSheet.Cells(CCnt, 46).value = Trim(STR(Round((Val(INV_RS(12)) / Val(INV_RS(7))), 2))) 'Trim(STR(INV_RS(12)))
'            End If
'            xlSheet.Cells(CCnt, 47).value = Trim(STR(Round((Val(INV_RS(12)) / Val(INV_RS(7))), 2))) 'Trim(STR(INV_RS(12)))                      'calculated field in xl
'            xlSheet.Cells(CCnt, 48).value = Trim(STR(INV_RS(16)))
'            xlSheet.Cells(CCnt, 49).value = Trim(INV_RS(15))
