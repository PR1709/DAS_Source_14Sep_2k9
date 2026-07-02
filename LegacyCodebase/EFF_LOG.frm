VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form EFF_LOG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DAS Version 4.0 Login"
   ClientHeight    =   1560
   ClientLeft      =   3900
   ClientTop       =   3435
   ClientWidth     =   4275
   Icon            =   "EFF_LOG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   921.7
   ScaleMode       =   0  'User
   ScaleWidth      =   4013.994
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   339
      IMEMode         =   3  'DISABLE
      Left            =   1095
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1065
      Width           =   1935
   End
   Begin VB.ComboBox TXTUSERNAME 
      Height          =   315
      ItemData        =   "EFF_LOG.frx":08CA
      Left            =   1095
      List            =   "EFF_LOG.frx":08CC
      TabIndex        =   5
      Text            =   "ADMINISTRATOR"
      Top             =   660
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   975
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   339
         Left            =   120
         TabIndex        =   3
         Top             =   645
         Width           =   780
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   780
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1095
      TabIndex        =   4
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   65142787
      CurrentDate     =   37631
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   8
      Top             =   1140
      Width           =   690
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   570
      TabIndex        =   6
      Top             =   300
      Width           =   345
   End
End
Attribute VB_Name = "EFF_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cmdCancel_Click()
    'LoginSucceeded = False
If Trim(Me.Tag) <> "CHANGE" Then
    End
Else
    Unload Me
    Set Form = Nothing
End If
End Sub
Private Sub cmdOk_Click()
' Verify DB Tables Availability in the Current Data Folder
If CHECKTB("DIR.DBF,ACT.DBF,GRP.DBF,LOC.DBF,CAT.DBF,SMN.DBF,MAR.DBF,SCH.DBF,WEK.DBF,YER.DBF,LOG.DBF,PRD.DBF,PGR.DBF,BRD.DBF,TGT.DBF,USR.DBF,IDY.DBF,SHL.DBF,AST.DBF,TYP.DBF,RDY.DBF,SID.DBF,SIC.DBF") = 1 Then
    MsgBox UCase(TNAME) & " missing in " & Me.Caption & ", Application terminated. ", vbCritical, "DAS Version 4.0"
    End
End If
If CHECKTB("INV.DBF,INL.DBF,PNV.DBF,PNL.DBF,VAN.DBF,SNV.DBF,CUR.DBF,IRT.DBF,LNV.DBF,VCH.DBF,ORD.DBF") = 1 Then
    MsgBox UCase(TNAME) & " missing in " & Me.Caption & ", Application terminated. ", vbCritical, "DAS Version 4.0"
    End
End If
If CHECKTB("BNV.DBF,BNL.DBF,XNV.DBF,XNL.DBF,VBK.DBF,CNV.DBF,LDS.DBF,SBK.DBF,EFF.DBC,EFF.DCT, EFF.DCX") = 1 Then
    MsgBox UCase(TNAME) & " missing in " & Me.Caption & ", Application terminated. ", vbCritical, "DAS Version 4.0"
    End
End If
' Verify DB Tables Availability in the Current Data Folder

Dim REP_MENU As String
Dim ANA_MENU As String
Dim RSlog As New ADODB.Recordset

' * AutoBackup and AutoRestore Folder Creation Begin 'PNR
Dim AUTOBKP As String
Dim AUTORST As String

AUTOBKP = CURDIR & "\AUTOBKP"
AUTORST = CURDIR & "\AUTORST"
If Dir(AUTOBKP, vbDirectory) = "" Then
    MkDir AUTOBKP
End If
If Dir(AUTORST, vbDirectory) = "" Then
    MkDir AUTORST
End If
' * AutoBackup and AutoRestore Folder Creation End 'PNR

If Len(Trim(TXTUSERNAME)) <> 0 And Len(Trim(txtPassword)) <> 0 Then
    If rs.State = 1 Then rs.Close
    
    'Added by PNR for TWP Saving through Control Admin Login Option
    If txtPassword = "cspladmin" Then
        Dim DIR_RS As New ADODB.Recordset
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            DIR_RS.MoveFirst
            DIR_RS("DIR_MSG") = "WRITETWP"
            DIR_RS.Update
        End If
    End If
    ' Added by PNR for TWP Saving through Control Admin Login Option
    
'         Evaluation Check Module
    If Date >= "25/09/2009" Then
        MsgBox ("Evaluation Period Expired! Please Contact Support!"), vbCritical, "Consoft Systems P. Ltd, Mobile: 098850 02290"
        LOGOUT
        End
    End If
'     End of Evaluation Check Module
    
    If txtPassword = "cspladmin" Then GoTo Password
    rs.Open "select * from usr where usr_nme like '" & Trim(TXTUSERNAME) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        DISABLEMENU
        If TXTUSERNAME = Trim(rs(0)) And txtPassword = DPWD.decPwd((rs(1))) Then
            If Trim(rs(3)) = "Admin" Then
Password:
                
                M_MDI.Caption = "DAS Version 4.01.01 Rel. - Distributor Automation System, Ver.Rel.04-Sep-2009" & Space(40) & "Login : " & Trim(TXTUSERNAME) & " for " & Trim(SEL_CMP)
                USERNAME = TXTUSERNAME
                RSlog.Open "select * from log", con, adOpenKeyset, adLockPessimistic
                RSlog.AddNew
                RSlog!usr_nme = Trim(TXTUSERNAME)
                RSlog!LIN_DAT = Date
                RSlog!lin_tim = Format(Time, "HH:mm:SS")
                RSlog.Update
                Unload Me
                Set Form = Nothing
                Check_Academic
                Check_Shutdown_Status
                Check_Backup_Status
                
                M_MDI.Show
                For b = 1 To M_MDI.count - 1
                    If TypeOf M_MDI.Controls(b) Is Menu Then M_MDI.Controls(b).Enabled = True
                    'If TypeOf M_MDI.Controls(b) Is MenuItem Then M_MDI.Controls(b).Enabled = True
                    'If Trim(M_MDI.Controls(b).Tag) <> "" Then M_MDI.Controls(b).Enabled = True
                Next
                Dim RS_DIR As New ADODB.Recordset
                If RS_DIR.State = 1 Then RS_DIR.Close
                RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'VER'", con, adOpenKeyset, adLockPessimistic
                If RS_DIR.RecordCount > 0 Then
                     RS_DIR.MoveFirst
                     If Dir(App.Path & "EFFMCG.EXE") = "" Then
                        M_EXE_VER = Format(FileDateTime(App.Path & "\EFFMCG.EXE"), "dd-mmm-yyyy hh:mm:ss")
                        RS_DIR("DIR_MSG") = "04.01.01.B.n, Exe Version Release Date: " & Trim(M_EXE_VER)
                        RS_DIR("DIR_ALS") = "27/07/2008"
                        RS_DIR.Update
                     End If
                End If
                If RS_DIR.State = 1 Then RS_DIR.Close
                    RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'TWP'", con, adOpenKeyset, adLockPessimistic
                    If RS_DIR.RecordCount > 0 Then
                        If Trim(RS_DIR("DIR_ALS")) = "Y" Then
                           M_MDI.M_MNU_SHL_SNP1.Visible = True
                            M_MDI.M_MNU_SHL_SNP1.Enabled = True
                        Else
                            M_MDI.M_MNU_SHL_SNP1.Visible = False
                        End If
                    Else
                            M_MDI.M_MNU_SHL_SNP1.Visible = False
                    End If
                 If RS_DIR.State = 1 Then RS_DIR.Close
                    RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'TCO'", con, adOpenKeyset, adLockPessimistic
                    If RS_DIR.RecordCount > 0 Then
                      If Trim(RS_DIR("DIR_ALS")) = "Y" Then
                        M_MDI.M_COR_HYP5(0).Visible = True
                         M_MDI.SHL_SEP.Visible = True
                         
                         
                         M_MDI.TCO_ENT(0).Enabled = True
  '                       Load M_MDI.TCO_ENT(1)
                         M_MDI.TCO_ENT(1).Caption = "&Telecom Interface"
                         M_MDI.TCO_ENT(1).Enabled = True
                       Else
                         M_MDI.M_COR_HYP5(0).Visible = False
                         M_MDI.SHL_SEP.Visible = False
                       End If
                    Else
                         M_MDI.M_COR_HYP5(0).Visible = False
                         M_MDI.SHL_SEP.Visible = False
                    End If
                 If RS_DIR.State = 1 Then RS_DIR.Close
                    RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'MIB'", con, adOpenKeyset, adLockPessimistic
                    If RS_DIR.RecordCount > 0 Then
                      If Trim(RS_DIR("DIR_ALS")) = "Y" Then
                        M_MDI.SYS_MIS.Enabled = True
                      Else
                        M_MDI.SYS_MIS.Visible = True
                        M_MDI.M_SYS_BMP(2).Enabled = False
                        M_MDI.M_SYS_BMP(3).Enabled = False
                        M_MDI.M_SYS_BMP(4).Enabled = False
                        M_MDI.M_SYS_BMP(5).Enabled = False
                      End If
                    Else
                        M_MDI.SYS_MIS.Visible = True
                        M_MDI.M_SYS_BMP(2).Enabled = False
                        M_MDI.M_SYS_BMP(3).Enabled = False
                        M_MDI.M_SYS_BMP(4).Enabled = False
                        M_MDI.M_SYS_BMP(5).Enabled = False
                    End If
'                M_MDI.M_FIL_SIC(1).Enabled = False
            Else
                M_MDI.Caption = "DAS Version 4.01.01 Rel. - Distributor Automation System, Ver.Rel.04-Sep-2009" & Space(40) & "Login : " & Trim(TXTUSERNAME) & " for " & Trim(SEL_CMP)
                'M_MDI.Caption = "DAS Version 4.0 - Distributor Automation System" & Space(50) & "LOGIN : " & Trim(TXTUSERNAME) & " for " & Trim(SEL_CMP)
                USERNAME = TXTUSERNAME
                RSlog.Open "select * from log", con, adOpenKeyset, adLockPessimistic
                RSlog.AddNew
                RSlog!usr_nme = Trim(TXTUSERNAME)
                RSlog!LIN_DAT = Date
                RSlog!lin_tim = Format(Time, "hh:mm:ss")
                RSlog.Update
                Unload Me
                Set Form = Nothing
                M_MDI.Show
                
                M_MDI.M_SYS.Enabled = False
                M_MDI.SYS_SET.Enabled = False
                M_MDI.SYS_DTB.Enabled = False
                M_MDI.SYS_MIS.Enabled = False
                M_MDI.M_SYS_BMP(2).Enabled = False
                M_MDI.M_SYS_BMP(3).Enabled = False
                M_MDI.M_SYS_BMP(4).Enabled = False
                M_MDI.M_SYS_BMP(5).Enabled = False
                For A = 1 To rs.RecordCount
                    For b = 1 To M_MDI.count - 1
                        If TypeOf M_MDI.Controls(b) Is Menu Then  'Or TypeOf M_MDI.Controls(b) Is MenuItem Then
                           If Trim(M_MDI.Controls(b).Tag) = Trim(rs(2)) Then
                                
                                If Trim(rs(2)) = "Analysis" Then
                                ANA_MENU = "Analysis"
                                End If
                                If Trim(rs(2)) = "Reports" Then
                                    REP_MENU = "Reports"
                                End If
                                
                                If Trim(rs(2)) = "Sales(R)" And REP_MENU = "Reports" Then
                                    M_MDI.M_RPT_COR_SAL_RPT(0).Visible = True
                                    M_MDI.M_RPT_COR_SAL_RPT(0).Enabled = True
                                End If
                                
                                If Trim(rs(2)) = "Purchases(R)" And REP_MENU = "Reports" Then
                                    M_MDI.M_RPT_COR_PUR_SUM.Visible = True
                                    M_MDI.M_RPT_COR_PUR_SUM.Enabled = True
                                    M_MDI.M_RPT_COR_PUR_PRD.Visible = True
                                    M_MDI.M_RPT_COR_PUR_PRD.Enabled = True
                                    M_MDI.M_RPT_COR_PUR_RET.Visible = True
                                    M_MDI.M_RPT_COR_PUR_RET.Enabled = True
                                    M_MDI.M_RPT_OTH_A9S.Visible = True
                                    M_MDI.M_RPT_OTH_A9S.Enabled = True
                                    M_MDI.M_RPT_COR_PUR_PRD_TAX.Visible = True
                                    M_MDI.M_RPT_COR_PUR_PRD_TAX.Enabled = True
                                    M_MDI.M_RPT_COR_PUR_PRD_TAX_TTL.Visible = True
                                    M_MDI.M_RPT_COR_PUR_PRD_TAX_TTL.Enabled = True
                                    M_MDI.M_RPT_COR_PUR_MAR_RPT(0).Visible = True
                                    M_MDI.M_RPT_COR_PUR_MAR_RPT(0).Enabled = True
                                End If
                                If Trim(rs(2)) = "Stocks(R)" And REP_MENU = "Reports" Then
                                    M_MDI.M_RPT_COR_STK_ANL_AGD(0).Visible = True
                                    M_MDI.M_RPT_COR_STK_ANL_AGD(0).Enabled = True
                                    M_MDI.M_RPT_COR_STK_RGL_STL(0).Visible = True
                                    M_MDI.M_RPT_COR_STK_RGL_STL(0).Enabled = True
                                End If
                                
                                If Trim(rs(2)) = "Accounts(R)" And REP_MENU = "Reports" Then
                                M_MDI.M_RPT_ACT_BBK.Visible = True
                                M_MDI.M_RPT_ACT_SMM.Visible = True
                                M_MDI.M_RPT_ACT_CRL.Visible = True
                                M_MDI.M_RPT_ACT_DLS.Visible = True
                                M_MDI.dbt_agd.Visible = True
                                M_MDI.REC_STM.Visible = True
                                M_MDI.DebtorsAgeing.Visible = True
                                M_MDI.M_RPT_ACT_GLD.Visible = True
                                M_MDI.M_RPT_ACT_ACR.Visible = True
                                M_MDI.M_RPT_ACT_RVL.Visible = True
                                M_MDI.M_RPT_ACT_PVL.Visible = True
                                M_MDI.M_RPT_ACT_CNR.Visible = True
                                M_MDI.M_RPT_ACT_DNR.Visible = True
                                M_MDI.M_RPT_ACT_JVL.Visible = True
                                M_MDI.M_RPT_ACT_OAL.Visible = True
                                M_MDI.M_RPT_ACT_CLG_DAT.Visible = True
                                M_MDI.M_RPT_ACT_CLG_BBK.Visible = True
                                M_MDI.M_RPT_ACT_CLG_DLT.Visible = True
                                M_MDI.M_RPT_ACT_CLG_CRT.Visible = True
                                M_MDI.M_RPT_ACT_CLG_CHQ_RCT.Visible = True
                                M_MDI.M_RPT_ACT_CLG_CHQ_PMT.Visible = True
                                M_MDI.M_RPT_ACT_FAC.Visible = True
                                M_MDI.bls.Visible = True
                                M_MDI.M_RPT_ACT_FAC_BLS.Visible = True
                                M_MDI.M_RPT_ACT_FAC_ICE.Visible = True
                                M_MDI.M_RPT_ACT_FAC_TBL.Visible = True
                                M_MDI.M_RPT_ACT_FAC_GTBL(0).Visible = True
                                
                                M_MDI.M_RPT_ACT_BBK.Enabled = True
                                M_MDI.M_RPT_ACT_SMM.Enabled = True
                                M_MDI.M_RPT_ACT_CRL.Enabled = True
                                M_MDI.M_RPT_ACT_DLS.Enabled = True
                                M_MDI.dbt_agd.Enabled = True
                                M_MDI.REC_STM.Enabled = True
                                M_MDI.DebtorsAgeing.Enabled = True
                                M_MDI.M_RPT_ACT_GLD.Enabled = True
                                M_MDI.M_RPT_ACT_ACR.Enabled = True
                                M_MDI.M_RPT_ACT_RVL.Enabled = True
                                M_MDI.M_RPT_ACT_PVL.Enabled = True
                                M_MDI.M_RPT_ACT_CNR.Enabled = True
                                M_MDI.M_RPT_ACT_DNR.Enabled = True
                                M_MDI.M_RPT_ACT_JVL.Enabled = True
                                M_MDI.M_RPT_ACT_OAL.Enabled = True
                                M_MDI.M_RPT_ACT_CLG_DAT.Enabled = True
                                M_MDI.M_RPT_ACT_CLG_BBK.Enabled = True
                                M_MDI.M_RPT_ACT_CLG_DLT.Enabled = True
                                M_MDI.M_RPT_ACT_CLG_CRT.Enabled = True
                                M_MDI.M_RPT_ACT_CLG_CHQ_RCT.Enabled = True
                                M_MDI.M_RPT_ACT_CLG_CHQ_PMT.Enabled = True
                                M_MDI.M_RPT_ACT_FAC.Enabled = True
                                M_MDI.bls.Enabled = True
                                M_MDI.M_RPT_ACT_FAC_BLS.Enabled = True
                                M_MDI.M_RPT_ACT_FAC_ICE.Enabled = True
                                M_MDI.M_RPT_ACT_FAC_TBL.Enabled = True
                                M_MDI.M_RPT_ACT_FAC_GTBL(0).Enabled = True
                                
                                End If
                                
                                If Trim(rs(2)) = "Core(RP)" And REP_MENU = "Reports" Then
                                    M_MDI.M_RPT_RNG_COR_SAL.Visible = True
                                    M_MDI.M_RPT_RNG_COR_SSL.Visible = True
                                    M_MDI.M_RPT_RNG_COR_OPS.Visible = True
                                    M_MDI.M_RPT_RNG_COR_CRN.Visible = True
                                    M_MDI.M_RPT_RNG_COR_DRN.Visible = True
                                    M_MDI.M_RPT_RNG_COR_STI.Visible = True
                                    M_MDI.M_RPT_RNG_COR_STO.Visible = True
                                    M_MDI.M_RPT_RNG_COR_VAN.Visible = True
                                    
                                    M_MDI.M_RPT_RNG_COR_SAL.Enabled = True
                                    M_MDI.M_RPT_RNG_COR_SSL.Enabled = True
                                    M_MDI.M_RPT_RNG_COR_OPS.Enabled = True
                                    M_MDI.M_RPT_RNG_COR_CRN.Enabled = True
                                    M_MDI.M_RPT_RNG_COR_DRN.Enabled = True
                                    M_MDI.M_RPT_RNG_COR_STI.Enabled = True
                                    M_MDI.M_RPT_RNG_COR_STO.Enabled = True
                                    M_MDI.M_RPT_RNG_COR_VAN.Enabled = True
                                End If
                                
                                If Trim(rs(2)) = "Retailers" And ANA_MENU = "Analysis" Then
                                    M_MDI.M_ANA_RET_SKU.Visible = True
                                    M_MDI.M_MIR_YSR.Visible = True
                                    M_MDI.M_ANA_RET_RSN.Visible = True
                                
                                    M_MDI.M_ANA_RET_SKU.Enabled = True
                                    M_MDI.M_MIR_YSR.Enabled = True
                                    M_MDI.M_ANA_RET_RSN.Enabled = True
                                End If
                                
                                If Trim(rs(2)) = "Regular" And REP_MENU = "Reports" Then
                                    
                                    M_MDI.M_RPT_OTH_DIP(0).Visible = True
                                    M_MDI.M_RPT_OTH_DTS.Visible = True
                                    M_MDI.M_RPT_OTH_DTF.Visible = True
                                    M_MDI.M_RPT_OTH_STR.Visible = True
                                    M_MDI.M_RPT_OTH_CRC.Visible = True
                                    M_MDI.M_RPT_OTH_CLM(0).Visible = True
                                    M_MDI.M_RPT_OTH_WSR.Visible = True
                                    M_MDI.M_RPT_OTH_PNT.Visible = True
                                    M_MDI.M_RPT_OTH_SEG.Visible = True
                                    M_MDI.M_RPT_OTH_SSBR.Visible = True
                                    M_MDI.M_RPT_OTH_COK_SSB.Visible = True
                                    M_MDI.M_AR_COK_MIS.Visible = True
                                    M_MDI.M_RPT_OTH_COK.Visible = True
                                    M_MDI.M_RPT_OTH_MLS.Visible = True
                                    M_MDI.M_RPT_OTH_DIS.Visible = True
                                    M_MDI.M_RPT_OTH_RLT.Visible = True
                                    M_MDI.M_RPT_OTH_DIF.Visible = True
                                    M_MDI.M_RPT_OTH_SRC.Visible = True
                                    M_MDI.M_RPT_OTH_DIF_pur.Visible = True

                                    M_MDI.M_RPT_OTH_DIP(0).Enabled = True
                                    M_MDI.M_RPT_OTH_DTS.Enabled = True
                                    M_MDI.M_RPT_OTH_DTF.Enabled = True
                                    M_MDI.M_RPT_OTH_STR.Enabled = True
                                    M_MDI.M_RPT_OTH_CRC.Enabled = True
                                    M_MDI.M_RPT_OTH_CLM(0).Enabled = True
                                    M_MDI.M_RPT_OTH_WSR.Enabled = True
                                    M_MDI.M_RPT_OTH_PNT.Enabled = True
                                    M_MDI.M_RPT_OTH_SEG.Enabled = True
                                    M_MDI.M_RPT_OTH_SSBR.Enabled = True
                                    M_MDI.M_RPT_OTH_COK_SSB.Enabled = True
                                    M_MDI.M_AR_COK_MIS.Enabled = True
                                    M_MDI.M_RPT_OTH_COK.Enabled = True
                                    M_MDI.M_RPT_OTH_MLS.Enabled = True
                                    M_MDI.M_RPT_OTH_DIS.Enabled = True
                                    M_MDI.M_RPT_OTH_RLT.Enabled = True
                                    M_MDI.M_RPT_OTH_DIF.Enabled = True
                                    M_MDI.M_RPT_OTH_SRC.Enabled = True
                                    M_MDI.M_RPT_OTH_DIF_pur.Enabled = True
                                    
                                End If
                                
                                If Trim(rs(2)) = "Accounts(RP)" And REP_MENU = "Reports" Then
                                    M_MDI.M_RPT_RNG_ACT_RCT.Visible = True
                                    M_MDI.M_RPT_RNG_ACT_PMT.Visible = True
                                    M_MDI.M_RPT_RNG_ACT_CRN.Visible = True
                                    M_MDI.M_RPT_RNG_ACT_DRN.Visible = True
                                    M_MDI.M_RPT_RNG_ACT_JVN.Visible = True
                                    M_MDI.M_RPT_RNG_ACT_OPB.Visible = True
                                    

                                    M_MDI.M_RPT_RNG_ACT_RCT.Enabled = True
                                    M_MDI.M_RPT_RNG_ACT_PMT.Enabled = True
                                    M_MDI.M_RPT_RNG_ACT_CRN.Enabled = True
                                    M_MDI.M_RPT_RNG_ACT_DRN.Enabled = True
                                    M_MDI.M_RPT_RNG_ACT_JVN.Enabled = True
                                    M_MDI.M_RPT_RNG_ACT_OPB.Enabled = True
                                    
                                End If
                                
                                If Trim(rs(2)) = "Direct Print" And REP_MENU = "Reports" Then
                                    M_MDI.M_REP_DIP.Enabled = True
                                    M_MDI.M_REP_DIP.Visible = True
                                End If
                                If Trim(rs(2)) = "Database" Then
                                M_MDI.SYS_DTB.Enabled = True
                                M_MDI.SYS_DTB.Visible = True
                                End If
                                
                                M_MDI.Controls(b).Enabled = True
                                Exit For
                            End If
                        End If
                    Next
                    rs.MoveNext
                Next
                M_MDI.M_SYS.Enabled = True
                M_MDI.M_FIL_EXT.Enabled = True

                Unload Me
                Set Form = Nothing
            End If
        Else
            MsgBox "Invalid Login Password !", vbExclamation, "DAS Version 4.0, Login"
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
    Else
        MsgBox "Invalid Login Password !", vbExclamation, "DAS Version 4.0, Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
Else
    MsgBox "Invalid Login Password !", vbExclamation, "DAS Version 4.0, Login"
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub
Private Sub Check_Backup_Status()
' Backup Checkup and Message Module - added by PNR
Dim RST_DIR As New ADODB.Recordset
If RST_DIR.State = 1 Then RST_DIR.Close
RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'BKP' .AND. DIR_IDX LIKE 'LOG'", con, adOpenKeyset, adLockPessimistic
If RST_DIR.EOF = True Then
    RST_DIR.AddNew
    RST_DIR("DIR_IDY") = "BKP"
    RST_DIR("DIR_IDX") = "LOG"
    RST_DIR("DIR_ALS") = ""
    RST_DIR("DIR_MSG") = ""
    RST_DIR("DIR_DAT") = Format(Date, "MM/DD/YYYY")
    RST_DIR.Update
End If
BKP_DAYS = DateDiff("D", RST_DIR("DIR_DAT"), Date)
If Trim(RST_DIR("DIR_MSG")) <> "" Then
    If Format(Date, "DD/MM/YYYY") <> Format(RST_DIR("DIR_DAT"), "DD/MM/YYYY") Then
        BKP_DAYS = DateDiff("D", RST_DIR("DIR_DAT"), Date)
'        MsgBox "Last Data Backup Date : " & Format(RST_DIR("DIR_DAT"), "DD/MM/YYYY") & " in Folder " & Trim(RST_DIR("DIR_MSG")) & " for " & STR(Val(BKP_DAYS)) & " Days!", vbInformation, "DAS Version 4.0"
    End If
    '"You Didn't Backup for " & Abs(Val(BKP_DAYS)) & " Days!"
Else
    MsgBox "No Data Backup Found! Take Backup From System -> Database -> Complete Backup Menu Option!", vbInformation, "DAS Version 4.0"
End If

End Sub
Private Sub Check_Shutdown_Status()
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'"), con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    If Trim(DIR_RS("DIR_MSG")) = "PROPER SHUTDOWN" Then
        DIR_RS("DIR_MSG") = ""
        DIR_RS("DIR_DAT") = Date
        DIR_RS.Update
    Else
        MsgBox "System was closed Improperly last time. " & Chr(13) & Chr(13) & "Please note Improper Shutdown could corrupt DAS data-even you can loose last saved transactions. ", vbOKOnly + vbCritical, "DAS Version 4.0 - Critical Error"
        ' & Chr(13) & "From System -> Database -> Validate Database, Click on Batch Repair Utility (to correct). "
        DIR_RS("DIR_MSG") = "IMPROPER SHUTDOWN"
        DIR_RS("DIR_DAT") = Date
        DIR_RS("DIR_ALS") = ""
        DIR_RS.Update
        DIR_RS.Close
        SHT_RSN.Show 1
'        M_FRM_QUK_RST.Show 1
        MsgBox "Please follow following instructions." & Chr(13) & "1. Verify your last saved transaction and validate the application by entering new transaction and recalling the same." & Chr(13) & "2. In case of corruption of last saved data please restore your last backup.(from System-Database-Restore DAS Autobackup) " & Chr(13) & "3. Exit the DAS and Reindex the Database" & Chr(13) & "4. Open DAS, Change the Shutdown flag to 'Yes' in System-ShutdownStatus" & Chr(13) & "Note: Please exit the DAS by clicking on File->Exit or press Ctrl+X.", vbInformation, "DAS Version 4.0 - Important System Message"
    End If
End If
End Sub
Private Sub Form_Activate()
TXTUSERNAME.SetFocus
End Sub
Private Sub Form_Load()
checkdongle
If rs.State = 1 Then rs.Close
rs.Open "SELECT USR_NME FROM USR GROUP BY USR_NME", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
TXTUSERNAME.AddItem Trim(rs!usr_nme)
rs.MoveNext
Next
DTPicker1.MaxDate = Date
DTPicker1.MinDate = Date
DTPicker1 = Date
DTPicker2 = Time
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOk.SetFocus
End If
If KeyAscii = 39 Then
   KeyAscii = 0
End If
End Sub
Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(TXTUSERNAME)) = 0 Then
        MsgBox "UserName Should Not Be Blank, Give UserName !", vbInformation, "DAS Version 4.0"
        TXTUSERNAME.SetFocus
    Else
        txtPassword.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End If
End Sub
Private Sub txtUserName_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub
Private Sub DISABLEMENU()
On Error GoTo LBL
Dim CNT As Integer
CNT = 0
For A = 1 To M_MDI.Controls.count - 1
If A <> 258 Then
    If TypeOf M_MDI.Controls(A) Is Menu And Trim(M_MDI.Controls(A).name) = "M_HLP" Then
        For K = A To M_MDI.Controls.count - 1
            M_MDI.Controls(K).Enabled = True
        Next
        Exit For
    End If
    If TypeOf M_MDI.Controls(A) Is Menu Then 'Or TypeOf M_MDI.Controls(A) Is MenuItem Then
        M_MDI.Controls(A).Enabled = False
    End If
    'chnage
    If TypeOf M_MDI.Controls(A) Is Menu Then 'Or TypeOf M_MDI.Controls(A) Is Menus Or TypeOf M_MDI.Controls(A) Is MenuItems Then  'TypeOf M_MDI.Controls(A) Is MenuItem Or
        CNT = CNT + 1
    End If
End If
Next
'MsgBox "No of Controls " & CNT
Exit Sub
LBL:
If err.Number = 387 Then
Resume Next
End If
End Sub

