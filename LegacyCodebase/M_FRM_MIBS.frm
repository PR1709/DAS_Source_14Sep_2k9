VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form M_FRM_MIBS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Mobile Data XChange"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "M_FRM_MIBS.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7395
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   345
      Left            =   5955
      TabIndex        =   10
      Top             =   15
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6390
      Top             =   4800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   420
      Left            =   4995
      TabIndex        =   9
      Top             =   45
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   630
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5175
      Top             =   4815
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   390
      Left            =   4740
      TabIndex        =   7
      Top             =   600
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Update DAS"
      Height          =   480
      Left            =   5205
      TabIndex        =   3
      Top             =   2685
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   480
      Left            =   5205
      TabIndex        =   4
      Top             =   3195
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Retailers Data"
      Height          =   480
      Left            =   5220
      TabIndex        =   0
      Top             =   1635
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "DAS Mobile"
      Height          =   2940
      Left            =   4965
      TabIndex        =   5
      Top             =   1132
      Width           =   1995
      Begin VB.CommandButton Command4 
         Caption         =   "&Products Data"
         Height          =   510
         Left            =   240
         TabIndex        =   1
         Top             =   1020
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   15
      Picture         =   "M_FRM_MIBS.frx":0442
      ScaleHeight     =   4995
      ScaleWidth      =   4440
      TabIndex        =   2
      Top             =   75
      Width           =   4500
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All Trademarks and Copyrights Acknowledged"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   15
         TabIndex        =   6
         Top             =   4800
         Width           =   3255
      End
   End
   Begin MSComctlLib.ProgressBar PBar2 
      Height          =   390
      Left            =   4755
      TabIndex        =   8
      Top             =   4245
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "M_FRM_MIBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Command1.Enabled = False
Dim HHT_INC As String
Dim SFLAG As Boolean
SFLAG = True
'Command2.Enabled = False
Dim TOT_REC, rec_len, pb1 As Integer
Dim RESULT As Variant
Dim in_str, input_data, FILE_NAME, t1 As Variant
TOT_REC = 0
rec_len = 0
PBar1.Value = 0
PBar1.Max = 100
PBar1.Min = 0
MSComm1.InBufferCount = 0
Call settings
FILE_NAME = CURDIR & "\HHT\INL_STR.TXT"
Do
    dumm = DoEvents()
    COM_ERR
    If MSComm1.PortOpen = True Then  'CHANGE BY PKB
    
        If MSComm1.Input = "$" Then
            MSComm1.Output = "@"
            Exit Do
        ElseIf MSComm1.Input = "^" Then
        MSComm1.Output = "`"
        Command1.Enabled = True
    '    Command1.SetFocus
      
        SFLAG = False
       MsgBox "Process failed due to invalid(Product) selection on Hand Held!", vbCritical, "DAS Version 4.0"
       If MSComm1.Input = "$" Then
       Command1.SetFocus
        Exit Do
        End If
        End If
    Else 'CHANGE BY PKB
    Exit Sub
   ' MsgBox "Connection Failed", vbCritical, "DAS Version 4.0" 'CHANGE BY PKB
    End If 'CHANGE BY PKB
   
Loop
If SFLAG = True Then
MSComm1.InBufferCount = 0
For I = 1 To 10
    Do
    Loop While MSComm1.InBufferCount = 0
    HHT_INC = HHT_INC & MSComm1.Input
Next I
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from HHT where HHT_INC like '" & (HHT_INC) & "'", PNL, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
For I = 1 To 4
    Do
    Loop While MSComm1.InBufferCount = 0
    TOT_REC = TOT_REC & MSComm1.Input
Next I

PBar1.Visible = True
PBar1.Max = Val(TOT_REC)
For I = 1 To 3
    Do
    Loop While MSComm1.InBufferCount = 0
    rec_len = rec_len & MSComm1.Input
Next I
delay (1)
MSComm1.Output = "!"
'Data
Open FILE_NAME For Output As #1
For I = 1 To TOT_REC
    in_str = ""
    input_data = ""
    For J = 1 To Val(rec_len)
        Do
        Loop While MSComm1.InBufferCount = 0
        input_data = input_data & MSComm1.Input
    Next J
    Debug.Print input_data
    Do
    Loop While MSComm1.Input <> Chr(10)
    Print #1, input_data
    pb1 = pb1 + 1
    PBar1.Value = pb1
    MSComm1.Output = "!"
Next I
Close #1
MsgBox "Retailers Data Received Successfully from Hand Held to PC!", vbInformation, "DAS Version 4.0"
PBar1.Visible = False
Command4.Enabled = True
'Command2.Enabled = True
Else
MsgBox "HHT NOT BELONGS TO THIS MACHINE!", vbInformation, "DAS Version 4.0"

End If
Else
 ' Unload Me
End If
End Sub

Private Sub Command1_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
Dim INL_STR As Variant
Dim RS1 As New ADODB.Recordset
Dim ACT_CHK As New ADODB.Recordset
Dim S, EPRD_UOM, ETOT_PCS, I, ENET_AMT, EMRP_RAT, NRS, TOTREC, LCST_PCG, CONFAC, SUBUNT  As Integer
Dim ss, INL, STKCHECK, STR As String
Dim f
Dim LINV_IDY, LCST_IDY, LGRS_AMT, LNET_AMT, LSMN_IDY, LTXN_TYP As String
Dim LROU_IDY, LUSR_IDY, LCSD_AMT, LSYS_DAT, LSYS_TIM, LSRN_AMT, EINV_IDY, EINV_DAT, EPRD_IDY, EPRD_QTY, EPRD_RAT, ERATAPP, ECOL_AMT, EORD_IDY, ETXN_TYP, EFST_FLG, EBAT_IDY, ESUB_UNT, ILDS_IDY As String
Dim rs As New ADODB.Recordset
Dim ERSI_IDY, EvAN_CLS, ECAS_AMT, ECHQ_AMT, EDEP_RCD, EDEP_REF As String
Dim PRS As New ADODB.Recordset
Dim LOGRS As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim ARS As New ADODB.Recordset
Dim PRS1 As New ADODB.Recordset
Dim PGRRS As New ADODB.Recordset
Dim BRS As New ADODB.Recordset
Dim PNLRS As New ADODB.Recordset
Dim ACTRS As New ADODB.Recordset
Dim IRTRS As New ADODB.Recordset
'' End of Record Set Declarations
Dim LOCRS As New ADODB.Recordset
Dim MARRS As New ADODB.Recordset
Dim CATRS As New ADODB.Recordset
Dim SMNRS As New ADODB.Recordset
Dim DIRRS As New ADODB.Recordset
Dim DIRRS1 As New ADODB.Recordset
Dim DIRRS2 As New ADODB.Recordset
Dim DIRRS3 As New ADODB.Recordset
Dim DIRRS4 As New ADODB.Recordset
Dim DIRRS5 As New ADODB.Recordset
Dim DIRRS6 As New ADODB.Recordset
Dim DIRRS7 As New ADODB.Recordset
Dim DIRRS8 As New ADODB.Recordset
Dim DIRRS9 As New ADODB.Recordset
Dim SHLRS As New ADODB.Recordset
Dim PRDRS As New ADODB.Recordset
Dim INLRS As New ADODB.Recordset
Dim CHKSHL As New ADODB.Recordset
   
Dim ECST_NME, ECST_AD1, ECST_AD2, ECST_AD3, ECSTAD4, ECST_CST, ECST_GST
Dim EPRD_NME, EBRD_NME, EGRP_IDY, EGRP_CLS, EGRP_CAT, ELST_RAT, ELST_DIS, EBAT_NME As String

' File Total Records, File Name and Input Data Declarations
Dim TOT_REC As Integer
Dim FILE_NAME As String
Dim input_data As Variant
' End of File Declarations
If CHECKTB("INL.DBF,INV.DBF,PNL.DBF,PNV.DBF,SNV.DBF,VAN.DBF,IRT.DBF,ACT.DBF,LOG.DBF") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & "DAS Version 4.0"
    Exit Sub
End If
Dim TOTCASAMT, TOTCHQAMT, TOTDEPRCD, TOTDEPREF As Integer
TOTCASAMT = 0
TOTCHQAMT = 0
TOTDEPRCD = 0
TOTDEPREF = 0
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists(CURDIR & "\HHT\INL_STR.TXT") = True Then
FILE_NAME = CURDIR & "\HHT\INL_STR.TXT"
Open FILE_NAME For Input As #1
If EOF(1) Then
    MsgBox "No Bill Value Records for Import!", vbCritical, "DAS Version 4.0"
    Exit Sub
End If
    Do
        Line Input #1, input_data
        TOT_REC = TOT_REC + 1
    Loop While Not EOF(1)
Close #1
Else
   MsgBox "Retailers Data Not Available! Export Retailers Data from Hand Held!", vbCritical, "DAS Version 4.0"
   Exit Sub
End If
Dim LGOD_CAS, LROD_CAS, P As Integer
Dim FSO As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
Set fil1 = FSO.GetFile(CURDIR & "\HHT\INL_STR.TXT")
Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
For I = 0 To TOT_REC - (P + 1)
    INL_STR = TXT_STREAM.ReadLine
    ENET_AMT = Trim(Mid(INL_STR, 156, 10))
    EDEP_REF = Trim(Mid(INL_STR, 228, 10))
    EDEP_RCD = Trim(Mid(INL_STR, 248, 10))
    ECAS_AMT = Trim(Mid(INL_STR, 186, 10))
    ECHQ_AMT = Trim(Mid(INL_STR, 196, 10))
    If Val(ENET_AMT) > 0 Or Val(EDEP_REF) > 0 Or Val(EDEP_RCD) > 0 Or Val(ECAS_AMT) > 0 Or Val(ECHQ_AMT) Then
    If Val(EDEP_REF) > 0 Then
    ECAS_AMT = Val(ECAS_AMT) - Val(EDEP_REF)
    Else
    ECAS_AMT = Val(ECAS_AMT)
    End If
    LGOD_AMT = Val(EDEP_RCD)
    ERSI_IDY = Trim(Mid(INL_STR, 106, 10))
    
    EDEP_RCD = Trim(Mid(INL_STR, 248, 10))
    
    LINV_IDY = Trim(Mid(INL_STR, 1, 10))
    EORD_IDY = Trim(Mid(INL_STR, 106, 10))
    EINV_DAT = Trim(Mid(INL_STR, 11, 10))
    EINV_DAT = Format(EINV_DAT, "DD/MM/YYYY")
    EINV_DAT = Format(EINV_DAT, "MM/DD/YYYY")
    LCST_IDY = Trim(Mid(INL_STR, 21, 10))
    LGRS_AMT = Trim(Mid(INL_STR, 116, 10))
    LCSD_AMT = Trim(Mid(INL_STR, 136, 10))
    wCSD_AMT = Mid(INL_STR, 136, 10)
    wGRS_AMT = Mid(INL_STR, 116, 10)
    EPKG_CST = Trim(Mid(INL_STR, 146, 10))
    LNET_AMT = Trim(Mid(INL_STR, 156, 10))
    LSMN_IDY = Trim(Mid(INL_STR, 61, 10))
    LROU_IDY = Trim(Mid(INL_STR, 96, 10))
    LSYS_DAT = Trim(Mid(INL_STR, 11, 10))
    LGOD_CAS = Mid(INL_STR, 238, 10)
    
    LROD_CAS = Mid(INL_STR, 218, 10)
    LSYS_TIM = Trim(Mid(INL_STR, 259, 10))
    LSRN_AMT = Trim(Mid(INL_STR, 126, 10))
    STKCHECK = Trim(Mid(EORD_IDY, 1, 3))
    ENET_AMT = Trim(Mid(INL_STR, 156, 10))
    EPTY_IDY = Trim(Mid(INL_STR, 21, 10))
    Close
    TOTCASAMT = TOTCASAMT + Val(ECAS_AMT)
    TOTCHQAMT = TOTCHQAMT + Val(ECHQ_AMT)
    TOTDEPRCD = TOTDEPRCD + Val(EDEP_RCD)
    TOTDEPREF = TOTDEPREF + Val(EDEP_REF)
    
    
     'CHECK CONDITION FOR CANCELL BILL
    If ACT_CHK.State = 1 Then ACT_CHK.Close
    ACT_CHK.Open "SELECT * FROM ACT where ACT_IDY = '" & LCST_IDY & "'", CON, adOpenKeyset, adLockPessimistic
    'If ACT_CHK.RecordCount > 0 And LINV_IDY <> "COLLECTION" Then
    If ACT_CHK.RecordCount > 0 Then
    '= START OF VAN UPDATE ='
        
    If rs.State = 1 Then rs.Close
    If MOD1 = True Then
    End If
    
    rs.Open "select * from VAN WHERE INV_IDY = '" & EORD_IDY & "' ", CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If rs2.State = 1 Then rs2.Close
        rs2.Open "select * from inl where inv_idy like '" & Trim(LINV_IDY) & "' and txn_typ like 'SAL'", CON, adOpenKeyset, adLockPessimistic
        If Not rs2.RecordCount > 0 Then
'            rs("DIS_AMT") = rs("DIS_AMT") - LCSD_AMT
'        End If
                rs.MoveFirst
                
                While Not rs.EOF
                    If Not rs("VAN_CLS") = "C" Then
                        rs("CAS_AMT") = 0
                        rs("CHQ_AMT") = 0
                        rs("DEP_RCD") = 0
                        rs("DEP_REF") = 0
                        rs("CAS_AMT") = TOTCASAMT
                        rs("CHQ_AMT") = TOTCHQAMT
                        rs("DEP_RCD") = TOTDEPRCD
                        rs("DEP_REF") = TOTDEPREF
                        rs("DIS_AMT") = rs("DIS_AMT") + LCSD_AMT 'CHANGED FOR DISPLAYING DISCOUNT
                        rs.Update
                    End If
                    rs.MoveNext
                Wend
            End If
    Else
    End If
    '= END OF VAN UPDATE ='
 
    '= START OF INL UPDATE ='
    If LINV_IDY <> "COLLECTION" Then 'START OF TO AVOID COLLECTION CONDITION
    LTXN_TYP = "SAL"
    If Len(EORD_IDY) > 0 And STKCHECK = "RSI" Then
        ESTK_UPD = "N"
    Else
        ESTK_UPD = "Y"
    End If
    ESTK_TYP = "Salable"
    EFST_FLG = "N"
    ETXN_TYP = "SAL"
    
    
    If MOD1 = True Then
    End If
    If LOGRS.State = 1 Then LOGRS.Close
    LOGRS.Open "SELECT * FROM LOG", CON, adOpenKeyset, adLockPessimistic
    LOGRS.MoveLast
    LUSR_IDY = LOGRS(0)
       
    If rs.State = 1 Then rs.Close
    If ARS.State = 1 Then ARS.Close
    If MOD1 = True Then
    End If
    
    STR = "Sales"
    ARS.Open "select ACT_IDY from ACT WHERE GRP_CLF = '" & STR & "'  ", CON, adOpenKeyset, adLockPessimistic
    If ARS.RecordCount > 0 Then
        M_SAL_IDY = ARS(0)
    Else
        M_SAL_IDY = Space(10)
    End If
    EACT_IDY = M_SAL_IDY
    If MOD1 = True Then
    End If
    If rs2.State = 1 Then rs2.Close
    rs2.Open "select * from inl where inv_idy like '" & Trim(LINV_IDY) & "' and txn_typ like 'SAL'", CON, adOpenKeyset, adLockPessimistic
    If Not rs2.RecordCount > 0 Then
        rs2.AddNew
        rs2("INV_IDY") = LINV_IDY
        rs2("INV_DAT") = EINV_DAT
        rs2("ORD_IDY") = EORD_IDY
        rs2("ACT_IDY") = M_SAL_IDY
        rs2("CST_IDY") = LCST_IDY
        rs2("GRS_AMT") = LGRS_AMT
        If Not Val(LCSD_AMT) > 0 Then
            rs2("CSD_AMT") = "0.00"
            rs2("CSD_PCG") = "0.00"
        Else
            rs2("CSD_AMT") = LCSD_AMT
        'If Mid(wCSD_AMT, 9, 2) > 0 Then
            LCSD_PCG = LCSD_AMT / LGRS_AMT
            LCSD_PCG = Val(LCSD_PCG) * 100
            rs2("CSD_PCG") = Mid(LCSD_PCG, 1, 7)
        'End If
 
        End If
        
        rs2("NET_AMT") = LNET_AMT
        rs2("SMN_IDY") = LSMN_IDY
        rs2("TXN_TYP") = LTXN_TYP
        rs2("ROU_IDY") = LROU_IDY
        rs2("USR_IDY") = LUSR_IDY
        rs2("SYS_DAT") = EINV_DAT
        rs2("SYS_TIM") = LSYS_TIM
        rs2("SRN_AMT") = LSRN_AMT
        
        rs2("DCH_IDY") = "         "
        rs2("LRN_IDY") = "         "
        rs2("WBL_IDY") = "          "
        rs2("VCH_IDY") = "          "
        rs2("SCH_IDY") = "          "
        rs2("LRN_AMT") = "0.00"
        rs2("SPD_PCG") = "0.00"
        rs2("SPD_AMT") = "0.00"
        rs2("EXD_PCG") = "0.00"
        rs2("EXD_AMT") = "0.00"
        rs2("TAX_PCG") = "0.00"
        rs2("TAX_AMT") = "0.00"
        rs2("FRT_AMT") = "0.00"
        rs2("LRN_DAT") = EINV_DAT
        rs2("WBL_DAT") = EINV_DAT
        rs2("PKG_CST") = EPKG_CST
        rs2("TTL_CAS") = "0.00"
        rs2("TXN_RMK") = "         "
        rs2("TXN_TPT") = "         "
        rs2("ROU_OFF") = "0.00"
        rs2("CRT_DAY") = "0"
        rs2("DUE_DAT") = EINV_DAT
        rs2("TXN_RM1") = "          "
        rs2("TXN_RM2") = "         "
        rs2("SCH_CHK") = "Y"
        rs2("SYS_DAT") = EINV_DAT
        rs2.Update
    Else
        'MsgBox "No Records for Update"
    End If
    '= END OF INL UPDATE ='

    '= START OF IRT UPDATE ='
    If Len(EINL_DAT) > 0 Then
       ESTK_UPD = "N"
    Else
       ESTK_UPD = "Y"
    End If
        ESTK_TYP = "Salable"
        EFST_FLG = "N"
        ETXN_TYP = "SAL"
        ESTK_TYP = Trim(ESTK_TYP)
        If IRTRS.State = 1 Then IRTRS.Close
        If MOD1 = True Then
        End If
        IRTRS.Open "SELECT * FROM IRT WHERE INV_IDY LIKE '" & Trim(LINV_IDY) & "'", CON, adOpenKeyset, adLockPessimistic
        If Not IRTRS.RecordCount > 0 Then
'        IRTRS.MoveLast
        IRTRS.AddNew
        IRTRS("INV_IDY") = LINV_IDY
        IRTRS("INV_DAT") = EINV_DAT
        IRTRS("NET_AMT") = LNET_AMT
        IRTRS("SMN_IDY") = LSMN_IDY
        IRTRS("TXN_TYP") = LTXN_TYP
        IRTRS("PTY_IDY") = EPTY_IDY
        IRTRS("ACT_IDY") = "          "
        IRTRS("DUE_DAT") = EINV_DAT
        IRTRS.Update
    Else
        'MsgBox "No Records for Update"
    End If
    '= END OF IRT UPDATE ='
    
    '= START OF PNL UPDATE ='
    If rs.State = 1 Then rs.Close
    If ARS.State = 1 Then ARS.Close
    If MOD1 = True Then
    End If
    If LOGRS.State = 1 Then LOGRS.Close
    If MOD1 = True Then
    End If
    LOGRS.Open "select * from LOG ", CON, adOpenKeyset, adLockPessimistic
    LOGRS.MoveLast
    EUSR_IDY = LOGRS(0)
    If ACTRS.State = 1 Then ACTRS.Close
    If MOD1 = True Then
    End If
    ACTRS.Open "select * from ACT WHERE ACT_IDY = '" & LCST_IDY & "'  ", CON, adOpenKeyset, adLockPessimistic
    If ACTRS.RecordCount > 0 Then
    ECST_ADD1 = ACTRS("ACT_AD1")
    ECST_ADD2 = ACTRS("ACT_AD2")
    ELOC_IDY = ACTRS("LOC_IDY")
    EMAR_IDY = ACTRS("MAR_IDY")
    ECAT_IDY = ACTRS("CAT_IDY")
    Else
    ECST_ADD1 = ""
    ECST_ADD2 = ""
    ELOC_IDY = ""
    EMAR_IDY = ""
    ECAT_IDY = ""
    End If
    If LOCRS.State = 1 Then LOCRS.Close
    If MOD1 = True Then
    End If
    LOCRS.Open "select * from LOC WHERE LOC_IDY = '" & ELOC_IDY & "'  ", CON, adOpenKeyset, adLockPessimistic
    If MARRS.State = 1 Then MARRS.Close
    If MOD1 = True Then
    End If
    MARRS.Open "select * from MAR WHERE MAR_IDY = '" & EMAR_IDY & "'  ", CON, adOpenKeyset, adLockPessimistic
    If CATRS.State = 1 Then CATRS.Close
    If MOD1 = True Then
    End If
    CATRS.Open "select * from CAT WHERE CAT_IDY = '" & ECAT_IDY & "'  ", CON, adOpenKeyset, adLockPessimistic
    If SMNRS.State = 1 Then SMNRS.Close
    If MOD1 = True Then
    End If
    SMNRS.Open "select * from SMN WHERE SMN_IDY = '" & LSMN_IDY & "'  ", CON, adOpenKeyset, adLockPessimistic
    
    If ARS.State = 1 Then ARS.Close
    If MOD1 = True Then
    End If
   
    STR = "Sales"
    ARS.Open "select ACT_IDY from ACT WHERE GRP_CLF = '" & STR & "'  ", CON, adOpenKeyset, adLockPessimistic
    M_SAL_IDY = ARS(0)
    EACT_IDY = M_SAL_IDY

    If DIRRS1.State = 1 Then DIRRS1.Close
    If MOD1 = True Then
    End If

    If DIRRS.State = 1 Then DIRRS.Close
    If MOD1 = True Then
    End If
    If DIRRS2.State = 1 Then DIRRS2.Close
    If MOD1 = True Then
    End If
    If DIRRS3.State = 1 Then DIRRS3.Close
    If MOD1 = True Then
    End If
    If DIRRS4.State = 1 Then DIRRS4.Close
    If MOD1 = True Then
    End If
    If DIRRS5.State = 1 Then DIRRS5.Close
    If MOD1 = True Then
    End If
    If DIRRS6.State = 1 Then DIRRS6.Close
    If MOD1 = True Then
    End If
    If DIRRS7.State = 1 Then DIRRS7.Close
    If MOD1 = True Then
    End If
    If DIRRS8.State = 1 Then DIRRS8.Close
    If MOD1 = True Then
    End If
    If DIRRS9.State = 1 Then DIRRS9.Close
    If MOD1 = True Then
    End If

    Dim STR1, STR2, STR3, STR4, STR5, STR6, STR7, STR8, STR9, ESYS_RM1, ESYS_RM2, ESYS_RM3, ESYS_RM4, ESYS_GST, ESYS_CST As String
    STR1 = "SYS"
    STR2 = "TT1"
    STR3 = "TT2"
    STR4 = "RM1"
    STR5 = "RM2"
    STR6 = "RM3"
    STR7 = "RM4"
    STR8 = "GST"
    STR9 = "CST"
    DIRRS.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR2 & "'", CON, adOpenKeyset, adLockPessimistic
    DIRRS1.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR3 & "'", CON, adOpenKeyset, adLockPessimistic
    DIRRS2.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR4 & "'", CON, adOpenKeyset, adLockPessimistic
    DIRRS3.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR5 & "'", CON, adOpenKeyset, adLockPessimistic
    DIRRS4.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR6 & "'", CON, adOpenKeyset, adLockPessimistic
    DIRRS5.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR7 & "'", CON, adOpenKeyset, adLockPessimistic
    DIRRS6.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR8 & "'", CON, adOpenKeyset, adLockPessimistic
    DIRRS7.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR9 & "'", CON, adOpenKeyset, adLockPessimistic
    EINS_NME = DIRRS(3)
    EINS_ADR = DIRRS1(3)
    ESYS_RM1 = DIRRS2(3)
    ESYS_RM2 = DIRRS3(3)
    ESYS_RM3 = DIRRS4(3)
    ESYS_RM4 = DIRRS5(3)
    ESYS_GST = DIRRS6(3)
    ESYS_CST = DIRRS7(3)
'    ESYS_RM1 = DIRRS8(3)

    EACT_IDY = M_SAL_IDY

    If rs2.State = 1 Then rs2.Close
    If MOD1 = True Then
    End If
    ECST_NME = " "
    ECST_AD1 = " "
    ECST_AD2 = " "
    ECST_AD3 = " "
    ECST_AD4 = " "
    ECST_CST = " "
    ECST_GST = " "
    ECST_ZIP = "0"
    ELOC_IDY = " "
    ECAT_IDY = ""
    EMAR_IDY = ""
    rs2.Open "select * from INL where INV_IDY = ''", CON, adOpenKeyset, adLockPessimistic
    If ACTRS.RecordCount > 0 Then
    ECST_NME = ACTRS("ACT_NME")
    ECST_AD1 = ACTRS("ACT_AD1")
    ECST_AD2 = ACTRS("ACT_AD2")
    ECST_AD3 = ACTRS("ACT_AD3")
    ECST_AD4 = ACTRS("ACT_AD4")
    ECST_CST = ACTRS("ACT_CST")
    ECST_GST = ACTRS("ACT_GST")
    ECST_ZIP = ACTRS("ACT_ZIP")
    ELOC_IDY = ACTRS("LOC_IDY")
    ECAT_IDY = ACTRS("CAT_IDY")
    EMAR_IDY = ACTRS("MAR_IDY")
    Else
    End If
    ESMN_NME = " "
    If SMNRS.RecordCount > o Then
    
    ESMN_NME = SMNRS("SMN_NME")
    Else
    ESMN_NME = " "
    End If
    If CATRS.RecordCount > 0 Then
    
    ECAT_NME = CATRS("CAT_NME")
    Else
    ECAT_NME = ""
    End If
    If MARRS.RecordCount > 0 Then
    
    EMAR_NME = MARRS("MAR_NME")
    Else
    EMAR_NME = ""
    End If
    If LOCRS.RecordCount > 0 Then
    
    ELOC_NME = LOCRS("LOC_NME")
    Else
    ELOC_NME = ""
    End If
    
    EACTIDY = LCST_IDY
    
    If CHECKTB("PNL.dbf") = 1 Then
    End If
    If PNLRS.State = 1 Then PNLRS.Close
    If MOD1 = True Then
    End If
    
    PNLRS.Open "SELECT * FROM PNL WHERE INV_IDY LIKE '" & Trim(LINV_IDY) & "' and txn_typ like 'SAL'", CON, adOpenKeyset, adLockPessimistic
    If Not PNLRS.RecordCount > 0 Then
'        PNLRS.MoveLast
        PNLRS.AddNew
        PNLRS("INV_IDY") = LINV_IDY
        PNLRS("INV_DAT") = EINV_DAT
        PNLRS("ORD_IDY") = EORD_IDY
        PNLRS("CST_IDY") = LCST_IDY
If Not ECST_NME = "" Then
PNLRS("CST_NME") = ECST_NME
Else
PNLRS("CST_NME") = "         "
End If


If Not ECST_AD1 = "" Then
  PNLRS("CST_AD1") = ECST_AD1
Else
  PNLRS("CST_AD1") = " "
End If

If Not ECST_AD2 = "" Then
  PNLRS("CST_AD2") = ECST_AD2
Else
  PNLRS("CST_AD2") = " "
End If
If Not ECST_AD3 = "" Then
  PNLRS("CST_AD3") = ECST_AD3
Else
  PNLRS("CST_AD3") = " "
End If

If Not ECST_AD4 = "" Then
  PNLRS("CST_AD4") = ECST_AD4
Else
  PNLRS("CST_AD4") = " "
End If
If Not ECST_CST = "" Then
PNLRS("CST_CST") = ECST_CST
Else
PNLRS("CST_CST") = " "
End If
        PNLRS("CST_GST") = ECST_GST
        PNLRS("CST_ZIP") = ECST_ZIP
        PNLRS("SMN_NME") = ESMN_NME
        PNLRS("CAT_NME") = ECAT_NME
        PNLRS("LOC_IDY") = ELOC_IDY
        PNLRS("CAT_IDY") = ECAT_IDY
        PNLRS("GRS_AMT") = LGRS_AMT
        PNLRS("CSD_AMT") = LCSD_AMT
        PNLRS("NET_AMT") = LNET_AMT
        PNLRS("SMN_IDY") = LSMN_IDY
        PNLRS("TXN_TYP") = LTXN_TYP
        PNLRS("ROU_IDY") = LROU_IDY
        PNLRS("USR_NME") = EUSR_IDY
        PNLRS("SYS_DAT") = EINV_DAT
        PNLRS("SYS_DAT") = EINV_DAT
        PNLRS("SYS_TIM") = LSYS_TIM
        PNLRS("SRN_AMT") = LSRN_AMT
        PNLRS("CSD_PCG") = "0.00"
        PNLRS("DCH_IDY") = "         "
        PNLRS("LRN_IDY") = "         "
        PNLRS("LRN_AMT") = "00"
        PNLRS("WBL_IDY") = "          "
        PNLRS("VCH_IDY") = "          "
        PNLRS("SCH_IDY") = "          "
        PNLRS("ACT_IDY") = EACT_IDY
        PNLRS("MAR_NME") = EMAR_NME
        PNLRS("SMN_NME") = ESMN_NME
        PNLRS("LOC_NME") = ELOC_NME
        'PNLRS("CAT_NME") = "              "
        PNLRS("SPD_PCG") = "00"
        PNLRS("SPD_AMT") = "000000000"
        PNLRS("EXD_PCG") = "00"
        PNLRS("EXD_CAL") = " "
        PNLRS("EXD_AMT") = "000000000000"
        PNLRS("TAX_PCG") = "0000"
        PNLRS("TAX_AMT") = "0000"
        PNLRS("FRT_AMT") = "0000"
        PNLRS("LRN_AMT") = "000000000000"
        PNLRS("PKG_CST") = EPKG_CST
        PNLRS("TTL_CAS") = "000000000000"
        PNLRS("ROU_OFF") = "0000"
        PNLRS("CRT_DAY") = "0000"
        PNLRS("SRN_AMT") = "000000000000"
        PNLRS("FRT_TYP") = " "
        PNLRS("DCH_IDY") = "          "
        PNLRS("DCH_DAT") = EINV_DAT
        PNLRS("LRN_IDY") = "       "
        PNLRS("LRN_DAT") = EINV_DAT
        PNLRS("WBL_IDY") = "       "
        PNLRS("WBL_DAT") = EINV_DAT
        PNLRS("TXN_RMK") = "       "
        PNLRS("TXN_TPT") = "       "
        PNLRS("TAX_ACT") = "       "
        PNLRS("ROU_ACT") = "       "
        PNLRS("FRT_ACT") = "       "
        PNLRS("DUE_DAT") = EINV_DAT
        PNLRS("DES_MOD") = " "
        PNLRS("STO_RMK") = "    "
        'PNLRS("ROU_IDY") = "       "
        PNLRS("TXN_RM1") = "       "
        PNLRS("TXN_RM2") = "       "
        PNLRS("VCH_IDY") = "       "
        PNLRS("VCH_DAT") = EINV_DAT
        PNLRS("SCH_CHK") = " "
        PNLRS("SCH_IDY") = "       "
        PNLRS("FRM_DAT") = EINV_DAT
        PNLRS("TO_DAT") = EINV_DAT
        PNLRS("DIS_CHK") = " "
        PNLRS("MAR_IDY") = EMAR_IDY
        PNLRS("INS_NME") = EINS_NME
        PNLRS("INS_ADR") = EINS_ADR
        ESYS_RM1 = Trim(ESYS_RM1)
        If ESYS_RM1 = "" Then
        PNLRS("SYS_RM1") = "    "
        Else
        PNLRS("SYS_RM1") = ESYS_RM1
        End If
        ESYS_RM2 = Trim(Mid(ESYS_RM2, 40))
        ESYS_RM3 = Trim(Mid(ESYS_RM3, 40))
        ESYS_RM4 = Trim(Mid(ESYS_RM4, 40))
        ESYS_GST = Trim(Mid(ESYS_GST, 30))
        ESYS_CST = Trim(Mid(ESYS_CST, 30))
        If ESYS_RM2 = "" Then
        PNLRS("SYS_RM2") = "    "
        Else
        PNLRS("SYS_RM2") = ESYS_RM2
        End If
        If ESYS_RM3 = "" Then
        PNLRS("SYS_RM3") = "    "
        Else
        PNLRS("SYS_RM3") = ESYS_RM3
        End If
        If ESYS_RM4 = "" Then
        PNLRS("SYS_RM4") = "    "
        Else
        PNLRS("SYS_RM4") = ESYS_RM4
        End If
        If ESYS_GST = "" Then
        PNLRS("INS_GST") = "     "
        Else
        PNLRS("INS_GST") = ESYS_GST
        End If
        If ESYS_CST = "" Then
        PNLRS("INS_CST") = "    "
        Else
        PNLRS("INS_CST") = ESYS_CST
        End If
        PNLRS.Update
    Else
        'MsgBox "No Records to Update"
    End If
    'End If
    '= END OF PNL TABLE =
    
    '= START OF SNV = '
    'Dim SHLRS As New ADODB.Recordset
    ERATAPP = "B"
    If rs.State = 1 Then rs.Close
   ' If RS1.State = 1 Then RS1.Close

    If MOD1 = True Then
    End If
    If SHLRS.State = 1 Then SHLRS.Close
    If MOD1 = True Then
    End If
    SHLRS.Open "select * from SHL", CON, adOpenKeyset, adLockPessimistic
    If SHLRS.RecordCount > 0 Then
    If SHLRS.RecordCount > 1 Then  ' START OF SHL CONDITION
        SHLRS.MoveFirst
        ESHL_IDY = SHLRS("SHL_IDY")
        EPRD_IDY = SHLRS("PRD_IDY")
    ElseIf SHLRS.RecordCount = 1 Then
        ESHL_IDY = SHLRS("SHL_IDY")
        EPRD_IDY = SHLRS("PRD_IDY")
    Else
    End If
    Dim SNVRS1 As New ADODB.Recordset
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from SNV where INV_IDY like '" & Trim(EORD_IDY) & "'AND CST_IDY LIKE '" & LCST_IDY & "' ", CON, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    If PRDRS.State = 1 Then PRDRS.Close
    If MOD1 = True Then
    End If
    PRDRS.Open "select * from PRD where PRD_IDY like '" & Trim(EPRD_IDY) & "'", CON, adOpenKeyset, adLockPessimistic
    If INLRS.State = 1 Then INLRS.Close
    If MOD1 = True Then
    End If
    INLRS.Open "select * from INL where INV_IDY like '" & Trim(LINV_IDY) & "' ", CON, adOpenKeyset, adLockPessimistic
    If CHKSHL.State = 1 Then CHKSHL.Close
    If MOD1 = True Then
    End If
    CHKSHL.Open "SELECT * FROM SHL WHERE PRD_IDY LIKE '" & EPRD_IDY & " '", CON, adOpenKeyset, adLockPessimistic
   Dim ETXN_IDY As String
   Dim IDY As Integer
'    If Not RS1.RecordCount > 0 Then
    If Val(LGOD_CAS) > 0 Then
    Dim IDYRS As New ADODB.Recordset
    If IDYRS.State = 1 Then IDYRS.Close
    If MOD1 = True Then
    End If
    IDY = 0
    IDYRS.Open "SELECT * FROM IDY", CON, adOpenKeyset, adLockPessimistic
    IDY = Val(IDYRS("IDY_SHT")) + 1
   ' Dim IDYLEN As Integer
    ETXN_IDY = IDY
    While Len(Trim(ETXN_IDY)) < 7
    ETXN_IDY = "0" & ETXN_IDY
    Wend
    ETXN_IDY = "SHT" & ETXN_IDY

    IDYRS("IDY_SHT") = IDY
    IDYRS.Update
'    RS1.MoveLast
    RS1.AddNew
    RS1("INV_IDY") = EORD_IDY
    RS1("INV_DAT") = EINV_DAT
    If SHLRS.RecordCount > 0 Then
        RS1("PRD_IDY") = SHLRS("PRD_IDY")
    Else
        RS1("PRD_IDY") = EPRD_IDY
    End If
    RS1("TXN_TYP") = "DPI"
    RS1("SHL_IDY") = SHLRS("SHL_IDY")
    RS1("I_SHL_NOS") = "0"
    RS1("I_PCK_NOS") = "0"
    RS1("D_SHL_NOS") = Mid(LGOD_CAS, 1, 6)
    ESUBUNT = Mid(LGOD_CAS, 8, 3)
    ER_PCK_NOS = RS1("D_SHL_NOS") * PRDRS("SUB_UNT")
    ER_PCK_NOS = ER_PCK_NOS + ESUBUNT
    RS1("D_PCK_NOS") = ER_PCK_NOS
    RS1("DP_SHL_NOS") = Mid(LGOD_CAS, 1, 6)
    RS1("DP_PCK_NOS") = ER_PCK_NOS
    RS1("DEP_AMT") = Val(LGOD_AMT)
    RS1("TXN_DAT") = EINV_DAT
    RS1("STK_UPD") = "Y"
    RS1("CST_IDY") = EPTY_IDY
    RS1("PRD_UPC") = PRDRS("SUB_UNT")
    RS1("PRD_PCK") = PRDRS("PRD_PCK")
    RS1("R_SHL_NOS") = "0"
    RS1("R_PCK_NOS") = "0"
    RS1("BRK_SHL") = "0"
    RS1("BRK_PCK") = "0"
    RS1("TXN_IDY") = ETXN_IDY
    RS1("smn_idy") = LSMN_IDY
    RS1.Update
    Else
    End If
    'End If    'END OF SHL CONDITION
    If Val(LROD_CAS) > 0 Then
    If RS1.State = 1 Then RS1.Close
'    RS1.Open "select * from SNV where INV_IDY like '" & Trim(EORD_IDY) & "'", CON, adOpenKeyset, adlockpessimistic
    RS1.Open "select * from SNV where INV_IDY like '" & Trim(EORD_IDY) & "'AND CST_IDY LIKE '" & LCST_IDY & "' ", CON, adOpenKeyset, adLockPessimistic

    'Dim IDYRS As New ADODB.Recordset
    IDY = 0
    If IDYRS.State = 1 Then IDYRS.Close
    If MOD1 = True Then
    End If
    IDYRS.Open "SELECT * FROM IDY", CON, adOpenKeyset, adLockPessimistic
    IDY = Val(IDYRS("IDY_SHT")) + 1
    ETXN_IDY = IDY
    While Len(Trim(ETXN_IDY)) < 7
    ETXN_IDY = "0" & ETXN_IDY
    Wend
    ETXN_IDY = "SHT" & ETXN_IDY
    IDYRS("IDY_SHT") = IDY
    IDYRS.Update
    If RS1.RecordCount > 0 Then
    RS1.MoveLast
    End If
    RS1.AddNew
        RS1("INV_IDY") = EORD_IDY
        RS1("INV_DAT") = EINV_DAT
        If SHLRS.RecordCount > 0 Then
        RS1("PRD_IDY") = SHLRS("PRD_IDY")
        Else
        RS1("PRD_IDY") = EPRD_IDY
        End If
        RS1("TXN_TYP") = "DPO"
        RS1("SHL_IDY") = SHLRS("SHL_IDY")
        RS1("I_SHL_NOS") = "0"
        RS1("I_PCK_NOS") = "0"
        RS1("D_SHL_NOS") = Mid(LROD_CAS, 1, 6)
        ESUBUNT = Mid(LROD_CAS, 8, 3)
        ER_PCK_NOS = RS1("D_SHL_NOS") * PRDRS("SUB_UNT")
        ER_PCK_NOS = ER_PCK_NOS + ESUBUNT
        RS1("D_PCK_NOS") = ER_PCK_NOS
        RS1("DP_SHL_NOS") = RS1("D_SHL_NOS")
        RS1("DP_PCK_NOS") = ER_PCK_NOS
        RS1("DEP_AMT") = EDEP_REF
        RS1("TXN_DAT") = EINV_DAT
        RS1("STK_UPD") = "Y"
        RS1("CST_IDY") = EPTY_IDY
        RS1("PRD_UPC") = PRDRS("SUB_UNT")
        RS1("PRD_PCK") = PRDRS("PRD_PCK")
        RS1("R_SHL_NOS") = "0"
        RS1("R_PCK_NOS") = "0"
        RS1("BRK_SHL") = "0"
        RS1("BRK_PCK") = "0"
        RS1("smn_idy") = LSMN_IDY
        RS1("TXN_IDY") = ETXN_IDY
        RS1.Update
    Else
    End If

'    ELSE
 '    End If
'
    '= END OF SNV WITH GOD_CAS =
    Else
    End If          ' END OF SHLCONDITION
    End If
    End If
    End If ' END OF TO AVOID COLLECTION CONDITION
    TXT_STREAM.Skip (0)
    Next I
    TXT_STREAM.Close
    '= END OF INL_STR.TXT UPDATE =' End of INL update from Hand Held Generated Text File

    '= START OF INV_STR.TXT UPDATE ='
    TOT_REC = 0
    FILE_NAME = CURDIR & "\HHT\INV_STR.TXT"
    Open FILE_NAME For Input As #1
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Dim INVFSO As New FileSystemObject, txtfile1, fil1INV As File, ts1 As TextStream
    Set fil1INV = INVFSO.GetFile(CURDIR & "\HHT\INV_STR.TXT")
    Set INV_STREAM = fil1INV.OpenAsTextStream(ForReading)
    For I = 0 To TOT_REC - 1
        INV_STR = INV_STREAM.ReadLine
        EINV_IDY = Trim(Mid(INV_STR, 1, 10))
        EBAT_IDY = Trim(Mid(INV_STR, 11, 10))
        EPRD_IDY = Trim(Mid(INV_STR, 21, 10))
        EPRD_QTY = Trim(Mid(INV_STR, 52, 7))
        EPRD_RAT = Trim(Mid(INV_STR, 69, 10))
        ECOL_AMT = Trim(Mid(INV_STR, 79, 10))
        ERSHL_NOS = Trim(Mid(INV_STR, 89, 7))
        EPRD_UNT = Trim(Mid(INV_STR, 104, 3))
        EPRD_PCS = Trim(Mid(INV_STR, 107, 3))
        ETOT_PCS = Trim(Mid(INV_STR, 99, 5))
        REFSTK_TYP = Trim(Mid(INV_STR, 51, 1))
        ILDS_IDY = "         "
        ERATAPP = "B"
    Close
    Dim PRD_CHK As New ADODB.Recordset
    If PRD_CHK.State = 1 Then PRD_CHK.Close
    PRD_CHK.Open "SELECT * FROM PRD where PRD_IDY = '" & EPRD_IDY & "'", CON, adOpenKeyset, adLockPessimistic
    If Not PRD_CHK.EOF And EINV_IDY <> "COLLECTION" Then
    
   ' Dim ESUBUNT As Integer
    ER_SHL_NOS = Mid(ERSHL_NOS, 1, 3)
    ESUBUNT = Mid(ERSHL_NOS, 4, 3)
   
    If Len(EORD_IDY) > 0 And STKCHECK = "RSI" Then
        ESTK_UPD = "N"
        Else
        ESTK_UPD = "Y"
    End If
        ESTK_TYP = "Salable"
        EFST_FLG = "N"
        ETXN_TYP = "SAL"
    If REFSTK_TYP = "S" Then
       ESTK_TYP = "Salable"
    ElseIf REFSTK_TYP = "T" Then
       ESTK_TYP = "TakeBack"
       EPRD_RAT = "-" & EPRD_RAT
       Else
       ESTK_TYP = "Free"
       EPRD_RAT = "0.00"
    End If
    ESTK_TYP = Trim(ESTK_TYP)
    ERATAPP = "B"

    If CHECKTB("INV.dbf") = 1 Then
    End If

    If RS1.State = 1 Then RS1.Close

    If MOD1 = True Then
    End If
    RS1.Open "select * from INV where INV_IDY like '" & Trim(EINV_IDY) & "' and TXN_TYP like 'SAL' and BAT_IDY like '" & Trim(EBAT_IDY) & "' and STK_TYP like '" & Trim(ESTK_TYP) & "'", CON, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    PRS.Open "select * from PRD WHERE PRD_IDY like '" & EPRD_IDY & "'", CON, adOpenKeyset, adLockPessimistic
    EPRD_UOM = PRS("CON_FAC") / PRS("SUB_UNT")
    EPRD_UOM = EPRD_UOM * ETOT_PCS
    ESUB_UNT = PRS("SUB_UNT")
    TOTREC = RS1.RecordCount
    If Not RS1.RecordCount > 0 Then
        RS1.AddNew
        RS1("INV_IDY") = EINV_IDY
        RS1("BAT_IDY") = EBAT_IDY
        RS1("PRD_QTY") = EPRD_QTY
        RS1("PRD_IDY") = EPRD_IDY
        RS1("PRD_RAT") = EPRD_RAT
        RS1("COL_AMT") = ECOL_AMT
        RS1("BAT_IDY") = EBAT_IDY
        RS1("RAT_APP") = ERATAPP
        RS1("ORD_IDY") = EORD_IDY
        RS1("INV_DAT") = EINV_DAT
        RS1("STK_UPD") = ESTK_UPD
        RS1("STK_TYP") = ESTK_TYP
        RS1("FST_FLG") = EFST_FLG
        RS1("TXN_TYP") = ETXN_TYP
        RS1("PRD_UNT") = EPRD_UNT
        RS1("PRD_PCS") = EPRD_PCS
        RS1("TOT_PCS") = ETOT_PCS
        RS1("PRD_UOM") = EPRD_UOM
        RS1("SUB_UNT") = ESUB_UNT
        RS1.Update
     Else
        'MsgBox "No Records to Update"
    End If
    '= END OF INV UPDATE ='
    '= START OF PNV UPDATE ='

    ERATAPP = "B"
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    If MOD1 = True Then
    End If
    RS1.Open "select * from PNV where INV_IDY like '" & Trim(EINV_IDY) & "' and TXN_TYP like 'SAL' and BAT_IDY like '" & Trim(EBAT_IDY) & "' and STK_TYP like '" & Trim(ESTK_TYP) & "'", CON, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    PRS.Open "select * from PRD WHERE PRD_IDY = '" & EPRD_IDY & " ' ", CON, adOpenKeyset, adLockPessimistic
    If BRS.State = 1 Then BRS.Close
    If MOD1 = True Then
    End If
    If PRS1.State = 1 Then PRS1.Close
    If MOD1 = True Then
    End If
    EBRD_NME = PRS("BRD_NME")
    BRS.Open "select * from BRD WHERE PRD_IDY = '" & EPRD_IDY & " ' ", CON, adOpenKeyset, adLockPessimistic
    PRS1.Open "select * from PRD WHERE BRD_NME = '" & EBRD_NME & "'", CON, adOpenKeyset, adLockPessimistic
    EPRD_NME = PRS("PRD_NME")
    EGRP_IDY = PRS("GRP_IDY")
    If PGRRS.State = 1 Then PGRRS.Close
    If MOD1 = True Then
    End If
    PGRRS.Open "select * from PGR WHERE GRP_IDY = '" & EGRP_IDY & "'", CON, adOpenKeyset, adLockPessimistic
    EGRP_CAT = PRS("BRD_NME")
    EGRP_CLS = PGRRS("GRP_CLS")
    ELST_RAT = "0.00"
    ELST_DIS = "0.00"
    EBAT_NME = BRS("BAT_NME")
    EMRP_RAT = BRS("PRD_MRP")
    EPRD_UOM = PRS("CON_FAC") / PRS("SUB_UNT")
    EPRD_UOM = EPRD_UOM * ETOT_PCS
    ESUB_UNT = PRS("SUB_UNT")
    If Not RS1.RecordCount > 0 Then
        RS1.AddNew
        RS1("INV_IDY") = EINV_IDY
        RS1("BAT_IDY") = EBAT_IDY
        RS1("PRD_QTY") = EPRD_QTY
        RS1("PRD_IDY") = EPRD_IDY
        RS1("PRD_RAT") = EPRD_RAT
        RS1("COL_AMT") = ECOL_AMT
        RS1("BAT_IDY") = EBAT_IDY
        RS1("RAT_APP") = ERATAPP
        If Trim(Mid(EORD_IDY, 1, 3)) = "RSI" Then
        RS1("ORD_IDY") = EORD_IDY
        Else
        RS1("ORD_IDY") = "0"
        End If
        RS1("INV_DAT") = EINV_DAT
        RS1("STK_UPD") = ESTK_UPD
        RS1("STK_TYP") = ESTK_TYP
        RS1("FST_FLG") = EFST_FLG
        RS1("TXN_TYP") = ETXN_TYP
        RS1("PRD_UNT") = EPRD_UNT
        RS1("PRD_PCS") = EPRD_PCS
        RS1("TOT_PCS") = ETOT_PCS
        RS1("PRD_UOM") = EPRD_UOM
        RS1("SUB_UNT") = ESUB_UNT
        RS1("PRD_NME") = EPRD_NME
        RS1("GRP_IDY") = EGRP_IDY
        RS1("GRP_CLS") = EGRP_CLS
        RS1("GRP_CAT") = EGRP_CAT
        RS1("LST_RAT") = ELST_RAT
        RS1("LST_DIS") = ELST_DIS
        RS1("BAT_NME") = EBAT_NME
        RS1("MRP_RAT") = EMRP_RAT
        'RS1("GRP_CAT") = "        "
        RS1("RET_QTY") = "0.000"
        RS1("CPN_RAT") = "0.00"
        RS1("CPN_AMT") = "0.00"
        RS1("ADD_DIS") = "0.00"
        RS1("ADD_AMT") = "0.00"
        RS1("TRD_DIS") = "0.00"
        RS1("DIS_AMT") = "0.00"
        RS1("TRD_TAX") = "0.00"
        RS1("TAX_AMT") = "0.00"
        RS1("OPN_QTY") = "0.00"
        RS1("ISS_QTY") = "0.000"
        RS1("TOT_QTY") = "0.000"
        RS1("SCH_AMT") = "0.00"
        RS1("SCH_IDY") = " "
        RS1.Update
    Else
        'MsgBox "No Records to Update"
   End If
   
   '= START OF SNV UPDATE ='
    ERATAPP = "B"
    If rs.State = 1 Then rs.Close

    If SHLRS.State = 1 Then SHLRS.Close
    SHLRS.Open "select * from SHL where PRD_IDY like '" & Trim(EPRD_IDY) & "' ", CON, adOpenKeyset, adLockPessimistic
    Dim SNVRS As New ADODB.Recordset
    If SNVRS.State = 1 Then SNVRS.Close
    
    SNVRS.Open "select * from SNV where INV_IDY like '" & Trim(EINV_IDY) & "' AND PRD_IDY like '" & Trim(EPRD_IDY) & "'", CON, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close

    If PRDRS.State = 1 Then PRDRS.Close

    PRDRS.Open "select * from PRD where PRD_IDY like '" & Trim(EPRD_IDY) & "' ", CON, adOpenKeyset, adLockPessimistic
    If INLRS.State = 1 Then INLRS.Close

    INLRS.Open "select * from INL where INV_IDY like '" & Trim(EINV_IDY) & "' ", CON, adOpenKeyset, adLockPessimistic
    If CHKSHL.State = 1 Then CHKSHL.Close
    
    CHKSHL.Open "SELECT * FROM SHL WHERE PRD_IDY LIKE '" & EPRD_IDY & " '", CON, adOpenKeyset, adLockPessimistic
    If CHKSHL.RecordCount > 0 Then
    If Not SNVRS.RecordCount > 0 Then
        SNVRS.AddNew
        SNVRS("INV_IDY") = EINV_IDY
        SNVRS("INV_DAT") = EINV_DAT
        SNVRS("PRD_IDY") = EPRD_IDY
        SNVRS("TXN_TYP") = "SAL"
        SNVRS("SHL_IDY") = SHLRS("SHL_IDY")
        If REFSTK_TYP = "T" Then
        SNVRS("I_SHL_NOS") = "-" & EPRD_UNT
        SNVRS("I_PCK_NOS") = "-" & ETOT_PCS
        Else
         SNVRS("I_SHL_NOS") = EPRD_UNT
        SNVRS("I_PCK_NOS") = ETOT_PCS
        End If
        SNVRS("R_SHL_NOS") = ER_SHL_NOS
        
        ER_PCK_NOS = ER_SHL_NOS * PRDRS("SUB_UNT")
        ER_PCK_NOS = ER_PCK_NOS + ESUBUNT
        SNVRS("R_PCK_NOS") = ER_PCK_NOS
        SNVRS("D_SHL_NOS") = SNVRS("I_SHL_NOS") - SNVRS("R_SHL_NOS")
        SNVRS("D_PCK_NOS") = ETOT_PCS - ER_PCK_NOS
        SNVRS("TXN_DAT") = EINV_DAT
        SNVRS("TXN_DAT") = EINV_DAT
        SNVRS("TXN_IDY") = EORD_IDY
        If "EORDIDY" = "" Then
        SNVRS("STK_UPD") = "Y"
        Else
        SNVRS("STK_UPD") = "N"
        End If
        SNVRS("PRD_UPC") = PRDRS("SUB_UNT")
        SNVRS("PRD_PCK") = PRDRS("PRD_PCK")
        If INLRS.RecordCount > 0 Then
        SNVRS("CST_IDY") = INLRS("CST_IDY")
        Else
        SNVRS("CST_IDY") = " "
        End If
         If INLRS.RecordCount > 0 Then
        SNVRS("SMN_IDY") = INLRS("SMN_IDY")
        Else
        SNVRS("SMN_IDY") = " "
        End If
         If INLRS.RecordCount > 0 Then
        SNVRS("ROU_IDY") = INLRS("ROU_IDY")
        Else
        SNVRS("ROU_IDY") = " "
        End If
        SNVRS("PCK_IDY") = "          "
        SNVRS("DEP_AMT") = "0"
        SNVRS("DP_SHL_NOS") = "0"
        SNVRS("DP_PCK_NOS") = "0"
        SNVRS("BRK_SHL") = "0"
        SNVRS("BRK_PCK") = "0"
        SNVRS.Update
    Else
    End If
    Else
    End If
    End If
    INV_STREAM.Skip (0)
    Next I
'    Command10.Enabled = True
    MsgBox "Bill Information from Hand Held updated into DAS Successfully!", vbInformation, "DAS Version 4.0"
    INV_STREAM.Close
'Command4.Enabled = True
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Command4_Click()
Command4.Enabled = False
Dim SFLAG As Boolean
SFLAG = True
Dim TOT_REC, rec_len, pb1 As Integer
Dim in_str, input_data, FILE_NAME As Variant
TOT_REC = 0
rec_len = 0
PBar2.Value = 0
PBar2.Max = 100
PBar2.Min = 0
MSComm1.InBufferCount = 0
Call settings
FILE_NAME = CURDIR & "\HHT\INV_STR.TXT"
'Do
'    dumm = DoEvents()
'    COM_ERR
'    If MSComm1.Input = "$" Then
'        MSComm1.Output = "@"
'        Exit Do
'    End If
'    'delay (1)
'Loop

Do
    dumm = DoEvents()
    COM_ERR
    If MSComm1.PortOpen = True Then  'CHANGE BY PKB
    
    
   If MSComm1.Input = "^" Then
            MSComm1.Output = "~"
            Exit Do
        ElseIf MSComm1.Input = "$" Then
      MSComm1.Output = "`"
        Command4.Enabled = True
      '  Command4.SetFocus
      
        SFLAG = False
       MsgBox "Process failed due to invalid(Retailer)selection on Hand Held!", vbCritical, "DAS Version 4.0"
       If MSComm1.Input = "$" Then
       Command4.SetFocus
        Exit Do
        End If
    End If
'        If MSComm1.Input = "$" Then
'            MSComm1.Output = "@"
'            Exit Do
'        End If
    Else 'CHANGE BY PKB
    Exit Sub
   ' MsgBox "Connection Failed", vbCritical, "DAS Version 4.0" 'CHANGE BY PKB
    End If 'CHANGE BY PKB
Loop


MSComm1.InBufferCount = 0

For I = 1 To 10
    Do
    Loop While MSComm1.InBufferCount = 0
    HHT_INC = HHT_INC & MSComm1.Input
Next I
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from HHT where HHT_INC like '" & (HHT_INC) & "'", PNL, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then


For I = 1 To 4
    Do
    Loop While MSComm1.InBufferCount = 0
    TOT_REC = TOT_REC & MSComm1.Input
Next I
If SFLAG = True Then
PBar2.Visible = True
PBar2.Max = Val(TOT_REC)
For I = 1 To 3
    Do
    Loop While MSComm1.InBufferCount = 0
    rec_len = rec_len & MSComm1.Input
Next I
delay (1)
MSComm1.Output = "!"

'Data
Open FILE_NAME For Output As #1
For I = 1 To TOT_REC
    in_str = ""
    input_data = ""
    For J = 1 To Val(rec_len)
        Do
        Loop While MSComm1.InBufferCount = 0
        input_data = input_data & MSComm1.Input
    Next J
    Debug.Print input_data
    Do
    Loop While MSComm1.Input <> Chr(10)
    Print #1, input_data
    pb1 = pb1 + 1
    PBar2.Value = pb1
    MSComm1.Output = "!"
Next I
Close #1
MsgBox "Product Data Received Successfully from Hand Held to PC!", vbInformation, "DAS Version 4.0"
PBar2.Visible = False
Command3.Enabled = True
Command3.SetFocus

Else
MsgBox "HHT NOT BELONGS TO THIS MACHINE!", vbInformation, "DAS Version 4.0"

End If
End If
End Sub

Private Sub Command4_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Command5_Click()
Dim INL_STR As Variant
Dim RS1 As New ADODB.Recordset
Dim ACT_CHK As New ADODB.Recordset
Dim S, EPRD_UOM, ETOT_PCS, I, ENET_AMT, EMRP_RAT, NRS, TOTREC, CONFAC, SUBUNT As Integer
Dim ss, INL, STKCHECK, STR As String
Dim f
Dim LINV_IDY, LCST_IDY, LGRS_AMT, LCSD_AMT, LNET_AMT, LSMN_IDY, LTXN_TYP, LROU_IDY, LUSR_IDY, LSYS_DAT, LSYS_TIM, LSRN_AMT, EINV_IDY, EINV_DAT, EPRD_IDY, EPRD_QTY, EPRD_RAT, ERATAPP, ECOL_AMT, EORD_IDY, ETXN_TYP, EFST_FLG, EBAT_IDY, ESUB_UNT, ILDS_IDY As String
Dim rs As New ADODB.Recordset
Dim ERSI_IDY, EvAN_CLS, ECAS_AMT, ECHQ_AMT, EDEP_RCD, EDEP_REF As String
Dim PRS As New ADODB.Recordset
Dim LOGRS As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim ARS As New ADODB.Recordset
Dim PRS1 As New ADODB.Recordset
Dim PGRRS As New ADODB.Recordset
Dim BRS As New ADODB.Recordset
Dim PNLRS As New ADODB.Recordset
Dim ACTRS As New ADODB.Recordset
Dim IRTRS As New ADODB.Recordset
'' End of Record Set Declarations
   Dim LOCRS As New ADODB.Recordset
   Dim MARRS As New ADODB.Recordset
   Dim CATRS As New ADODB.Recordset
   Dim SMNRS As New ADODB.Recordset
   Dim DIRRS As New ADODB.Recordset
   Dim DIRRS1 As New ADODB.Recordset
   Dim DIRRS2 As New ADODB.Recordset
   Dim DIRRS3 As New ADODB.Recordset
   Dim DIRRS4 As New ADODB.Recordset
   Dim DIRRS5 As New ADODB.Recordset
   Dim DIRRS6 As New ADODB.Recordset
   Dim DIRRS7 As New ADODB.Recordset
   Dim DIRRS8 As New ADODB.Recordset
   Dim DIRRS9 As New ADODB.Recordset
   Dim SHLRS As New ADODB.Recordset
   Dim PRDRS As New ADODB.Recordset
   Dim INLRS As New ADODB.Recordset
   Dim CHKSHL As New ADODB.Recordset
   
Dim ECST_NME, ECST_AD1, ECST_AD2, ECST_AD3, ECSTAD4, ECST_CST, ECST_GST
Dim EPRD_NME, EBRD_NME, EGRP_IDY, EGRP_CLS, EGRP_CAT, ELST_RAT, ELST_DIS, EBAT_NME As String

' File Total Records, File Name and Input Data Declarations
Dim TOT_REC As Integer
Dim FILE_NAME As String
Dim input_data As Variant
' End of File Declarations
If CHECKTB("INL.DBF,INV.DBF,PNL.DBF,PNV.DBF,SNV.DBF,VAN.DBF,IRT.DBF,ACT.DBF,LOG.DBF") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & "DAS Version 4.0"
    Exit Sub
End If
Dim TOTCASAMT, TOTCHQAMT, TOTDEPRCD, TOTDEPREF As Integer
TOTCASAMT = 0
TOTCHQAMT = 0
TOTDEPRCD = 0
TOTDEPREF = 0
FILE_NAME = CURDIR & "\HHT\INL_STR.TXT"
Open FILE_NAME For Input As #1
If EOF(1) Then
    MsgBox "No Bill Value Records for Import!", vbCritical, "DAS Version 4.0"
    Exit Sub
End If

    Do
        Line Input #1, input_data
        TOT_REC = TOT_REC + 1
    Loop While Not EOF(1)
Close #1
'
Dim FSO As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
Set fil1 = FSO.GetFile(CURDIR & "\HHT\INL_STR.TXT")
Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
'INL_STR = input_data
'MsgBox "" & INL_STR
For I = 0 To TOT_REC - 1
    INL_STR = TXT_STREAM.ReadLine
    ERSI_IDY = Trim(Mid(INL_STR, 106, 10))
    ECAS_AMT = Trim(Mid(INL_STR, 186, 10))
    ECHQ_AMT = Trim(Mid(INL_STR, 196, 10))
    EDEP_RCD = Trim(Mid(INL_STR, 248, 10))
    EDEP_REF = Trim(Mid(INL_STR, 228, 10))
    LINV_IDY = Trim(Mid(INL_STR, 1, 10))
    EORD_IDY = Trim(Mid(INL_STR, 106, 10))
    MIBS_DAT = Trim(Mid(INL_STR, 11, 10))
    TEXT_DAT = Format(EINV_DAT, "DD/MM/YYYY")
    EINV_DAT = Format(EINV_DAT, "MM/DD/YYYY")
    LCST_IDY = Trim(Mid(INL_STR, 21, 10))
    LGRS_AMT = Trim(Mid(INL_STR, 116, 10))
    LCSD_AMT = Trim(Mid(INL_STR, 136, 10))
    EPKG_CST = Trim(Mid(INL_STR, 146, 10))
    LNET_AMT = Trim(Mid(INL_STR, 156, 10))
    LSMN_IDY = Trim(Mid(INL_STR, 61, 10))
    LROU_IDY = Trim(Mid(INL_STR, 96, 10))
    LSYS_DAT = Trim(Mid(INL_STR, 11, 10))
    LGOD_CAS = Mid(INL_STR, 238, 10)
    LGOD_AMT = Mid(INL_STR, 248, 10)
    LROD_CAS = Mid(INL_STR, 218, 10)
    LROD_AMT = Mid(INL_STR, 228, 10)
    LSYS_TIM = Trim(Mid(INL_STR, 259, 10))
    LSRN_AMT = Trim(Mid(INL_STR, 126, 10))
    STKCHECK = Trim(Mid(EORD_IDY, 1, 3))
    ENET_AMT = Trim(Mid(INL_STR, 156, 10))
    EPTY_IDY = Trim(Mid(INL_STR, 21, 10))
    Close
    TOTCASAMT = TOTCASAMT + Val(ECAS_AMT)
    TOTCHQAMT = TOTCHQAMT + Val(ECHQ_AMT)
    TOTDEPRCD = TOTDEPRCD + Val(EDEP_RCD)
    TOTDEPREF = TOTDEPREF + Val(EDEP_REF)
 
    '= START OF SNV = '
    ERATAPP = "B"
    If rs.State = 1 Then rs.Close

    If MOD1 = True Then
    End If
    If SHLRS.State = 1 Then SHLRS.Close
    If MOD1 = True Then
    End If
    SHLRS.Open "select * from SHL", CON, adOpenKeyset, adLockPessimistic
    If SHLRS.RecordCount > 1 Then
        SHLRS.MoveFirst
        ESHL_IDY = SHLRS("SHL_IDY")
        EPRD_IDY = SHLRS("PRD_IDY")
    ElseIf SHLRS.RecordCount = 1 Then
        ESHL_IDY = SHLRS("SHL_IDY")
        EPRD_IDY = SHLRS("PRD_IDY")
    Else
    End If
    If MOD1 = True Then
    End If
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    If PRDRS.State = 1 Then PRDRS.Close
    If MOD1 = True Then
    End If
    PRDRS.Open "select * from PRD where PRD_IDY like '" & Trim(EPRD_IDY) & "'", CON, adOpenKeyset, adLockPessimistic
    If INLRS.State = 1 Then INLRS.Close
    If MOD1 = True Then
    End If
    INLRS.Open "select * from INL where INV_IDY like '" & Trim(LINV_IDY) & "' ", CON, adOpenKeyset, adLockPessimistic

  
    If CHKSHL.State = 1 Then CHKSHL.Close
    If MOD1 = True Then
    End If
    CHKSHL.Open "SELECT * FROM SHL WHERE PRD_IDY LIKE '" & EPRD_IDY & " '", CON, adOpenKeyset, adLockPessimistic
    If MOD1 = True Then
    End If
    If RS1.State = 1 Then RS1.Close
    'If RS1.State = 1 Then RS1.Close
   
    RS1.Open "select * from SNV where INV_IDY like '" & Trim(EORD_IDY) & " ' ", CON, adOpenKeyset, adLockPessimistic
    If Not RS1.RecordCount > 0 Then
    If Val(LGOD_CAS) > 0 Then
    Dim IDYRS As New ADODB.Recordset
    If IDYRS.State = 1 Then IDYRS.Close
    If MOD1 = True Then
    End If
    IDYRS.Open "SELECT * FROM IDY", CON, adOpenKeyset, adLockPessimistic
    IDY = Val(IDYRS("IDY_SHT")) + 1
    If IDY < 10 Then
    ETXN_IDY = "SHT000000" & IDY
    ElseIf IDY > 9 Then
    ETXN_IDY = "SHT00000" & IDY
    ElseIf IDY > 99 Then
    ETXN_IDY = "SHT0000" & IDY
    ElseIf IDY > 999 Then
    ETXN_IDY = "SHT000" & IDY
    ElseIf IDY > 9999 Then
    ETXN_IDY = "SHT00" & IDY
    ElseIf IDY > 99999 Then
    ETXN_IDY = "SHT0" & IDY
    Else
    End If
    IDYRS("IDY_SHT") = IDY
    IDYRS.Update
    'RS1.MoveLast
    RS1.AddNew
    RS1("INV_IDY") = EORD_IDY
    RS1("INV_DAT") = EINV_DAT
    If SHLRS.RecordCount > 0 Then
        RS1("PRD_IDY") = SHLRS("PRD_IDY")
    Else
        RS1("PRD_IDY") = EPRD_IDY
    End If
    RS1("TXN_TYP") = "DPI"
    RS1("SHL_IDY") = SHLRS("SHL_IDY")
    RS1("I_SHL_NOS") = "0"
    RS1("I_PCK_NOS") = "0"
    RS1("D_SHL_NOS") = Mid(LGOD_CAS, 1, 6)
    ESUBUNT = Mid(LGOD_CAS, 5, 3)
    ER_PCK_NOS = RS1("D_SHL_NOS") * PRDRS("SUB_UNT")
    ER_PCK_NOS = ER_PCK_NOS + ESUBUNT
    RS1("D_PCK_NOS") = ER_PCK_NOS
    RS1("DP_SHL_NOS") = Mid(LGOD_CAS, 1, 6)
    RS1("DP_PCK_NOS") = ER_PCK_NOS
    RS1("DEP_AMT") = LGOD_CAS
    RS1("TXN_DAT") = EINV_DAT
    RS1("STK_UPD") = "Y"
    RS1("CST_IDY") = EPTY_IDY
    RS1("PRD_UPC") = PRDRS("SUB_UNT")
    RS1("PRD_PCK") = PRDRS("PRD_PCK")
    RS1("R_SHL_NOS") = "0"
    RS1("R_PCK_NOS") = "0"
    RS1("BRK_SHL") = "0"
    RS1("BRK_PCK") = "0"
    RS1("TXN_IDY") = ETXN_IDY
    RS1.Update
    Else
    End If
    If Val(LROD_CAS) > "0" Then
    'RS1.MoveLast
        RS1.AddNew
        RS1("INV_IDY") = EORD_IDY
        RS1("INV_DAT") = EINV_DAT
        If SHLRS.RecordCount > 0 Then
        RS1("PRD_IDY") = SHLRS("PRD_IDY")
        Else
        RS1("PRD_IDY") = EPRD_IDY
        End If
        RS1("TXN_TYP") = "DPO"
        RS1("SHL_IDY") = SHLRS("SHL_IDY")
        RS1("I_SHL_NOS") = "0"
        RS1("I_PCK_NOS") = "0"
        RS1("D_SHL_NOS") = Mid(LGOD_CAS, 1, 6)
        ESUBUNT = Mid(LGOD_CAS, 5, 3)
        ER_PCK_NOS = RS1("D_SHL_NOS") * PRDRS("SUB_UNT")
        ER_PCK_NOS = ER_PCK_NOS + ESUBUNT
        RS1("D_PCK_NOS") = ER_PCK_NOS
        RS1("DP_SHL_NOS") = RS1("D_SHL_NOS")
        RS1("DP_PCK_NOS") = ER_PCK_NOS
        RS1("DEP_AMT") = LROD_CAS
        RS1("TXN_DAT") = EINV_DAT
        RS1("STK_UPD") = "Y"
        RS1("CST_IDY") = EPTY_IDY
        RS1("PRD_UPC") = PRDRS("SUB_UNT")
        RS1("PRD_PCK") = PRDRS("PRD_PCK")
        RS1("R_SHL_NOS") = "0"
        RS1("R_PCK_NOS") = "0"
        RS1("BRK_SHL") = "0"
        RS1("BRK_PCK") = "0"
        RS1("TXN_IDY") = ETXN_IDY
        RS1.Update
    Else
    End If

    Else
    End If
    
    '= END OF SNV WITH GOD_CAS =
    
    TXT_STREAM.Skip (0)
    Next I
    TXT_STREAM.Close
    '= END OF SNV WITH GOD_CAS =
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
PBar1.Visible = False
PBar2.Visible = False
'Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
End Sub
Private Sub settings()
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    If Option1.Value = True Then
        Comport = 1
    Else
        Comport = 2
    End If
    With MSComm1
        .CommPort = Comport
        .InBufferSize = 3072
        .OutBufferSize = 512
        .InBufferCount = 0
        .InputLen = 1
        .settings = "57600,n,8,1"
    End With
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
End Sub

Public Sub delay(dlytime As Variant)
    Dim delay_timer
    Dim dli
    delay_timer = Timer
    dli = 0
    Do
    dli = dli + 1
       If Timer > (delay_timer + dlytime) Then
          Exit Do
       End If
    Loop
End Sub
''Public Sub COM_ERR()
''    Dim RESULT As Integer
''    Select Case MSComm1.CommEvent
''      Case comEventBreak   ' A Break was received.
''      RESULT = MsgBox("A COM Break Detected! Check COM Port and Related Connections!", vbYesNo + vbQuestion, "DAS Version 4.0")
''      If RESULT <> "YES" Then
''      Unload Me
''      Else
''      End If
''
''      Case comEventFrame   ' Framing Error
''      MsgBox "Framing Error! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEventOverrun   ' Data Lost.
''      MsgBox "Data Transmission Loss! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEventRxOver   ' Receive buffer overflow.
''      MsgBox "COM Port Data Receive Buffer Overflow!, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEventRxParity   ' Parity Error.
''      MsgBox "COM Port Parity Error! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEventTxFull   ' Transmit buffer full.
''      MsgBox "COM Transmit Buffer Full! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEventDCB   ' Unexpected error retrieving DCB
''      MsgBox "Error Retrieving DCB, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEvCD   ' Change in the CD line.
''      MsgBox "Change in CD Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEvCTS   ' Change in the CTS line.
''      MsgBox "Change in CTS Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEvDSR   ' Change in the DSR line.
''      MsgBox "Change in DSR Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEvRing   ' Change in the Ring Indicator.
''      MsgBox "Change in Ring Indicator, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEvReceive   ' Received RThreshold # of
''      MsgBox "Receive Error, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEvSend   ' There are SThreshold number of
''                     ' characters in the transmit
''                     ' buffer.
''      MsgBox "Data is in Transmit Buffer! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
''      Case comEvEOF   ' An EOF charater was found in
''                     ' the input stream
''      MsgBox "EOF in Input Stream! Check COM Connections!", vbCritical, "DAS Version 4.0"
''   End Select
''End Sub


Public Sub TIMEOUT()
'Timer1.Interval = 15000
'End Sub
End Sub

Public Sub COM_ERR()
    Dim RESULT As Variant
    Select Case MSComm1.CommEvent
      Case comEventBreak   ' A Break was received.
      RESULT = MsgBox("A COM Break Detected! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventFrame   ' Framing Error
      RESULT = MsgBox("Framing Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventOverrun   ' Data Lost.
      RESULT = MsgBox("Data Transmission Loss! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventRxOver   ' Receive buffer overflow.
      RESULT = MsgBox("COM Port Data Receive Buffer Overflow! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventRxParity   ' Parity Error.
      RESULT = MsgBox("COM Port Parity Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventTxFull   ' Transmit buffer full.
      RESULT = MsgBox("COM Transmit Buffer Full! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventDCB   ' Unexpected error retrieving DCB
      RESULT = MsgBox("Error Retrieving DCB! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvCD   ' Change in the CD line.
      RESULT = MsgBox("Change in CD Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvCTS   ' Change in the CTS line.
      RESULT = MsgBox("Change in CTS Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvDSR   ' Change in the DSR line.
      RESULT = MsgBox("Change in DSR Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvRing   ' Change in the Ring Indicator.
      RESULT = MsgBox("Change in Ring Indicator! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvReceive   ' Received RThreshold # of
      RESULT = MsgBox("Receive Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      RESULT = MsgBox("Data is in Transmit Buffer! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
      RESULT = MsgBox("EOF in Input Stream! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
   End Select
   If RESULT = 4 Or RESULT = "" Then
   'Unload Me
   ElseIf RESULT = 2 Then
   Unload Me
   'M_FRM_MIBS.Visible = True
   End If
End Sub

Function TimeIt(n)
   Dim StartTime, EndTime
   StartTime = Timer
   For I = 1 To n
   Next
   EndTime = Timer
   TimeIt = EndTime - StartTime
End Function

