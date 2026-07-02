VERSION 5.00
Begin VB.Form M_FRM_STR 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Version 4.0.B.n"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_STR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   3240
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   4605
      Width           =   6855
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright - Consoft Systems Pvt. Ltd."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   6855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Visit us at www.consoftsystems.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DasXtract"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2790
      TabIndex        =   8
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Version 4.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   5055
      TabIndex        =   7
      Top             =   600
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   2670
      Picture         =   "M_FRM_STR.frx":08CA
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distributor Automation System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   855
      TabIndex        =   4
      Top             =   120
      Width           =   5130
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   3720
      Width           =   1545
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year 2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:4.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "M_FRM_STR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Shell App.Path & "\MSINFO32.EXE", vbMaximizedFocus
End Sub

Private Sub Form_Load()
    Friday_MMExport 'pnr MODULE
    End
'    FST_BKP
End Sub

Private Sub Friday_MMExport()
Dim CON As New ADODB.Connection
CON.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CurDir & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"

' * Distributor ID Identification
        If Dir("C:\DEX\OutBox\*.ZIP") = "" Then
        Else
            Kill "C:\DEX\OutBox\*.ZIP"               'Delete Existing ZIP Files in C:\DEX\OutBox Folder
        End If
        Dim DIR_RS As New ADODB.Recordset
        Dim M_WDS_IDY As String
        Dim M_WDS_SRL As String
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", CON, adOpenKeyset, adLockPessimistic
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
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'SYS' and DIR_IDX = 'VER'", CON, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                DIR_RS("DIR_MSG") = Trim("01.01.01")
                DIR_RS.Update
            End If
        End If
' * Distributor ID Identification
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        'Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker4, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker3, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", CON, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        Pending_RS.Open "SELECT * FROM DIR WHERE CDOW(DIR_DAT) = 'Friday' AND TRIM(DIR_MSG) = 'Y'", CON, adOpenKeyset, adLockPessimistic
        
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveLast
            ST_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
            Pending_RS.MoveLast
            EN_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
        End If

If Pending_RS.RecordCount > 0 Then
        
    ' * : Beginning of DAS_SECONDARY.DAT Creation Code for Data Transfer
    
    Dim F_INV
    Dim F_INL
    Dim F_VAN
    Dim F_EIS
    
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_INL = CreateObject("Scripting.filesystemobject")
    
'    If FS_SOU.FolderExists(CurDir & "\HHT") = False Then
'       FS_SOU.CreateFolder (CurDir & "\HHT")
'    End If
    
    If FS_INL.FileExists(CurDir & "\DAS_PURCHASE.DAT") = True Then
        FS_INL.DeleteFile (CurDir & "\DAS_PURCHASE.DAT")
    End If
    
    Set F_INL = FS_INL.OpenTextFile(CurDir & "\DAS_PURCHASE.DAT", 2, True)
    
    Dim K As Double
    Dim r As Double
    Dim S As Double
    Dim t As Double
    K = 0
    r = 0
    S = 0
    t = 0
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
       
    If DIR_RS.State = 1 Then DIR_RS.Close
'    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) like 'DAY' AND alltrim(DIR_MSG) like 'Y' AND DIR_DAT >= CTOD('" & DTPicker4 & "') AND DIR_DAT <= CTOD('" & DTPicker3 & "')", CON, adOpenKeyset, adLockPessimistic
    DIR_RS.Open "SELECT * FROM DIR WHERE CDOW(DIR_DAT) = 'Friday' AND TRIM(DIR_MSG) = 'Y'", CON, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveLast
        ST_DAT = DIR_RS("DIR_DAT")
        BEG_DAT = DIR_RS("DIR_DAT")
        DIR_RS.MoveLast
        EN_DAT = DIR_RS("DIR_DAT")
        END_DAT = DIR_RS("DIR_DAT")
    End If
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", CON, adOpenKeyset, adLockPessimistic
    INV_RS.Open "SELECT * FROM INV WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", CON, adOpenKeyset, adLockPessimistic
    PRD_RS.Open "SELECT * FROM PRD", CON, adOpenKeyset, adLockPessimistic
    BAT_RS.Open "SELECT * FROM BRD", CON, adOpenKeyset, adLockPessimistic
    SCH_RS.Open "SELECT * FROM SCH", CON, adOpenKeyset, adLockPessimistic
    
    
    If INL_RS.State = 1 Then INL_RS.Close
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", CON, adOpenKeyset, adLockPessimistic
    If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
    While Not INL_RS.EOF
        M_ADD_AMT = 0
        M_DIS_AMT = 0
        M_TAX_AMT = 0
        T_SKU_QTY = 0
        If INV_RS.State = 1 Then INV_RS.Close
        INV_RS.Open "SELECT * FROM INV WHERE INV.INV_IDY = '" & Trim(INL_RS("INV_IDY")) & "'", CON, adOpenKeyset, adLockPessimistic
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
                T_SKU_QTY = Val(T_SKU_QTY) + Val(INV_RS("TOT_PCS"))
            End If
            INV_RS.MoveNext
        Wend
        If Trim(INL_RS("TXN_TYP")) <> "OPS" Then
            If INV_RS.RecordCount > 0 Then
                If Trim(INL_RS("TXN_TYP")) = "PUR" Then
                    'F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & AMT(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & AMT(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & AMT(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & Trim(INL_RS("ORD_IDY"))
                    F_INL.WriteLine Format(INL_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(M_WDS_IDY) & "|" & Trim(INL_RS("INV_IDY")) & "|" & Trim(INL_RS("ORD_IDY")) & "|" & Format(INL_RS("DCH_DAT"), "YYYYMMDD") & "|" & Val(T_SKU_QTY) & "|" & Val(INL_RS("NET_AMT"))
                    r = r + 1
                Else
                    'F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & AMT(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & AMT(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & AMT(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & ""
                    'r = r + 1
                End If
            End If
        End If
        INL_RS.MoveNext
    Wend
    F_INL.Close
    
'    cancel

    ' * : Beginning of DAS_STOCK.DAT Creation Code for Data Transfer

    Dim F_STK
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CurDir & "\HHT") = False Then
       FS_SOU.CreateFolder (CurDir & "\HHT")
    End If
    
    If FS_SOU.FileExists(CurDir & "\DAS_STOCK.DAT") = True Then
        FS_SOU.DeleteFile (CurDir & "\DAS_STOCK.DAT")
    End If
    
    Set F_STK = FS_SOU.OpenTextFile(CurDir & "\DAS_STOCK.DAT", 2, True)
    Dim l As Integer
    l = 0
    LDD = DateDiff("D", BEG_DAT, END_DAT)
    LDD = LDD + 1
    Do While LDD <> 0
            STKTXTDEX (BEG_DAT)
            Dim CLD_RS As New ADODB.Recordset
            
            If CLD_RS.State = 1 Then CLD_RS.Close
            If PRD_RS.State = 1 Then PRD_RS.Close
            If BAT_RS.State = 1 Then BAT_RS.Close
            CLD_RS.Open "SELECT PRD.PRD_IDY, SUM(CLS_STK) FROM CLD, BRD, PRD WHERE CLD.PRD_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY ORDER BY PRD.PRD_IDY GROUP BY PRD.PRD_IDY", CON, adOpenKeyset, adLockPessimistic
            PRD_RS.Open "SELECT * FROM PRD", CON, adOpenKeyset, adLockPessimistic
            BAT_RS.Open "SELECT * FROM BRD", CON, adOpenKeyset, adLockPessimistic
            
            
            If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
            
                While Not CLD_RS.EOF
                               
                    If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
                    BAT_RS.Find "BAT_IDY='" & Trim(CLD_RS("PRD_IDY")) & "'"
                    If BAT_RS.EOF = False Then
                       L_BAT_NME = BAT_RS("BAT_NME")
                    End If
                    
                    If PRD_RS.State = 1 Then PRD_RS.Close
                    PRD_RS.Open "SELECT * FROM PRD WHERE PRD.PRD_IDY = '" & Trim(CLD_RS(0)) & "'", CON, adOpenKeyset, adLockPessimistic
                    If PRD_RS.EOF = False Then
                       L_PRD_COD = PRD_RS("ALT_IDY")
                    End If
                                
        '            F_STK.WriteLine Trim(M_WDS_IDY) & "|" & Trim(L_PRD_COD) & "|" & Trim(STR(CLD_RS("OPN_STK"))) & "|" & Trim(STR(CLD_RS("PUR_STK"))) & "|" & Trim(STR(CLD_RS("PRT_STK"))) & "|" & Trim(STR(CLD_RS("SAL_STK"))) & "|" & Trim(STR(CLD_RS("SRN_STK"))) & "|" & Trim(STR(CLD_RS("STI_STK"))) & "|" & Trim(STR(CLD_RS("STO_STK"))) & "|" & Trim(STR(CLD_RS("CLS_STK"))) & "|" & Trim(STR(BAT_RS("PRD_MRP"))) & "|" & Trim(BAT_RS("BAT_NME")) & "|" & BAT_RS("BRD_DOM")
                    'F_STK.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(L_PRD_COD) & "|" & Trim(L_BAT_NME) & "|" & BAT_RS("BRD_DOM") & "|" & AMT(BAT_RS("PRD_MRP")) & "|" & QTY(STM(CLD_RS("OPN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PUR_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PRT_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SAL_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SRN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STI_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STO_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("CLS_STK"), PRD_RS("SUB_UNT")))
                    F_STK.WriteLine Format(BEG_DAT, "YYYYMMDD") & "|" & Format(END_DAT, "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & "|" & Val(CLD_RS(1)) & "|" & amt(Val(CLD_RS(1)) / PRD_RS("SUB_UNT"))
                    CLD_RS.MoveNext
                    l = l + 1
                Wend

        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_STK.Close
' * : End of DAS_STOCK.DAT Creation Code for Data Transfer

' Beginning of Distributor Information for MAPI Data Xfer
    
    Dim M_NME, M_ADR, M_TFX, M_WDS As String
    Dim RST_DIR As New ADODB.Recordset
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT1'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_NME = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT2'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_ADR = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT3'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_TFX = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'WDS' .AND. DIR_IDX LIKE 'IDY'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_WDS = Trim(RST_DIR("DIR_ALS"))
    End If

    If Dir(CurDir & "\DST_INF.DAT") = "" Then
        Set F_VER = FS_SOU.OpenTextFile(CurDir & "\DST_INF.DAT", 2, True)
        F_VER.WriteLine "Distributor : " & M_NME
        F_VER.WriteLine "Address     : " & M_ADR
        F_VER.WriteLine "Telephone   : " & M_TFX
        F_VER.WriteLine "DAS Code-ID : " & M_WDS
        F_VER.Close
    End If

' End of Distributor Information for MAPI Data Xfer

    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", CON, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then DIR_RS.MoveFirst
       DIR_RS("DIR_MSG") = Trim(M_WDS_SRL)
       DIR_RS.Update
    If DIR_RS.State = 1 Then DIR_RS.Close
    cancel
End If
End Sub
Private Sub cancel()
End
End Sub
Private Sub STKTXTDEX(DEN_DATE As Date)
Dim CON As New ADODB.Connection
CON.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CurDir & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
    EN_DAT = DEN_DATE
    If RS.State = 1 Then RS.Close
    CON.Execute "DELETE FROM CLD"
    RS.Open "SELECT * FROM CLD", CON, adOpenKeyset, adLockPessimistic
    If Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adLockPessimistic 'Modified by Radhika and Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adLockPessimistic
    End If
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!opn_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = 0
                RS!sal_stk = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
            CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!opn_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = 0
                RS!sal_stk = Val(RS1(0)) * -1
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "OPS" Then
            CON.Execute "UPDATE CLD SET OPN_STK = OPN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = 0
                RS!opn_stk = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "PUR" Then
        CON.Execute "UPDATE CLD SET PUR_STK = PUR_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS!PUR_STK = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "PRT" Then
        CON.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS!PRT_STK = Val(RS1(0))
                RS!PUR_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "STI" Then
        CON.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS!STI_STK = Val(RS1(0))
                RS!PRT_STK = 0
                RS!PUR_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "SRN" Then
        CON.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        CON.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS(7) = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS!PUR_STK = 0
                RS!cls_stk = 0
                RS.Update
            End If
        End If
        RS1.MoveNext
    Next
    RS.Close
    CON.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
    CON.Execute "UPDATE CLD SET OPN_STK=CLS_STK"
    CON.Execute "UPDATE CLD SET PUR_STK=0"
    CON.Execute "UPDATE CLD SET PRT_STK=0"
    CON.Execute "UPDATE CLD SET SAL_STK=0"
    CON.Execute "UPDATE CLD SET STO_STK=0"
    CON.Execute "UPDATE CLD SET SRN_STK=0"
    CON.Execute "UPDATE CLD SET STI_STK=0"
    CON.Execute "UPDATE CLD SET CLS_STK=0"
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM CLD", CON, adOpenKeyset, adLockPessimistic
    If Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    End If
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!opn_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = 0
                RS!sal_stk = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
            CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!opn_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = 0
                RS!sal_stk = Val(RS1(0)) * -1
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "OPS" Then
            CON.Execute "UPDATE CLD SET OPN_STK = OPN_STK +" & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = 0
                RS!opn_stk = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "PUR" Then
        CON.Execute "UPDATE CLD SET PUR_STK = PUR_STK +" & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS!PUR_STK = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "PRT" Then
        CON.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS!PRT_STK = Val(RS1(0))
                RS!PUR_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "STI" Then
        CON.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS!STI_STK = Val(RS1(0))
                RS!PRT_STK = 0
                RS!PUR_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "SRN" Then
        CON.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!PUR_STK = 0
                RS!srn_stk = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS(7) = 0
                RS!cls_stk = 0
                RS.Update
            End If
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        CON.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                RS.AddNew
                RS!prd_idy = Trim(RS1!BAT_IDY)
                RS!sal_stk = 0
                RS!opn_stk = 0
                RS!srn_stk = 0
                RS(7) = Val(RS1(0))
                RS!PRT_STK = 0
                RS!STI_STK = 0
                RS!PUR_STK = 0
                RS!cls_stk = 0
                RS.Update
            End If
        End If
        RS1.MoveNext
    Next
    RS.Close
    CON.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
End Sub

