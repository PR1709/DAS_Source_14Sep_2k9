VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form checkfiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validate Database"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ClipControls    =   0   'False
   Icon            =   "checkfiles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "&Batch Repair Utility"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   2580
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Activity Log"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   465
      Width           =   1695
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1800
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Fast Backup"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   885
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Data &Path"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1305
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Reset Database"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2145
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Show Missing Files"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1725
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2790
      ItemData        =   "checkfiles.frx":0442
      Left            =   120
      List            =   "checkfiles.frx":0444
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Missing Files List in the Current Directory Path"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   3330
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "checkfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim strPath As String
Private Sub Command1_Click()
List2.CLEAR
'Label1.Visible = False
List2.Visible = False
For i = 0 To List1.ListCount - 1
    strPath = CURDIR & "\" & List1.List(i)
    If Dir(strPath) = "" Then
        List2.AddItem List1.List(i)
    End If
Next
If List2.ListCount = 0 Then
    MsgBox "No Missing files Found", vbInformation, "DAS Version 4.0"
Else
'    Label1.Visible = True
    List2.Visible = True
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
    Dim FS_SOU As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")

    If FS_SOU.FileExists(App.Path & "\SET.LZH") = True Then
        CopyFile App.Path & "\SET.LZH", CURDIR & "\SET.LZH", 0
        CopyFile App.Path & "\LHA.EXE", CURDIR & "\LHA.EXE", 0
        If FS_SOU.FileExists(CURDIR & "\PNL.DBF") = True Then
            FS_SOU.DeleteFile (CURDIR & "\PNL.DBF")
        End If
        If FS_SOU.FileExists(CURDIR & "\PNL.CDX") = True Then
            FS_SOU.DeleteFile (CURDIR & "\PNL.CDX")
        End If
        If FS_SOU.FileExists(CURDIR & "\PNV.DBF") = True Then
            FS_SOU.DeleteFile (CURDIR & "\PNV.DBF")
        End If
        If FS_SOU.FileExists(CURDIR & "\PNV.CDX") = True Then
            FS_SOU.DeleteFile (CURDIR & "\PNV.CDX")
        End If
        Shell "LHA E SET.LZH", vbHide
        MsgBox ("Corrupted Intermediate File Reset! "), vbInformation, "DAS Version 4.0"
    Else
        MsgBox ("Corrupted Intermediate File Cannot Be Reset! SET.LZH File missing in Current Directory!"), vbInformation, "DAS Version 4.0"
    End If
End Sub
Private Sub Command4_Click()
MsgBox CURDIR, vbInformation, "DAS Version 4.0 - Current Data Path"
End Sub
Private Sub Command5_Click()
Me.MousePointer = vbHourglass
Dim DIR_RS As New ADODB.Recordset
  If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then DIR_RS.MoveFirst
       DIR_RS("DIR_MSG") = Trim(M_WDS_SRL)
       DIR_RS.Update
       M_WDS_IDY = Trim(DIR_RS!DIR_ALS)
    If DIR_RS.State = 1 Then DIR_RS.Close
    'Zip File Creation
            Dim DEX As String
            DEX = "C:\DEX"
            If Dir(DEX, vbDirectory) = "" Then
                MkDir DEX
            End If
            Dim FS_SOU, F_VER As Object
            Set FS_SOU = CreateObject("Scripting.filesystemobject")
            
            If Dir(App.Path & "\VER.BAT") = "" Then
                Set F_VER = FS_SOU.OpenTextFile(App.Path & "\VER.BAT", 2, True)
                F_VER.WriteLine "DIR " & App.Path & "\*.* /OG > VER_APP.TXT"
                F_VER.WriteLine "DIR " & CURDIR & "\*.* /OG > VER_DAT.TXT"
                F_VER.Close
            Else
                CopyFile App.Path & "\VER.BAT", CURDIR & "\VER.BAT", 0
            End If
            
            Shell CURDIR & "\VER.BAT", vbHide
            
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

            If Dir(CURDIR & "\DIR_INF.TXT") = "" Then
                Set F_VER = FS_SOU.OpenTextFile(CURDIR & "\DIR_INF.TXT", 2, True)
                F_VER.WriteLine "Distributor : " & M_NME
                F_VER.WriteLine "Address     : " & M_ADR
                F_VER.WriteLine "Telephone   : " & M_TFX
                F_VER.WriteLine "DAS Code-ID : " & M_WDS
                F_VER.Close
            End If
'            Friday_MMExport
            
            If Dir(CURDIR & "\CHGDBF.BAT") = "" Then
                MsgBox "Rename (CHGDBF.BAT) DBF Utility Batch File Not Found ! Quitting Application !", vbCritical, "DAS Version 4.0"
                con.Close
                End
            Else
                Shell "CHGDBF.BAT", vbMinimizedFocus
            End If
            
            M_FIL_NME = "\" & "FBKPDAS" & Trim(M_WDS_IDY) & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".ZIP"
            Dim Myzip As ColaZip.GSColaZip32
            Set Myzip = New ColaZip.GSColaZip32
            Myzip.RootDirectory = "\"
            Myzip.ZipFileName = "C:\DEX" & M_FIL_NME
            Myzip.AddFile (CURDIR & "\EFF.*")
            Myzip.AddFile (CURDIR & "\*.DBF")
            Myzip.AddFile (CURDIR & "\*.FPT")
            Myzip.AddFile (CURDIR & "\VER_APP.TXT")
            Myzip.AddFile (CURDIR & "\VER_DAT.TXT")
            Myzip.AddFile (CURDIR & "\DIR_INF.TXT")
            Myzip.AddFile (CURDIR & "\SHUTDOWN.LOG")
            
            If Company.SysInfo1.OSPlatform = 1 Then
                Myzip.AddFile "c:\windows\system\EFFMCG4.txt"
            ElseIf Company.SysInfo1.OSPlatform = 2 Then
                Myzip.AddFile "c:\winnt\system\EFFMCG4.txt"
            Else
                Myzip.AddFile "c:\windows\system\EFFMCG4.txt"
            End If
            
            Myzip.MakeColaZip ("bkp")   'creation of Zip File with Password bkp (in lower case)
            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
            ' Zip File Creation - Completed
            
            If Dir(CURDIR & "\MAKDBF.BAT") = "" Then
                MsgBox "Rename (MAKDBF.BAT) DBF Utility Batch File Not Found ! Quitting Application !", vbCritical, "DAS Version 4.0"
                con.Close
                End
            Else
                Shell "MAKDBF.BAT", vbMinimizedFocus
            End If
            
            Me.MousePointer = vbDefault
            M_MSG_STR = "Zipped Data From : " & Trim(M_NME) & " , " & Trim(M_ADR) & " , " & Trim(M_TFX) & " , " & Trim(M_WDS)
            M_MSG_EML = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'")
            M_MSG_SUB = "Zip Data Xfer Fast Backup - Data Extract : " & M_WDS
            
            ' Zip File Send by eMail Option
            M_MAIL_ATTMNT = "C:\DEX" & Trim(M_FIL_NME)
            If MsgBox("Send Zip File Data to Help Desk Now ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                If Trim(M_MSG_EML) <> "" Then
                    MailSend Trim(M_MSG_EML), Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
                Else
                    MailSend "effmcg@consoftsystems.com", Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
                End If
                MsgBox ("Data Send via MAPI Executed!"), vbInformation, "DAS Version 4.0"
                Unload Me
            End If
End Sub

Private Sub Command6_Click()
M_LOG_INF.Show 1
End Sub

Private Sub Command7_Click()
'DB_Reset
con.Close
If Dir(CURDIR & "\REPAIR.BAT") = "" Then
    MsgBox "Repair Batch Utility Not Found!", vbCritical, "DAS Version 4.0"
    con.Open
    Command2.SetFocus
Else
    Shell "REPAIR.BAT", vbMinimizedFocus
    End
'    If MsgBox("Set Active Connection ? Choosing Cancel Will Quit DAS Now with Improper ShutDown Status!", vbOKCancel, "DAS Version 4.0") = vbOK Then
'        CON.Open
'        Unload Me
'    Else
'        End
'    End If
End If
End Sub

Private Sub Form_Activate()
Command7.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
With List1
    .AddItem "Brd_mis.RP1"
    .AddItem "Cmp_rpt.RP1"
    .AddItem "pkt_mis.RP1"
    .AddItem "Cmp_mis.RP1"
    .AddItem "Cat_mis.RP1"
    .AddItem "Seg_mis.RP1"
    .AddItem "Pck_mis.RP1"
    .AddItem "Smn_mis.RP1"
    .AddItem "Loc_mis.RP1"
    .AddItem "Sku_mis.RP1"
    .AddItem "Rou_mis.RP1"
    .AddItem "Mar_mis.RP1"
    .AddItem "rtc_mis.RP1"
    .AddItem "Ret_mis.RP1"
    .AddItem "Sal_prg.RP1"
    .AddItem "Aging.RP1"
    .AddItem "cr.RP1"
    .AddItem "Ledger.RP1"
    .AddItem "Rep_lib.RP1"
    .AddItem "vouch.RP1"
    .AddItem "Dos_vch.RP1"
    .AddItem "Lists.RP1"
    .AddItem "report.RP1"
    .AddItem "Assets.RP1"
    .AddItem "Finance.RP1"
    .AddItem "Mir.RP1"
    .AddItem "Stock.RP1"
    .AddItem "Sal_prg.RP1"
    .AddItem "snn.RP1"
    .AddItem "MYP.RP1"
    .AddItem "SNP.RP1"
    .AddItem "KOR.RP1"
    .AddItem "ind.exe"
    .AddItem "cnd.exe"
    .AddItem "RRUNTIME.EXE"
    .AddItem "RRUN.EXE"
    .AddItem "RR.CNF"
    .AddItem "PRI.EXE"
    .AddItem "PRB.EXE"
    .AddItem "HHT.DBF"
    .AddItem "Eff.dbc"
    .AddItem "Eff.dct"
    .AddItem "Eff.dcx"
    .AddItem "Act.cdx"
    .AddItem "Agd.cdx"
    .AddItem "Asm.cdx"
    .AddItem "Ast.cdx"
    .AddItem "Att.cdx"
    .AddItem "Brd.cdx"
    .AddItem "Cat.cdx"
    .AddItem "Cld.cdx"
    .AddItem "Dir.cdx"
    .AddItem "Grp.cdx"
    .AddItem "Inl.cdx"
    .AddItem "Inv.cdx"
    .AddItem "Inv.cdx"
    .AddItem "Irt.cdx"
    .AddItem "Lds.cdx"
    .AddItem "Lev.cdx"
    .AddItem "Led.cdx"
    .AddItem "Lnv.cdx"
    .AddItem "Loc.cdx"
    .AddItem "Mar.cdx"
    .AddItem "Mbr.cdx"
    .AddItem "Mcl.cdx"
    .AddItem "Opa.cdx"
    .AddItem "Ops.cdx"
    .AddItem "Ord.cdx"
    .AddItem "Pgr.cdx"
    .AddItem "Pnl.cdx"
    .AddItem "Pnv.cdx"
    .AddItem "Prd.cdx"
    .AddItem "Sch.cdx"
    '.AddItem "Sid.cdx"
    .AddItem "Shl.cdx"
    .AddItem "Smn.cdx"
    .AddItem "Snv.cdx"
    .AddItem "Sta.cdx"
    .AddItem "Stl.cdx"
    .AddItem "Tpd.cdx"
    .AddItem "Typ.cdx"
    .AddItem "Van.cdx"
    .AddItem "cur.cdx"
    .AddItem "Vch.cdx"
    .AddItem "Yer.cdx"
    .AddItem "Act.dbf"
    .AddItem "Agd.dbf"
    .AddItem "Asm.dbf"
    .AddItem "Ast.dbf"
    .AddItem "Att.dbf"
    .AddItem "aca.dbf"
    .AddItem "Brd.dbf"
    .AddItem "Cat.dbf"
    .AddItem "Cld.dbf"
    .AddItem "Csr.dbf"
    .AddItem "sic.dbf"
    .AddItem "RP1.dbf"
    .AddItem "rp2.dbf"
    .AddItem "Dir.dbf"
    .AddItem "Dsu.dbf"
    .AddItem "fun.dbf"
    .AddItem "Grp.dbf"
    .AddItem "idy.dbf"
    .AddItem "Ine.dbf"
    .AddItem "Inl.dbf"
    .AddItem "Inv.dbf"
    .AddItem "Irt.dbf"
    '.AddItem "INC.dbf"
    '.AddItem "COM.dbf"
    .AddItem "QPS.dbf"
    '.AddItem "BRK.dbf"
    .AddItem "RDY.dbf"
    .AddItem "Lds.dbf"
    .AddItem "Led.dbf"
    .AddItem "Lev.dbf"
    .AddItem "Lnv.dbf"
    .AddItem "Loc.dbf"
    .AddItem "Log.dbf"
    .AddItem "Mar.dbf"
    .AddItem "Mbr.dbf"
    .AddItem "Mcl.dbf"
    .AddItem "Npo.dbf"
    .AddItem "Opa.dbf"
    .AddItem "Ops.dbf"
    .AddItem "Ord.dbf"
    .AddItem "Pgr.dbf"
    .AddItem "Pnl.dbf"
    .AddItem "Pnv.dbf"
    .AddItem "ppc.dbf"
    .AddItem "Prd.dbf"
    .AddItem "pre.dbf"
    .AddItem "Prt.dbf"
    .AddItem "rri.dbf"
    .AddItem "Rtp.dbf"
    .AddItem "Sch.dbf"
    .AddItem "Sid.dbf"
    .AddItem "Scl.dbf"
    .AddItem "Shl.dbf"
    .AddItem "smn.dbf"
    .AddItem "Snv.dbf"
    .AddItem "Sta.dbf"
    .AddItem "Stl.dbf"
    .AddItem "Sts.dbf"
    .AddItem "Tgt.dbf"
    .AddItem "Tgb.dbf"
    .AddItem "Tgc.dbf"
    .AddItem "Tgs.dbf"
    .AddItem "ASM.FPT"
    .AddItem "top.dbf"
    .AddItem "Tpd.dbf"
    .AddItem "Trt.dbf"
    .AddItem "Typ.dbf"
    .AddItem "usr.dbf"
    .AddItem "Van.dbf"
    .AddItem "cur.dbf"
    .AddItem "Vch.dbf"
    .AddItem "Wek.dbf"
    .AddItem "Yer.dbf"
    .AddItem "BNV.DBF"
    .AddItem "BNL.DBF"
    .AddItem "CNV.DBF"
    .AddItem "VBK.DBF"
    .AddItem "SBK.DBF"
    .AddItem "CBK.DBF"
    .AddItem "XNV.DBF"
    .AddItem "XNL.DBF"
End With
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

Private Sub DB_Reset()
Dim INV_RS As New ADODB.Recordset
Dim PNV_RS As New ADODB.Recordset
Dim INL_RS As New ADODB.Recordset
Dim PNL_RS As New ADODB.Recordset
Dim DIR_RS As New ADODB.Recordset
Dim ACT_RS As New ADODB.Recordset
Dim Pgr_RS As New ADODB.Recordset
Dim PRD_RS As New ADODB.Recordset
Dim BAT_RS As New ADODB.Recordset

con.Execute "DELETE FROM PNV"
con.Execute "DELETE FROM PNL"

INV_RS.Open "SELECT * FROM INV", con, adOpenKeyset, adLockPessimistic
If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
While Not INV_RS.EOF
    If PNV_RS.State = 1 Then PNV_RS.Close
    PNV_RS.Open "SELECT * FROM PNV", con, adOpenKeyset, adLockPessimistic
    If Len(Trim(INV_RS("PRD_IDY"))) And Len(Trim(INV_RS("BAT_IDY"))) > 0 Then 'FOR FAULTY RECORDS FILTERATION
    PNV_RS.AddNew
    PNV_RS("INV_IDY") = INV_RS("INV_IDY")
    PNV_RS("INV_DAT") = INV_RS("INV_DAT")
    PNV_RS("PRD_IDY") = INV_RS("PRD_IDY")
    M_GRP_IDY = GetValue("GRP_IDY", "PRD WHERE PRD_IDY LIKE '" & Trim(INV_RS("PRD_IDY")) & "'")
    PNV_RS("PRD_NME") = GetValue("PRD_NME", "PRD WHERE PRD_IDY LIKE '" & Trim(INV_RS("PRD_IDY")) & "'")
    PNV_RS("GRP_IDY") = GetValue("GRP_IDY", "PRD WHERE PRD_IDY LIKE '" & Trim(INV_RS("PRD_IDY")) & "'")
    PNV_RS("GRP_CLS") = GetValue("GRP_CLS", "PGR WHERE GRP_IDY LIKE '" & Trim(M_GRP_IDY) & "'")
    PNV_RS("GRP_CAT") = GetValue("GRP_CAT", "PGR WHERE GRP_IDY LIKE '" & Trim(M_GRP_IDY) & "'")
    PNV_RS("LST_RAT") = INV_RS("LST_RAT")
    PNV_RS("LST_DIS") = INV_RS("LST_DIS")
    PNV_RS("PRD_QTY") = INV_RS("PRD_QTY")
    If IsNull(INV_RS("RET_QTY")) = True Then
        PNV_RS("RET_QTY") = 0#
    Else
        PNV_RS("RET_QTY") = INV_RS("RET_QTY")
    End If
    If IsNull(INV_RS("RAT_APP")) = True Then
        PNV_RS("RAT_APP") = ""
    Else
        PNV_RS("RAT_APP") = INV_RS("RAT_APP")
    End If
    PNV_RS("PRD_RAT") = INV_RS("PRD_RAT")
    If IsNull(INV_RS("CPN_RAT")) = True Then
        PNV_RS("CPN_RAT") = 0#
    Else
        PNV_RS("CPN_RAT") = INV_RS("CPN_RAT")
    End If
    If IsNull(INV_RS("CPN_AMT")) = True Then
        PNV_RS("CPN_AMT") = 0#
    Else
        PNV_RS("CPN_AMT") = INV_RS("CPN_AMT")
    End If
    If IsNull(INV_RS("ADD_DIS")) = True Then
        PNV_RS("ADD_DIS") = 0#
    Else
        PNV_RS("ADD_DIS") = INV_RS("ADD_DIS")
    End If
    If IsNull(INV_RS("ADD_AMT")) = True Then
        PNV_RS("ADD_AMT") = 0#
    Else
        PNV_RS("ADD_AMT") = INV_RS("ADD_AMT")
    End If
    If IsNull(INV_RS("TRD_DIS")) = True Then
        PNV_RS("TRD_DIS") = 0#
    Else
        PNV_RS("TRD_DIS") = INV_RS("TRD_DIS")
    End If
    If IsNull(INV_RS("DIS_AMT")) = True Then
        PNV_RS("DIS_AMT") = 0#
    Else
        PNV_RS("DIS_AMT") = INV_RS("DIS_AMT")
    End If
    If IsNull(INV_RS("TRD_TAX")) = True Then
        PNV_RS("TRD_TAX") = 0#
    Else
        PNV_RS("TRD_TAX") = INV_RS("TRD_TAX")
    End If
    If IsNull(INV_RS("TAX_AMT")) = True Then
        PNV_RS("TAX_AMT") = 0#
    Else
        PNV_RS("TAX_AMT") = INV_RS("TAX_AMT")
    End If
    PNV_RS("COL_AMT") = INV_RS("COL_AMT")
    PNV_RS("STK_UPD") = INV_RS("STK_UPD")
    PNV_RS("TXN_TYP") = INV_RS("TXN_TYP")
    If IsNull(INV_RS("FST_FLG")) = True Then
        PNV_RS("FST_FLG") = ""
    Else
        PNV_RS("FST_FLG") = INV_RS("FST_FLG")
    End If
    If IsNull(INV_RS("OPN_QTY")) = True Then
        PNV_RS("OPN_QTY") = 0#
    Else
        PNV_RS("OPN_QTY") = INV_RS("OPN_QTY")
    End If
    If IsNull(INV_RS("ISS_QTY")) = True Then
        PNV_RS("ISS_QTY") = 0#
    Else
        PNV_RS("ISS_QTY") = INV_RS("ISS_QTY")
    End If
    If IsNull(INV_RS("TOT_QTY")) = True Then
        PNV_RS("TOT_QTY") = 0#
    Else
        PNV_RS("TOT_QTY") = INV_RS("TOT_QTY")
    End If
    If IsNull(INV_RS("SCH_AMT")) = True Then
        PNV_RS("SCH_AMT") = 0#
    Else
        PNV_RS("SCH_AMT") = INV_RS("SCH_AMT")
    End If
    PNV_RS("ORD_IDY") = INV_RS("ORD_IDY")
    PNV_RS("BAT_IDY") = INV_RS("BAT_IDY")
    PNV_RS("BAT_NME") = GetValue("BAT_NME", "BRD WHERE BAT_IDY LIKE '" & Trim(INV_RS("BAT_IDY")) & "'")
    PNV_RS("PRD_UOM") = INV_RS("PRD_UOM")
    PNV_RS("SUB_UNT") = INV_RS("SUB_UNT")
    PNV_RS("PRD_UNT") = INV_RS("PRD_UNT")
    PNV_RS("PRD_PCS") = INV_RS("PRD_PCS")
    PNV_RS("TOT_PCS") = INV_RS("TOT_PCS")
    PNV_RS("STK_TYP") = INV_RS("STK_TYP")
    If IsNull(INV_RS("SCH_IDY")) = True Then
        PNV_RS("SCH_IDY") = ""
    Else
        PNV_RS("SCH_IDY") = INV_RS("SCH_IDY")
    End If
    
    M_MRP_RAT = GetValue("PRD_MRP", "BRD WHERE BAT_IDY LIKE '" & Trim(INV_RS("BAT_IDY")) & "'")
    If IsNull(M_MRP_RAT) = True Then
        PNV_RS("MRP_RAT") = 0
    Else
        PNV_RS("MRP_RAT") = M_MRP_RAT
    End If
    If IsNull(INV_RS("INC_TTL")) = True Then
        PNV_RS("INC_TTL") = 0#
    Else
        PNV_RS("INC_TTL") = INV_RS("INC_TTL")
    End If
    PNV_RS("MAR_PGR") = ""
    PNV_RS.Update
    End If
    INV_RS.MoveNext
Wend
Set PNV_RS = Nothing
Set INV_RS = Nothing

INL_RS.Open "SELECT * FROM INL", con, adOpenKeyset, adLockPessimistic
If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
While Not INL_RS.EOF
    If PNL_RS.State = 1 Then PNL_RS.Close
    PNL_RS.Open "SELECT * FROM PNL", con, adOpenKeyset, adLockPessimistic
    PNL_RS.AddNew
    PNL_RS("INV_IDY") = INL_RS("INV_IDY")
    PNL_RS("INV_DAT") = INL_RS("INV_DAT")
    PNL_RS("ORD_IDY") = INL_RS("ORD_IDY")
    PNL_RS("CST_IDY") = INL_RS("CST_IDY")
    M_MAR_IDY = GetValue("MAR_IDY", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    M_SMN_IDY = GetValue("SMN_IDY", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    M_CAT_IDY = GetValue("CAT_IDY", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    M_LOC_IDY = GetValue("LOC_IDY", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    M_ROU_IDY = GetValue("ROU_IDY", "ACT WHERE ROU_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_NME") = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_AD1") = GetValue("ACT_AD1", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_AD2") = GetValue("ACT_AD2", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_AD3") = GetValue("ACT_AD3", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_AD4") = GetValue("ACT_AD4", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_CST") = GetValue("ACT_CST", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_GST") = GetValue("ACT_GST", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("CST_ZIP") = GetValue("ACT_ZIP", "ACT WHERE ACT_IDY LIKE '" & Trim(INL_RS("CST_IDY")) & "'")
    PNL_RS("MAR_NME") = GetValue("MAR_NME", "MAR WHERE MAR_IDY LIKE '" & Trim(M_MAR_IDY) & "'")
    PNL_RS("SMN_NME") = GetValue("SMN_NME", "SMN WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "'")
    PNL_RS("LOC_NME") = GetValue("LOC_NME", "LOC WHERE LOC_IDY LIKE '" & Trim(M_LOC_IDY) & "'")
    PNL_RS("CAT_NME") = GetValue("CAT_NME", "CAT WHERE CAT_IDY LIKE '" & Trim(M_CAT_IDY) & "'")
    PNL_RS("GRS_AMT") = INL_RS("GRS_AMT")
    PNL_RS("SPD_PCG") = INL_RS("SPD_PCG")
    PNL_RS("SPD_AMT") = INL_RS("SPD_AMT")
    PNL_RS("CSD_PCG") = INL_RS("CSD_PCG")
    PNL_RS("CSD_AMT") = INL_RS("CSD_AMT")
    PNL_RS("EXD_PCG") = INL_RS("EXD_PCG")
    If IsNull(INL_RS("EXD_CAL")) = True Then
        PNL_RS("EXD_CAL") = ""
    Else
        PNL_RS("EXD_CAL") = INL_RS("EXD_CAL")
    End If
    PNL_RS("EXD_AMT") = INL_RS("EXD_AMT")
    PNL_RS("TAX_PCG") = INL_RS("TAX_PCG")
    PNL_RS("TAX_AMT") = INL_RS("TAX_AMT")
    PNL_RS("NET_AMT") = INL_RS("NET_AMT")
    PNL_RS("SMN_IDY") = Trim(M_SMN_IDY)
    PNL_RS("LOC_IDY") = Trim(M_LOC_IDY)
    PNL_RS("MAR_IDY") = Trim(M_MAR_IDY)
    PNL_RS("CAT_IDY") = Trim(M_CAT_IDY)
    PNL_RS("TXN_TYP") = INL_RS("TXN_TYP")
    If IsNull(INL_RS("FRT_TYP")) = True Then
        PNL_RS("FRT_TYP") = ""
    Else
        PNL_RS("FRT_TYP") = INL_RS("FRT_TYP")
    End If
    PNL_RS("FRT_AMT") = INL_RS("FRT_AMT")
    PNL_RS("DCH_IDY") = INL_RS("DCH_IDY")
    If IsNull(INL_RS("DCH_DAT")) = True Then
        PNL_RS("DCH_DAT") = INL_RS("INV_DAT")
    Else
        PNL_RS("DCH_DAT") = INL_RS("DCH_DAT")
    End If
    PNL_RS("LRN_IDY") = INL_RS("LRN_IDY")
    PNL_RS("LRN_DAT") = INL_RS("LRN_DAT")
    PNL_RS("LRN_AMT") = INL_RS("LRN_AMT")
    PNL_RS("WBL_IDY") = INL_RS("WBL_IDY")
    PNL_RS("WBL_DAT") = INL_RS("WBL_DAT")
    PNL_RS("PKG_CST") = INL_RS("PKG_CST")
    PNL_RS("TTL_CAS") = INL_RS("TTL_CAS")
    PNL_RS("TXN_RMK") = INL_RS("TXN_RMK")
    PNL_RS("TXN_TPT") = INL_RS("TXN_TPT")
    If IsNull(INL_RS("TAX_ACT")) = True Then
        PNL_RS("TAX_ACT") = ""
    Else
        PNL_RS("TAX_ACT") = INL_RS("TAX_ACT")
    End If
    If IsNull(INL_RS("ROU_ACT")) = True Then
        PNL_RS("ROU_ACT") = ""
    Else
        PNL_RS("ROU_ACT") = INL_RS("ROU_ACT")
    End If
    PNL_RS("ROU_OFF") = INL_RS("ROU_OFF")
    If IsNull(INL_RS("FRT_ACT")) = True Then
        PNL_RS("FRT_ACT") = ""
    Else
        PNL_RS("FRT_ACT") = INL_RS("FRT_ACT")
    End If
    PNL_RS("CRT_DAY") = INL_RS("CRT_DAY")
    PNL_RS("DUE_DAT") = INL_RS("DUE_DAT")
    If IsNull(INL_RS("DES_MOD")) = True Then
        PNL_RS("DES_MOD") = ""
    Else
        PNL_RS("DES_MOD") = INL_RS("DES_MOD")
    End If
    If IsNull(INL_RS("STO_RMK")) = True Then
        PNL_RS("STO_RMK") = ""
    Else
        PNL_RS("STO_RMK") = INL_RS("STO_RMK")
    End If
    PNL_RS("ROU_IDY") = INL_RS("ROU_IDY")
    PNL_RS("TXN_RM1") = Trim(INL_RS("TXN_RM1"))
    PNL_RS("TXN_RM2") = Trim(INL_RS("TXN_RM2"))
    PNL_RS("VCH_IDY") = INL_RS("VCH_IDY")
    If IsNull(INL_RS("VCH_DAT")) = True Then
        PNL_RS("VCH_DAT") = INL_RS("INV_DAT")
    Else
        PNL_RS("VCH_DAT") = INL_RS("VCH_DAT")
    End If
    PNL_RS("SCH_CHK") = INL_RS("SCH_CHK")
    PNL_RS("SRN_AMT") = INL_RS("SRN_AMT")
    PNL_RS("SCH_IDY") = INL_RS("SCH_IDY")
    PNL_RS("USR_NME") = INL_RS("USR_IDY")
    PNL_RS("SYS_DAT") = INL_RS("SYS_DAT")
    PNL_RS("SYS_TIM") = INL_RS("SYS_TIM")
    If IsNull(INL_RS("DIS_CHK")) = True Then
        PNL_RS("DIS_CHK") = ""
    Else
        PNL_RS("DIS_CHK") = INL_RS("DIS_CHK")
    End If
    PNL_RS("MAR_DOP") = 0#
    PNL_RS("INS_NME") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT1'")
    PNL_RS("INS_ADR") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT2'")
    PNL_RS("SYS_RM1") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'RM1'")
    PNL_RS("SYS_RM2") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'RM2'")
    PNL_RS("SYS_RM3") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'RM3'")
    PNL_RS("SYS_RM4") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'RM4'")
    PNL_RS("INS_GST") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'GST'")
    PNL_RS("INS_CST") = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'CST'")
    PNL_RS("FRM_DAT") = INL_RS("INV_DAT")
    PNL_RS("TO_DAT") = INL_RS("INV_DAT")
    PNL_RS("ACT_IDY") = INL_RS("ACT_IDY")
    PNL_RS.Update
    INL_RS.MoveNext
Wend
Set PNL_RS = Nothing
Set INL_RS = Nothing
MsgBox ("Tmp DB Restore Successful! Please Exit and run IndexData before you work on DAS!"), vbInformation, "DAS Version 4.0"
End Sub
Private Sub Friday_MMExport()
' * Distributor ID Identification
'        If Dir("C:\DEX\OutBox\*.DAT") = "" Then
'        Else
'            Kill "C:\DEX\OutBox\*.DAT"               'Delete Existing ZIP Files in C:\DEX\OutBox Folder
'        End If
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
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        'Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker4, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker3, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", CON, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        Pending_RS.Open "SELECT * FROM DIR WHERE CDOW(DIR_DAT) = 'Friday' AND TRIM(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
        
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveLast
            ST_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
            Pending_RS.MoveLast
            EN_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
        End If

If Pending_RS.RecordCount > 0 Then
        
    ' * : Beginning of DAS_CUSTOMER.DAT Creation Code for Data Transfer
    Me.MousePointer = vbHourglass
    Dim F_OMS, F_LOC, F_MAR, F_SMN
    Dim FS_SOU, FS_LOC, FS_MAR, FS_SMN As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_LOC = CreateObject("Scripting.filesystemobject")
    Set FS_MAR = CreateObject("SCripting.filesystemobject")
    Set FS_SMN = CreateObject("SCripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_CUSTOMER.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_CUSTOMER.DAT")
    End If
'    If FS_LOC.FileExists(CURDIR & "\HHT\LOCATION.TXT") = True Then
'        FS_LOC.DeleteFile (CURDIR & "\HHT\LOCATION.TXT")
'    End If
'    If FS_MAR.FileExists(CURDIR & "\HHT\MARKET.TXT") = True Then
'        FS_MAR.DeleteFile (CURDIR & "\HHT\MARKET.TXT")
'    End If
'    If FS_SMN.FileExists(CURDIR & "\HHT\SALESMAN.TXT") = True Then
'        FS_SMN.DeleteFile (CURDIR & "\HHT\SALESMAN.TXT")
'    End If
    
    Set F_OMS = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_CUSTOMER.DAT", 2, True)
'    Set F_LOC = FS_LOC.OpenTextFile(CURDIR & "\HHT\LOCATION.TXT", 2, True)
'    Set F_MAR = FS_MAR.OpenTextFile(CURDIR & "\HHT\MARKET.TXT", 2, True)
'    Set F_SMN = FS_SMN.OpenTextFile(CURDIR & "\HHT\SALESMAN.TXT", 2, True)
    
    Dim J As Integer
    J = 0
    
    
    Dim Category_RS As New ADODB.Recordset
    Dim Location_RS As New ADODB.Recordset
    Dim Market_RS As New ADODB.Recordset
    Dim Salesman_RS As New ADODB.Recordset
    Dim Invoice_Chk_RS As New ADODB.Recordset
    
    Dim Retailer_RS As New ADODB.Recordset
    If Retailer_RS.State = 1 Then Retailer_RS.Close
    Retailer_RS.Open "select * from ACT where ACT_TYP = 'SAL' or ACT_TYP = 'PUR'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
    
    If Retailer_RS.RecordCount > 0 Then Retailer_RS.MoveFirst
       
       Dim L_CAT_NME As String
       Dim L_LOC_NME As String
       Dim L_MAR_NME As String
       Dim L_DST_NME As String
       Dim L_MAR_GRD As String
       
       If Category_RS.State = 1 Then Category_RS.Close
       Category_RS.Open "SELECT * FROM CAT", con, adOpenKeyset, adLockPessimistic
       If Location_RS.State = 1 Then Location_RS.Close
       Location_RS.Open "SELECT * FROM LOC", con, adOpenKeyset, adLockPessimistic
       If Market_RS.State = 1 Then Market_RS.Close
       Market_RS.Open "SELECT * FROM MAR", con, adOpenKeyset, adLockPessimistic
       
       While Not Retailer_RS.EOF
       
          If Category_RS.RecordCount > 0 Then Category_RS.MoveFirst
          Category_RS.Find "CAT_IDY='" & Trim(Retailer_RS("CAT_IDY")) & "'"
          If Category_RS.EOF = False Then
             L_CAT_NME = Category_RS("CAT_NME")
          End If
          
          If Location_RS.RecordCount > 0 Then Location_RS.MoveFirst
          Location_RS.Find "LOC_IDY='" & Trim(Retailer_RS("LOC_IDY")) & "'"
          If Location_RS.EOF = False Then
             L_LOC_NME = Location_RS("LOC_NME")
          End If
          
          If Market_RS.RecordCount > 0 Then Market_RS.MoveFirst
          Market_RS.Find "MAR_IDY='" & Trim(Retailer_RS("MAR_IDY")) & "'"
          If Market_RS.EOF = False Then
             L_MAR_NME = Market_RS("MAR_NME")
             L_DST_NME = Market_RS("MAR_PGR")
             L_MAR_GRD = Market_RS("MAR_GRD")
             L_MAR_CSG = Market_RS("MAR_CSG")
             L_MAR_CLU = Market_RS("MAR_CLU")
             L_MAR_ROC = Market_RS("MAR_ROC")
             L_MAR_TWC = Market_RS("MAR_TWC")
             L_MAR_TWN = Market_RS("MAR_TWN")
             L_MAR_DST = Market_RS("MAR_DST")
             L_MAR_STA = Market_RS("MAR_STA")
             L_MAR_COU = Market_RS("MAR_COU")
          End If
          
'          F_OMS.WriteLine Trim(M_WDS_IDY) & "|" & Trim(Retailer_RS("ACT_IDY")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(L_DST_NME) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("DBT_TYP"))
          If Trim(Retailer_RS("ACT_TYP")) = "PUR" Then
                'F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "S" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                'J = J + 1
          Else
                'F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "C" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                If Invoice_Chk_RS.State = 1 Then Invoice_Chk_RS.Close
                Invoice_Chk_RS.Open "SELECT * FROM INL WHERE CST_IDY LIKE '" & Trim(Retailer_RS("ACT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Invoice_Chk_RS.RecordCount > 0 Then
                    F_OMS.WriteLine Trim(M_WDS_IDY) & Right(Trim(Retailer_RS("ACT_IDY")), 4) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & "I" & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_ZIP")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_FAX")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(L_MAR_COU) & "|" & Trim(L_MAR_STA) & "|" & Trim(L_MAR_DST) & "|" & Trim(L_MAR_TWN) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_CSG) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(L_MAR_CLU) & "|" & Trim(L_MAR_ROC) & "|" & Trim(L_MAR_TWC) & "|" & Trim(Retailer_RS("DBT_TYP"))
                    J = J + 1
                End If
          End If
          Retailer_RS.MoveNext
       Wend
'''       MsgBox "" & J & ": Outlet Records Created for Data Transfer!"
'''    Else
'''       MsgBox ("No Outlet Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'''       cancel
'''    End If
'       Dim n As Integer
'       n = 0
'       If Location_RS.State = 1 Then Location_RS.Close
'       Location_RS.Open "SELECT * FROM LOC", CON, adOpenKeyset, adlockpessimistic
'       While Not Location_RS.EOF
'          F_LOC.WriteLine Trim(Location_RS("LOC_IDY")) & "|" & Trim(Location_RS("LOC_NME"))
'          Location_RS.MoveNext
'          n = n + 1
'       Wend
'
'       Dim o As Integer
'       o = 0
'       If Market_RS.State = 1 Then Market_RS.Close
'       Market_RS.Open "SELECT * FROM MAR", CON, adOpenKeyset, adlockpessimistic
'       While Not Market_RS.EOF
'          F_MAR.WriteLine Trim(Market_RS("MAR_IDY")) & "|" & Trim(Market_RS("MAR_NME")) & "|" & Trim(Market_RS("MAR_PGR")) & "|" & Trim(Market_RS("MAR_GRD"))
'          Market_RS.MoveNext
'          o = o + 1
'       Wend
'
'       Dim P As Integer
'       P = 0
'       If Salesman_RS.State = 1 Then Salesman_RS.Close
'       Salesman_RS.Open "SELECT * FROM SMN", CON, adOpenKeyset, adlockpessimistic
'       While Not Salesman_RS.EOF()
'          F_SMN.WriteLine Trim(Salesman_RS("SMN_IDY")) & "|" & Trim(Salesman_RS("SMN_NME"))
'          Salesman_RS.MoveNext
'          P = P + 1
'       Wend
'
        F_OMS.Close
'        F_LOC.Close
'        F_MAR.Close
'        F_SMN.Close
'    cancel
    
    ' * : End of DAS_CUSTOMER.DAT Creation Code for Data Transfer
    
    ' * : Beginning of DAS_SECONDARY.DAT Creation Code for Data Transfer
    
    Dim F_INV
    Dim F_INL
    Dim F_VAN
    Dim F_EIS
    
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_INL = CreateObject("Scripting.filesystemobject")
'    Set FS_VAN = CreateObject("Scripting.filesystemobject")
'    Set FS_EIS = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_SECONDARY.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_SECONDARY.DAT")
    End If
    
    If FS_INL.FileExists(CURDIR & "\HHT\DAS_PURCHASE.DAT") = True Then
        FS_INL.DeleteFile (CURDIR & "\HHT\DAS_PURCHASE.DAT")
    End If
    
'    If FS_VAN.FileExists(CURDIR & "\HHT\LOAD_SHEET.TXT") = True Then
'        FS_VAN.DeleteFile (CURDIR & "\HHT\LOAD_SHEET.TXT")
'    End If
'
'    If FS_EIS.FileExists(CURDIR & "\HHT\EMPTY_ISSUE.TXT") = True Then
'        FS_EIS.DeleteFile (CURDIR & "\HHT\EMPTY_ISSUE.TXT")
'    End If
    
    Set F_INV = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_SECONDARY.DAT", 2, True)
    Set F_INL = FS_INL.OpenTextFile(CURDIR & "\HHT\DAS_PURCHASE.DAT", 2, True)
'    Set F_VAN = FS_VAN.OpenTextFile(CURDIR & "\HHT\LOAD_SHEET.TXT", 2, True)
'    Set F_EIS = FS_EIS.OpenTextFile(CURDIR & "\HHT\EMPTY_ISSUE.TXT", 2, True)
    
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
       
    If DIR_RS.State = 1 Then DIR_RS.Close
'    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) like 'DAY' AND alltrim(DIR_MSG) like 'Y' AND DIR_DAT >= CTOD('" & DTPicker4 & "') AND DIR_DAT <= CTOD('" & DTPicker3 & "')", CON, adOpenKeyset, adLockPessimistic
    DIR_RS.Open "SELECT * FROM DIR WHERE CDOW(DIR_DAT) = 'Friday' AND TRIM(DIR_MSG) = 'Y' ORDER BY DIR_DAT", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveLast
        ST_DAT = DIR_RS("DIR_DAT")
        BEG_DAT = DIR_RS("DIR_DAT")
        DIR_RS.MoveLast
        EN_DAT = DIR_RS("DIR_DAT")
        END_DAT = DIR_RS("DIR_DAT")
    End If
    ST_DAT = (ST_DAT - 7)
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    INV_RS.Open "SELECT * FROM INV WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    SCH_RS.Open "SELECT * FROM SCH", con, adOpenKeyset, adLockPessimistic
    
    If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
    
        While Not INV_RS.EOF

            If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
            PRD_RS.Find "PRD_IDY='" & Trim(INV_RS("PRD_IDY")) & "'"
            If PRD_RS.EOF = False Then
               L_PRD_COD = PRD_RS("ALT_IDY")
            End If
            
            If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
            BAT_RS.Find "BAT_IDY='" & Trim(INV_RS("BAT_IDY")) & "'"
            If BAT_RS.EOF = False Then
               L_BAT_NME = BAT_RS("BAT_NME")
               L_PUR_LPP = Val(BAT_RS("U_PUR_LPP"))
               L_SAL_LPP = Val(BAT_RS("U_SAL_LPP"))
               L_MRP_RAT = (Val(BAT_RS("PRD_MRP")) / Val(BAT_RS("SKU_SBU")))
               L_RMM_UNT = (Val(BAT_RS("PRD_MRP")) / Val(BAT_RS("SKU_SBU"))) - Val(BAT_RS("U_SAL_LPP"))
               L_DMM_UNT = Val(BAT_RS("SAL_LPP")) - Val(BAT_RS("PUR_LPP")) 'VALUE TAKEN ON PER CASE BASIS - PNR
               PL_TOT_VAT = (Val(BAT_RS("P_TOT_VAT")) / Val(BAT_RS("SKU_SBU")))
               SL_TOT_VAT = (Val(BAT_RS("S_TOT_VAT")) / Val(BAT_RS("SKU_SBU")))
            End If
            
            If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
            INL_RS.Find "INV_IDY='" & Trim(INV_RS("INV_IDY")) & "'"
            If INL_RS.EOF = False Then
               L_CST_IDY = INL_RS("CST_IDY")
               L_ORD_IDY = INL_RS("ORD_IDY")
               L_LRN_IDY = INL_RS("LRN_IDY")
               L_LRN_DAT = INL_RS("LRN_DAT")
            End If
            
            If SCH_RS.RecordCount > 0 Then SCH_RS.MoveFirst
            SCH_RS.Find "SCH_IDY='" & Trim(INV_RS("SCH_IDY")) & "'"
            If SCH_RS.EOF = False Then
                L_MER_NUM = SCH_RS("MER_NUM")
            End If
            
            'M_WDS_IDY & "|"Trim(STR(Inv_RS("COL_AMT"))) & "|" &
            If Trim(INV_RS("TXN_TYP")) = "SAL" Then
                If Trim(INV_RS("STK_TYP")) = "Salable" Then
                    'F_INV.WriteLine Trim(INV_RS("INV_IDY")) & "|" & Format(INV_RS("INV_DAT"), "DD/MM/YYYY") & "|" & Trim(INV_RS("TXN_TYP")) & "|" & Trim(L_CST_IDY) & "|" & Trim(L_PRD_COD) & "|" & QTY(INV_RS("PRD_QTY")) & "|" & AMT(INV_RS("PRD_RAT")) & "|" & Trim(L_BAT_NME) & "|" & Format(BAT_RS("BRD_DOM"), "DD/MM/YYYY") & "|" & AMT(BAT_RS("PRD_MRP")) & "|" & Trim(INV_RS("STK_TYP")) & "|" & Trim(L_MER_NUM) & "|" & Trim(L_LRN_IDY) & "|" & Format(L_LRN_DAT, "DD/MM/YYYY") & "|" & Trim(INL_RS("TXN_TPT"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIV" & "|" & Val(INV_RS("TOT_PCS")) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIR" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) & "|" & Trim(INV_RS("INV_IDY"))
                    K = K + 5
                End If
                If Trim(INV_RS("STK_TYP")) = "Free" Then
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIV" & "|" & Val(INV_RS("TOT_PCS")) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIR" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) & "|" & Trim(INV_RS("INV_IDY"))
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_SAL_LPP)) & "|" & Trim(INV_RS("INV_IDY"))
                    K = K + 6
                    
                    If LNV_RS.State = 1 Then LNV_RS.Close
                    LNV_RS.Open "SELECT * FROM LNV WHERE INV_IDY LIKE '" & Trim(INV_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                    If LNV_RS.RecordCount > 0 Then
                        While Not LNV_RS.EOF
                            If Val(LNV_RS("LNV_AMT")) > 0 Then
                                F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & Val(LNV_RS("LNV_AMT")) & "|" & Trim(INV_RS("INV_IDY"))
                                K = K + 1
                            End If
                            LNV_RS.MoveNext
                        Wend
                    End If
                End If
            End If
            If Trim(INV_RS("TXN_TYP")) = "SRN" Then
                If Trim(INV_RS("STK_TYP")) = "Salable" Then
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRV" & "|" & (Val(INV_RS("TOT_PCS")) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRR" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    K = K + 5
                End If
                If Trim(INV_RS("STK_TYP")) = "Free" Then
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRV" & "|" & (Val(INV_RS("TOT_PCS")) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRR" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_SAL_LPP)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                    K = K + 6
                End If
            End If
            INV_RS.MoveNext
        Wend
'        MsgBox "" & K & ": Stock Records Created for Data Transfer!"
'    Else
'       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'       cancel
'    End If
    F_INV.Close
    
    If INL_RS.State = 1 Then INL_RS.Close
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
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
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_STOCK.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_STOCK.DAT")
    End If
    
    Set F_STK = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_STOCK.DAT", 2, True)
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
            CLD_RS.Open "SELECT PRD.PRD_IDY, SUM(CLS_STK) FROM CLD, BRD, PRD WHERE CLD.PRD_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY AND CLS.CLS_STK > 0 ORDER BY PRD.PRD_IDY GROUP BY PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
            PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
            BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
            
            
            If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
            
                While Not CLD_RS.EOF
                               
                    If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
                    BAT_RS.Find "BAT_IDY='" & Trim(CLD_RS("PRD_IDY")) & "'"
                    If BAT_RS.EOF = False Then
                       L_BAT_NME = BAT_RS("BAT_NME")
                    End If
                    
                    If PRD_RS.State = 1 Then PRD_RS.Close
                    PRD_RS.Open "SELECT * FROM PRD WHERE PRD.PRD_IDY = '" & Trim(CLD_RS(0)) & "'", con, adOpenKeyset, adLockPessimistic
                    If PRD_RS.EOF = False Then
                       L_PRD_COD = PRD_RS("ALT_IDY")
                    End If
                                
        '            F_STK.WriteLine Trim(M_WDS_IDY) & "|" & Trim(L_PRD_COD) & "|" & Trim(STR(CLD_RS("OPN_STK"))) & "|" & Trim(STR(CLD_RS("PUR_STK"))) & "|" & Trim(STR(CLD_RS("PRT_STK"))) & "|" & Trim(STR(CLD_RS("SAL_STK"))) & "|" & Trim(STR(CLD_RS("SRN_STK"))) & "|" & Trim(STR(CLD_RS("STI_STK"))) & "|" & Trim(STR(CLD_RS("STO_STK"))) & "|" & Trim(STR(CLD_RS("CLS_STK"))) & "|" & Trim(STR(BAT_RS("PRD_MRP"))) & "|" & Trim(BAT_RS("BAT_NME")) & "|" & BAT_RS("BRD_DOM")
                    'F_STK.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(L_PRD_COD) & "|" & Trim(L_BAT_NME) & "|" & BAT_RS("BRD_DOM") & "|" & AMT(BAT_RS("PRD_MRP")) & "|" & QTY(STM(CLD_RS("OPN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PUR_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PRT_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SAL_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SRN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STI_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STO_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("CLS_STK"), PRD_RS("SUB_UNT")))
                    F_STK.WriteLine Format(BEG_DAT, "YYYYMMDD") & "|" & Format(END_DAT, "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & "|" & Val(CLD_RS(1))
                    CLD_RS.MoveNext
                    l = l + 1
                Wend
        '        MsgBox "" & L & ": SKU Records Created for Data Transfer!"
        '    Else
        '       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
        '       cancel
        '    End If
           
        '    cancel
            ' * : End of DAS_STOCK.DAT Creation Code for Data Transfer
        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_STK.Close
    
    ' * : Log Text File Creation Begin Process
    
    Dim F_VER
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_VERSION.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_VERSION.DAT")
    End If
    
    Dim Q As Integer
    Dim LDC_DAT As Date
    Q = 0
    Set F_VER = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_VERSION.DAT", 2, True)
    
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) like 'DAY' AND alltrim(DIR_MSG) like 'Y'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveLast
        LDC_DAT = DIR_RS("DIR_DAT")
    Else
        LDC_DAT = Date
    End If
    M_EXE_VER = Format(FileDateTime(App.Path & "\EFFMCG.EXE"), "dd/mmm/yyyy hh:mm")
    F_VER.WriteLine Trim(M_WDS_IDY) & "|" & Trim("01.01.01") & "|" & Format(DTPicker4, "dd/MM/yyyy") & "|" & Format(DTPicker3, "dd/MM/yyyy") & "|" & Format(LDC_DAT, "dd/MM/yyyy") & "|" & Trim(M_EXE_VER)
    Q = Q + 1
    
    F_VER.Close
    
    Dim F_LOG
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_LOG.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_LOG.DAT")
    End If
    
    Set F_LOG = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_LOG.DAT", 2, True)
    
    F_LOG.WriteLine "DAS_CUSTOMER.DAT" & "|" & str(J)
    F_LOG.WriteLine "DAS_SECONDARY.DAT" & "|" & str(K)
    F_LOG.WriteLine "DAS_PURCHASE.DAT" & "|" & str(r)
    F_LOG.WriteLine "DAS_STOCK.DAT" & "|" & str(l)
    F_LOG.WriteLine "DAS_VERSION.DAT" & "|" & str(Q)
    
    Dim MDD As String
    Dim MMM As String
    Dim MYY As String
    MDD = Trim(str(Day(EN_DAT)))
    MMM = Trim(str(Month(EN_DAT)))
    MYY = Trim(str(Year(EN_DAT)))
'    F_LOG.WriteLine MDD & MMM & MYY & M_WDS_SRL 'Refers to Day Closure Date, Last Line in Log File is Blocked - PNR
    F_LOG.Close
    ' * : Log Text File Creation End Process
    Me.MousePointer = vbDefault
'    MsgBox ("File Creation Process Successful! Log File Generated!"), vbInformation, "DAS Version 4.0"
    
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
            
'            CopyFile CURDIR & "\HHT" & "\DAS_CUSTOMER.DAT", CDO & "\DAS_CUSTOMER.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_SECONDARY.DAT", CDO & "\DAS_SECONDARY.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_PURCHASE.DAT", CDO & "\DAS_PURCHASE.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_STOCK.DAT", CDO & "\DAS_STOCK.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_VERSION.DAT", CDO & "\DAS_VERSION.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_LOG.DAT", CDO & "\DAS_LOG.DAT", 0
'            CopyFile CURDIR & "\DST_INF.DAT", CDO & "\DST_INF.DAT", 0
            
'            CopyFile CURDIR & "\HHT" & "\DAS_CUSTOMER.DAT", DEX & "\DAS_CUSTOMER.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_SECONDARY.DAT", DEX & "\DAS_SECONDARY.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DAS_PURCHASE.DAT", DEX & "\DAS_PURCHASE.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DAS_STOCK.DAT", DEX & "\DAS_STOCK.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_VERSION.DAT", DEX & "\DAS_VERSION.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_LOG.DAT", DEX & "\DAS_LOG.DAT", 0
'            CopyFile CURDIR & "\DST_INF.DAT", DEX & "\DST_INF.DAT", 0
                        
                        
            
'            CopyFile App.Path & "\" & "Zip32.dll", CURDIR & "\" & "Zip32.dll", 0
'            M_EML_IDY = GetValue("DIR_MSG", "DIR where DIR_IDY like 'EML' AND DIR_IDX LIKE 'IDY'")
'
'            M_FIL_NME = "\" & "DAS" & Trim(M_WDS_IDY) & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".ZIP"
'            Dim Myzip As ColaZip.GSColaZip32
'            Set Myzip = New ColaZip.GSColaZip32
'            Myzip.RootDirectory = "\"
'            Myzip.ZipFileName = "C:\DEX" & M_FIL_NME
''            Myzip.AddFile ("C:\DEX" & "\DAS_CUSTOMER.DAT")
''            Myzip.AddFile ("C:\DEX" & "\DAS_SECONDARY.DAT")
''            Myzip.AddFile ("C:\DEX" & "\DAS_PURCHASE.DAT")
'            Myzip.AddFile ("C:\DEX" & "\DAS_STOCK.DAT")
'            Myzip.AddFile ("C:\DEX" & "\DAS_VERSION.DAT")
'            Myzip.AddFile ("C:\DEX" & "\DAS_LOG.DAT")
'            Myzip.AddFile ("C:\DEX" & "\DST_INF.DAT")
            
'            Myzip.MakeColaZip ("DAS")
            
'            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
                      
'            M_MSG_STR = "Margin Minder Zipped Data From : " & Trim(M_NME) & " , " & Trim(M_ADR) & " , " & Trim(M_TFX) & " , " & Trim(M_WDS) & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
'            M_MSG_EML = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'")
'            M_MSG_SUB = "Zip Data Xfer Margin Minder Data Extract : " & M_WDS & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
'
'            M_MAIL_ATTMNT = "C:\DEX" & M_FIL_NME
'            If MsgBox("Send Zip File Data to DAS MIS Analytics Desk Now ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
'                If Trim(M_EML_IDY) = "" Then
'                    MailSend "effmcg@consoftsystems.com", Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
'                    MsgBox ("Data Send via MAPI Executed! Email ID: effmcg@consoftsystems.com" & Trim(M_EML_IDY)), vbInformation, "DAS Version 4.0"
'                Else
'                    MailSend Trim(M_EML_IDY), Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
'                    MsgBox ("Data Send via MAPI Executed! Email ID: " & Trim(M_EML_IDY)), vbInformation, "DAS Version 4.0"
'                End If
'            Else
'                cancel
'            End If
'            CopyFile DEX & M_FIL_NME, CDO & M_FIL_NME, 0
            CopyFile CURDIR & "\HHT" & "\DAS_STOCK.DAT", CDO & "\DAS_STOCK.DAT", 0
    ' Zip File Creation
            
    cancel
'Else
'    DTPicker4.SetFocus
End If
End Sub
Private Sub cancel()

End Sub
Private Sub STKTXTDEX(DEN_DATE As Date)
    EN_DAT = DEN_DATE
    If rs.State = 1 Then rs.Close
    con.Execute "DELETE FROM CLD"
    rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
    If Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika and Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic
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
            con.Execute "UPDATE CLD SET OPN_STK = OPN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET PUR_STK = PUR_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
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
        con.Execute "UPDATE CLD SET PUR_STK = PUR_STK +" & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
        con.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
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
End Sub
