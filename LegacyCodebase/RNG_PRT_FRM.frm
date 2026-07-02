VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RNG_PRT_FRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Range Print"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "RNG_PRT_FRM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "System Generated Invoice &TAX"
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "System Generated Invoice &SAL"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Height          =   1200
      Left            =   5880
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   700
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Range Numbers"
      Height          =   1455
      Left            =   3120
      TabIndex        =   7
      Top             =   945
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   840
         MaxLength       =   10
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   840
         MaxLength       =   10
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   915
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   435
         Width           =   345
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Number"
      Height          =   195
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Range Dates"
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   945
      Width           =   2415
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   720
         TabIndex        =   2
         Top             =   390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   85327875
         CurrentDate     =   37337
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   720
         TabIndex        =   5
         Top             =   870
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   85327875
         CurrentDate     =   37337
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Date"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Value           =   -1  'True
      Width           =   735
   End
End
Attribute VB_Name = "RNG_PRT_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_GotFocus()
If Check2.Value = 1 Then
    Check2.Value = 0
End If
End Sub
Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then Option1.SetFocus
End Sub
Private Sub Check2_GotFocus()
If Check1.Value = 1 Then
    Check1.Value = 0
End If
End Sub
Private Sub Command1_Click()
Command1.Enabled = False
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If Option1.Value = True Then
    con.Execute "UPDATE DIR WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT' SET DIR_ALS ='" & Format(DTPicker1, "DD/MM/YYYY") & "',DIR_MSG='" & Format(DTPicker2, "DD/MM/YYYY") & "'"
    If RNG_PRT_FRM.Tag = "SAL" Then
        If Check1.Value = 0 Then
            SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "').AND.pnl->TXN_TYP='SAL'.AND.SUBSTR(pnl->INV_IDY,1,3)<>'SAL'"
        Else
            SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "').AND.pnl->TXN_TYP='SAL'.AND.SUBSTR(pnl->INV_IDY,1,3)='SAL'"
        End If
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            If Check1.Value = 0 Then
                SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "').AND.pnl->TXN_TYP='SAL'.AND.SUBSTR(pnl->INV_IDY,1,3)<>'SAL'"
            Else
                SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "').AND.pnl->TXN_TYP='SAL'.AND.SUBSTR(pnl->INV_IDY,1,3)='SAL'"
            End If
            If Check2.Value = 0 Then
                SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "').AND.pnl->TXN_TYP='SAL'.AND.SUBSTR(pnl->INV_IDY,1,3)<>'TAX'"
            Else
                SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "').AND.pnl->TXN_TYP='SAL'.AND.SUBSTR(pnl->INV_IDY,1,3)='TAX'"
            End If
            DOSREPORT SAL_QRY, 6
            'Shell CURDIR & "\RRUNTIME RRI 7"
            If Dir(CURDIR & "\PRI.EXE") = "" Then
                MsgBox "PRI.EXE not found", vbExclamation
            Else
                Shell CURDIR & "\PRI.EXE"
            End If
        Else
            If REC_SET.State = 1 Then REC_SET.Close
            REC_SET.Open "select * from dir where dir_idy like 'AUT' AND DIR_IDX LIKE 'IND'", con, adOpenKeyset, adLockPessimistic
            If REC_SET.RecordCount > 0 Then
                If Trim(REC_SET("DIR_MSG")) = "" Then
                    OpenReport CURDIR & "\dos_vch.RP1", "Invoice", 2, SAL_QRY, 6
                Else
                    OpenReport CURDIR & "\dos_vch.RP1", Trim(REC_SET("DIR_MSG")), 2, SAL_QRY, 6
                End If
            Else
                OpenReport CURDIR & "\dos_vch.RP1", "Invoice", 2, SAL_QRY, 6
            End If
        End If
    ElseIf RNG_PRT_FRM.Tag = "PUR" Then
        SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='PUR'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
        SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "')"
            DOSREPORT SAL_QRY, 0
            Shell CURDIR & "\RRUNTIME RRI 1"
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Purchase Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "RCT" Then
        SAL_QRY = "VCH->TXN_dat>=ctod('" & DTPicker1 & "').and.VCH->TXN_dat<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='RCT'"
        OpenReport CURDIR & "\VOUCH.RP1", "New Rct Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "PMT" Then
        SAL_QRY = "VCH->TXN_dat>=ctod('" & DTPicker1 & "').and.VCH->TXN_dat<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='PMT'"
        OpenReport CURDIR & "\VOUCH.RP1", "New Pmt Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "PRT" Then
        SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='PRT'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "')"
            DOSREPORT SAL_QRY, 5
            Shell CURDIR & "\RRUNTIME RRI 6", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Purchase Return Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "SRN" Then
        SAL_QRY = "PNL->INV_DAT>=ctod('" & DTPicker1 & "').and.PNL->INV_DAT<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='SRN'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "')"
            DOSREPORT SAL_QRY, 2
            Shell CURDIR & "\RRUNTIME RRI 3", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Sale Return Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "OPA" Then
        SAL_QRY = "VCH->TXN_dat>=ctod('" & DTPicker1 & "').and.VCH->TXN_dat<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='OPN'"
        OpenReport CURDIR & "\VOUCH.RP1", "New Opn Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "JRN" Then
        SAL_QRY = "VCH->TXN_dat>=ctod('" & DTPicker1 & "').and.VCH->TXN_dat<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='JRN'"
        OpenReport CURDIR & "\VOUCH.RP1", "New JRN Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "OPS" Then
        SAL_QRY = "PNL->INV_DAT>=ctod('" & DTPicker1 & "').and.PNL->INV_DAT<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='OPS'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
        SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "')"
            DOSREPORT SAL_QRY, 1
            Shell CURDIR & "\RRUNTIME RRI 2", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Opening Stock Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "STI" Then
        SAL_QRY = "PNL->INV_DAT>=ctod('" & DTPicker1 & "').and.PNL->INV_DAT<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='STI'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "')"
            DOSREPORT SAL_QRY, 3
            Shell CURDIR & "\RRUNTIME RRI 4", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Transfer Inward Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "STO" Then
        SAL_QRY = "PNL->INV_DAT>=ctod('" & DTPicker1 & "').and.PNL->INV_DAT<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='STO'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            SAL_QRY = "pnl->inv_dat>=ctod('" & Format(DTPicker1, "DD/MM/YYYY") & "').and.pnl->inv_dat<=ctod('" & Format(DTPicker2, "DD/MM/YYYY") & "')"
            DOSREPORT SAL_QRY, 4
            Shell CURDIR & "\RRUNTIME RRI 5", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Transfer Outward Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "CRN" Then
        SAL_QRY = "VCH->TXN_dat>=ctod('" & DTPicker1 & "').and.VCH->TXN_dat<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='CRN'"
        OpenReport CURDIR & "\VOUCH.RP1", "NEW CRT Note", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "DRN" Then
        SAL_QRY = "VCH->TXN_dat>=ctod('" & DTPicker1 & "').and.VCH->TXN_dat<=ctod('" & DTPicker2 & "')"
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='DBN'"
        OpenReport CURDIR & "\VOUCH.RP1", "NEW DBT Note", 2, SAL_QRY, 6
    End If
Else
    If RNG_PRT_FRM.Tag = "SAL" Then
        If Check1.Value = 0 Or Check2.Value = 0 Then
            SAL_QRY = "VAL(pnl->INV_IDY)>=" & Val(Text1) & ".AND.VAL(pnl->INV_IDY)<=" & Val(Text2) & ".AND.pnl->TXN_TYP='SAL'"
        Else
            SAL_QRY = "VAL(SUBSTR(pnl->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(pnl->INV_IDY,4,7))<=" & Val(Text2) & ".AND.pnl->TXN_TYP='SAL'"
        End If
        MsgBox SAL_QRY
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
        DOSREPORT SAL_QRY, 6
'            If REC_SET.State = 1 Then REC_SET.Close
'            REC_SET.Open "SELECT * FROM RRI", PNL, adOpenKeyset, adlockpessimistic
'            If REC_SET.RecordCount > 0 Then
'                REC_SET.Move 6
'                REC_SET!RI_QUERY = "O"
'                REC_SET!ri_filter = SAL_QRY
'                REC_SET!ri_copies = 1
'                REC_SET!ri_begpage = 1
'                REC_SET!ri_endpage = 9999999
'                REC_SET!ri_printer = ""
'                REC_SET!ri_mindex = CURDIR & "\PNV.CDX"
'                REC_SET.Update
'            End If
'            REC_SET.Close
            'Shell CURDIR & "\RRUNTIME RRI 7"
            If Dir(CURDIR & "\PRI.EXE") = "" Then
                MsgBox "PRI.EXE not found", vbExclamation
            Else
                Shell CURDIR & "\PRI.EXE"
            End If
        Else
            If REC_SET.State = 1 Then REC_SET.Close
            REC_SET.Open "select * from dir where dir_idy like 'AUT' AND DIR_IDX LIKE 'IND'", con, adOpenKeyset, adLockPessimistic
            If REC_SET.RecordCount > 0 Then
                If Trim(REC_SET("DIR_MSG")) = "" Then
                    OpenReport CURDIR & "\dos_vch.RP1", "Invoice", 2, SAL_QRY, 6
                Else
                    OpenReport CURDIR & "\dos_vch.RP1", Trim(REC_SET("DIR_MSG")), 2, SAL_QRY, 6
                End If
            Else
                OpenReport CURDIR & "\dos_vch.RP1", "Invoice", 2, SAL_QRY, 6
            End If
        End If
    ElseIf RNG_PRT_FRM.Tag = "PUR" Then
        SAL_QRY = "VAL(SUBSTR(PNL->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(PNL->INV_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='PUR'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            DOSREPORT SAL_QRY, 0
            Shell CURDIR & "\RRUNTIME RRI 1"
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Purchase Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "RCT" Then
        SAL_QRY = "VAL(SUBSTR(VCH->TXN_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(VCH->TXN_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='RCT'"
        OpenReport CURDIR & "\VOUCH.RP1", "New Rct Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "PMT" Then
        SAL_QRY = "VAL(SUBSTR(VCH->TXN_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(VCH->TXN_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='PMT'"
        OpenReport CURDIR & "\VOUCH.RP1", "New Pmt Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "PRT" Then
        SAL_QRY = "VAL(SUBSTR(PNL->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(PNL->INV_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='PRT'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            DOSREPORT SAL_QRY, 5
            Shell CURDIR & "\RRUNTIME RRI 6", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Purchase Return Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "SRN" Then
        SAL_QRY = "VAL(SUBSTR(PNL->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(PNL->INV_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='SRN'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            DOSREPORT SAL_QRY, 2
            Shell CURDIR & "\RRUNTIME RRI 3", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Sale Return Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "OPA" Then
        SAL_QRY = "VAL(SUBSTR(VCH->TXN_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(VCH->TXN_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='OPN'"
        OpenReport CURDIR & "\VOUCH.RP1", "New Opn Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "JRN" Then
        SAL_QRY = "VAL(SUBSTR(VCH->TXN_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(VCH->TXN_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='JRN'"
        OpenReport CURDIR & "\VOUCH.RP1", "New JRN Voucher", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "OPS" Then
        SAL_QRY = "VAL(SUBSTR(PNL->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(PNL->INV_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='OPS'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            DOSREPORT SAL_QRY, 1
            Shell CURDIR & "\RRUNTIME RRI 2", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Opening Stock Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "STI" Then
        SAL_QRY = "VAL(SUBSTR(PNL->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(PNL->INV_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='STI'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            DOSREPORT SAL_QRY, 3
            Shell CURDIR & "\RRUNTIME RRI 4", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Transfer Inward Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "STO" Then
        SAL_QRY = "VAL(SUBSTR(PNL->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(PNL->INV_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.PNL->TXN_TYP='STO'"
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
            DOSREPORT SAL_QRY, 4
            Shell CURDIR & "\RRUNTIME RRI 5", vbHide
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Transfer Outward Note", 2, SAL_QRY, 6
        End If
    ElseIf RNG_PRT_FRM.Tag = "CRN" Then
        SAL_QRY = "VAL(SUBSTR(VCH->TXN_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(VCH->TXN_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='CRN'"
        OpenReport CURDIR & "\VOUCH.RP1", "NEW CRT Note", 2, SAL_QRY, 6
    ElseIf RNG_PRT_FRM.Tag = "DRN" Then
        SAL_QRY = "VAL(SUBSTR(VCH->TXN_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(VCH->TXN_IDY,4,7))<=" & Val(Text2)
        SAL_QRY = SAL_QRY & ".AND.VCH->TXN_TYP='DBN'"
        OpenReport CURDIR & "\VOUCH.RP1", "NEW DBT Note", 2, SAL_QRY, 6
    End If
End If
    SAL_QRY = ""
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_CloseUp()
DTPicker1.SetFocus
End Sub

Private Sub DTPicker2_CloseUp()
DTPicker2.SetFocus
End Sub

Private Sub DTPicker1_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then DTPicker2.SetFocus
If K = 38 And s = 1 Then DTPicker1.SetFocus
End Sub

Private Sub DTPicker2_LostFocus()
If DTPicker1 > DTPicker2 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker1.SetFocus
    Exit Sub
End If
End Sub

Private Sub DTPicker2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then Command1.SetFocus
End Sub
Private Sub Form_Activate()
Frame1.Enabled = False
Frame2.Enabled = False
If Me.Tag = "LSS" Or Me.Tag = "SAL" Then

Else
    Option1.SetFocus
End If
SETYEAROPEN RNG_PRT_FRM
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Option1_KeyPress(K As Integer)
If K = 13 Then
    Frame1.Enabled = True
    Frame2.Enabled = False
    DTPicker1.SetFocus
End If

End Sub

Private Sub Option2_KeyPress(K As Integer)
If K = 13 Then
    Frame2.Enabled = True
    Frame1.Enabled = False
    Text1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 13 Then
Else
K = 0
End If
If K = 13 Then
    If Len(Trim(Text1)) = 0 Then
        MsgBox "Enter Start Number !", vbExclamation, "DAS Version 4.0"
        Text1.SetFocus
        Exit Sub
    End If
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 13 Then
Else
K = 0
End If
If K = 13 Then
    If Len(Trim(Text2)) = 0 Then
        MsgBox "Enter End Number !", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
        Exit Sub
    ElseIf Val(Text2) < Val(Text1) Then
        MsgBox "End Number Should Be Greater Than Start Number !", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
        Exit Sub
    End If
    Command1.SetFocus
End If
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

Public Sub VAN_RET_NUM()
Dim INV_DAT As New ADODB.Recordset
Dim VAN_DAT As New ADODB.Recordset
Dim RSI_DAT As New ADODB.Recordset
Dim RSR_DAT As New ADODB.Recordset
Me.MousePointer = vbHourglass
con.Execute "DELETE FROM VAN"
VAN_DAT.Open "SELECT * FROM VAN", con, adOpenKeyset, adLockPessimistic
SAL_QRY = "VAL(SUBSTR(PNL->INV_IDY,4,7))>=" & Val(Text1) & ".AND.VAL(SUBSTR(PNL->INV_IDY,4,7))<=" & Val(Text2)
INV_DAT.Open "SELECT * FROM INV WHERE " & SAL_QRY & " AND TXN_TYP LIKE 'RSI'", con, adOpenKeyset, adLockPessimistic
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


Public Sub DOSREPORT(qry As String, MOVEREC As Integer)
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM RRI", PNL, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then
    REC_SET.Move MOVEREC
    REC_SET!RI_QUERY = "O"
    REC_SET!ri_filter = Trim(qry)
    REC_SET!ri_scope = ""
    REC_SET!ri_loscope = ""
    REC_SET!ri_hiscope = ""
    REC_SET!ri_copies = 1
    REC_SET!ri_begpage = 1
    REC_SET!ri_endpage = 9999999
    REC_SET!ri_printer = ""
    REC_SET!ri_mindex = CURDIR & "\PNV.CDX"
    REC_SET.Update
End If
REC_SET.Close
End Sub
