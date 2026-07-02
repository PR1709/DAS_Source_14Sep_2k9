VERSION 5.00
Begin VB.Form M_FRM_IFD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import DAS Data"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_IFD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   23
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   22
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Select Applicable Addons "
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   3015
      Begin VB.CheckBox chkCal 
         Caption         =   "Calen&dar"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   680
         Width           =   1335
      End
      Begin VB.CheckBox chkTgt 
         Caption         =   "&Targets"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkSch 
         Caption         =   "Sc&hemes"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkRdy 
         Caption         =   "&Route Plan "
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkShells 
         Caption         =   "&Modules"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   680
         Width           =   1455
      End
      Begin VB.CheckBox chkAst 
         Caption         =   "&Assets"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Version "
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
      Begin VB.OptionButton optEarVer 
         Caption         =   "&Earlier Versions"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optCurVer 
         Caption         =   "Current &Version 4.0"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Select"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton optComData 
         Caption         =   "Co&mplete Data"
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optFileData 
         Caption         =   "&File Data"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Select Source Folder"
      Height          =   4575
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   5175
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   2415
      End
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label lblvalidfolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton opt_eff 
         Caption         =   "DAS"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opt_efx 
         Caption         =   "Efflinx"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   6000
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processing Selection... Please Wait"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   2520
   End
End
Attribute VB_Name = "M_FRM_IFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sConn As New ADODB.Connection    ' Source
Dim dConn As New ADODB.Connection    ' Destination
Dim tConn As New ADODB.Recordset     ' used to access tables from the eff database
Dim SRS As New ADODB.Recordset
Dim DRS As New ADODB.Recordset
Dim TB As TableDef
Dim tb1 As TableDef
Dim TNAME As String
Dim i As Integer
Private Sub chkAst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then chkShells.SetFocus
End Sub
Private Sub chkCal_Click()
If chkCal.Value = 1 Then
    list1.AddItem "wek"
Else
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "wek" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
End If
End Sub

Private Sub chkCal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Command1.Enabled = True Then Command1.SetFocus
End If
End Sub

Private Sub chkRdy_Click()
If chkRdy.Value = 1 Then
    list1.AddItem "rdy"
Else
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "rdy" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
End If
End Sub
Private Sub chkRdy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then chkSch.SetFocus
End Sub
Private Sub chkSch_Click()
If chkSch.Value = 1 Then
    list1.AddItem "sch"
    If Dir(CURDIR + "\sid.dbf") <> "" Then
        list1.AddItem "sid"
    End If
Else
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "sch" Then
            list1.RemoveItem K                 ' Scheme Number to be imported into IDY as per Prince Chopra feedback
            Exit For
        End If
    Next
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "sid" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
End If
End Sub
Private Sub chkSch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then chkTgt.SetFocus
End Sub
Private Sub chkShells_Click()
If chkShells.Value = 1 Then
    list1.AddItem "avt"
    list1.AddItem "shl"
    'list1.AddItem "snv"
Else
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "avt" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "shl" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
'    For K = 0 To list1.ListCount - 1
'        If Trim(list1.List(K)) = "snv" Then
'            list1.RemoveItem K
'            Exit For
'        End If
'    Next
End If
End Sub
Private Sub chkShells_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then chkRdy.SetFocus
End Sub
Private Sub chkTgt_Click()
If chkTgt.Value = 1 Then
    list1.AddItem "tgt"
    list1.AddItem "tgb"
    list1.AddItem "tgc"
    list1.AddItem "tgs"
    list1.AddItem "rtg"
Else
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "tgt" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "tgb" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "tgc" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "tgs" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To list1.ListCount - 1
        If Trim(list1.List(K)) = "rtg" Then
            list1.RemoveItem K
            Exit For
        End If
    Next
End If
End Sub
Private Sub chkTgt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then chkCal.SetFocus
End Sub
Private Sub Command1_Click()
On Error GoTo lab
Dim tmp As Integer
Dim FLD As String
If Frame5.Enabled = False Then
    If optFileData.Value = True And optEarVer.Value = True Then
        IMP_BKP
        'MASTERS FROM EARLIER VERSION
        EmasData
    ElseIf optFileData.Value = True And optCurVer.Value = True Then
        IMP_BKP
        CmasData
    End If
    If optComData.Value = True And optEarVer.Value = True Then
        IMP_BKP
        'COMPLETE DATA FROM EARLIER VERSION
        etrsData
    ElseIf optComData.Value = True And optCurVer.Value = True Then
        IMP_BKP
        ctrsData
    End If
End If
If Dir1.Path = App.Path Then
    MsgBox "Current Folder Path Selected! Invalid Folder Path to import File Data", vbCritical, "DAS Version 4.0"
    Exit Sub
End If
If Len(Dir1.Path) <> 3 Then
    If Dir(Dir1.Path & "\EFF.DBC") = "" Then
        MsgBox "Not A Valid DAS Version 4.0 Database Folder Path", vbCritical, "DAS Version 4.0"
        Exit Sub
    End If
Else
    If Dir(Dir1.Path & "EFF.DBC") = "" Then
        MsgBox "Not A Valid DAS Version 4.0 Database Folder Path", vbCritical, "DAS Version 4.0"
        Exit Sub
    End If
End If
If MsgBox("This Action Will Delete Existing Data ! Continue?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
    Exit Sub
Else
    If MsgBox("Confirm Data Import From " & Dir1.Path, vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then Exit Sub
End If
Me.MousePointer = vbHourglass
Label4.Visible = True
Label4.Refresh
If sConn.State = 1 Then sConn.Close ' Source Connection
sConn.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & Dir1.Path & "\EFF.DBC;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
If dConn.State = 1 Then dConn.Close ' Destination Connection
dConn.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\EFF.DBC;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
Set tConn = sConn.OpenSchema(adSchemaTables)    ' Setting tables to the Recordset Object
dConn.BeginTrans
While Not tConn.EOF
    For i = 0 To list1.ListCount - 1
        If Trim(tConn!TABLE_NAME) = Trim(list1.List(i)) Then
            dConn.Execute "delete from " & Trim(tConn!TABLE_NAME)
            If SRS.State = 1 Then SRS.Close
            If DRS.State = 1 Then DRS.Close
            SRS.Open "select * from " & Trim(tConn!TABLE_NAME), sConn, adOpenKeyset, adLockPessimistic
            DRS.Open "select * from " & Trim(tConn!TABLE_NAME), dConn, adOpenKeyset, adLockPessimistic
            If Trim(tConn!TABLE_NAME) = "act" Then
                    While Not SRS.EOF
                        DRS.AddNew
                            For K = 0 To DRS.Fields.count - 1
                                If DRS(K).Type = adVarChar Or DRS(K).Type = adBSTR Or DRS(K).Type = adChar Or DRS(K).Type = adLongVarChar Or DRS(K).Type = adSmallInt Then
                                    DRS(K) = ""
                                ElseIf DRS(K).Type = adInteger Or DRS(K).Type = adNumeric Or DRS(K).Type = adBigInt Or DRS(K).Type = adSingle Or DRS(K).Type = adDouble Then
                                    DRS(K) = 0
                                ElseIf DRS(K).Type = adDate Or DRS(K).Type = adDBDate Or DRS(K).Type = adDBTimeStamp Then
                                    DRS(K) = Date
                                End If
                            Next
                            
                            For K = 0 To SRS.Fields.count - 1
                                FLD = SRS.Fields(K).name
                                If Trim(FLD) = "cst_typ" Then
                                    If Trim(SRS(FLD)) = "" Then
                                        DRS(FLD) = "Retailer"
                                    Else
                                        DRS(FLD) = SRS(FLD)
                                    End If
                                'Making Key_act value to "N" if the value is Null from Earlier version Import
                                ElseIf Trim(FLD) = "Key_act" Then
                                    If IsNull(SRS(FLD)) = True Then
                                        DRS(FLD) = "N"
                                    End If
                                ElseIf Trim(FLD) = "dbt_rap" Then
                                    If Trim(SRS(FLD)) = "" Then
                                        DRS(FLD) = "B"
                                    Else
                                        'If CHKfLD(tConn!TABLE_NAME, FLD) = True Then DRS(FLD) = SRS(FLD)
                                        DRS(FLD) = SRS(FLD)
                                    End If
                                Else
                                    'If CHKfLD(tConn!TABLE_NAME, FLD) = True Then DRS(FLD) = SRS(FLD)
                                    If Trim(SRS(FLD)) = "" Then
                                    DRS(FLD) = ""
                                    Else
                                    DRS(FLD) = SRS(FLD)
                                    End If
                                End If
                            Next
                        DRS.Update
                        SRS.MoveNext
                    Wend
                Exit For
            End If
            While Not SRS.EOF
                DRS.AddNew
                For K = 0 To DRS.Fields.count - 1
                    If DRS(K).Type = adVarChar Or DRS(K).Type = adBSTR Or DRS(K).Type = adChar Or DRS(K).Type = adLongVarChar Or DRS(K).Type = adSmallInt Then
                        DRS(K) = ""
                    ElseIf DRS(K).Type = adInteger Or DRS(K).Type = adNumeric Or DRS(K).Type = adBigInt Or DRS(K).Type = adSingle Or DRS(K).Type = adDouble Then
                        DRS(K) = 0
                    ElseIf DRS(K).Type = adDate Or DRS(K).Type = adDBDate Or DRS(K).Type = adDBTimeStamp Then
                        DRS(K) = Date
                    End If
                Next
                For K = 0 To SRS.Fields.count - 1
                    FLD = SRS.Fields(K).name
                    'If CHKfLD(tConn!TABLE_NAME, FLD) = True Then DRS(FLD) = SRS(FLD)
                    If Trim(SRS(FLD)) = "" Then
                    DRS(FLD) = ""
                    Else
                    DRS(FLD) = SRS(FLD)
                    End If
                Next
                DRS.Update
                SRS.MoveNext
            Wend
            Exit For
        End If
    Next
    tConn.MoveNext
Wend
If Frame5.Enabled = False Then
' Updation According to Version
    If optEarVer.Value = True Then
        If SRS.State = 1 Then SRS.Close
        If DRS.State = 1 Then DRS.Close
        SRS.Open "SELECT * FROM PRD", dConn, adOpenDynamic, adLockPessimistic
        DRS.Open "SELECT * FROM BRD", dConn, adOpenDynamic, adLockPessimistic
        dConn.Execute "delete from brd"
        While Not SRS.EOF
            DRS.AddNew
            For K = 0 To DRS.Fields.count - 1
                If DRS(K).Type = adVarChar Or DRS(K).Type = adBSTR Or DRS(K).Type = adChar Or DRS(K).Type = adLongVarChar Or DRS(K).Type = adSmallInt Then
                    DRS(K) = ""
                ElseIf DRS(K).Type = adInteger Or DRS(K).Type = adNumeric Or DRS(K).Type = adBigInt Or DRS(K).Type = adSingle Or DRS(K).Type = adDouble Then
                    DRS(K) = 0
                ElseIf DRS(K).Type = adDate Or DRS(K).Type = adDBDate Or DRS(K).Type = adDBTimeStamp Then
                    DRS(K) = Date
                End If
            Next
            DRS("prd_idy") = SRS("prd_idy")
            DRS!Bat_nme = ""
            DRS("PRD_PDR") = SRS("pur_rat")
            DRS("prd_sdr") = SRS("sal_rat")
            DRS("prd_mrp") = SRS("mrp_rat")
            DRS("prd_qty") = SRS("cur_lev") '0 is replaced with "cur_lev" by pkb
            DRS("bat_idy") = "BAT" & Right(SRS("prd_idy"), 7)
            DRS("brd_dom") = Date$
            DRS("brd_doe") = Date$
            DRS.Update
            SRS.MoveNext
        Wend
        dConn.Execute "UPDATE IDY SET IDY_BAT=''"
        dConn.Execute "UPDATE IDY SET IDY_BAT=IDY_PRD"
        If SRS.State = 1 Then SRS.Close
        If DRS.State = 1 Then DRS.Close
        DRS.Open "select * from brd", dConn, adOpenDynamic, adLockPessimistic
        While Not DRS.EOF
            dConn.Execute "update inv set bat_idy='" & Trim(DRS("bat_idy")) & "'" & " where prd_idy like '" & Trim(DRS("prd_idy")) & "'"
            DRS.MoveNext
        Wend
    ElseIf optCurVer.Value = True Then
        If optFileData.Value = True Then
            If DRS.State = 1 Then DRS.Close
            DRS.Open "SELECT * FROM IDY", dConn, adOpenKeyset, adLockPessimistic
            If Not DRS.EOF Then
                list1.AddItem "bat" 'edited
                For n = 0 To DRS.Fields.count - 1
                tmp = 0
                    For K = 0 To list1.ListCount - 1
                        If Mid(Trim(DRS.Fields(n).name), 5, 3) = Trim(list1.List(K)) Then
                            tmp = 1
                            Exit For
                        End If
                    Next
                    If tmp = 0 Then
                        If Trim(Mid(Trim(DRS.Fields(n).name), 5, 3)) = "sch" Then
                        Else
                            dConn.Execute "UPDATE IDY SET IDY_" & Trim(Mid(Trim(DRS.Fields(n).name), 5, 3)) & " = '0'"
                        End If
                    End If
                Next
                list1.RemoveItem Val(list1.ListCount - 1)   ' Edited
            End If
        End If
    End If
    Me.MousePointer = vbNormal
    Label4.Visible = False
    If MsgBox("Process Completed! Proceed for Addons ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        Frame5.Enabled = True
        chkAst.SetFocus
        list1.CLEAR
    Else
        Unload Me
    End If
Else
    MsgBox "Process Completed!", vbInformation, "DAS Version 4.0"
    Unload Me
End If
dConn.CommitTrans
Label4.Visible = False
Me.MousePointer = vbNormal
Exit Sub
lab:
If err.Number = 3265 Then
    Resume Next
ElseIf err.Number = -2147467259 Then
    MsgBox "Invalid Database! Validate Database for Missing Files to Import", vbExclamation, "DAS Version 4.0"
    Me.MousePointer = vbNormal
    dConn.RollbackTrans
    If dConn.State = 1 Then dConn.Close
    Exit Sub
Else
    MsgBox err.Number & " " & err.Description & " " & Trim(tConn!TABLE_NAME)
    dConn.RollbackTrans
    If dConn.State = 1 Then dConn.Close
    Exit Sub
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Dir1_Change()
st = Dir1.Path
If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
st = Mid(Dir1.Path, 1, Len(Dir1.Path) - 1)
End If
If Dir(st & "\EFF.DBC") = "" Then
    lblvalidfolder = UCase("Invalid DAS Version 4.0 Data folder")
Else
    lblvalidfolder = UCase("Valid DAS Version 4.0 Data folder")
End If
End Sub
Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.Enabled = True
    Command1.SetFocus
End If
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Dir1.SetFocus
End Sub

Private Sub Form_Activate()
optFileData.SetFocus
End Sub

Private Sub EmasData()
With list1
    .CLEAR
    .AddItem "act"
    .AddItem "cat"
    .AddItem "csr"
    .AddItem "dir"
    .AddItem "grp"
    .AddItem "idy" 'When master data import is selected all the fields where transaction record number should be made to zero
    .AddItem "loc"
    .AddItem "mar"
    .AddItem "pgr"
    .AddItem "prd"
    .AddItem "pre"
    .AddItem "prt"
    .AddItem "smn"
End With
End Sub
Private Sub etrsData()
With list1
    .CLEAR
    EmasData
    .AddItem "act"
    .AddItem "cat"
    .AddItem "csr"
    .AddItem "dir"
    .AddItem "grp"
    .AddItem "idy"
    .AddItem "inl"
    .AddItem "inv"
    .AddItem "lev"
    .AddItem "lnv"
    .AddItem "loc"
    .AddItem "mar"
    .AddItem "pgr"
    .AddItem "prd"
    .AddItem "pre"
    .AddItem "prt"
    .AddItem "rtp"
    .AddItem "smn"
    .AddItem "van"
    .AddItem "vch"
 '   .AddItem "lds" 'Changed by pkb to be verified with PNR Sir
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub optComData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    optCurVer.SetFocus
End If
End Sub

Private Sub optCurVer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Drive1.SetFocus
End Sub

Private Sub optEarVer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Drive1.SetFocus
End If
End Sub

Private Sub optFileData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    optCurVer.SetFocus
End If
End Sub
Private Sub CmasData()
With list1
    .CLEAR
    .AddItem "act"
    .AddItem "cat"
    .AddItem "csr"
    .AddItem "dir"
    .AddItem "grp"
    .AddItem "idy" 'When master data import is selected all the fields where transaction record number should be made to zero
    .AddItem "loc"
    .AddItem "mar"
    .AddItem "pgr"
    .AddItem "prd"
    .AddItem "brd"
    .AddItem "pre"
    .AddItem "prt"
    .AddItem "smn"
    .AddItem "shl"
    .AddItem "irt"
    .AddItem "aca"
    .AddItem "asm"
    .AddItem "ast"
    .AddItem "att"
    .AddItem "com" 'Added by PNR
'    .AddItem "avt"
'    .AddItem "rtg"
End With
End Sub
Private Sub ctrsData()
With list1
    .CLEAR
    CmasData
    .AddItem "act"
    .AddItem "cat"
    .AddItem "csr"
    .AddItem "dir"
    .AddItem "grp"
    .AddItem "idy"
    .AddItem "inl"
    .AddItem "inv"
    .AddItem "lev"
    .AddItem "lnv"
    .AddItem "loc"
    .AddItem "mar"
    .AddItem "pgr"
    .AddItem "prd"
    .AddItem "pre"
    .AddItem "prt"
    .AddItem "rtp"
    .AddItem "smn"
    .AddItem "van"
    .AddItem "vch"
    .AddItem "shl"
    .AddItem "snv"
    .AddItem "irt"
    .AddItem "aca"
    .AddItem "asm"
    .AddItem "ast"
    .AddItem "att"
    .AddItem "cur" 'Added by PNR
    .AddItem "com" 'Added by PNR
    .AddItem "lds" 'added by pkb to be check with PNR Sir CONFIRMED FROM TSR
'    .AddItem "avt"
'    .AddItem "rtg"
End With
End Sub
Private Function CHKFLD(TB As String, FNAME As String) As Boolean
Dim FD As Integer
Dim rs As New ADODB.Recordset
FD = 0
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM " & TB, con, adOpenStatic
For f = 0 To rs.Fields.count - 1
    If Trim(rs.Fields(f).name) = Trim(FNAME) Then
        FD = 1
        Exit For
    End If
Next
If FD = 1 Then
    CHKFLD = True
Else
    CHKFLD = False
End If
End Function
