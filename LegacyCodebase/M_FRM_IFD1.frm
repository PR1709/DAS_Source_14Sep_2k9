VERSION 5.00
Begin VB.Form M_FRM_IFD1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Effmcg Data"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_IFD1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
         Caption         =   "Shells n &Packs"
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
         Width           =   1575
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
         Top             =   240
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
         Caption         =   "Effmcg"
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
Attribute VB_Name = "M_FRM_IFD1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sConn As New ADODB.Connection    ' Source
Dim dConn As New ADODB.Connection ' Destination
Dim tConn As New ADODB.Recordset   ' used to access tables from the eff database
Dim sRs As New ADODB.Recordset
Dim dRs As New ADODB.Recordset
Dim TB As TableDef
Dim tb1 As TableDef
Dim TNAME As String
Dim I As Integer

Private Sub chkAst_Click()
If chkAst.Value = 1 Then
    List1.AddItem "aca"
    List1.AddItem "asm"
    List1.AddItem "ast"
    List1.AddItem "att"
Else
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "aca" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "asm" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "ast" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "att" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
End If
End Sub

Private Sub chkAst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then chkShells.SetFocus
End Sub

Private Sub chkCal_Click()
If chkCal.Value = 1 Then
    List1.AddItem "wek"
Else
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "wek" Then
            List1.RemoveItem K
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
    List1.AddItem "rdy"
Else
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "rdy" Then
            List1.RemoveItem K
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
    List1.AddItem "sch"
Else
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "sch" Then
            List1.RemoveItem K
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
    List1.AddItem "shl"
Else
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "shl" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
End If
End Sub

Private Sub chkShells_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then chkRdy.SetFocus
End Sub

Private Sub chkTgt_Click()
If chkTgt.Value = 1 Then
    List1.AddItem "tgt"
    List1.AddItem "tgb"
    List1.AddItem "tgc"
    List1.AddItem "tgs"
Else
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "tgt" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "tgb" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "tgc" Then
            List1.RemoveItem K
            Exit For
        End If
    Next
    For K = 0 To List1.ListCount - 1
        If Trim(List1.list(K)) = "tgs" Then
            List1.RemoveItem K
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
        EmasData
    ElseIf optFileData.Value = True And optCurVer.Value = True Then
        CmasData
    End If
    If optComData.Value = True And optEarVer.Value = True Then
        etrsData
    ElseIf optComData.Value = True And optCurVer.Value = True Then
        ctrsData
    End If
End If
If Dir1.Path = App.Path Then
    MsgBox "Current Folder Path Selected! Invalid Folder Path to import File Data", vbCritical, "Effmcg"
    Exit Sub
End If
If Len(Dir1.Path) <> 3 Then
    If Dir(Dir1.Path & "\EFF.DBC") = "" Then
        MsgBox "Not A Valid Effmcg Database Folder Path", vbCritical, "Effmcg"
        Exit Sub
    End If
Else
    If Dir(Dir1.Path & "EFF.DBC") = "" Then
        MsgBox "Not A Valid Effmcg Database Folder Path", vbCritical, "Effmcg"
        Exit Sub
    End If
End If
If MsgBox("This Action Will Delete Existing Data ! Continue?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
Else
    If MsgBox("Confirm Data Import From " & Dir1.Path, vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If
Me.MousePointer = vbHourglass
Label4.Visible = True
Label4.Refresh
If sConn.State = 1 Then sConn.Close ' Source Connection
sConn.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & Dir1.Path & "\Eff.dbc;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
If dConn.State = 1 Then dConn.Close ' Destination Connection
dConn.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\Eff.dbc;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
Set tConn = sConn.OpenSchema(adSchemaTables)    ' Setting tables to the Recordset Object
While Not tConn.EOF
    For I = 0 To List1.ListCount - 1
        If Trim(tConn!TABLE_NAME) = Trim(List1.list(I)) Then
            dConn.Execute "delete from " & Trim(tConn!TABLE_NAME)
            If sRs.State = 1 Then sRs.Close
            If dRs.State = 1 Then dRs.Close
            sRs.Open "select * from " & Trim(tConn!TABLE_NAME), sConn, adOpenKeyset, adLockOptimistic
            dRs.Open "select * from " & Trim(tConn!TABLE_NAME), dConn, adOpenKeyset, adLockOptimistic
            If Trim(tConn!TABLE_NAME) = "act" Then
                While Not sRs.EOF
                    dRs.AddNew
                        For K = 0 To dRs.Fields.Count - 1
                            If dRs(K).Type = adVarChar Or dRs(K).Type = adBSTR Or dRs(K).Type = adChar Or dRs(K).Type = adLongVarChar Or dRs(K).Type = adSmallInt Then
                                dRs(K) = ""
                            ElseIf dRs(K).Type = adInteger Or dRs(K).Type = adNumeric Or dRs(K).Type = adBigInt Or dRs(K).Type = adSingle Or dRs(K).Type = adDouble Then
                                dRs(K) = 0
                            ElseIf dRs(K).Type = adDate Or dRs(K).Type = adDBDate Or dRs(K).Type = adDBTimeStamp Then
                                dRs(K) = Date
                            End If
                        Next
                        For K = 0 To sRs.Fields.Count - 1
                            FLD = sRs.Fields(K).name
                            If Trim(FLD) = "cst_typ" Then
                                If Trim(sRs(FLD)) = "" Then
                                    dRs(FLD) = "Retailer"
                                Else
                                    dRs(FLD) = sRs(FLD)
                                End If
                            ElseIf Trim(FLD) = "dbt_rap" Then
                                If Trim(sRs(FLD)) = "" Then
                                    dRs(FLD) = "B"
                                Else
                                    If CHKfLD(tConn!TABLE_NAME, FLD) = True Then dRs(FLD) = sRs(FLD)
                                End If
                            Else
                                If CHKfLD(tConn!TABLE_NAME, FLD) = True Then dRs(FLD) = sRs(FLD)
                            End If
                        Next
                    dRs.Update
                    sRs.MoveNext
                Wend
                Exit For
            End If
            While Not sRs.EOF
                dRs.AddNew
                For K = 0 To dRs.Fields.Count - 1
                    If dRs(K).Type = adVarChar Or dRs(K).Type = adBSTR Or dRs(K).Type = adChar Or dRs(K).Type = adLongVarChar Or dRs(K).Type = adSmallInt Then
                        dRs(K) = ""
                    ElseIf dRs(K).Type = adInteger Or dRs(K).Type = adNumeric Or dRs(K).Type = adBigInt Or dRs(K).Type = adSingle Or dRs(K).Type = adDouble Then
                        dRs(K) = 0
                    ElseIf dRs(K).Type = adDate Or dRs(K).Type = adDBDate Or dRs(K).Type = adDBTimeStamp Then
                        dRs(K) = Date
                    End If
                Next
                For K = 0 To sRs.Fields.Count - 1
                    FLD = sRs.Fields(K).name
                    If CHKfLD(tConn!TABLE_NAME, FLD) = True Then dRs(FLD) = sRs(FLD)
                Next
                dRs.Update
                sRs.MoveNext
            Wend
            Exit For
        End If
    Next
    tConn.MoveNext
Wend
If Frame5.Enabled = False Then
'    If Dir(Dir1.Path + "\rdy.dbf") = "" Then
'        If sRs.State = 1 Then sRs.Close
'        If dRs.State = 1 Then dRs.Close
'        sRs.Open "select act_idy,smn_idy,rou_idy from act where act_typ='SAL'", sConn, adOpenKeyset, adLockOptimistic
'        dRs.Open "select * from rdy where act_idy=''", sConn, adOpenKeyset, adLockOptimistic
'        While Not sRs.EOF
'            For K = 1 To 7
'                dRs("act_idy") = sRs("act_idy")
'                dRs("day_idy") = K
'                dRs("smn_idy") = sRs("smn_idy")
'                dRs("rou_idy") = sRs("rou_idy")
'                dRs("seq_num") = 0
'                dRs("rou_dst") = 0
'                dRs("tim_inn") = 0
'                dRs("tim_out") = 0
'            Next
'            sRs.MoveNext
'        Wend
'    End If
' Updation According to Version
    If optEarVer.Value = True Then
        If sRs.State = 1 Then sRs.Close
        If dRs.State = 1 Then dRs.Close
        sRs.Open "SELECT * FROM PRD", dConn, adOpenDynamic, adLockOptimistic
        dRs.Open "SELECT * FROM BRD", dConn, adOpenDynamic, adLockOptimistic
        dConn.Execute "delete from brd"
        While Not sRs.EOF
            dRs.AddNew
            For K = 0 To dRs.Fields.Count - 1
                If dRs(K).Type = adVarChar Or dRs(K).Type = adBSTR Or dRs(K).Type = adChar Or dRs(K).Type = adLongVarChar Or dRs(K).Type = adSmallInt Then
                    dRs(K) = ""
                ElseIf dRs(K).Type = adInteger Or dRs(K).Type = adNumeric Or dRs(K).Type = adBigInt Or dRs(K).Type = adSingle Or dRs(K).Type = adDouble Then
                    dRs(K) = 0
                ElseIf dRs(K).Type = adDate Or dRs(K).Type = adDBDate Or dRs(K).Type = adDBTimeStamp Then
                    dRs(K) = Date
                End If
            Next
            dRs("prd_idy") = sRs("prd_idy")
            dRs!Bat_nme = ""
            dRs("PRD_PDR") = sRs("pur_rat")
            dRs("prd_sdr") = sRs("sal_rat")
            dRs("prd_mrp") = sRs("mrp_rat")
            dRs("prd_qty") = 0
            dRs("bat_idy") = "BAT" & Right(sRs("prd_idy"), 7)
            dRs("brd_dom") = Date$
            dRs("brd_doe") = Date$
            dRs.Update
            sRs.MoveNext
        Wend
        dConn.Execute "UPDATE IDY SET IDY_BAT=''"
        dConn.Execute "UPDATE IDY SET IDY_BAT=IDY_PRD"
        If sRs.State = 1 Then sRs.Close
        If dRs.State = 1 Then dRs.Close
        dRs.Open "select * from brd", dConn, adOpenDynamic, adLockOptimistic
        While Not dRs.EOF
            dConn.Execute "update inv set bat_idy='" & Trim(dRs("bat_idy")) & "'" & " where prd_idy like '" & Trim(dRs("prd_idy")) & "'"
            dRs.MoveNext
        Wend
    ElseIf optCurVer.Value = True Then
        If optFileData.Value = True Then
            If dRs.State = 1 Then dRs.Close
            dRs.Open "SELECT * FROM IDY", dConn, adOpenKeyset, adLockOptimistic
            If Not dRs.EOF Then
                For n = 0 To dRs.Fields.Count - 1
                tmp = 0
                    For K = 0 To List1.ListCount - 1
                        If Mid(Trim(dRs.Fields(n).name), 5, 3) = Trim(List1.list(K)) Then
                            tmp = 1
                            Exit For
                        End If
                    Next
                    If tmp = 0 Then
                        dConn.Execute "UPDATE IDY SET IDY_" & Trim(Mid(Trim(dRs.Fields(n).name), 5, 3)) & " = '0'"
                    End If
                Next
            End If
        End If
    End If
    Me.MousePointer = vbNormal
    Label4.Visible = False
    If MsgBox("Process Completed! Proceed for Addons ?", vbQuestion + vbYesNo, "Effmcg") = vbYes Then
        Frame5.Enabled = True
        chkAst.SetFocus
        List1.Clear
    Else
        Unload Me
    End If
Else
    MsgBox "Process Completed!", vbExclamation, "Effmcg"
    Unload Me
End If
Label4.Visible = False
Me.MousePointer = vbNormal
Exit Sub
lab:
If err.Number = 3265 Then
    MsgBox err.Description
    'Resume Next
Else
    MsgBox err.Description
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
If Dir(st & "\eff.dbc") = "" Then
    lblvalidfolder = UCase("Invalid Effmcg Data folder")
Else
    lblvalidfolder = UCase("Valid Effmcg Data folder")
End If
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.Enabled = True
    Command1.SetFocus
End If
End Sub

Private Sub Drive1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Dir1.SetFocus
End Sub

Private Sub Form_Activate()
optFileData.SetFocus
End Sub

Private Sub EmasData()
With List1
    .Clear
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
With List1
    .Clear
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
With List1
    .Clear
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
End With
End Sub
Private Sub ctrsData()
With List1
    .Clear
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
End With
End Sub


Private Function CHKfLD(TB As String, FNAME As String) As Boolean
Dim FD As Integer
Dim RS As New ADODB.Recordset
FD = 0
If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM " & TB, CON, adOpenStatic
For f = 0 To RS.Fields.Count - 1
    If Trim(RS.Fields(f).name) = Trim(FNAME) Then
        FD = 1
        Exit For
    End If
Next
If FD = 1 Then
    CHKfLD = True
Else
    CHKfLD = False
End If
End Function
