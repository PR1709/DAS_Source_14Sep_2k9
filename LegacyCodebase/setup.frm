VERSION 5.00
Begin VB.Form setup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS New Company"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   Icon            =   "setup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   2640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3750
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton finish 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Select Destination Path"
      Height          =   195
      Left            =   825
      TabIndex        =   9
      Top             =   1620
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      Height          =   195
      Left            =   1365
      TabIndex        =   5
      Top             =   668
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter File Folder Name"
      Height          =   195
      Left            =   885
      TabIndex        =   3
      Top             =   1155
      Width           =   1605
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Sub cancel_Click()
CHK_EXT = MsgBox("Exit Without Company Creation?", vbYesNo, "DAS Version 4.0")
If CHK_EXT = 6 Then
Unload Me
End If
End Sub

Private Sub Drive1_Change()
On Error GoTo err
    Dir1.Path = Drive1.Drive   ' Set directory path.
    Exit Sub
err:
    MsgBox "Please Insert Disk !", vbCritical, "DAS Version 4.0"
    Drive1.SetFocus
End Sub

Private Sub Dir1_Change()
Text2 = Dir1.Path   ' Set file path.
Text2.SelStart = Len(Text1)
Text2.SetFocus
End Sub

Private Sub finish_click()
On Error GoTo lab:
Me.MousePointer = vbHourglass
finish.Enabled = False
If Len(Trim(Text1)) = 0 Then
    MsgBox "Please Enter Valid Company Name (10) Characters!", vbCritical, "DAS Version 4.0"
    Me.MousePointer = vbNormal
    finish.Enabled = True
    Text1.SetFocus
    Exit Sub
End If
If Len(Trim(Text7)) = 0 Then
    MsgBox "Invalid Folder Name", vbCritical, "DAS Version 4.0"
    Me.MousePointer = vbNormal
    finish.Enabled = True
    Text7.SetFocus
    Exit Sub
End If
If Len(Trim(Text2)) = 0 Then
    MsgBox "Invalid Destination Drive n Folder!", vbCritical, "DAS Version 4.0"
    Me.MousePointer = vbNormal
    finish.Enabled = True
    Text1.SetFocus
    Exit Sub
End If
If company.SysInfo1.OSPlatform = 1 Then
    Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
ElseIf company.SysInfo1.OSPlatform = 2 Then
    Open "c:\winnt\system\EFFMCG4.txt" For Input As #1   'Open file for output.
Else
    Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
End If
Do While Not EOF(1)
    Line Input #1, textline  ' Read line into variable.
    A = Split(textline, "@", -1, vbTextCompare)
    If UCase(Mid(Trim(A(0)), 2, Len(A(0)) - 1)) = UCase(Trim(Text1)) Then
    MsgBox "Company Already Exists!", vbInformation, "DAS Version 4.0"
    Me.MousePointer = vbNormal
    finish.Enabled = True
    Close #1
    Exit Sub
    End If
Loop
Close #1
If Dir(App.Path & "\bkp", vbDirectory) = "" Then MkDir App.Path & "\bkp"
If Len(Trim(Text2)) = 3 Then
Destdir = Text2 & Text7
Else
Destdir = Text2 & "\" & Text7
End If
MkDir Destdir
If company.SysInfo1.OSPlatform = 1 Then
    Open "c:\windows\system\EFFMCG4.txt" For Append As #1   'Open file for output.
ElseIf company.SysInfo1.OSPlatform = 2 Then
    Open "c:\winnt\system\EFFMCG4.txt" For Append As #1   'Open file for output.
Else
    Open "c:\windows\system\EFFMCG4.txt" For Append As #1   'Open file for output.
End If
If LCase(Text2) <> "c:\" Then
    rec = Trim(Text1) & "@" & Trim(Text7) & "@" & Destdir
Else
    rec = Trim(Text1) & "@" & Trim(Text7) & "@" & Trim(Text2) & Text7
End If
Write #1, rec
Close #1
Set FS = CreateObject("SCRIPTING.FILESYSTEMOBJECT")
    FS.CopyFolder App.Path & "\bkp", Destdir
    
    If cn.State = 1 Then cn.Close
    cn.Open "Driver={Microsoft Visual Foxpro Driver};UID=;PWD=;SourceDB=" & Destdir & "\Eff.dbc;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    
    Set f = FS.GetFolder(Destdir)
    Set FN = f.Files
    
    For Each F1 In FN
        tp = F1.Type
        If Right(Trim(F1.name), 4) = ".dbf" Then
            If Trim(F1.name) = "dir.dbf" Or Trim(F1.name) = "pre.dbf" Or Trim(F1.name) = "idy.dbf" Or Trim(F1.name) = "cat.dbf" Or Trim(F1.name) = "loc.dbf" Or Trim(F1.name) = "grp.dbf" Or Trim(F1.name) = "typ.dbf" Or Trim(F1.name) = "fun.dbf" Or Trim(F1.name) = "rri.dbf" Or Trim(F1.name) = "foxuser.dbf" Or Trim(F1.name) = "Dkcontrl.dbf" Or Trim(F1.name) = "Optional.dbf" Or Trim(F1.name) = "Required.dbf" Or Trim(F1.name) = "str_doc.dbf" Or Trim(F1.name) = "pgr.dbf" Or Trim(F1.name) = "Reports.dbf" Or Trim(F1.name) = "yer.dbf" Or Trim(F1.name) = "pre.dbf" Then
            ElseIf Trim(F1.name) = "usr.dbf" Then
                cn.Execute "delete from " & Trim(Trim(F1.name)) & " where usr_nme<>'ADMINISTRATOR'"
            Else
                cn.Execute "delete from " & Trim(Trim(F1.name))
            End If
        End If
    Next
     
    cn.Execute "update idy set idy_sal='0'"
    cn.Execute "update idy set idy_act='0'"
    cn.Execute "update idy set idy_prd='0'"
'    CN.Execute "update idy set idy_pgr='0'"
    cn.Execute "update idy set idy_pur='0'"
    cn.Execute "update idy set idy_srn='0'"
    cn.Execute "update idy set idy_prt='0'"
    cn.Execute "update idy set idy_ops='0'"
    cn.Execute "update idy set idy_rct='0'"
    cn.Execute "update idy set idy_pay='0'"
    cn.Execute "update idy set idy_jrn='0'"
    cn.Execute "update idy set idy_opn='0'"
    cn.Execute "update idy set idy_rsi='0'"
    cn.Execute "update idy set idy_mar='0'"
    cn.Execute "update idy set idy_smn='0'"
    cn.Execute "update idy set idy_chn='0'"
    cn.Execute "update idy set idy_ord='0'"
    cn.Execute "update idy set idy_otc='0'"
    cn.Execute "update idy set idy_irt='0'"
    cn.Execute "update idy set idy_CRN='0'"
    cn.Execute "update idy set idy_DBN='0'"
    cn.Execute "update idy set idy_STI='0'"
    cn.Execute "update idy set idy_STO='0'"
    cn.Execute "update idy set idy_LDS='0'"
    cn.Execute "update idy set idy_SCH='0'"
    cn.Execute "update idy set idy_BAT='0'"
    cn.Execute "update idy set idy_TXN='0'"
    cn.Execute "update idy set idy_SHL='0'"
    cn.Execute "update idy set idy_SHT='0'"
    cn.Execute "update idy set idy_adj='0'"
    cn.Execute "update fun set fun_two=''"
    cn.Execute "update fun set fun_thr=''"
    cn.Execute "update fun set fun_for=''"
    cn.Execute "update fun set fun_fiv=''"
    cn.Execute "update fun set fun_six=''"
    cn.Execute "update fun set fun_sev=''"
    cn.Execute "update fun set fun_eig=''"
    MsgBox "New Company Creation Successful!", vbExclamation
    company.Tag = "NEW"
Unload Me
Me.MousePointer = vbNormal
finish.Enabled = True
Exit Sub
lab:
If err.Number = 76 Or err.Number = 75 Then
    MsgBox "Directory (File Folder) Already Exists!", vbExclamation
    finish.Enabled = True
    Me.MousePointer = vbNormal
    Text7.SetFocus
Else
    MsgBox err.Description
End If
End Sub
Private Sub Form_Activate()
Text1.SetFocus
Text2 = Dir1.Path
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 58 Or K = 92 Then K = 0
If K > 96 And K < 123 Then
    K = K - 32
ElseIf K = 13 Then
    If Len(Trim(Text1)) = 0 Then
        MsgBox "Valid Company Name Entry Required! (10) Characters!", vbInformation, "DAS Version 4.0"
        Text1.SetFocus
    Else
        Text7.SetFocus
    End If
ElseIf K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
ElseIf K >= 65 And K <= 91 Then
Else
    K = 0
End If

'If K = 13 Then
'    If Len(Trim(Text1)) = 0 Then
'        MsgBox "Invalid Value", vbInformation
'        Text1.SetFocus
'    Else
'        Text7.SetFocus
'    End If
'End If
End Sub

Private Sub Text7_keypress(K As Integer)
If K = 58 Or K = 92 Then K = 0
If K > 96 And K < 123 Then
    K = K - 32
ElseIf K = 13 Then
    If Len(Trim(Text7)) = 0 Then
        MsgBox "Invalid Value", vbInformation
        Text7.SetFocus
    Else
        finish.SetFocus
    End If
ElseIf K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
ElseIf K >= 65 And K <= 91 Then
Else
    K = 0
End If
End Sub
Public Sub createcompany()
If company.SysInfo1.OSPlatform = 1 Then
    If Dir("c:\windows\system\EFFMCG4.txt", vbNormal) <> "" Then
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1
        Do While Not EOF(1)
        Line Input #1, textline  ' Read line into variable.
    Loop
      If Dir("c:\" & Text1, vbDirectory) = "" Then
            If MsgBox("Create New Company ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                MkDir "c:\" & Text1
                ChDir "c:\" & Text1
                Shell "c:\" & Text1 & "\lha e c:\" & Text1 & "\emg.lzh", vbHide
                MsgBox "New Company Created, Exit and Restart DAS Version 4.0 Software", vbExclamation, "DAS Version 4.0"
                Unload Me
                Exit Sub
            Else
                Text1 = ""
                Text1.SetFocus
                Exit Sub
            End If
        Else
            strattributes = "DBQ=c:\" & Text1 & "\Eff.dbc" & vbCr & "Description=Effmcg Driver" & vbCr & "SourceType=DBC"
            DBEngine.RegisterDatabase "EFF", "Microsoft Visual Foxpro Driver", True, strattributes
        End If
End If
ElseIf company.SysInfo1.OSPlatform = 2 Then
If Dir("c:\winnt\system\EFFMCG4.txt", vbNormal) <> "" Then
    Open "c:\winnt\system\EFFMCG4.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, textline  ' Read line into variable.
    Loop
      If Dir("c:\" & Text1, vbDirectory) = "" Then
            If MsgBox("Create New Company", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
                MkDir "c:\" & Text1
                ChDir "c:\" & Text1
                Shell "c:\" & Text1 & "\lha e c:\" & Text1 & "\emg.lzh", vbHide
                MsgBox "New Company Created, Exit & Restart DAS Version 4.0 Software", vbExclamation, "DAS Version 4.0"
                Unload Me
                Exit Sub
            Else
                Text1 = ""
                Text1.SetFocus
                Exit Sub
            End If
        Else
            strattributes = "DBQ=c:\" & Text1 & "\Eff.dbc" & vbCr & "Description=Effmcg Driver" & vbCr & "SourceType=DBC"
            DBEngine.RegisterDatabase "EFF", "Microsoft Visual Foxpro Driver", True, strattributes
        End If
End If
Else
    If Dir("c:\windows\system\EFFMCG4.txt", vbNormal) <> "" Then
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1
        Do While Not EOF(1)
        Line Input #1, textline  ' Read line into variable.
    Loop
      If Dir("c:\" & Text1, vbDirectory) = "" Then
            If MsgBox("Create New Company", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
                MkDir "c:\" & Text1
                ChDir "c:\" & Text1
                Shell "c:\" & Text1 & "\lha e c:\" & Text1 & "\emg.lzh", vbHide
                MsgBox "New Company Created, Exit & Restart DAS Version 4.0 Software", vbExclamation, "DAS Version 4.0"
                Unload Me
                Exit Sub
            Else
                Text1 = ""
                Text1.SetFocus
                Exit Sub
            End If
        Else
            strattributes = "DBQ=c:\" & Text1 & "\Eff.dbc" & vbCr & "Description=Effmcg Driver" & vbCr & "SourceType=DBC"
            DBEngine.RegisterDatabase "EFF", "Microsoft Visual Foxpro Driver", True, strattributes
        End If
End If
End If
frmLogin.Caption = Text1
Unload company
frmLogin.Show 1
End Sub
