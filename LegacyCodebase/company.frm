VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form company 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Select Company"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ClipControls    =   0   'False
   Icon            =   "company.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1688
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   540
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&DAS Support"
      Height          =   375
      Left            =   4628
      TabIndex        =   4
      ToolTipText     =   "Distributor Automation System Help and Support Information"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   308
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1628
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&New"
      Height          =   375
      Left            =   3188
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Company"
      Height          =   195
      Left            =   368
      TabIndex        =   5
      Top             =   600
      Width           =   1155
   End
End
Attribute VB_Name = "company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Sub Combo1_keypress(K As Integer)
If K = 13 Then
    Command1_Click
End If
End Sub

Private Sub Command1_Click()
On Error GoTo LBL
createcompany "ASDF"
If Me.Tag = "Change" Then
    If SysInfo1.OSPlatform = 1 Then
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    ElseIf SysInfo1.OSPlatform = 2 Then
        Open "c:\winnt\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    Else
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    End If
    Do While Not EOF(1)
        Line Input #1, textline  ' Read line into variable.
        A = Split(textline, "@", -1, vbTextCompare)
        If UCase(Mid(Trim(A(0)), 2, Len(A(0)) - 1)) = Trim(Combo1) Then
            Close #1
            Text = UCase(Mid(Trim(A(2)), 1, Len(Trim(A(2))) - 1))
            CURDIR = Text
            If PNL.State = 1 Then PNL.Close
            PNL.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
            ChDir Text
            con.Close
            con.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
            SEL_CMP = Trim(Combo1)
            Unload Company
            Unload M_FRM_STR
            Exit Sub
        End If
    Loop
    MsgBox "Invalid Company Selection !", vbInformation, "DAS Version 4.0"
    Combo1.SetFocus
Else
If Len(Trim(Combo1)) = 0 Then
    MsgBox "Select Company Name from Drop Down Menu, Press Down Arrow Key!", vbInformation, "DAS Version 4.0"
    Combo1.SetFocus
Else
    If SysInfo1.OSPlatform = 1 Then
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    ElseIf SysInfo1.OSPlatform = 2 Then
        Open "c:\winnt\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    Else
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    End If
    Do While Not EOF(1)
        Line Input #1, textline  ' Read line into variable.
        A = Split(textline, "@", -1, vbTextCompare)
        If UCase(Mid(Trim(A(0)), 2, Len(A(0)) - 1)) = Trim(Combo1) Then
            Close #1
            Text = UCase(Mid(Trim(A(2)), 1, Len(Trim(A(2))) - 1))
            CURDIR = Text
            If PNL.State = 1 Then PNL.Close
            PNL.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
            ChDir Text
            If con.State = 1 Then con.Close
            con.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
            SEL_CMP = Trim(Combo1)
            Unload Company
            Unload M_FRM_STR
            Load M_MDI
            M_MDI.Show
            ' Backup Checkup and Message Module - added by PNR
            Dim DIR_RS As New ADODB.Recordset
            If DIR_RS.State = 1 Then DIR_RS.Close
            DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'SAY'"), con, adOpenKeyset, adLockPessimistic
            If DIR_RS.RecordCount > 0 Then
                DIR_RS.MoveFirst
                If Trim(DIR_RS("DIR_MSG")) = "T" Then
                    M_LOG_INF.Show 1
                End If
            End If
            EFF_LOG.Show 1
                
                ' * Routine to Create Login Information into ShutDown Batch File 'PNR
                Dim DEX As String
                DEX = "C:\DEX"
                If Dir(DEX, vbDirectory) = "" Then
                    MkDir DEX
                End If
                Dim FS_SOU, F_VER As Object
                Set FS_SOU = CreateObject("Scripting.filesystemobject")
                
                If Dir(App.Path & "\SHUTDOWN.LOG") = "" Then
                    Set F_VER = FS_SOU.OpenTextFile(App.Path & "\SHUTDOWN.LOG", 2, True)
                    F_VER.WriteLine "REM *** DAS Normal Login ***"
                    F_VER.WriteLine "Date " & Date$ & " : Time " & Time$
                    F_VER.Close
                Else
                    Set F_VER = FS_SOU.OpenTextFile(App.Path & "\SHUTDOWN.LOG", 8, True)
                    F_VER.WriteLine "REM *** DAS Normal Login ***"
                    F_VER.WriteLine "Date " & Date$ & " : Time " & Time$
                    F_VER.Close
                End If
                CopyFile App.Path & "\SHUTDOWN.LOG", CURDIR & "\SHUTDOWN.LOG", 0
            
                ' * End of Login Information into ShutDown Batch File 'PNR
            
            Exit Sub
        End If
    Loop
    MsgBox "Invalid Company Selection !", vbCritical, "DAS Version 4.0"
    Combo1.SetFocus
End If
End If
Close #1
Exit Sub
LBL:
MsgBox err.Description & "  " & err.Number & "  Unable to Create Connection with the Database!", vbCritical, "DAS Version 4.0"
End Sub
Private Sub Command2_Click()
If Me.Tag = "NEW" Then
End
End If
Unload Me
Unload M_FRM_STR
End Sub
Private Sub Command3_Click()
If Dir(App.Path & "\BKP\EFF.DBC") = "" Then
    MsgBox "Operation Not Valid Now! Login With Existing Company Required to Create New Company", vbCritical, "DAS Version 4.0"
    Exit Sub
Else
    setup.Show 1
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 92 Or KeyAscii = 47 Then KeyAscii = 0
End Sub
Private Sub Command4_Click()
If Dir("C:\Program Files\Effmcg4\HELPDESK.TXT", vbNormal) <> "" = True Then
    Shell "C:\PROGRAM FILES\INTERNET EXPLORER\iexplore.exe C:\Program Files\Effmcg4\HELPDESK.TXT", vbMaximizedFocus
Else
    MsgBox "Request for DAS HelpDesk Support Information, Send EMAIL to effmcg@consoftsystems.com or das@apac.ko.com", vbCritical, "DAS Version 4.0"
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And Me.Tag = "Change" Then
    Unload Me
    Unload M_FRM_STR
ElseIf KeyAscii = 27 And Me.Tag = "" Then
    End
End If
End Sub
Private Sub Form_Load()
'Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=C:\coke\COKE\Eff.dbc;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;
If App.PrevInstance = True Then
    MsgBox "Already Opened"
    End
    Exit Sub
End If
LOGIN
osversion = SysInfo1.osversion
If Dir(App.Path & "\eff.dbc") = "" Then
    MsgBox "DAS Version 4.0 installation is incomplete,Run DISK144\DISK1\SETUP.EXE for complete installation", vbCritical, "DAS Version 4.0"
    End
Else
    If SysInfo1.OSPlatform = 1 Then
        If Dir("c:\windows\system\EFFMCG4.txt") = "" Then
            Open "c:\windows\system\EFFMCG4.txt" For Output As #1   'Open file for output.
            Write #1, "Existing Company" & "@" & "Path" & "@" & App.Path
            Close #1
        End If
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    ElseIf SysInfo1.OSPlatform = 2 Then
        M_XP_DIR = "C:\WINNT"
        M_XP_SYS = "\SYSTEM"
        If Dir("c:\winnt\system", vbDirectory) = "" Then
            MkDir M_XP_DIR
            MkDir M_XP_DIR & M_XP_SYS
        End If
        If Dir("c:\winnt\system\EFFMCG4.txt") = "" Then
        Open "c:\winnt\system\EFFMCG4.txt" For Output As #1   'Open file for output.
        Write #1, "Existing Company" & "@" & "Path" & "@" & App.Path
        Close #1
        End If
        Open "c:\winnt\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    Else
        If Dir("c:\windows\system\EFFMCG4.txt") = "" Then
        Open "c:\windows\system\EFFMCG4.txt" For Output As #1   'Open file for output.
        Write #1, "Existing Company" & "@" & "Path" & "@" & App.Path
        Close #1
        End If
        Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
    End If
End If
Combo1.CLEAR
Do While Not EOF(1)
    Line Input #1, textline  ' Read line into variable.
    A = Split(textline, "@", -1, vbTextCompare)
    Combo1.AddItem UCase(Mid(Trim(A(0)), 2, Len(A(0)) - 1))
Loop
Close #1
End Sub
