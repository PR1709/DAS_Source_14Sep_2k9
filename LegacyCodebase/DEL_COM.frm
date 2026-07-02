VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Begin VB.Form DEL_COM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Select Company"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "DEL_COM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&New"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1455
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   540
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Company"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   600
      Width           =   1155
   End
End
Attribute VB_Name = "DEL_COM"
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
Command1.Enabled = False
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
            If CON.State = 1 Then CON.Close
            CON.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
            SEL_CMP = Trim(Combo1)
            Unload Me
            Unload M_FRM_STR
            Exit Sub
        End If
    Loop
    MsgBox "Company Does Not Exist!", vbInformation, "DAS Version 4.0"
    Combo1.SetFocus
Else
If Len(Trim(Combo1)) = 0 Then
    MsgBox "Select Company Name!", vbInformation, "DAS Version 4.0"
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
            CON.Close
            CON.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
            SEL_CMP = Trim(Combo1)
            Unload Me
            Unload M_FRM_STR
            Dim DIR_RS As New ADODB.Recordset
            If DIR_RS.State = 1 Then DIR_RS.Close
            DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'SAY'"), CON, adOpenKeyset, adLockPessimistic
            If DIR_RS.RecordCount > 0 Then
                DIR_RS.MoveFirst
                If Trim(DIR_RS("DIR_MSG")) = "T" Then
                    M_LOG_INF.Show 1
                End If
            End If
            EFF_LOG.Show 1
            Exit Sub
        End If
    Loop
    MsgBox "Company Does Not Exist! ", vbInformation, "DAS Version 4.0"
    Combo1.SetFocus
End If
End If
Close #1
Exit Sub
LBL:
MsgBox "Unable to Create Connection with the Database...", vbCritical, "DAS Version 4.0 - Critical Error"
End Sub

Private Sub Command2_Click()
Unload Me
Unload M_FRM_STR
End Sub

Private Sub Command3_Click()
setup.Show 1
End Sub

'''''Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
'''''If KeyCode = 13 Then
'''''If Len(Trim(Text1)) = 0 Then
'''''    MsgBox "Enter Company Name", vbExclamation, "DAS Version 4.0"
'''''    Text1.SetFocus
'''''Else
'''''    Command1.SetFocus
'''''End If
'''''ElseIf KeyCode = 27 Then
'''''    End
'''''End If
'''''End Sub
'''''
'''''Private Sub Text1_KeyPress(KeyAscii As Integer)
'''''If KeyAscii = 92 Or KeyAscii = 47 Then KeyAscii = 0
'''''End Sub

Private Sub Command4_Click()
If Len(Trim(Combo1)) = 0 Then
    MsgBox "Select Company Name", vbInformation, "DAS Version 4.0"
    Combo1.SetFocus
Else
    If MsgBox("Remove Company Information ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
        If Trim(Combo1) = "EXISTING COMPANY" Then
        MsgBox "DAS Version 4.0 Default Company ! Cannot Remove!", vbCritical, "DAS Version 4.0"
        Exit Sub
        End If
        If SysInfo1.OSPlatform = 1 Then
            Name "c:\windows\system\EFFMCG4.txt" As "c:\windows\system\EFFMCG41.txt"
            Open "c:\windows\system\EFFMCG41.txt" For Input As #1   'Open file for output.
            Open "c:\windows\system\EFFMCG4.txt" For Output As #2   'Open file for output.
        ElseIf SysInfo1.OSPlatform = 2 Then
            Name "c:\winnt\system\EFFMCG4.txt" As "c:\winnt\system\EFFMCG41.txt"
            Open "c:\winnt\system\EFFMCG41.txt" For Input As #1   'Open file for output.
            Open "c:\winnt\system\EFFMCG4.txt" For Output As #2   'Open file for output.
        Else
            Name "c:\windows\system\EFFMCG4.txt" As "c:\windows\system\EFFMCG41.txt"
            Open "c:\windows\system\EFFMCG41.txt" For Input As #1   'Open file for output.
            Open "c:\windows\system\EFFMCG4.txt" For Output As #2   'Open file for output.
        End If
        Do While Not EOF(1)
            Line Input #1, textline  ' Read line into variable.
            A = Split(textline, "@", -1, vbTextCompare)
            If UCase(Mid(Trim(A(0)), 2, Len(A(0)) - 1)) = Trim(Combo1) Then
            Else
                Write #2, Trim(Mid(Trim(textline), 2, Len(Trim(textline)) - 2))
                A = Split(textline, "@", -1, vbTextCompare)
                Combo1.AddItem UCase(Mid(Trim(A(0)), 2, Len(A(0)) - 1))
            End If
        Loop
        Close #2
        Close #1
        If SysInfo1.OSPlatform = 1 Then
            Kill "c:\windows\system\EFFMCG41.txt"
            Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
        ElseIf SysInfo1.OSPlatform = 2 Then
            Kill "c:\winnt\system\EFFMCG41.txt"
            Open "c:\winnt\system\EFFMCG4.txt" For Input As #1   'Open file for output.
        Else
            Kill "c:\windows\system\EFFMCG41.txt"
            Open "c:\windows\system\EFFMCG4.txt" For Input As #1   'Open file for output.
        End If
        Combo1.CLEAR
        Do While Not EOF(1)
            Line Input #1, textline  ' Read line into variable.
            A = Split(textline, "@", -1, vbTextCompare)
            Combo1.AddItem UCase(Mid(Trim(A(0)), 2, Len(A(0)) - 1))
        Loop
    End If
End If
Close #1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
Unload M_FRM_STR
End If
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "Already Opened"
    Unload Me
    Exit Sub
End If
LOGIN
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
