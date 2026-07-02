VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form CMP_SEL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Version 4.0 Select Company for Additional Utilities"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ClipControls    =   0   'False
   Icon            =   "CMP_SEL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Company"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   1155
   End
End
Attribute VB_Name = "CMP_SEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub cmdOk_Click()
    a = Split(Combo1.Text, "@", -1, vbTextCompare)
    EFF_LIB1.CompanyName = Trim(a(0))
    EFF_LIB1.CompanyFolderPath = Trim(a(1))
    Unload M_FRM_STR
    Unload Me
    M_MDI.Show
End Sub

Private Sub Combo1_KEYPRESS(K As Integer)
If K = 13 Then
    cmdOk_Click
End If
End Sub

Private Sub Command2_Click()
If Me.Tag = "FIRST" Then
    End
Else
    Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And Me.Tag = "Change" Then
    End
ElseIf KeyAscii = 27 And Me.Tag = "" Then
    End
End If
End Sub

Private Sub Form_Load()
'checkdongle

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
Combo1.CLEAR
Do While Not EOF(1)
    Line Input #1, textline  ' Read line into variable.
    a = Split(textline, "@", -1, vbTextCompare)
    Combo1.AddItem UCase(Mid(Trim(a(0)), 2, Len(a(0)) - 1)) & "                                                                                                                                                                        @ " & Replace(a(2), """", "")
Loop
Close #1
End Sub

Private Sub Form_Unload(cancel As Integer)
'LOGOUT
End Sub

