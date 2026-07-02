VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BAC_UP_FRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DAS Data Backup"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BAC_UP_FRM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2708
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1508
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4440
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2670
      Width           =   4695
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   4695
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Min             =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copying Files to Destination Folder Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Drive and Folder Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1335
      TabIndex        =   1
      Top             =   240
      Width           =   2475
   End
End
Attribute VB_Name = "BAC_UP_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Dim DIR_PAT As String

Private Sub Command1_Click()
Command1.Enabled = False
Text1_KeyPress 13
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_KeyPress(K As Integer)
If K = 13 Then Dir1_Change
End Sub

Private Sub Form_Activate()
Text1.Enabled = True
Drive1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If ProgressBar1.Visible = True Then
    Else
    Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Drive1_Change()
On Error GoTo err
    Dir1.Path = Drive1.Drive   ' Set directory path.
    Exit Sub
err:
    MsgBox "Please Insert Disk !", vbExclamation, "DAS Version 4.0"
    Drive1.SetFocus
End Sub
Private Sub Dir1_Change()
Text1 = Dir1.Path   ' Set file path.
Text1.SelStart = Len(Text1)
Text1.SetFocus
End Sub
Private Sub Text1_KeyPress(K As Integer)
On Error GoTo er1
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If Len(Trim(Text1)) = 0 Then
        MsgBox "Enter Destination Directory !", vbExclamation, "DAS Version 4.0"
        Command1.Enabled = True
        Text1.Locked = False
        Text1.SetFocus
        Exit Sub
    ElseIf Mid(Trim(Text1), 2, 2) <> ":\" Then
        MsgBox "Invalid Destination, Verify !", vbExclamation, "DAS Version 4.0"
        SendKeys "{home}+{end}"
        Command1.Enabled = True
        Text1.Locked = False
        Text1.SetFocus
        Exit Sub
    ElseIf Len(Trim(Text1)) <> 3 And InStr(Text1, "/") = Len(Text1) Then
        Text1 = Mid(Trim(Text1), 1, Len(Text1) - 1)
    End If
    Text1.Locked = True
    If Dir(Trim(Text1), vbDirectory) = "" Then
        If MsgBox("Destination Directory Not Found, Create New ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
            MkDir Trim(Text1)
        Else
            Text1.SetFocus
            Text1.Locked = False
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    End If
   
    Timer1.Interval = 100
    Timer1.Enabled = True
End If
Exit Sub
er1:
    If err.Number = 76 Or err.Number = 52 Then
        MsgBox "Drive Does Not Exist, Verify !", vbExclamation, "DAS Version 4.0"
        Command1.Enabled = True
    Else
        MsgBox err.Description
        Command1.Enabled = True
    End If
        SendKeys "{home}+{end}"
        Text1.Locked = False
        Text1.SetFocus
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
BAC_UP_FRM.MousePointer = vbHourglass
Destdir = Trim(Text1)
Label2.Refresh
If Len(Trim(Text1)) = 3 And Right(Trim(Text1), 2) = ":\" Then
    Destdir = Trim(Text1) & "BackUp" & Date$
    If Dir(Destdir, vbDirectory) = "" Then
        MkDir Destdir
    Else
        Kill Destdir & "\*.*"
        RmDir Destdir
        MkDir Destdir
    End If
End If
If Right(Trim(Destdir), 2) = ":\" Or Right(Trim(Destdir), 1) = ":" Or Right(Trim(Destdir), 1) = "*" Or Right(Trim(Destdir), 1) = "<" Or Right(Trim(Destdir), 1) = ">" Or Right(Trim(Destdir), 1) = "|" Then
    MsgBox "Select a valid Data Folder!", vbInformation + vbOKOnly, "DAS Version 4.0"
    Text1.SetFocus
    SendKeys "{home}+{end}"
    Timer1.Interval = 0
    Me.MousePointer = vbNormal
    Command1.Enabled = True
    Text1.Locked = False
    Exit Sub
End If

Label2.Visible = True ' = ProgressBar1.Value & "%"
Label2.Refresh

Me.MousePointer = vbHourglass
Dim FS As Object, f As Object, F1 As Object, FN As Object
Set FS = CreateObject("Scripting.FileSystemObject")

SPATH = CURDIR & "\???.DBF"
DPATH = Destdir
FS.CopyFile SPATH, DPATH

SPATH = CURDIR & "\EFF.*"
DPATH = Destdir
FS.CopyFile SPATH, DPATH

SPATH = CURDIR & "\*.CDX"
DPATH = Destdir
FS.CopyFile SPATH, DPATH

SPATH = CURDIR & "\*.FPT"
DPATH = Destdir
FS.CopyFile SPATH, DPATH

SPATH = App.Path & "\*.RP1"
DPATH = Destdir
FS.CopyFile SPATH, DPATH

SPATH = App.Path & "\*.RRW"
DPATH = Destdir
FS.CopyFile SPATH, DPATH

CopyFile App.Path & "\IND.EXE", Destdir & "\IND.EXE", 0
CopyFile App.Path & "\CND.EXE", Destdir & "\CND.EXE", 0
CopyFile App.Path & "\RRUNTIME.EXE", Destdir & "\RRUNTIME.EXE", 0
CopyFile App.Path & "\RRUN.EXE", Destdir & "\RRUN.EXE", 0
CopyFile App.Path & "\RR.CNF", Destdir & "\RR.CNF", 0
CopyFile App.Path & "\PRI.EXE", Destdir & "\PRI.EXE", 0
CopyFile App.Path & "\PRB.EXE", Destdir & "\PRB.EXE", 0
BAC_UP_FRM.MousePointer = vbDefault
Timer2.Enabled = False
Label2 = ""
Text1.Locked = False
MsgBox "Data Backup to : " & Destdir & " Successful!", vbInformation, "DAS Version 4.0"
Me.MousePointer = vbDefault

' Backup Date and Path Record - added by PNR
    Dim RST_DIR As New ADODB.Recordset
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'BKP' .AND. DIR_IDX LIKE 'LOG'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        RST_DIR.MoveFirst
        RST_DIR("DIR_MSG") = Trim(Text1)
        RST_DIR("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        RST_DIR.Update
    Else
        RST_DIR.AddNew
        RST_DIR("DIR_IDY") = "BKP"
        RST_DIR("DIR_IDX") = "LOG"
        RST_DIR("DIR_ALS") = ""
        RST_DIR("DIR_MSG") = Trim(Text1)
        RST_DIR("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        RST_DIR.Update
    End If
' Backup Date and Path Record - added by PNR
ChDir CURDIR
Unload Me
Set Form = Nothing
End Sub

Public Sub INC()
On Error Resume Next
Label2 = Int(ProgressBar1.Value)
Label2.Refresh
ProgressBar1.Value = ProgressBar1.Value + 2.5
End Sub
