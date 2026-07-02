VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_RST_BKP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore DB From AutoBackup"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   ClipControls    =   0   'False
   Icon            =   "M_FRM_RST_BKP.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirm"
      Height          =   375
      Left            =   1665
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   345
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Restore DB in Progress"
      Height          =   195
      Left            =   1380
      TabIndex        =   1
      Top             =   240
      Width           =   1650
   End
End
Attribute VB_Name = "M_FRM_RST_BKP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COM_CTR As Integer
Dim CTR As Integer
Private Sub Command1_Click()
Command1.Enabled = False
Dim SPATH As String
Dim DPATH As String
CTR = 8
' *
    Me.AutoRedraw = True
    Me.Refresh
    ProgressBar1.Visible = True
    Me.MousePointer = vbHourglass
    AUTOBKP = CURDIR & "\AUTOBKP"
    If Dir(AUTOBKP, vbDirectory) = "" Then
        MsgBox "AutoBackup Folder Information Missing !", vbCritical, "DAS Version 4.0"
        Exit Sub
    End If
    AUTORST = CURDIR & "\AUTORST"
    If Dir(AUTORST, vbDirectory) = "" Then
        MkDir AUTORST
    End If
    INC_PCB
    Dim FS As Object, f As Object, F1 As Object, FN As Object
    Set FS = CreateObject("Scripting.FileSystemObject")
    SPATH = AUTOBKP & "\???.DBF"
    CPATH = CURDIR & "\???.DBF"
    FS.CopyFile CPATH, AUTORST
    FS.CopyFile SPATH, CURDIR
    INC_PCB
    SPATH = AUTOBKP & "\EFF.*"
    CPATH = CURDIR & "\EFF.*"
    FS.CopyFile CPATH, AUTORST
    FS.CopyFile SPATH, CURDIR
    INC_PCB
    SPATH = AUTOBKP & "\*.CDX"
    CPATH = CURDIR & "\*.CDX"
    FS.CopyFile CPATH, AUTORST
    FS.CopyFile SPATH, CURDIR
    INC_PCB
    SPATH = AUTOBKP & "\*.FPT"
    CPATH = CURDIR & "\*.FPT"
    FS.CopyFile CPATH, AUTORST
    FS.CopyFile SPATH, CURDIR
    INC_PCB
    SPATH = App.Path & "\*.RP1"
    CPATH = CURDIR & "\*.RP1"
    FS.CopyFile CPATH, AUTORST
    FS.CopyFile SPATH, CURDIR
    INC_PCB
    SPATH = App.Path & "\*.RRW"
    CPATH = CURDIR & "\*.RRW"
    FS.CopyFile CPATH, AUTORST
    FS.CopyFile SPATH, CURDIR
    INC_PCB
    CopyFile App.Path & "\IND.EXE", CURDIR & "\IND.EXE", 0
    CopyFile App.Path & "\CND.EXE", CURDIR & "\CND.EXE", 0
    CopyFile App.Path & "\RRUNTIME.EXE", CURDIR & "\RRUNTIME.EXE", 0
    CopyFile App.Path & "\RRUN.EXE", CURDIR & "\RRUN.EXE", 0
    CopyFile App.Path & "\RR.CNF", CURDIR & "\RR.CNF", 0
    CopyFile App.Path & "\PRI.EXE", CURDIR & "\PRI.EXE", 0
    CopyFile App.Path & "\PRB.EXE", CURDIR & "\PRB.EXE", 0
    INC_PCB
    Me.MousePointer = vbDefault
    MsgBox ("Data Restore Complete from AutoBackup Folder !"), vbInformation, "DAS Version 4.0"
    Unload Me
    Set Form = Nothing
End Sub
Public Sub INC_PCB()
COM_CTR = COM_CTR + 1
ProgressBar1.Value = (100 / CTR) * COM_CTR
End Sub
