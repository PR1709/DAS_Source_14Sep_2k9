VERSION 5.00
Begin VB.Form M_FRM_DRP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Text File to Print (Character Mode)"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "M_FRM_DRP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H8000000F&
      Height          =   2040
      Left            =   2760
      Pattern         =   "*.txt"
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2880
      Width           =   855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H8000000F&
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2880
      Width           =   855
   End
End
Attribute VB_Name = "M_FRM_DRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CopyFile File1.Path & File1.FileName, "PRN", A
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Dir1_Change()
File1.FileName = "*.TXT"
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
On Error GoTo err
    Dir1.Path = Drive1.Drive   ' Set directory path.
    Exit Sub
err:
    MsgBox "Please Insert Disk !", vbExclamation, "DAS Version 4.0"
    Drive1.SetFocus
End Sub
Private Sub File1_Click()
Text1 = File1.Path & "\" & File1.FileName
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

