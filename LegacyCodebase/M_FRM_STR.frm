VERSION 5.00
Begin VB.Form M_FRM_STR 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS - Distributor Automation System - Version 4.0.2bn"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_STR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2200
      Left            =   0
      Top             =   3840
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   4320
      Width           =   6860
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   3405
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   60
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Version Release Date"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1575
         TabIndex        =   7
         Top             =   60
         Width           =   1785
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   0
         Picture         =   "M_FRM_STR.frx":08CA
         Stretch         =   -1  'True
         Top             =   440
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   6870
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usage of this software is subject to license notice contained within the user profile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   585
         TabIndex        =   4
         Top             =   720
         Width           =   5760
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 1999 - 2007. Consoft Systems Pvt Ltd, INDIA. All rights reserved."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   705
         TabIndex        =   1
         Top             =   480
         Width           =   5475
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Network Version"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   4920
      TabIndex        =   5
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Image Image4 
      Height          =   1785
      Left            =   2303
      Picture         =   "M_FRM_STR.frx":199D
      Top             =   1770
      Width           =   2235
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distributor Automation System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   750
      TabIndex        =   3
      Top             =   765
      Width           =   5340
   End
   Begin VB.Image Image3 
      Height          =   1785
      Left            =   -120
      Picture         =   "M_FRM_STR.frx":EA1F
      Top             =   -3480
      Width           =   2235
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:4.0.2bn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   4980
      TabIndex        =   2
      Top             =   1245
      Width           =   1110
   End
End
Attribute VB_Name = "M_FRM_STR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
    Dim lngLocale As Long
    lngLocale = GetSystemDefaultLCID()
    SetLocaleInfo lnglocal, LOCALE_SSHORTDATE, "MM/dd/yyyy"
    On Error Resume Next
    Text1.Visible = False
    If Dir(App.Path & "EFFMCG.EXE") = "" Then
        M_EXE_VER = Format(FileDateTime(App.Path & "\EFFMCG.EXE"), "dd-mmm-yyyy hh:mm:ss")
        Text1 = Trim(M_EXE_VER)
        Text1.Visible = True
    End If
End Sub
Private Sub Timer1_Timer()
    Company.Show 1
End Sub
