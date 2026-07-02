VERSION 5.00
Begin VB.Form M_FRM_ABT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About DAS - Distributor Automation System Version 4.0"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_ABT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "DAS &Support"
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
      Left            =   5640
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S&ys Info"
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
      Left            =   5640
      TabIndex        =   19
      Top             =   3900
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
      Left            =   5640
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6795
      TabIndex        =   16
      Top             =   4680
      Width           =   6855
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   $"M_FRM_ABT.frx":0442
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   17
         Top             =   -120
         Width           =   6855
      End
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1205
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   11
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   10
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.Image Image3 
      Height          =   4680
      Left            =   0
      Picture         =   "M_FRM_ABT.frx":04F7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   240
      Picture         =   "M_FRM_ABT.frx":45E8D
      Top             =   -3480
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   1320
      Picture         =   "M_FRM_ABT.frx":4C54F
      Stretch         =   -1  'True
      Top             =   160
      Width           =   540
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.consoftsystems.com"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved"
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
      Left            =   1200
      TabIndex        =   15
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year   1999 - 2007"
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
      Left            =   1200
      TabIndex        =   14
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 4.0"
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
      Left            =   1200
      TabIndex        =   13
      Top             =   3600
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is Licensed to: "
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax: +91 40 3242 4031.   Email: effmcg@consoftsystems.com"
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
      Left            =   1350
      TabIndex        =   5
      Top             =   840
      Width           =   4395
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile: +91 98850 02290"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hyderabad - 500 007. Andhra Pradesh. INDIA."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright - Consoft Systems Pvt. Ltd."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   45
   End
End
Attribute VB_Name = "M_FRM_ABT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
If Dir(App.Path & "\MSINFO32.EXE") = "" Then
    MsgBox "Required Control File Missing!", vbExclamation
Else
    Shell App.Path & "\MSINFO32.EXE", vbMaximizedFocus
End If
End Sub

Private Sub Command4_Click()
If Dir("C:\Program Files\Effmcg4\HELPDESK.TXT", vbNormal) <> "" = True Then
    Shell "C:\PROGRAM FILES\INTERNET EXPLORER\iexplore.exe C:\Program Files\Effmcg4\HELPDESK.TXT", vbMaximizedFocus
Else
    MsgBox "Request for DAS HelpDesk Support Information, Send EMAIL to effmcg@consoftsystems.com or das@apac.ko.com", vbCritical, "DAS Version 4.0"
End If
End Sub

Private Sub Form_Activate()
Dim REC_SET As New ADODB.Recordset
REC_SET.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT1'", CON, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then Label8 = Trim(REC_SET("DIR_MSG"))
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT2'", CON, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then Label9 = Trim(REC_SET("DIR_MSG"))
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT3'", CON, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then Label10 = Trim(REC_SET("DIR_MSG"))
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT4'", CON, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then Label11 = Trim(REC_SET("DIR_MSG"))
End Sub

Private Sub Form_Load()
checkdongle
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label6.ForeColor = vbBlack
End Sub

Private Sub Label6_Click()
Shell "C:\PROGRAM FILES\INTERNET EXPLORER\iexplore.exe www.consoftsystems.com", vbMaximizedFocus
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label6.ForeColor = vbBlue
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label6.ForeColor = vbBlack
End Sub

