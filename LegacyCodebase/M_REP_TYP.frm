VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_TYP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Report Filters"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "M_REP_TYP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Select Group Type"
      TabPicture(0)   =   "M_REP_TYP.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_TYP.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   3015
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1080
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Group Type"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   488
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   3360
         TabIndex        =   4
         Top             =   960
         Width           =   1215
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   350
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   220
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   350
            Left            =   120
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   650
            Width           =   975
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&All"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   -74760
         TabIndex        =   1
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Type"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   4215
      End
   End
End
Attribute VB_Name = "M_REP_TYP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Command1_Click
ElseIf Check1.Value = 0 Then
    Check1.Enabled = False
    Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
If Check1.Value = 0 Then
    OpenReport CURDIR & "\lists.RP1", "Group Accounts list", 2, "GRP->GRP_CLF='" & Trim(Text1) & "'", 6
Else
    OpenReport CURDIR & "\lists.RP1", "Group Accounts list", 0, "" & Trim(Text1) & "'", 6
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Check1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "Listview1" Then
        SSTab1.Tab = 0
        Text1.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
End Sub

Private Sub ListView1_DblClick()
ListView1_KeyPress 13
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ListView1.ListItems.count > 0 Then
        Text1 = ListView1.SelectedItem
        SSTab1.Tab = 0
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii = 13 Then
    Dim rs As New ADODB.Recordset
    rs.Open "select * from typ where typ_def like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
    ListView1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        ListView1.ListItems.Add , , Trim(rs("typ_def"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    ListView1.SetFocus
End If
End Sub
