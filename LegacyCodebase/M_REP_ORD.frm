VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_ORD 
   Caption         =   "Select Asset"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "M_REP_ORD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Select Asset"
      TabPicture(0)   =   "M_REP_ORD.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_ORD.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "ListView1"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   3360
         TabIndex        =   4
         Top             =   720
         Width           =   1215
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   870
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   2
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2143
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order ID"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order Date"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Asset Code"
         Height          =   195
         Left            =   315
         TabIndex        =   8
         Top             =   1395
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Select Asset"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   945
         Width           =   885
      End
   End
End
Attribute VB_Name = "M_REP_ORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check1.Value = 0 Then
        Check1.Enabled = False
        Text1.SetFocus
    Else
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
If Check1.Value = 0 Then
    'Openreport CURDIR & "\lists.rp1", "Group Accounts list", 2, "GRP->GRP_CLF='" & Trim(Text1) & "'", 6
    Openreport CURDIR & "\Report.rp1", "Supplier Order", 2, "ORD->ORD_IDY='" & Trim(Text1) & "'", 6
Else
    'Openreport CURDIR & "\lists.rp1", "Group Accounts list", 0, "" & Trim(Text1) & "'", 6
    Openreport CURDIR & "\Report.rp1", "Supplier Order", 0, "'" & Trim(Text1) & "'", 6
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Check1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, S As Integer)
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
    If ListView1.ListItems.Count > 0 Then
        Text1 = ListView1.SelectedItem
        SSTab1.Tab = 0
        Check1.Value = False
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii = 13 Then
    Dim rs As New ADODB.Recordset
    rs.Open "select DISTINCT ORD_IDY, ORD_DAT from ord where ord_idy like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adLockOptimistic
    ListView1.ListItems.Clear
    For A = 1 To rs.RecordCount
         Set ls = ListView1.ListItems.Add(, , Trim(rs("ORD_IDY")))
         ls.ListSubItems.Add , , Trim(rs("ORD_DAT") & "")
         rs.MoveNext
    Next
    SSTab1.Tab = 1
    ListView1.SetFocus
End If
End Sub

