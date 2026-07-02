VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_RET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "M_REP_RET.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Report Filters"
      TabPicture(0)   =   "M_REP_RET.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Check7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_RET.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   5400
         TabIndex        =   25
         Top             =   3480
         Width           =   1095
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All"
         Height          =   375
         Left            =   960
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   4095
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   5055
         Begin VB.ComboBox list2 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   40
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1695
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   40
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1215
            Width           =   2295
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   40
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   735
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   40
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   40
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   2662
            Width           =   2295
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            MaxLength       =   40
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   3120
            Width           =   2295
         End
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "M_REP_RET.frx":047A
            Left            =   2040
            List            =   "M_REP_RET.frx":0487
            TabIndex        =   18
            Top             =   3637
            Width           =   495
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Sale Sla&b"
            Height          =   375
            Left            =   720
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CheckBox Check8 
            Caption         =   "D&iscount"
            Height          =   375
            Left            =   720
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   3098
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Location"
            Height          =   375
            Left            =   720
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   218
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Ca&tegory"
            Height          =   375
            Left            =   720
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   713
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "&Market"
            Height          =   375
            Left            =   720
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1193
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "&Salesman"
            Height          =   375
            Left            =   720
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1673
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Ro&ute"
            Height          =   375
            Left            =   720
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2130
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "&Retailer"
            Height          =   375
            Left            =   720
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   2640
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
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
            Text            =   "Market Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Market"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   6225
      End
   End
End
Attribute VB_Name = "M_REP_RET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then
If Check1.Value = 1 Then
    Check7.Value = 0
    Text1.Enabled = True
    Text1.SetFocus
Else
    Text1.Enabled = False
    Text1 = ""
    Check2.SetFocus
End If
End If
End Sub

Private Sub CHECK2_KeyPress(K As Integer)
If K = 13 Then
If Check2.Value = 1 Then
    Check7.Value = 0
    Text2.Enabled = True
    Text2.SetFocus
Else
    Text2 = ""
    Text2.Enabled = False
    Check3.SetFocus
End If
End If
End Sub

Private Sub CHECK3_KeyPress(K As Integer)
If K = 13 Then
If Check3.Value = 1 Then
    Check7.Value = 0
    Text3.Enabled = True
    Text3.SetFocus
Else
    Text3 = ""
    Text3.Enabled = False
    Check4.SetFocus
End If
End If
End Sub

Private Sub Check4_keypress(K As Integer)
If K = 13 Then
If Check4.Value = 1 Then
    Check7.Value = 0
    Text4.Enabled = True
    Text4.SetFocus
Else
    Text4.Enabled = False
    Text4 = ""
    Check5.SetFocus
End If
End If
End Sub

Private Sub Check5_keypress(K As Integer)
If K = 13 Then
If Check5.Value = 1 Then
    Check7.Value = 0
    List2.Enabled = True
    List2.SetFocus
Else
    List2.Enabled = False
    List6 = ""
    Check6.SetFocus
End If
End If
End Sub

Private Sub Check6_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check6.Value = 1 Then
        Check7.Value = 0
        Text6.Enabled = True
        Text6.SetFocus
    Else
        Text6.Enabled = True
        Text6 = ""
        Check8.SetFocus
    End If
End If
End Sub

Private Sub Check7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check7.Value = 1 Then
        Frame1.Enabled = False
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        If List2.ListCount > 0 Then List2.ListIndex = 0
        Command1.SetFocus
    Else
        Frame1.Enabled = True
        Check7.Enabled = False
        Check1.SetFocus
    End If
End If
End Sub

Private Sub Check9_KeyPress(K As Integer)
If K = 13 Then
If Check9.Value = 1 Then
    Check7.Value = 0
    list3.Enabled = True
    list3.SetFocus
Else
    list3.Enabled = False
    list3 = ""
    Command1.SetFocus
End If
End If
End Sub
Private Sub Check8_KeyPress(K As Integer)
If K = 13 Then
If Check8.Value = 1 Then
    Check7.Value = 0
    Text5.Enabled = True
    Text5.SetFocus
    Text5 = "0.00"
Else
    Text5.Enabled = False
    Text5 = ""
    Check9.SetFocus
End If
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'LOC' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text1) <> "" Then
        rs("DIR_ALS") = Trim(Label3) & ""
        rs("DIR_MSG") = Trim(Text1) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->LOC_IDY='" & Trim(Label3) & "'"
        Else
            SAL_QRY = "ACT->LOC_IDY='" & Trim(Label3) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'CAT' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text2) <> "" Then
        rs("DIR_ALS") = Trim(Label4) & ""
        rs("DIR_MSG") = Trim(Text2) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
        SAL_QRY = SAL_QRY & ".AND.ACT->CAT_IDY='" & Trim(Label4) & "'"
        Else
        SAL_QRY = "ACT->CAT_IDY='" & Trim(Label4) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
'
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'MAR' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
        If Trim(Text3) <> "" Then
        rs("DIR_ALS") = Trim(Label5) & ""
        rs("DIR_MSG") = Trim(Text3) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->MAR_IDY='" & Trim(Label5) & "'"
        Else
            SAL_QRY = "ACT->MAR_IDY='" & Trim(Label5) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'SMN' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text4) <> "" Then
        rs("DIR_ALS") = Trim(Label6) & ""
        rs("DIR_MSG") = Trim(Text4) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND. RDY->SMN_IDY='" & Trim(Label6) & "'"
        Else
            SAL_QRY = "RDY->SMN_IDY='" & Trim(Label6) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'ROU' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(List2.Text) <> "" Then
        rs("DIR_ALS") = Trim(List2) & ""
        rs("DIR_MSG") = Trim(List2) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.RDY->ROU_IDY='" & Trim(List2) & "'"
        Else
            SAL_QRY = "RDY->ROU_IDY='" & Trim(List2) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'DLR' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text6) <> "" Then
        rs("DIR_ALS") = Trim(Label7) & ""
        rs("DIR_MSG") = Trim(Text6) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->ACT_IDY='" & Trim(Label7) & "'"
        Else
            SAL_QRY = "ACT->ACT_IDY='" & Trim(Label7) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
rs.Update
If Check8.Value = 1 Then
    If Trim(Len(SAL_QRY)) <> 0 Then
        SAL_QRY = SAL_QRY & ".AND.ACT->dbt_dis=" & Val(Text5)
    Else
        SAL_QRY = "ACT->dbt_dis=" & Val(Text5)
    End If
End If

If Check9.Value = 1 Then
    If Trim(Len(SAL_QRY)) <> 0 Then
        SAL_QRY = SAL_QRY & ".AND.ACT->DBT_rap='" & Trim(list3) & "'"
    Else
        SAL_QRY = "ACT->DBT_rap='" & Trim(list3) & "'"
    End If
End If

If Trim(Len(SAL_QRY)) <> 0 Then
    SAL_QRY = SAL_QRY & ".AND.(ACT->ACT_TYP='SAL'.OR.ACT->ACT_TYP='BOT')"
Else
    SAL_QRY = "ACT->ACT_TYP='SAL'.OR.ACT->ACT_TYP='BOT'"
End If

If Me.Tag = "DIS" Then
    OpenReport CURDIR & "\LISTS.RP1", "Retailers List with Sales Slab", 2, SAL_QRY, 6
ElseIf Me.Tag = "LST" Then
    OpenReport CURDIR & "\LISTS.RP1", "Retailers List", 2, SAL_QRY, 6
ElseIf Me.Tag = "MLS" Then
    OpenReport CURDIR & "\LISTS.RP1", "Retailers Mailers List", 2, SAL_QRY, 6
End If
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check7.SetFocus
End Sub

Private Sub Form_Activate()
SETYEAROPEN M_REP_SQY
SSTab1.TabEnabled(1) = False
Check7.Value = 1
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check7.SetFocus
End Sub
Private Sub Form_Load()
SAL_QRY = ""
AddData List2, Me, "rou_idy", "rdy", "", "group by rou_idy"
End Sub
Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        If Label20.Caption = "Select Retailer" Then Text6.SetFocus
        If Label20.Caption = "Select Location" Then Text1.SetFocus
        If Label20.Caption = "Select Category" Then Text2.SetFocus
        If Label20.Caption = "Select Salesman" Then Text4.SetFocus
        If Label20.Caption = "Select Market" Then Text3.SetFocus
        SSTab1.Tab = 0
    Else
        Unload Me
    End If
End If
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If list1.ListItems.count > 0 Then
        If Label20 = "Select Location" Then
            Label3 = list1.SelectedItem.ListSubItems.Item(1)
            Text1 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check2.SetFocus
        ElseIf Label20 = "Select Category" Then
            Label4 = list1.SelectedItem.ListSubItems.Item(1)
            Text2 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check3.SetFocus
        ElseIf Label20 = "Select Market" Then
            Label5 = list1.SelectedItem.ListSubItems.Item(1)
            Text3 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check4.SetFocus
        ElseIf Label20 = "Select Salesman" Then
            Label6 = list1.SelectedItem.ListSubItems.Item(1)
            Text4 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check5.SetFocus
        ElseIf Label20 = "Select Retailer" Then
            Label7 = list1.SelectedItem.ListSubItems.Item(1)
            Text6 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check8.SetFocus
        End If
    End If
End If
End Sub

Private Sub List2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check6.SetFocus
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Text1_Change()
If Text1 = "" Then Label3 = ""
End Sub

Private Sub TEXT1_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("LOC", "LOC_IDY", Trim(Label3)) = True Then
Check2.SetFocus
Exit Sub
End If
rs.Open "select * from LOC where LOC_NME like '" & Trim(Text1) & "%'", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("LOC_NME")))
        ls.ListSubItems.Add , , Trim(rs("LOC_IDY") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Location"
    list1.ColumnHeaders(1).Text = "Location Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Location Not Found, Define Location !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text2_Change()
If Text2 = "" Then Label4 = ""
End Sub

Private Sub TEXT2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("CAT", "CAT_IDY", Trim(Label4)) = True Then
Check3.SetFocus
Exit Sub
End If

rs.Open "select * from CAT where cAT_nme like '" & Trim(Text2) & "%'", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("CAT_nme")))
        ls.ListSubItems.Add , , Trim(rs("cAT_idy") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Category"
    list1.ColumnHeaders(1).Text = "Category Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Category Not Found, Define Category...!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text2.Enabled = True
        Text2.SetFocus
    End If
End If


End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text3_Change()
If Text3 = "" Then Label5 = ""
End Sub

Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("MAR", "MAR_IDY", Trim(Label5)) = True Then
Check4.SetFocus
Exit Sub
End If

rs.Open "select * from mar where mar_nme like '" & Trim(Text3) & "%'", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("mar_nme")))
        ls.ListSubItems.Add , , Trim(rs("mar_idy") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Market"
    list1.ColumnHeaders(1).Text = "Market Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text3.Enabled = True
        Text3.SetFocus
    End If
End If
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text4_Change()
If Text4 = "" Then Label6 = ""
End Sub

Private Sub Text4_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("SMN", "SMN_IDY", Trim(Label6)) = True Then
Check5.SetFocus
Exit Sub
End If
rs.Open "select * from smn where smn_nme like '" & Trim(Text4) & "%'", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Salesman"
    list1.ColumnHeaders(1).Text = "Salesman Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text4.Enabled = True
        Text4.SetFocus
    End If
End If
End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text5_Change()
If Len(Trim(Text5)) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub TEXT5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text5 = discount(Text5)
    Check9.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text6_Change()
If Text6 = "" Then Label7 = ""
End Sub

Private Sub TEXT6_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
If check("ACT", "ACT_IDY", Trim(Label7)) = True Then
Check8.SetFocus
Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select * from act where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT')", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Retailer"
    list1.ColumnHeaders(1).Text = "Retailer Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Retailer Not Found, Define Retailer !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text6.Enabled = True
        Text6.SetFocus
    End If
End If
End If
End Sub

Private Sub Text6_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub
