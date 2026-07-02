VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_AQY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "M_FRM_AQY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   5027
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Report Filters"
      TabPicture(0)   =   "M_FRM_AQY.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_AQY.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   3735
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   2400
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16515075
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   600
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16515075
            CurrentDate     =   37258
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   315
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   1980
            TabIndex        =   16
            Top             =   315
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   4200
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         Caption         =   "&All Accounts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   3735
         Begin VB.OptionButton CHECK3 
            Appearance      =   0  'Flat
            Caption         =   "&Group Account"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   278
            Width           =   1575
         End
         Begin VB.OptionButton CHECK1 
            Appearance      =   0  'Flat
            Caption         =   "Accoun&t"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   765
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1800
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   3201
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3960
         TabIndex        =   3
         Top             =   1995
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   2
         Top             =   480
         Width           =   5265
      End
   End
End
Attribute VB_Name = "M_FRM_AQY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Check7.Value = 0
Check1.Value = True
Text2.Enabled = True
Text2.SetFocus
End Sub

Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then
Check7.Value = 0
Text2.Enabled = True
Text2.SetFocus
End If
End Sub

Private Sub Check3_Click()
Check7.Value = 0
Check3.Value = True
Text3.Enabled = True
Text3.SetFocus
End Sub

Private Sub CHECK3_KeyPress(K As Integer)
If K = 13 Then
Text3.Enabled = True
Text3.SetFocus
End If
End Sub

Private Sub Check7_KeyPress(K As Integer)
If K = 13 Then
    If Check7.Value = 1 Then
        Frame2.Enabled = False
        Command1_Click
    Else
        Frame2.Enabled = True
        Check7.Enabled = False
        Check3.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Command1.Enabled = False
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
GROUPID = ""
ACCOUNTTYP = ""
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    
    
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
'    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
'    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
End If
If Check7.Value = 1 Then
    If Me.Tag <> "GLD" Then
        SAL_QRY = "ACT->GRP_CLF='Bank Book'.OR.ACT->GRP_CLF='Cash Book'"
    End If
Else
    If Check1.Value = True Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy like 'PRD' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            If Len(Trim(Text2)) > 0 Then
                rs("DIR_ALS") = Trim(Label5) & ""
                rs("DIR_MSG") = Trim(Text2) & ""
                rs.Update
                SAL_QRY = "ACT->act_idy='" & Trim(Label5) & "'"
                If Me.Tag = "ACT" Then SAL_QRY = "ACT->ACT_IDY='" & Trim(Label5) & "'"
                
            End If
        End If
    ElseIf Check3.Value = True Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy like 'PGR' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            If Len(Trim(Text3)) > 0 Then
                rs("DIR_ALS") = Trim(Label5) & ""
                rs("DIR_MSG") = Trim(Text3) & ""
                SAL_QRY = "act->grp_idy='" & Trim(Label5) & "'"
                If Me.Tag = "ACT" Then SAL_QRY = "ACT->GRP_IDY='" & Trim(Label5) & "'"
                rs.Update
            End If
            rs.Update
        End If
    End If
End If
If Me.Tag <> "ACT" Then
    If Check1.Value = True Then
        FAupdate Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), Trim(Label5), "ACT" 'ACCOUNT
    ElseIf Check3.Value = True Then
        FAupdate Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), Trim(Label5), "GRP" 'GROUP ACCOUNT
    Else
        FAupdate Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), "", ""
    End If
End If

If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "SELECT ACT_IDY FROM OPA WHERE ACT_IDY NOT IN(SELECT ACT_IDY FROM LED)", con, adOpenKeyset, adLockPessimistic
RS1.Open "SELECT * FROM LED", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    RS1.AddNew
    RS1("INV_DAT") = DTPicker1
    RS1("ACT_IDY") = rs(0)
    RS1.Update
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "SELECT ACT_IDY FROM LED WHERE ACT_IDY NOT IN(SELECT ACT_IDY FROM OPA) GROUP BY ACT_IDY", con, adOpenKeyset, adLockPessimistic
RS1.Open "SELECT * FROM OPA", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    RS1.AddNew
    RS1("ACT_IDY") = Trim(rs(0) & "")
    RS1.Update
    rs.MoveNext
Next

If Me.Tag = "GLD" Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
    End If
    OpenReport CURDIR & "\Finance.RP1", "General Ledger", 2, SAL_QRY, 6
ElseIf Me.Tag = "ACT" Then
    If Check7.Value = 0 Then
        OpenReport CURDIR & "\lists.RP1", "Accounts list", 2, SAL_QRY, 6
    Else
        OpenReport CURDIR & "\lists.RP1", "Accounts list", 0, SAL_QRY, 6
    End If
ElseIf Me.Tag = "CLJ" Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
        rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
        rs.Update
    End If
    OpenReport CURDIR & "\Finance.RP1", "Claim Journal Report", 2, SAL_QRY, 6
Else
    OpenReport CURDIR & "\Finance.RP1", "Bank Book", 2, SAL_QRY, 6
End If
SAL_QRY = ""
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check7.Enabled = True Then
        Check7.SetFocus
    Else
        Command1.SetFocus
    End If
ElseIf KeyCode = 38 And Shift = 1 Then
    DTPicker1.SetFocus
End If
End Sub

Private Sub DTPicker2_LostFocus()
If DTPicker1 > DTPicker2 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker1.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
If Me.Tag <> "ACT" Then
    SETYEAROPEN M_FRM_AQY
    DTPicker1.SetFocus
Else
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    Check7.SetFocus
End If
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
SAL_QRY = ""
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        If Label20.Caption = "Select Group Account" Then Text3.SetFocus
        If Label20.Caption = "Select Account" Then Text2.SetFocus
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
        If Label20 = "Select Group Account" Then
            Text3 = Trim(list1.SelectedItem)
            Text2 = ""
            Text4 = ""
            Label5 = list1.SelectedItem.ListSubItems.Item(2)
            Check7.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Account" Then
            Label5 = list1.SelectedItem.ListSubItems.Item(2)
            Text2 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            Text3 = ""
            Check7.Value = False
            list1.Visible = False
            Command1.SetFocus
        End If
    End If
End If
End Sub

Private Sub TEXT2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("act", "act_idy", Trim(Label5)) = True Then
Else
    If rs.State = 1 Then rs.Close
If Me.Tag = "ACT" Or Me.Tag = "CLJ" Then
    rs.Open "select * from act where act_nme like '" & Trim(Text2) & "%'", con, adOpenStatic
ElseIf Me.Tag <> "GLD" Then
    rs.Open "select * from act where act_nme like '" & Trim(Text2) & "%' And grp_clf in('Cash Book','Bank Book')", con, adOpenStatic
Else
    rs.Open "select * from act where act_nme like '" & Trim(Text2) & "%'", con, adOpenStatic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
'        list1.ColumnHeaders.Clear
        Set ls = list1.ListItems.Add(, , Trim(rs("ACT_NME")))
        ls.ListSubItems.Add , , Trim(rs("GRP_CLF") & "")
        ls.ListSubItems.Add , , Trim(rs("ACT_IDY") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Account"
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Account Name"
    list1.ColumnHeaders(2).Text = "Type"
    list1.ColumnHeaders(3).Text = "ID"
    list1.ColumnHeaders(1).Width = 3000
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Account Not Found, Define Account!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text2.Enabled = True
        Text2.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub


Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("pgr", "grp_cat", Trim(Text3)) = True Then
    Command1.SetFocus
Else
If Me.Tag = "ACT" Or Me.Tag = "CLJ" Then
    rs.Open "select * from grp where GRP_nme like '" & Trim(Text3) & "%'", con, adOpenStatic
ElseIf Me.Tag <> "GLD" Then
    rs.Open "select * from grp where GRP_nme like '" & Trim(Text3) & "%'And grp_clf in('Cash Book','Bank Book')", con, adOpenStatic
Else
    rs.Open "select * from grp where GRP_nme like '" & Trim(Text3) & "%'", con, adOpenStatic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    Label20 = "Select Group Account"
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Group Name"
    list1.ColumnHeaders(2).Text = "Type"
    list1.ColumnHeaders(3).Text = "ID"
    list1.ColumnHeaders(1).Width = 3000
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("GRP_NME")))
        ls.ListSubItems.Add , , Trim(rs("GRP_CLF"))
        ls.ListSubItems.Add , , Trim(rs("GRP_IDY"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Group Account Not Found, Define Group !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text3.Enabled = True
        Text3.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub
