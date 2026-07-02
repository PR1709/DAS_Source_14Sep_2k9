VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_PAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "M_FRM_PAR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Report Filters"
      TabPicture(0)   =   "M_FRM_PAR.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTPicker1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DTPicker2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CHECK1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CHECK2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CHECK3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_PAR.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "list1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73440
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   3120
         Width           =   855
      End
      Begin VB.OptionButton CHECK3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Category"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1095
      End
      Begin VB.OptionButton CHECK2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Segment"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2085
         Width           =   1095
      End
      Begin VB.OptionButton CHECK1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Product"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2565
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73440
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73440
         MaxLength       =   30
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   375
         Left            =   -69240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -69240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   5400
         Width           =   855
      End
      Begin MSComctlLib.ListView list1 
         Height          =   4695
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   8281
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
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   -70680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24969219
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   -73200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24969219
         CurrentDate     =   37258
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Percentage (%)"
         Height          =   195
         Left            =   -74880
         TabIndex        =   17
         Top             =   3188
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   -73920
         TabIndex        =   9
         Top             =   555
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   -71340
         TabIndex        =   8
         Top             =   555
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   6225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -71040
         TabIndex        =   4
         Top             =   2595
         Width           =   45
      End
   End
End
Attribute VB_Name = "M_FRM_PAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then
Text2.Enabled = True
Text2.SetFocus
End If
End Sub

Private Sub CHECK2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.Enabled = True
    Text4.SetFocus
End If
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
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Command1_Click
    Else
        Check3.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
Dim RS As New ADODB.Recordset

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
SETYEAR M_FRM_PPQ
SSTab1.TabEnabled(1) = False
If M_FRM_PPQ.Tag = "TOP" Then
    Label1.Visible = True
    Label2.Visible = True
    DTPicker1.Visible = True
    DTPicker2.Visible = True
    DTPicker1.SetFocus
End If
SSTab1.Tab = 0
'Check7.Value = 1
End Sub

Private Sub Form_KEYPRESS(K As Integer)
If K = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        If Label20.Caption = "Select Product Category" Then Text3.SetFocus
        If Label20.Caption = "Select Product Segment" Then Text4.SetFocus
        If Label20.Caption = "Select Product" Then Text2.SetFocus
        If Label20.Caption = "Select Salesman" Then Text5.SetFocus
        If Label20.Caption = "Select Retailers Location" Then Text8.SetFocus
        If Label20.Caption = "Select Retailer" Then Text6.SetFocus
        If Label20.Caption = "Select Retailers Category" Then Text1.SetFocus
        If Label20.Caption = "Select Market" Then Text7.SetFocus
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
    If list1.ListItems.Count > 0 Then
        If Label20 = "Select Product Category" Then
            Text3 = list1.SelectedItem
            Text2 = ""
            Text4 = ""
            Label5 = ""
            SSTab1.Tab = 0
            list1.Visible = False
            command1.SetFocus
        ElseIf Label20 = "Select Product Segment" Then
            Text4 = list1.SelectedItem
            Text2 = ""
            Text3 = ""
            Label5 = ""
            SSTab1.Tab = 0
            list1.Visible = False
            command1.SetFocus
        ElseIf Label20 = "Select Product" Then
            Text2 = list1.SelectedItem.ListSubItems.Item(1)
            Label5 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            command1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text1_Keydown(KeyCode As Integer, S As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text18 = discount(Trim(Text18))
    command1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Or K = 45 Then
Else
K = 0
End If
End Sub

Private Sub Text2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("prd", "prd_idy", Trim(Text2)) = True Then
Check4.SetFocus
Else
RS.Open "select * from prd,pgr where prd_nme like '" & Trim(Text2) & "%' and prd.grp_idy like pgr.grp_idy", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        list1.ColumnHeaders.Clear
        Set ls = list1.ListItems.Add(, , Trim(RS("prd_nme")))
        ls.ListSubItems.Add , , Trim(RS("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(RS("grp_cls") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Product"
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Product Name"
    list1.ColumnHeaders(2).Text = "Product ID"
    list1.ColumnHeaders(3).Text = "Product Segement"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Product Not Found, Define Product!", vbExclamation, "Effmcg"
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
Dim RS As New ADODB.Recordset
If check("pgr", "grp_cat", Trim(Text3)) = True Then
    Check4.SetFocus
Else
RS.Open "select * from pgr where GRP_cat like '" & Trim(Text3) & "%' GROUP BY GRP_cat", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    Label20 = "Select Product Category"
    list1.ColumnHeaders.Clear
    list1.ColumnHeaders.Add , , , list1.Width - 100
    list1.ColumnHeaders(1).Text = "Product Category"
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("grp_cat")))
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Category Not Found, Define Category !", vbExclamation, "Effmcg"
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

Private Sub Text4_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("pgr", "grp_cls", Trim(Text4)) = True Then
Text4.SetFocus
Else
RS.Open "select * from pgr where grp_cls like '" & Trim(Text4) & "%' group by grp_cls", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    Label20 = "Select Product Segment"
    list1.ColumnHeaders.Clear
    list1.ColumnHeaders.Add , , , list1.Width - 100
    list1.ColumnHeaders(1).Text = "Product Segment"
    For A = 1 To RS.RecordCount
        list1.ListItems.Add , , Trim(RS("grp_cls"))
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text4.Enabled = True
        Text4.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub
