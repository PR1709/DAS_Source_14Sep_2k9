VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PUR_PQY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "M_PUR_PQY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Report Filters"
      TabPicture(0)   =   "M_PUR_PQY.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_PUR_PQY.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "list1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Select Dates"
         Height          =   735
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   5775
         Begin VB.CheckBox Check8 
            Caption         =   "&FMCG Calendar"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   270
            Width           =   1455
         End
         Begin VB.ComboBox cmbCaldates 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   4440
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74186755
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   2760
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74186755
            CurrentDate     =   37258
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   2280
            TabIndex        =   27
            Top             =   315
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   4100
            TabIndex        =   26
            Top             =   315
            Width           =   195
         End
      End
      Begin VB.Frame Frame3 
         Height          =   650
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   5775
         Begin VB.CheckBox Check6 
            Caption         =   "All &Suppliers"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2760
            MaxLength       =   40
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   200
            Width           =   2895
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Supplier"
            Height          =   195
            Left            =   1920
            TabIndex        =   20
            Top             =   270
            Width           =   570
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1200
         Left            =   4800
         TabIndex        =   14
         Top             =   2880
         Width           =   1335
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1080
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All Products"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1695
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   4455
         Begin VB.OptionButton CHECK3 
            Caption         =   "Ca&tegory"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   278
            Width           =   1095
         End
         Begin VB.OptionButton CHECK2 
            Caption         =   "Seg&ment"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   758
            Width           =   1095
         End
         Begin VB.OptionButton CHECK1 
            Caption         =   "&Product"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1238
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   2655
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5318
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
      Begin VB.Label Label9 
         Height          =   195
         Left            =   3120
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4440
         TabIndex        =   4
         Top             =   2715
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   1080
         Width           =   195
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
         Top             =   600
         Width           =   5850
      End
   End
End
Attribute VB_Name = "M_PUR_PQY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSWEK As New ADODB.Recordset
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


Private Sub Check6_Click()
    If Check6.Value = 0 Then
        Check6.Enabled = False
        Text6.Enabled = True
        Text6.SetFocus
    Else
        Check7.SetFocus
    End If
End Sub

Private Sub Check6_KeyDown(K As Integer, s As Integer)
    If K = 13 Then
        If Check6.Value = 1 Then
            Check7.SetFocus
        Else
            Text6.Enabled = True
            Text6.SetFocus
        End If
    End If
End Sub

Private Sub Check7_KeyPress(K As Integer)
    If K = 13 Then
        If Check7.Value = 1 Then
            Frame2.Enabled = False
            Check1.Value = 0
            Check2.Value = 0
            Check3.Value = 0
            Command1.SetFocus
        Else
            Frame2.Enabled = True
            Check7.Enabled = False
            Check3.SetFocus
        End If
    End If
End Sub


Private Sub cmbCaldates_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If RSWEK.State = 1 Then RSWEK.Close
        RSWEK.Open "select frm_dat,to_dat from wek where wek_idy=" & Val(cmbCaldates.Text), con, adOpenKeyset, adLockPessimistic
        If Not RSWEK.EOF Then
            DTPicker1.Enabled = True
            DTPicker2.Enabled = True
            DTPicker1.Value = Format(RSWEK("frm_dat"), "MM/DD/YYYY")
            DTPicker2.Value = Format(RSWEK("to_dat"), "MM/DD/YYYY")
            DTPicker1.Enabled = False
            DTPicker2.Enabled = False
            Check6.SetFocus
        End If
    End If
End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Command1.Enabled = False
If rs.State = 1 Then rs.Close
'dates
    rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
            rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
            rs.Update
            SAL_QRY = "inv->inv_dat>=ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "').and.inv->inv_dat<=ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "')"
        End If
'Products
If Check1.Value = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'PRD' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Label5)) > 0 Then
            rs("DIR_ALS") = Trim(Label5) & ""
            rs("DIR_MSG") = Trim(Text2) & ""
            rs.Update
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.prd->prd_idy='" & Trim(Label5) & "'"
            Else
                SAL_QRY = "prd->prd_IDY='" & Trim(Label5) & "'"
            End If
        End If
    End If
ElseIf Check2.Value = True Then
'Segment
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'SEG' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Text4)) > 0 Then
            rs("DIR_MSG") = Trim(Text4) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.pgr->grP_cls='" & Trim(Text4) & "'"
            Else
                SAL_QRY = "pgr->grp_cls='" & Trim(Text4) & "'"
            End If
            rs.Update
        End If
    End If
ElseIf Check3.Value = True Then
'Category
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'PGR' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Text3)) > 0 Then
            rs("DIR_MSG") = Trim(Text3) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.pgr->grp_cat='" & Trim(Text3) & "'"
            Else
                SAL_QRY = "PGR->grp_cat='" & Trim(Text3) & "'"
            End If
            rs.Update
        End If
        rs.Update
    End If
End If

'Supplier ID
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'SUP' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text6) <> "" Then
        rs("DIR_ALS") = Trim(Label9) & ""
        rs("DIR_MSG") = Trim(Text6) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->ACT_IDY='" & Trim(Label9) & "'"
        Else
            SAL_QRY = "ACT->ACT_IDY='" & Trim(Label9) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
rs.Update
If Me.Tag = "PRD_DIS" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='STO'.OR.INV->TXN_TYP='PRT').AND.(inv->add_amt>0.or.inv->dis_amt>0)"
    OpenReport CURDIR & "\REPORT.RP1", "Product Discount Retailer Base", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "PUR_SUM_VWT" Then
    SAL_QRY = SAL_QRY & ".AND.INV->TXN_TYP='PUR'"
    OpenReport CURDIR & "\REPORT.RP1", "Purchase Product Volume Weight", 2, SAL_QRY, 6
    Unload Me
Else
    SAL_QRY = SAL_QRY & ".AND.INV->TXN_TYP='PUR'"
    OpenReport CURDIR & "\REPORT.RP1", "Purchase Product", 2, SAL_QRY, 6
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Check6.Enabled = True Then
    Check6.SetFocus
Else
    Check7.SetFocus
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
SETYEAROPEN M_PUR_PQY
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
Check7.Value = 1
Check6.Value = 1
DTPicker1.SetFocus
Check8.SetFocus
End Sub

Private Sub Form_KeyPress(K As Integer)
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
'''''        If Label20.Caption = "Select Retailers Category" Then Text1.SetFocus
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
    If list1.ListItems.count > 0 Then
        If Label20 = "Select Product Category" Then
            Text3 = list1.SelectedItem
            Text2 = ""
            Text4 = ""
            Label5 = ""
            SSTab1.Tab = 0
            Check7.Value = False
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Product Segment" Then
            Text4 = list1.SelectedItem
            Text2 = ""
            Text3 = ""
            Label5 = ""
            SSTab1.Tab = 0
            Check7.Value = False
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Product" Then
            Label5 = list1.SelectedItem.ListSubItems.Item(1)
            Text2 = Trim(list1.SelectedItem)
            Text3 = ""
            Text4 = ""
            SSTab1.Tab = 0
            Check7.Value = False
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Supplier" Then
            Label9 = list1.SelectedItem.ListSubItems.Item(1)
            Text6 = Trim(list1.SelectedItem)
            Check6.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Check7.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text2_Change()
If Text2 = "" Then Label5 = ""
End Sub

Private Sub TEXT2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("prd", "prd_idy", Trim(Label5)) = True Then
Check4.SetFocus
Else
If Check6.Value = Unchecked Then
    rs.Open "select * from prd,pgr where cmp_idy like '" & Trim(Label9) & "%' AND prd_nme like '" & Trim(Text2) & "%' and prd.grp_idy like pgr.grp_idy", con, adOpenStatic
Else
    rs.Open "select * from prd,pgr where prd_nme like '" & Trim(Text2) & "%' and prd.grp_idy like pgr.grp_idy", con, adOpenStatic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
'        list1.ColumnHeaders.Clear
        Set ls = list1.ListItems.Add(, , Trim(rs("prd_nme")))
        ls.ListSubItems.Add , , Trim(rs("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(rs("grp_cls") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Product"
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Product Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.ColumnHeaders(3).Text = "Product Segement"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Product Not Found, Define Product!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text2.Enabled = True
        Text2.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("pgr", "grp_cat", Trim(Text3)) = True Then
    Check4.SetFocus
Else
rs.Open "select * from pgr where GRP_cat like '" & Trim(Text3) & "%' GROUP BY GRP_cat", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    Label20 = "Select Product Category"
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , , list1.Width - 100
    list1.ColumnHeaders(1).Text = "Product Category"
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("grp_cat")))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Category Not Found, Define Category !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text3.Enabled = True
        Text3.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text4_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("pgr", "grp_cls", Trim(Text4)) = True Then
Text4.SetFocus
Else
rs.Open "select * from pgr where grp_cls like '" & Trim(Text4) & "%' group by grp_cls", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    Label20 = "Select Product Segment"
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , , list1.Width - 100
    list1.ColumnHeaders(1).Text = "Product Segment"
    For A = 1 To rs.RecordCount
        list1.ListItems.Add , , Trim(rs("grp_cls"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
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
End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text6_Change()
If Text6 = "" Then Label9 = ""
End Sub

Private Sub TEXT6_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("act", "act_idy", Trim(Label9)) = True Then
    Check7.SetFocus
Else
    rs.Open "select * from act  where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'PUR' or act_typ like 'BOT')", con, adOpenStatic
    list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy") & "")
        rs.MoveNext
    Next
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    SSTab1.Tab = 1
    Label20 = "Select Supplier"
    list1.ColumnHeaders(1).Text = "Supplier Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text6.Enabled = True
        Text6.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub Text6_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub


Private Sub Check8_Click()
Check8_KeyDown 13, 0
End Sub

Private Sub Check8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check8.Value = 1 Then
        cmbCaldates.Enabled = True
        If RSWEK.State = 1 Then RSWEK.Close
        RSWEK.Open "select wek_idy from wek", con, adOpenKeyset, adLockPessimistic
        cmbCaldates.CLEAR
        If Not RSWEK.EOF Then
            While Not RSWEK.EOF
                cmbCaldates.AddItem RSWEK("wek_idy")
                RSWEK.MoveNext
            Wend
            cmbCaldates.SetFocus
        Else
            MsgBox "Sales Calendar does not exist!", vbExclamation, "DAS Version 4.0"
            Check8.SetFocus
            Check8.Value = 0
        End If
    Else
        cmbCaldates.Enabled = False
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
        DTPicker1.SetFocus
    End If
End If
End Sub

Private Sub cmbCaldates_Click()
'    cmbCaldates_KeyDown 13, 0
End Sub


