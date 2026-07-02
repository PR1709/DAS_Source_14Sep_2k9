VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_PPQ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "M_FRM_PPQ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Report Filters"
      TabPicture(0)   =   "M_FRM_PPQ.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DTPicker1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTPicker2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_PPQ.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox Check7 
         Caption         =   "&All Products"
         Height          =   255
         Left            =   135
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   960
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   3960
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3735
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton CHECK1 
            Caption         =   "&Product"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1238
            Width           =   855
         End
         Begin VB.OptionButton CHECK2 
            Caption         =   "&Segment"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   758
            Width           =   1095
         End
         Begin VB.OptionButton CHECK3 
            Caption         =   "Ca&tegory"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   278
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   4905
         _ExtentX        =   8652
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
         Left            =   2520
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74317827
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   600
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74317827
         CurrentDate     =   37258
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   555
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   2160
         TabIndex        =   12
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
         Left            =   -74760
         TabIndex        =   10
         Top             =   600
         Width           =   4905
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3960
         TabIndex        =   9
         Top             =   2480
         Visible         =   0   'False
         Width           =   45
      End
   End
End
Attribute VB_Name = "M_FRM_PPQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then
Check7.Value = 0
Text2.Enabled = True
Text2.SetFocus
End If
End Sub

Private Sub CHECK2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check7.Value = 0
    Text4.Enabled = True
    Text4.SetFocus
End If
End Sub

Private Sub CHECK3_KeyPress(K As Integer)
If K = 13 Then
Check7.Value = 0
Text3.Enabled = True
Text3.SetFocus
End If
End Sub

Private Sub Check7_KeyPress(K As Integer)
If K = 13 Then
    If Check7.Value = 1 Then
        Frame2.Enabled = False
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Command1_Click
    Else
        Check7.Enabled = False
        Frame2.Enabled = True
        Check3.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
SAL_QRY = ""
Command1.Enabled = False
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
End If

Dim STL_QRY As String
If rs.State = 1 Then rs.Close
'Products
If M_FRM_PPQ.Tag = "Balance Stock" Then SAL_QRY = ""
If DTPicker1.Visible = True Then SAL_QRY = "inv->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inv->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
If Check1.Value = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'PRD' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Text2)) > 0 Then
            rs("DIR_ALS") = Trim(Label5) & ""
            rs("DIR_MSG") = Trim(Text2) & ""
            rs.Update
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.prd->prd_idy='" & Trim(Label5) & "'"
                STL_QRY = " AND prd.prd_idy='" & Trim(Label5) & "'"
            Else
                SAL_QRY = " AND prd.prd_IDY='" & Trim(Label5) & "'"
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
                SAL_QRY = SAL_QRY & ".AND.pgr->grp_cls='" & Trim(Text4) & "'"
                STL_QRY = STL_QRY & " AND pgr.Grp_cls='" & Trim(Text4) & "'"
            Else
                STL_QRY = " AND pgr.grp_cls='" & Trim(Text4) & "'"
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
                STL_QRY = STL_QRY & " AND pgr.grp_cat='" & Trim(Text3) & "'"
                SAL_QRY = SAL_QRY & ".AND.pgr->grp_cat='" & Trim(Text3) & "'"
            Else
                STL_QRY = " AND PGR.grp_cat='" & Trim(Text3) & "'"
                SAL_QRY = "PGR->grp_cat='" & Trim(Text3) & "'"
            End If
            rs.Update
        End If
        rs.Update
    End If
End If
If M_FRM_PPQ.Tag = "Balance Stock" Then
    OpenReport CURDIR & "\stock.RP1", "Stock Balance", 2, SAL_QRY, 6
    SAL_QRY = ""
ElseIf Me.Tag = "PRD" Then
    OpenReport CURDIR & "\lists.RP1", "Products List", 2, SAL_QRY, 6
ElseIf Me.Tag = "BAT" Then
    OpenReport CURDIR & "\lists.RP1", "Batch List", 2, SAL_QRY, 6
ElseIf Me.Tag = "PRC" Then
    OpenReport CURDIR & "\lists.RP1", "Price List", 2, SAL_QRY, 6
ElseIf Me.Tag = "STK" Then
    OpenReport CURDIR & "\lists.RP1", "Products List (80) Column", 2, SAL_QRY, 6
ElseIf Me.Tag = "RAT" Then
    OpenReport CURDIR & "\lists.RP1", "Products Report", 2, SAL_QRY, 6
ElseIf Me.Tag = "BAL_VAL" Then
    OpenReport CURDIR & "\stock.RP1", "Stock Balance Supplier", 2, SAL_QRY, 6
ElseIf Me.Tag = "LST_IDY" Then
    OpenReport CURDIR & "\lists.RP1", "Products List ID", 2, SAL_QRY, 6
ElseIf Me.Tag = "Stock Ledger" Or Me.Tag = "Stock Ledger Test" Then
    STL_LED Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), STL_QRY, Me.Tag
ElseIf Me.Tag = "DMG_BRK_SEG" Then
    OpenReport CURDIR & "\stock.RP1", "Damage Breakup Segment", 2, SAL_QRY, 6
ElseIf Me.Tag = "DMG_BRK" Then
    OpenReport CURDIR & "\stock.RP1", "Damage Breakup", 2, SAL_QRY, 6
ElseIf Me.Tag = "PAR_MAR" Then
    SAL_QRY = SAL_QRY & " .and. INV->TXN_TYP = 'SAL'"
    OpenReport CURDIR & "\MIR.RP1", "Product Pickup Market", 2, SAL_QRY, 6
ElseIf Me.Tag = "PAR_ROU" Then
    SAL_QRY = SAL_QRY & " .and. INV->TXN_TYP = 'SAL'"
    OpenReport CURDIR & "\MIR.RP1", "Product Pickup Route", 2, SAL_QRY, 6
ElseIf Me.Tag = "PAR_CAT" Then
    SAL_QRY = SAL_QRY & " .and. INV->TXN_TYP = 'SAL'"
    OpenReport CURDIR & "\MIR.RP1", "Product Pickup Categories", 2, SAL_QRY, 6
ElseIf Me.Tag = "PAR_LOC" Then
    SAL_QRY = SAL_QRY & " .and. INV->TXN_TYP = 'SAL'"
    OpenReport CURDIR & "\MIR.RP1", "Product Pickup Locations", 2, SAL_QRY, 6
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
If KeyCode = 38 And Shift = 1 Then DTPicker1.SetFocus
End Sub

Private Sub DTPicker2_LostFocus()
If DTPicker1 > DTPicker2 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker1.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
SETYEAROPEN M_FRM_PPQ
SSTab1.TabEnabled(1) = False
If M_FRM_PPQ.Tag = "TOP" Or Me.Tag = "Stock Ledger" Or Me.Tag = "DMG_BRK" Or Me.Tag = "DMG_BRK_SEG" Or Me.Tag = "PAR_MAR" Or Me.Tag = "PAR_ROU" Or Me.Tag = "PAR_LOC" Or Me.Tag = "PAR_CAT" Or Me.Tag = "Stock Ledger Test" Then
    Label1.Visible = True
    Label2.Visible = True
    DTPicker1.Visible = True
    DTPicker2.Visible = True
    DTPicker1.SetFocus
End If
SSTab1.Tab = 0
Check7.Value = 1
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
            Check7.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Product Segment" Then
            Text4 = list1.SelectedItem
            Text2 = ""
            Text3 = ""
            Label5 = ""
            Check7.Value = False
            SSTab1.Tab = 0
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

Else
rs.Open "select * from prd,pgr where prd_nme like '" & Trim(Text2) & "%' and prd.grp_idy like pgr.grp_idy", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        list1.ColumnHeaders.CLEAR
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
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
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
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub
