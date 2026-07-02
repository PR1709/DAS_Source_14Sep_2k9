VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_PQY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "M_FRM_PQY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Select Supplier"
      TabPicture(0)   =   "M_FRM_PQY.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_PQY.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Select Dates"
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3855
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   2400
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16580611
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   600
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16580611
            CurrentDate     =   37258
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   2040
            TabIndex        =   15
            Top             =   315
            Width           =   195
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   315
            Width           =   345
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   3855
         Begin VB.CheckBox Check1 
            Caption         =   "&All"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   840
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Supplier"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   675
            Width           =   570
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   4080
         TabIndex        =   1
         Top             =   1320
         Width           =   1215
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   4
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2990
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
            Text            =   "Supplier Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Supplier ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Supplier"
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
         TabIndex        =   5
         Top             =   480
         Width           =   4935
      End
   End
End
Attribute VB_Name = "M_FRM_PQY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check1.Value = 1 Then
        Command1.SetFocus
    ElseIf Check1.Value = 0 Then
        Check1.Enabled = False
        Text1.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
If Check1.Value = 0 Then
    If Len(Trim(Text1)) = 0 Then
        MsgBox "Invalid Value", vbExclamation, "DAS Version 4.0"
        Text1.SetFocus
        Exit Sub
    End If
End If
Command1.Enabled = False
Dim REC_SET1 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    SAL_QRY = "inv->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inv->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    rs.Update
End If
If Check1.Value = 1 Then
    con.Execute "update dir where dir_idy like 'SUP' AND dir_idx like 'IDY' SET dir_als ='*'"
    Else
    con.Execute "update dir where dir_idy like 'SUP' AND dir_idx like 'IDY' SET dir_als ='" & Trim(Label5) & "'"
    con.Execute "update dir where dir_idy like 'SUP' AND dir_idx like 'IDY' SET dir_MSG ='" & Trim(Text1) & "'"
    SAL_QRY = SAL_QRY & ".AND.INL->CST_IDY='" & Trim(Label5) & "'"
End If

If M_FRM_PQY.Tag = "Purchase Summary" Then
    Unload Me
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PUR'"
    OpenReport CURDIR & "\report.RP1", "Purchase Report Date", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_PQY.Tag = "Transit Report" Then
    SAL_QRY = SAL_QRY & ".and. inv->txn_typ='PUR' .AND. INV->STK_TYP <> 'Salable'"
    OpenReport CURDIR & "\report.RP1", "Transit Report", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_PQY.Tag = "Purchase Report" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PUR'"
    OpenReport CURDIR & "\report.RP1", "Purchase Report", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_PQY.Tag = "Purchase Return Report" Then
    SAL_QRY = SAL_QRY & ".and.inv->txn_typ='PRT'"
    OpenReport CURDIR & "\report.RP1", "Purchase Return Report", 2, Trim(SAL_QRY), 6
ElseIf M_FRM_PQY.Tag = "Yearly Sales" Then
    SAL_QRY = "prd->cmp_idy='" & Trim(Label5) & "'"
    OpenReport CURDIR & "\MIR.RP1", "Yearly Sales", 2, SAL_QRY, 6
ElseIf M_FRM_PQY.Tag = "Yearly Products" Then
    SAL_QRY = "prd->cmp_idy='" & Trim(Label5) & "'"
    OpenReport CURDIR & "\MIR.RP1", "Yearly Products", 2, SAL_QRY, 6
ElseIf M_FRM_PQY.Tag = "Yearly Products Cases" Then
    SAL_QRY = "prd->cmp_idy='" & Trim(Label5) & "'"
    OpenReport CURDIR & "\MIR.RP1", "Yearly Products Cases", 2, SAL_QRY, 6
ElseIf Me.Tag = "Variance" Then
    On Error GoTo lab
    Me.MousePointer = vbHourglass
    con.Execute "DELETE FROM tpd"
    If REC_SET1.State = 1 Then REC_SET1.Close
    REC_SET1.Open "SELECT PRD.PRD_IDY,PRD_QTY,PRD_RAT FROM INV,INL,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND INL.INV_IDY LIKE INV.INV_IDY AND INL.CST_IDY LIKE '" & Trim(Label5) & "%' AND INV.TXN_TYP IN('PUR') AND PRD.BAS_RAT <> INV.PRD_RAT AND INV.INV_DAT <=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND INV.INV_DAT >=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "')", con, adOpenKeyset, adLockPessimistic
    While REC_SET1.EOF = False
        con.Execute "INSERT INTO tpd VALUES('" & Trim(REC_SET1(0) & "") & "'," & Val(REC_SET1(1) & "") & "," & Val(REC_SET1(2) & "") & ",0)"
        REC_SET1.MoveNext
    Wend
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\MIR.RP1", "Purchase Rate Changes Report", 2, "TPD->PRD_QTY > 0", 6
    Unload Me
Exit Sub
lab:
MsgBox "No Data Found", vbExclamation
Unload Me
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then Check1.SetFocus
If KeyAscii = 38 And s = 1 Then DTPicker1.SetFocus
End Sub

Private Sub DTPicker2_LostFocus()
If DTPicker1 > DTPicker2 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker1.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
If Me.Tag = "Yearly Sales" Or Me.Tag = "Yearly Products Cases" Or Me.Tag = "Yearly Products" Then
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Check1.SetFocus
Else
    DTPicker1.SetFocus
End If
SETYEAR M_FRM_PQY
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        Text1.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Sub List1_DblClick()
List1_keypress 13
End Sub

Private Sub List1_keypress(K As Integer)
If K = 13 Then
    Text1 = Trim(list1.SelectedItem)
    Label5 = Trim(list1.SelectedItem.ListSubItems.Item(1))
    list1.Visible = False
    Check1.Value = False
    SSTab1.Tab = 0
    Command1.SetFocus
End If
End Sub

Private Sub Text1_Change()
If Text1 = "" Then Label5 = ""
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If Len(Trim(Text1)) = 0 Then
        Supplierlist
    Else
        Dim rs As New ADODB.Recordset
        rs.Open "select * from act where (act_typ like 'BOT' OR ACT_TYP LIKE 'PUR') AND ACT_IDY LIKE '" & Sincrement(Text1) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            Label5 = Sincrement(Text1)
            Text1 = Trim(rs("ACT_NME"))
            Command1.SetFocus
        Else
            Supplierlist
        End If
    End If
End If
End Sub

Private Sub Supplierlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from act where (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'PUR') and act_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("ACT_nme")))
        ls.ListSubItems.Add , , Trim(rs("ACT_idy"))
        ls.ListSubItems.Add , , Trim(rs("ACT_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text1.SetFocus
    End If
End If
End Sub
