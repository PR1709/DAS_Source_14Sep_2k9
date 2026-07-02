VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_SQY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "M_REP_SQY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   10583
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Report Control Filters"
      TabPicture(0)   =   "M_REP_SQY.frx":0442
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
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_SQY.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   3600
         TabIndex        =   38
         Top             =   1210
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox txtNoDays 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   960
            MaxLength       =   3
            TabIndex        =   39
            Top             =   200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblNoDays 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No of Days"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Select Dates"
         Height          =   735
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   6375
         Begin VB.CheckBox Check8 
            Caption         =   "&FMCG Calendar"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   1455
         End
         Begin VB.ComboBox cmbCaldates 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   4920
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16646147
            CurrentDate     =   40009
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   3000
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16646147
            CurrentDate     =   40009
         End
         Begin VB.Label FDAT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   2520
            TabIndex        =   37
            Top             =   315
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   4560
            TabIndex        =   36
            Top             =   315
            Width           =   195
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   240
         TabIndex        =   28
         Top             =   1210
         Width           =   3255
         Begin VB.ListBox PTY_TYP 
            Height          =   285
            ItemData        =   "M_REP_SQY.frx":047A
            Left            =   1080
            List            =   "M_REP_SQY.frx":04A2
            Style           =   1  'Checkbox
            TabIndex        =   29
            Top             =   210
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Party Type"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   255
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   1680
         TabIndex        =   24
         Top             =   1850
         Visible         =   0   'False
         Width           =   3495
         Begin VB.OptionButton Option2 
            Caption         =   "&Invoice"
            Height          =   195
            Left            =   1440
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Product"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   5280
         TabIndex        =   21
         Top             =   4420
         Width           =   1335
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1970
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   3375
         Left            =   240
         TabIndex        =   9
         Top             =   2500
         Width           =   4935
         Begin VB.ComboBox list2 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   2280
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   855
            Width           =   2655
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1335
            Width           =   2655
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2775
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Location"
            Height          =   375
            Left            =   360
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Ca&tegory"
            Height          =   375
            Left            =   360
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "&Market"
            Height          =   375
            Left            =   360
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "&Salesman"
            Height          =   375
            Left            =   360
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Ro&ute"
            Height          =   375
            Left            =   360
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "&Retailer"
            Height          =   375
            Left            =   360
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   2760
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   6376
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Market Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   3840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   2880
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
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
Attribute VB_Name = "M_REP_SQY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSWEK As New ADODB.Recordset
Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then
If Check1.Value = 1 Then
    Check7.Value = 0
    Text1.Enabled = True
    Text1.SetFocus
Else
    Text1.Enabled = False
    Text1 = ""
    Label3 = ""
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
    Label4 = ""
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
    Label5 = ""
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
    Label6 = ""
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

Private Sub Check6_keypress(K As Integer)
If K = 13 Then
If Check6.Value = 1 Then
    Check7.Value = 0
    Text6.Enabled = True
    Text6.SetFocus
Else
    Text6.Enabled = False
    Text6 = ""
    Label7 = ""
    Command1.SetFocus
End If
End If
End Sub

Private Sub Check7_KeyPress(K As Integer)
If K = 13 Then
    If Check7.Value = 1 Then
        Frame1.Enabled = False
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = 0
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        Command1_Click
    Else
        Check7.Enabled = False
        Frame1.Enabled = True
        Check1.SetFocus
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
        PTY_TYP.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    If Me.Tag = "SAL_N14" Or Me.Tag = "SAL_BPS" Or Me.Tag = "SAL_SGA" Then
        rs("DIR_DAT") = Format(DTPicker2, "MM/DD/YYYY")
    Else
        rs("DIR_DAT") = Format(DTPicker2, "DD/MM/YYYY")
    End If
    rs.Update
    If Me.Tag = "TOP" Or Me.Tag = "DRL" Or Me.Tag = "CLG_DLT" Then
    SAL_QRY = ""
    Else
        If Me.Tag = "CFR" Or Me.Tag = "STR" Or Me.Tag = "MCFR" Then
        SAL_QRY = "INL->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inl->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        ElseIf Me.Tag = "dbt_agd" Then
        SAL_QRY = "irt->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.irt->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "').AND.IRT->DAY_BTW>=" & Val(txtNoDays.Text)
        ElseIf Me.Tag = "REC_STM" Then
        SAL_QRY = "INL->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and. INL->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        ElseIf Me.Tag = "SMN_COM_PYB" Then
        SAL_QRY = "VAN->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and. VAN->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "') .and. (VAN_CLS = 'S' .OR. VAN_CLS ='C')"
        ElseIf Me.Tag = "ROU_GP_RPT" Then
        SAL_QRY = "VAN->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and. VAN->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        Else
        SAL_QRY = "INL->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.INL->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        End If
    End If
End If
con.Execute "update dir set dir_als='" & Val(txtNoDays.Text) & "' where dir_idy='SYS' AND DIR_IDX='DEM'"
If Me.Tag = "Shell" Then
    SAL_QRY = ""
End If

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
        If Len(Trim(SAL_QRY)) <> 0 Then
            If Me.Tag = "TOP" Or Me.Tag = "NOP" Or Me.Tag = "Shell" Then
                SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label6) & "'"
            ElseIf Me.Tag = "dbt_agd" Then
               SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label6) & "'"
'               SAL_QRY = SAL_QRY & ".AND.ACT1->SMN_IDY='" & Trim(Label6) & "'"
            ElseIf Me.Tag = "REC_STM" Then
               SAL_QRY = SAL_QRY & ".AND. INL->SMN_IDY='" & Trim(Label6) & "'"
            ElseIf Me.Tag = "SMN_COM_PYB" Then
                SAL_QRY = SAL_QRY & ".AND. VAN->SMN_IDY='" & Trim(Label6) & "'"
            Else
                SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label6) & "'"
            End If
        Else
            If Me.Tag = "NOP" Or Me.Tag = "Shell" Then      'REMOVED "TOP" TAG AND PLACED ELSE IFCONDITION BELOW INSPITE OF ELSE
                SAL_QRY = "INL->SMN_IDY='" & Trim(Label6) & "'"
            ElseIf Me.Tag <> "TOP" And Me.Tag <> "CLG_DLT" Then
                SAL_QRY = "INL->SMN_IDY='" & Trim(Label6) & "'"
            End If
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If

If Me.Tag <> "Trip Sheet" Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'ROU' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Trim(List2.Text) <> "" Then
            rs("DIR_ALS") = Trim(List2) & ""
            rs("DIR_MSG") = Trim(List2) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                If Me.Tag = "TOP" Or Me.Tag = "NOP" Then
                    SAL_QRY = SAL_QRY & ".AND.INL->ROU_IDY='" & Trim(List2) & "'"
                ElseIf Me.Tag = "SMN_COM_PYB" Then
                    SAL_QRY = SAL_QRY & ".AND. VAN->ROU_IDY='" & Trim(List2) & "'"
                Else
                    SAL_QRY = SAL_QRY & ".AND.INL->ROU_IDY='" & Trim(List2) & "'"
                End If
            Else
                If Me.Tag = "NOP" Or Me.Tag = "Shell" Then  'REMOVED "TOP" TAG AND PLACED ELSE IFCONDITION BELOW INSPITE OF ELSE
                    SAL_QRY = "INL->ROU_IDY='" & Trim(List2) & "'"
                ElseIf Me.Tag <> "TOP" And Me.Tag <> "CLG_DLT" Then
                    SAL_QRY = "INL->ROU_IDY='" & Trim(List2) & "'"
                End If
            End If
        Else
            rs("DIR_ALS") = "*"
            rs("DIR_MSG") = "All"
        End If
        rs.Update
    End If
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'DLR' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text6) <> "" Then
        rs("DIR_ALS") = Trim(Label7) & ""
        rs("DIR_MSG") = Trim(Text6) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            If Me.Tag = "dbt_agd" Then
               SAL_QRY = SAL_QRY & ".AND.ACT1->ACT_IDY='" & Trim(Label7) & "'"
            ElseIf Me.Tag = "REC_STM" Then
               SAL_QRY = SAL_QRY & ".AND. INL->CST_IDY = '" & Trim(Label7) & "'"
            ElseIf Me.Tag = "Shell" Then
                SAL_QRY = SAL_QRY & ".AND.ACT->ACT_IDY='" & Trim(Label7) & "'"
            Else
                SAL_QRY = SAL_QRY & ".AND.ACT->ACT_IDY='" & Trim(Label7) & "'"
            End If
        Else
            SAL_QRY = "ACT->ACT_IDY='" & Trim(Label7) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
If PTY_TYP.Selected(0) = False Then
    For A = 1 To 4
        If PTY_TYP.Selected(A) = True Then
            If SALQRY = "" Then
                SALQRY = "act->cst_typ='" & PTY_TYP.List(A) & "'"
            Else
                SALQRY = SALQRY & ".or.act->cst_typ='" & PTY_TYP.List(A) & "'"
            End If
        End If
    Next
    If Trim(SALQRY) <> "" Then
        SALQRY = "(" & SALQRY & ")"
        If Len(Trim(SAL_QRY)) = 0 Then
            SAL_QRY = SALQRY
        Else
            SAL_QRY = SAL_QRY & ".and." & SALQRY
        End If
    End If
End If
If Me.Tag = "PPD_SEG" Then
   SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP = 'SAL'"
   OpenReport CURDIR & "\MIR.RP1", "Segment Pickup Retailer", 2, SAL_QRY, 6
   Unload Me
ElseIf Me.Tag = "Shell" Then
    M_CNF = MsgBox("Update Shells Account Before Print?", vbYesNo + vbQuestion, "DAS Version 4.0")
    If M_CNF = vbYes Then
        Me.MousePointer = vbhourglasss
        SHLPARTYACCOUNT
        Me.MousePointer = vbDefault
    End If
    If Len(Trim(SAL_QRY)) = 0 Then
        SAL_QRY = "(SAL_STK<>0.OR.FRE_STK<>0).and.shl_typ='Shell'"
    Else
        SAL_QRY = SAL_QRY & ".AND.(SAL_STK<>0.OR.FRE_STK<>0).and.shl_typ='Shell'"
    End If
    OpenReport CURDIR & "\LISTS.RP1", "Shells Party Wise", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "PPD_CAT" Then
   SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP = 'SAL'"
   OpenReport CURDIR & "\MIR.RP1", "Category Pickup Retailer", 2, SAL_QRY, 6
   Unload Me
ElseIf Me.Tag = "SOT" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\mir.RP1", "Secondary Offtake", 2, SAL_QRY, 6
    Unload M_REP_SQY
ElseIf Me.Tag = "Trip Sheet" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL')"
    M_FRM_RUT.Tag = Me.Tag
    Unload Me
    M_FRM_RUT.Show 1
ElseIf Me.Tag = "Scheme Value" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\REPORT.RP1", "Sales Scheme Value", 2, SAL_QRY, 6
    Unload M_REP_SQY
ElseIf M_REP_SQY.Tag = "DRL" Or M_REP_SQY.Tag = "ARR_COL" Then
    Dim DAT As Date
    C = Format(DTPicker2, "DD/MM/YYYY")
    DAT = DateAdd("d", 1, DTPicker2)
    TRAIL_BALANCE Format(DAT, "MM/DD/YYYY")
    If Len(Trim(SAL_QRY)) = 0 Then
        SAL_QRY = "OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
    Else
        SAL_QRY = SAL_QRY & ".AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
    End If
    con.Execute "update dir set dir_MSG='" & C & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
'    CON.Execute "update dir set DIR_DAT='" & Format(DTPicker2, "MM/DD/YYYY") & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
    If Len(Trim(Label5)) > "0" Then
       SAL_QRY = "ACT->MAR_IDY='" & Trim(Label5) & "'.AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
     ElseIf Len(Trim(Label3)) > "0" Then
       SAL_QRY = "ACT->LOC_IDY='" & Trim(Label3) & "'.AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
     ElseIf Len(Trim(Label4)) > "0" Then
       SAL_QRY = "ACT->CAT_IDY='" & Trim(Label4) & "'.AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
   
    ElseIf Len(Trim(Label7)) > "0" Then
       SAL_QRY = "ACT->ACT_IDY='" & Trim(Label7) & "'.AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
'    ElseIf Len(Trim(Label6)) > "0" Then
'       SAL_QRY = "INL->SMN_IDY='" & Trim(Label6) & "'.AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
'    ElseIf Len(Trim(LIST2)) > "0" Then
'       SAL_QRY = "INL->ROU_IDY='" & Trim(LIST2.Text) & "'.AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
  Else
        SAL_QRY = "OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
    End If
    If M_REP_SQY.Tag = "ARR_COL" Then
        OpenReport CURDIR & "\KOR.RP1", "Retailer AR and COL Report", 2, Trim(SAL_QRY), 6
    Else
        OpenReport CURDIR & "\Finance.RP1", "Debtors List", 2, Trim(SAL_QRY), 6
    End If
    Unload Me
    
ElseIf Me.Tag = "TOP" Then
    con.Execute "DELETE FROM TOP"
    If REC_SET.State = 1 Then REC_SET.Close
    If Len(Trim(Label6)) <> 0 Then          'CHANGED BY PKB
        REC_SET.Open "SELECT CST_IDY,SUM(GRS_AMT),SUM(NET_AMT) FROM INL WHERE SMN_IDY = '" & Label6 & "' AND INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV_DAT <= CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND TXN_TYP IN('LSL','SAL') GROUP BY CST_IDY", con, adOpenKeyset, adLockPessimistic   'CHANGED BY PKB
'    Else        'CHANGED BY PKB
'        REC_SET.Open "SELECT CST_IDY,SUM(GRS_AMT),SUM(NET_AMT) FROM INL WHERE INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV_DAT <= CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND TXN_TYP IN('LSL','SAL') GROUP BY CST_IDY", CON, adOpenKeyset, adlockpessimistic
'    End If      'CHANGED BY PKB
'    If REC_SET.State = 1 Then REC_SET.Close
    ElseIf Len(Trim(List2)) <> 0 Then          'CHANGED BY PKB
        REC_SET.Open "SELECT CST_IDY,SUM(GRS_AMT),SUM(NET_AMT) FROM INL WHERE ROU_IDY = '" & List2 & "' AND INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV_DAT <= CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND TXN_TYP IN('LSL','SAL') GROUP BY CST_IDY", con, adOpenKeyset, adLockPessimistic   'CHANGED BY PKB
    Else        'CHANGED BY PKB
        REC_SET.Open "SELECT CST_IDY,SUM(GRS_AMT),SUM(NET_AMT) FROM INL WHERE INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV_DAT <= CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND TXN_TYP IN('LSL','SAL') GROUP BY CST_IDY", con, adOpenKeyset, adLockPessimistic
    End If      'CHANGED BY PKB
    
    While REC_SET.EOF = False
        con.Execute "INSERT INTO TOP VALUES('" & Trim(REC_SET(0) & "") & "'," & Val(REC_SET(1)) & "," & Val(REC_SET(2)) & ")"
        REC_SET.MoveNext
    Wend
    OpenReport CURDIR & "\MIR.RP1", "Top Outlets", 2, SAL_QRY, 6
    Unload M_REP_SQY
ElseIf Me.Tag = "LBD" Then
    SAL_QRY = SAL_QRY & ".AND. ACT->DBT_TYP = 'Y' .AND.(INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\PNT_MIS.RP1", "Last Billing Date", 2, SAL_QRY, 6
    Unload M_REP_SQY
ElseIf Me.Tag = "DSCPC" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\PNT_MIS.RP1", "Discount Scheme CPC", 2, SAL_QRY, 6
    Unload M_REP_SQY
ElseIf Me.Tag = "BRKREP" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->STK_TYP='Breakage')"
    OpenReport CURDIR & "\report.RP1", "Breakage Report", 2, SAL_QRY, 6
    Unload M_REP_SQY
ElseIf Me.Tag = "LEKREP" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->STK_TYP='Leakage')"
    OpenReport CURDIR & "\report.RP1", "Leakage Report", 2, SAL_QRY, 6
    Unload M_REP_SQY
ElseIf Me.Tag = "SMR" Then
    Me.MousePointer = vbHourglass
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    OpenReport CURDIR & "\REP_LIB.RP1", "Stock Movement", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "VAN" Then
    OpenReport CURDIR & "\report.RP1", "Sales Report - Van", 2, Trim(SAL_QRY) & ".AND.LEN(LTRIM(RTRIM(INL->ORD_IDY)))=10.AND.INL->TXN_TYP='SAL'", 6
    Unload Me
ElseIf Me.Tag = "RAT" Then
On Error GoTo lab
    Me.MousePointer = vbHourglass
    con.Execute "DELETE FROM tpd"
    If REC_SET1.State = 1 Then REC_SET1.Close
    REC_SET1.Open "SELECT INV.PRD_IDY,INT(SUM(((INV.PRD_QTY-INT(INV.PRD_QTY))*1000)+(INT(INV.PRD_QTY)*INV.SUB_UNT))/INV.SUB_UNT)+(MOD(SUM(((INV.PRD_QTY-INT(INV.PRD_QTY))*1000)+(INT(INV.PRD_QTY)*INV.SUB_UNT)),INV.SUB_UNT)/1000) TOTAL,INV.PRD_RAT FROM INV,INL,PRD,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV.INV_IDY LIKE INL.INV_IDY  AND PRD.PRD_IDY LIKE INV.PRD_IDY AND INV.TXN_TYP IN('SAL') AND PRD.SAL_RAT <> INV.PRD_RAT AND ACT.MAR_IDY LIKE '" & Trim(Label5) & "%' AND ACT.LOC_IDY LIKE '" & Trim(Label3) & "%' AND ACT.CAT_IDY LIKE '" & Trim(Label4) & "%' AND INL.SMN_IDY LIKE '" & Trim(Label6) & "%' AND INL.ROU_IDY LIKE '" & Trim(List2) & "%' AND ACT.ACT_IDY LIKE '" & Trim(Label7) & "%' AND INV.INV_DAT >=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV.INV_DAT <=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') GROUP BY INV.PRD_IDY,INV.PRD_RAT", con, adOpenKeyset, adLockPessimistic
    While REC_SET1.EOF = False
        con.Execute "INSERT INTO tpd VALUES('" & Trim(REC_SET1(0) & "") & "'," & Val(REC_SET1(1) & "") & "," & Val(REC_SET1(2) & "") & ",0)"
        REC_SET1.MoveNext
    Wend
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\MIR.RP1", "Rate Changes Report", 2, "TPD->PRD_QTY > 0", 6
    Unload Me
Exit Sub
lab:
MsgBox "Cannot Process Report! Data Not Found!", vbExclamation, "DAS Version 4.0"
Unload Me
ElseIf Me.Tag = "CFR" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\MIR.RP1", "Coverage Frequency Report", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "MCFR" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\MIR.RP1", "Market Coverage Frequency", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "STR" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\MIR.RP1", "Salesman Turnover Report", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "RAR" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    OpenReport CURDIR & "\MIR.RP1", "Retailers Addition Report", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "dbt_agd" Then
    'Dim IRTRS As New ADODB.Recordset
    'IRTRS.Open "SELECT * FROM IRT", CON, adOpenKeyset, adlockpessimistic
    SAL_QRY = SAL_QRY & ".AND.(IRT->NET_AMT - IRT->RCD_AMT) <> 0"
    OpenReport CURDIR & "\REP_LIB.RP1", "Debtors Aging CHK", 2, SAL_QRY, 6
    Unload Me
'ElseIf Me.Tag = "dbt_agd_chk" Then
    'SAL_QRY = SAL_QRY & ".AND.(IRT->NET_AMT - IRT->RCD_AMT) <> 0"
    'Openreport CURDIR & "\REP_LIB.RP1", "Debtors Aging CHK", 2, SAL_QRY, 6
    'Unload Me
ElseIf Me.Tag = "REC_STM" Then
    SAL_QRY = SAL_QRY & ".AND. (IRT->NET_AMT <> IRT->RCD_AMT)"
    SAL_QRY = SAL_QRY & ".AND. INL->TXN_TYP = 'SAL'"
    OpenReport CURDIR & "\REP_LIB.RP1", "Debtors Account", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "Sales Report" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    OpenReport CURDIR & "\report.RP1", "Sales Report Date", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SSU_RCH" Then
    SAL_QRY = SAL_QRY & ".and. (INV->TXN_TYP='SAL' .and. INV->STK_TYP<>'TakeBack')"
    OpenReport CURDIR & "\REPORT.RP1", "Sales Summary Retail Channel", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SSU_RLO" Then
    SAL_QRY = SAL_QRY & ".and. (INV->TXN_TYP='SAL' .and. INV->STK_TYP<>'TakeBack')"
    OpenReport CURDIR & "\REPORT.RP1", "Sales Summary Retail Location", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SSU_RMA" Then
    SAL_QRY = SAL_QRY & ".and. (INV->TXN_TYP='SAL' .and. INV->STK_TYP<>'TakeBack')"
    OpenReport CURDIR & "\REPORT.RP1", "Sales Summary Retail Market", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SSU_RRO" Then
    SAL_QRY = SAL_QRY & ".and. (INV->TXN_TYP='SAL' .and. INV->STK_TYP<>'TakeBack')"
    OpenReport CURDIR & "\REPORT.RP1", "Sales Summary Retail Route", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SRP_DAC" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\REPORT.RP1", "Product Discount Retail Group", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SSU_RSM" Then
    SAL_QRY = SAL_QRY & ".and. (INV->TXN_TYP='SAL' .and. INV->STK_TYP<>'TakeBack')"
    OpenReport CURDIR & "\REPORT.RP1", "Sales Summary Retail Salesman", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "Load Reports" Then
    SAL_QRY = SAL_QRY & ".AND. (INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\report.RP1", "Load Report One", 2, Trim(SAL_QRY), 6
    OpenReport CURDIR & "\report.RP1", "Load Report Two", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "DPT" Then
    'If Option2.Value = True Then
        SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='SAL').AND.(INL->spd_AMT>0.OR.INL->csd_AMT>0.OR.INl->exd_AMT>0)"
        OpenReport CURDIR & "\report.RP1", "Discount Report", 2, Trim(SAL_QRY), 6
    'Else
    '    M_PUR_PQY.Tag = "PRD_DIS"
    '    M_PUR_PQY.Show 1
    'End If
    Unload Me
ElseIf Me.Tag = "Sales Summary" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    OpenReport CURDIR & "\report.RP1", "Sales Summary", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "Sales Summary(Customer)" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    OpenReport CURDIR & "\report.RP1", "Sales Report (Customer)", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "CLG_DLT" Then
    Dim DAT1 As Date
    C = Format(DTPicker2, "DD/MM/YYYY")
    DAT1 = DateAdd("d", 0, DTPicker2)
   ' TRAIL_BALANCECLG Format(DAT1, "MM/DD/YYYY")
    If Len(Trim(SAL_QRY)) = 0 Then 'Or Me.Tag = "CLG_DLT" Then
  '      SAL_QRY = "OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
'    ElseIf Me.Tag = "CLG_DLT" Then
    SAL_QRY = "OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')" '.AND.INL->SMN_IDY='" & Trim(Label6) & "' "
  '    SAL_QRY = SAL_QRY & ".AND. INL->SMN_IDY='" & Trim(Label6) & "'"
    Else
        SAL_QRY = SAL_QRY & ".AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
    End If
    con.Execute "update dir set dir_MSG='" & C & "' where  dir_idy like 'HDR' and dir_idx like 'DAT'"
    OpenReport CURDIR & "\Finance.RP1", "Debtors List", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SAL_SGA" Then
    SAL_QRY = SAL_QRY & ".AND. (INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\KOR.RP1", "Sales and SGA Productivity", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SAL_N14" Then
    SAL_QRY = SAL_QRY & ".AND. (INV->TXN_TYP='SAL') .AND. ACT->DBT_TYP = 'Y'"
    Dim CNT_RS As New ADODB.Recordset
    Dim ACT_RS As New ADODB.Recordset
    If CNT_RS.State = 1 Then CNT_RS.Close
    CNT_RS.Open "SELECT COUNT(DISTINCT PRD_IDY), CST_IDY FROM INV INNER JOIN INL ON INV.INV_IDY = INL.INV_IDY WHERE INV.TXN_TYP = 'SAL' .AND. INL->INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') .AND. INL->INV_DAT <= CTOD('" & Format(DateAdd("D", -14, DTPicker2), "MM/DD/YYYY") & "') GROUP BY INL.CST_IDY", con, adOpenKeyset, adLockPessimistic
    If ACT_RS.State = 1 Then ACT_RS.Close
    ACT_RS.Open "SELECT * FROM ACT", con, adOpenKeyset, adLockPessimistic
    If ACT_RS.RecordCount > 0 Then ACT_RS.MoveFirst
    While Not CNT_RS.EOF
        ACT_RS.Find "ACT_IDY='" & Trim(CNT_RS(1)) & "'"
        If ACT_RS.EOF = False Then
            ACT_RS("SKU_RNG") = Val(Trim(CNT_RS(0)))
            ACT_RS.Update
        End If
        CNT_RS.MoveNext
    Wend
    ACT_RS.Close
    CNT_RS.Close
    OpenReport CURDIR & "\KOR.RP1", "Not Billed Past 14 Days", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SAL_RPR" Then
'    AddScheduledCalls DTPicker1, DTPicker2, "SMN"
    FillScheduledCalls
    'SAL_QRY = SAL_QRY & ".AND. (INV->TXN_TYP='SAL') .AND. (ACT->DBT_TYP='Y')"
    SAL_QRY = SAL_QRY & ".AND. (INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\KOR.RP1", "Route Productivity Report", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SAL_BPS" Then
    SAL_QRY = SAL_QRY & ".AND. (INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\KOR.RP1", "Secondary Summary Sales Report", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "SMN_COM_PYB" Then
    SAL_QRY = SAL_QRY '& ".AND. (INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\KOR.RP1", "Commission Payable Report", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "ROU_GP_RPT" Then
    SAL_QRY = SAL_QRY '& ".AND. (INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\KOR.RP1", "Route GP Report", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "GDW_GP_RPT" Then
    SAL_QRY = "INL->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.INL->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    SAL_QRY = SAL_QRY & " .AND. INV->TXN_TYP = 'SAL' .AND. LEN(TRIM(INV->ORD_IDY)) = 0"
    OpenReport CURDIR & "\KOR.RP1", "Godown GP Report", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "HVOS" Then
    'MsgBox SAL_QRY & " .AND. Trim(CSR->DIS_TYP) = 'OnSpot' .AND. INV->ADD_AMT > 0"
    SAL_QRY = SAL_QRY & " .AND. (INV->ADD_AMT > 0) .AND. (Trim(ACT->CST_TYP) = 'HVO-NP' .OR. TRIM(ACT->CST_TYP) = 'HVO-NP-WO') .AND. TRIM(INV->TXN_TYP) = 'SAL'"
    
'    SAL_QRY = "INV->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and. INV->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')  .AND. (INV->ADD_AMT > 0) .AND. Trim(ACT->CST_TYP) = 'HVO-NP' .AND. TRIM(INV->TXN_TYP) = 'SAL'"
    OpenReport CURDIR & "\KOR.RP1", "HVO OnSpot Rpt", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "HVOP" Then
    'SAL_QRY = "INV->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and. INV->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')  .AND. Trim(CSR->DIS_TYP) = 'Periodic'"
'    HVO_Periodic_Rpt Format(DTPicker1, "MM/DD/YYYY"), Format(DTPicker2, "MM/DD/YYYY")
    'SAL_QRY = "INV->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and. INV->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "') .AND. (INV->STK_TYP = 'Salable') .AND. (INV->TXN_TYP = 'SAL') .and. CSR->DIS_TYP = 'Periodic' .AND. (CSR->B_BAS_AMT > 0)"
    SAL_QRY = SAL_QRY & " .AND. CSR->DIS_TYP = 'Periodic' .AND. INV->STK_TYP = 'Salable' .AND. INV->TOT_PCS > 0"
    '.AND. (CSR->B_BAS_AMT > 0)"
    OpenReport CURDIR & "\KOR.RP1", "HVO Periodic Rpt", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "FLRS" Then     'Fill Rate Detailed Report
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL'.OR.INL->TXN_TYP='SSO')"
    OpenReport CURDIR & "\ORD_LIB.RP1", "Fill Rate SKU Summary", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "FLRD" Then     'Fill Rate Summary Report
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL'.OR.INL->TXN_TYP='SSO')"
    OpenReport CURDIR & "\ORD_LIB.RP1", "Fill Rate SKU Details", 2, Trim(SAL_QRY), 6
    Unload Me
Else
    M_REP_IQY.Tag = M_REP_SQY.Tag
    Unload M_REP_SQY
    M_REP_IQY.Show 1
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
    If txtNoDays.Visible = True Then
        txtNoDays.SetFocus
    Else
        PTY_TYP.SetFocus
    End If
End If
If KeyCode = 38 And Shift = 1 And DTPicker1.Visible = True Then DTPicker1.SetFocus
End Sub

Private Sub DTPicker2_LostFocus()
If Me.Tag <> "DRL" And Me.Tag <> "CLG_DLT" Then
    If DTPicker1 > DTPicker2 Then
        MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
        DTPicker1.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Activate()
If Me.Tag = "dbt_agd" Then
    Frame6.Visible = True
    Frame4.Visible = False
    lblNoDays.Visible = True
    txtNoDays.Visible = True
ElseIf Me.Tag = "CLG_DLT" Or Me.Tag = "DRL" Then
    FDAT.Visible = False
    DTPicker1.Visible = False
    Label2.Visible = False
    Label1.Visible = True
End If
SETYEAROPEN M_REP_SQY
If PTY_TYP.ListCount > 0 Then PTY_TYP.Selected(0) = True
SSTab1.TabEnabled(1) = False
If Me.Tag = "Shell" Then
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    Frame3.Visible = False
    PTY_TYP.Visible = False
    Frame4.Visible = False
    Check7.SetFocus
    FDAT.Visible = False
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    Frame3.Visible = False
    Check7.SetFocus
ElseIf Me.Tag <> "DRL" And Me.Tag <> "CLG_DLT" Then
    DTPicker1.SetFocus
Else
    DTPicker1.Visible = False
'    Label1.Visible = False
    DTPicker2.SetFocus
End If
'Frame6.Visible = False
'If Me.Tag = "DPT" Then
'    Frame3.Visible = True
'End If
Check7.Value = 1
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
If Me.Tag = "CLG_DLT" Or Me.Tag = "DRL" Then
    DTPicker2.SetFocus
    Check8.Enabled = False
Else
Check8.SetFocus
End If
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

Private Sub Form_Load()
If PTY_TYP.ListCount > 0 Then PTY_TYP.Selected(0) = True
AddData List2, Me, "rou_idy", "rdy", "", "group by rou_idy"
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
            Command1.SetFocus
        ElseIf Label20 = "Select HVO OnSpot or Periodic" Then
            Label7 = list1.SelectedItem.ListSubItems.Item(1)
            Text6 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        End If
    End If
End If
End Sub


Private Sub List2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check6.SetFocus
End Sub

Private Sub PTY_TYP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If PTY_TYP.Selected(0) = True Then
        For M = 1 To PTY_TYP.ListCount - 1
            PTY_TYP.Selected(M) = False
        Next
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            Check1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(1) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            Check1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(2) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            Check1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(3) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            Check1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(4) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            Check1.SetFocus
        End If
    Else
        
    End If
End If
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
        MsgBox "Category Not Found, Define Category!", vbExclamation, "DAS Version 4.0"
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

Private Sub Text6_Change()
If Text6 = "" Then Label7 = ""
End Sub

Private Sub TEXT6_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If check("ACT", "ACT_IDY", Trim(Label7)) = True Then
        Command1.SetFocus
        Exit Sub
    End If
    Dim ls As ListItem
    Dim rs As New ADODB.Recordset
    If Me.Tag = "Shell" Then
        rs.Open "select * from act where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT' or act_typ like 'PUR')", con, adOpenStatic
    ElseIf Me.Tag = "HVOS" Then
        rs.Open "select * from act where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT') and (CST_TYP like 'HVO-NP' OR CST_TYP like 'HVO-NP-WO')", con, adOpenStatic
    ElseIf Me.Tag = "HVOP" Then
        rs.Open "select * from act where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT') and (CST_TYP like 'HVO-P'  OR CST_TYP like 'HVO-P-WO')", con, adOpenStatic
    Else
        rs.Open "select * from act where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT')", con, adOpenStatic
    End If
    list1.ListItems.CLEAR
    If rs.RecordCount > 0 Then
        For A = 1 To rs.RecordCount
            Set ls = list1.ListItems.Add(, , Trim(rs("ACT_NME")))
            ls.ListSubItems.Add , , Trim(rs("ACT_IDY") & "")
            ls.ListSubItems.Add , , Trim(rs("CST_TYP") & "")
            rs.MoveNext
        Next
        SSTab1.Tab = 1
        If Me.Tag = "HVOS" Or Me.Tag = "HVOP" Then
            Label20 = "Select HVO OnSpot or Periodic"
            list1.ColumnHeaders(1).Text = "HVO Name"
        Else
            Label20 = "Select Retailer"
            list1.ColumnHeaders(1).Text = "Retailer Name"
        End If
        list1.ColumnHeaders(2).Text = "ID"
        list1.ColumnHeaders(3).Text = "Type"
        list1.Visible = True
        list1.SetFocus
    Else
        If SSTab1.Tab = 0 Then
            MsgBox "Retailer - HVO Not Found, Define Retailer - HVO!", vbInformation, "DAS Version 4.0"
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
Public Sub clean()
End Sub
Private Sub txtNoDays_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check7.Enabled = True Then
        Check7.SetFocus
    Else
        Check1.SetFocus
    End If
End If
End Sub

Private Sub txtNoDays_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
    KeyAscii = 0
End If
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
        If DTPicker1.Visible = True Then
            DTPicker1.SetFocus
        Else
        If DTPicker2.Visible = True And DTPicker2.Enabled = True Then
            DTPicker2.SetFocus
        Else
            Check7.SetFocus
        End If
        End If
    End If
End If
End Sub
Private Sub cmbCaldates_Click()
'    cmbCaldates_KeyDown 13, 0
End Sub
Private Sub HVO_Periodic_Rpt(ST_DAT As Date, EN_DAT As Date)
On Error Resume Next
Dim M_BAS_AMT As Long
Dim CSR_RS As New ADODB.Recordset
Dim INV_RS As New ADODB.Recordset
Dim INL_RS As New ADODB.Recordset
Dim CPRD_RS As New ADODB.Recordset
If CSR_RS.State = 1 Then CSR_RS.Close
CSR_RS.Open "select B_PRD_PCS, B_PRD_SUT, B_PRD_RAT, B_BAS_AMT, ACT_IDY, B_PRD_IDY, DIS_AMT from CSR where CSR->DIS_TYP = 'Periodic'", con, adOpenKeyset, adLockPessimistic
If CSR_RS.RecordCount > 0 Then CSR_RS.MoveFirst
While Not CSR_RS.EOF
    CSR_RS("B_PRD_PCS") = 0
    CSR_RS("B_PRD_SUT") = 0
    CSR_RS("B_PRD_RAT") = 0
    CSR_RS("B_BAS_AMT") = 0
    CSR_RS.Update
    CSR_RS.MoveNext
Wend
If CSR_RS.State = 1 Then CSR_RS.Close
CSR_RS.Open "SELECT DISTINCT ACT_IDY FROM CSR", con, adOpenKeyset, adLockPessimistic
If CSR_RS.RecordCount > 0 Then CSR_RS.MoveFirst
While Not CSR_RS.EOF
    If INV_RS.State = 1 Then INV_RS.Close
    INV_RS.Open "SELECT * FROM INV, INL WHERE INV->INV_IDY = INL->INV_IDY .AND. INV->INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV->INV_DAT <= CTOD('" & EN_DAT & "') .AND. INV->TXN_TYP = 'SAL' .AND. INV->STK_TYP = 'Salable' .AND. INL->CST_IDY = '" & Trim(CSR_RS!ACT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
    While Not INV_RS.EOF
        If CPRD_RS.State = 1 Then CPRD_RS.Close
        CPRD_RS.Open "SELECT B_PRD_PCS, B_PRD_SUT, B_PRD_RAT, B_BAS_AMT, ACT_IDY, B_PRD_IDY, DIS_AMT FROM CSR WHERE CSR->ACT_IDY = '" & Trim(INV_RS!CST_IDY) & "' .AND. CSR->B_PRD_IDY = '" & Trim(INV_RS!prd_idy) & "'", con, adOpenKeyset, adLockOptimistic
        If CPRD_RS.EOF = False Then
            CPRD_RS!B_PRD_PCS = 0
            CPRD_RS!B_PRD_SUT = 0
            CPRD_RS!B_BAS_AMT = 0
            CPRD_RS.Update
        End If
        INV_RS.MoveNext
    Wend
    If INV_RS.State = 1 Then INV_RS.Close
    INV_RS.Open "SELECT * FROM INV, INL WHERE INV->INV_IDY = INL->INV_IDY .AND. INV->INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV->INV_DAT <= CTOD('" & EN_DAT & "') .AND. INV->TXN_TYP = 'SAL' .AND. INV->STK_TYP = 'Salable' .AND. INL->CST_IDY = '" & Trim(CSR_RS!ACT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
'    INV_RS.MoveFirst
    While Not INV_RS.EOF
        If CPRD_RS.State = 1 Then CPRD_RS.Close
        CPRD_RS.Open "SELECT B_PRD_PCS, B_PRD_SUT, B_PRD_RAT, B_BAS_AMT, ACT_IDY, B_PRD_IDY, DIS_AMT FROM CSR WHERE CSR->ACT_IDY = '" & Trim(INV_RS!CST_IDY) & "' .AND. CSR->B_PRD_IDY = '" & Trim(INV_RS!prd_idy) & "'", con, adOpenKeyset, adLockOptimistic
        If CPRD_RS.EOF = False Then
            CPRD_RS!B_PRD_PCS = Val(CPRD_RS!B_PRD_PCS) + Val(INV_RS!TOT_PCS)
            CPRD_RS!B_PRD_SUT = Val(CPRD_RS!B_PRD_SUT) + Val(INV_RS!SUB_UNT)
            M_BAS_AMT = (Val(INV_RS!TOT_PCS) * (Val(CPRD_RS!DIS_AMT) / Val(INV_RS!SUB_UNT)))
            CPRD_RS!B_BAS_AMT = Val(CPRD_RS!B_BAS_AMT) + M_BAS_AMT
            CPRD_RS.Update
        End If
        INV_RS.MoveNext
    Wend
    CSR_RS.MoveNext
Wend
Set CSR_RS = Nothing
Set CPRD_RS = Nothing
Set INV_RS = Nothing
End Sub

