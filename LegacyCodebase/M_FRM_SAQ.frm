VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SAQ 
   Caption         =   "Select Report Filters"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "M_FRM_SAQ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6750
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
      TabCaption(0)   =   "Report Control Filters"
      TabPicture(0)   =   "M_FRM_SAQ.frx":0442
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
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_SAQ.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   5280
         TabIndex        =   22
         Top             =   3480
         Width           =   1335
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1200
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   3375
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   4695
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1815
            Width           =   2655
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2775
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Location"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Ca&tegory"
            Height          =   375
            Left            =   360
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "&Market"
            Height          =   375
            Left            =   360
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "&Salesman"
            Height          =   375
            Left            =   360
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Ro&ute"
            Height          =   375
            Left            =   360
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "&Retailer"
            Height          =   375
            Left            =   360
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2760
            Width           =   855
         End
         Begin VB.ComboBox LIST2 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2280
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select Dates"
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   400
         Width           =   6375
         Begin VB.CheckBox Check8 
            Caption         =   "&FMCG Calendar"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   270
            Width           =   1455
         End
         Begin VB.ComboBox cmbCaldates 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   4920
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74186755
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Left            =   3180
            TabIndex        =   5
            Top             =   240
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74186755
            CurrentDate     =   38091
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   2760
            TabIndex        =   7
            Top             =   315
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   4560
            TabIndex        =   6
            Top             =   315
            Width           =   195
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   25
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
         TabIndex        =   31
         Top             =   3840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   30
         Top             =   2880
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   600
         Width           =   6225
      End
   End
End
Attribute VB_Name = "M_FRM_SAQ"
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

Private Sub Check6_keypress(K As Integer)
If K = 13 Then
If Check6.Value = 1 Then
    Check7.Value = 0
    Text6.Enabled = True
    Text6.SetFocus
Else
    Text6.Enabled = False
    Text6 = ""
    Command1.SetFocus
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
        Check7.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
Dim SLQRY
Command1.Enabled = False
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
    If Me.Tag = "TOP" Or Me.Tag = "DRL" Or Me.Tag = "NOP" Then
    SAL_QRY = ""
    Else
        If Me.Tag = "CFR" Or Me.Tag = "STR" Then
        SAL_QRY = "INL->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').AND.inl->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        ElseIf Me.Tag = "INW_REG" Then
        SLQRY = "avt->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.avt->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        ElseIf Me.Tag = "OUT_REG" Then
        SLQRY = "avt->sal_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.avt->sal_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        ElseIf Me.Tag = "AVT_RPT" Or Me.Tag = "CAF_REP" Or Me.Tag = "CPN_REP" Then
        SLQRY = "avt->avt_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.avt->avt_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        ElseIf Me.Tag = "CHN_REP" Then
        SLQRY = "avt->chr_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.avt->chr_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        Else
        SAL_QRY = "inv->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inv->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        SLQRY = "avt->sal_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.avt->sal_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
        End If
    End If
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'LOC' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text1) <> "" Then
        rs("DIR_ALS") = Trim(Label3) & ""
        rs("DIR_MSG") = Trim(Text1) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->LOC_IDY='" & Trim(Label3) & "'"
            SLQRY = SLQRY & ".AND.ACT->LOC_IDY='" & Trim(Label3) & "'"
        Else
            SLQRY = SLQRY & ".AND.ACT->LOC_IDY='" & Trim(Label3) & "'"
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
        SLQRY = SLQRY & ".AND.ACT->CAT_IDY='" & Trim(Label4) & "'"
        Else
        SLQRY = SLQRY & ".AND.ACT->CAT_IDY='" & Trim(Label4) & "'"
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
            SLQRY = SLQRY & ".AND.ACT->MAR_IDY='" & Trim(Label5) & "'"
        Else
            SLQRY = SLQRY & ".AND.ACT->MAR_IDY='" & Trim(Label5) & "'"
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
            If Me.Tag = "TOP" Then
                SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label6) & "'"
            ElseIf Me.Tag = "NOP" Then
                SAL_QRY = "ACT->SMN_IDY='" & Trim(Label6) & "'"
            Else
                SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label6) & "'"
            End If
            SLQRY = SLQRY & ".AND.AVT->SMN_IDY='" & Trim(Label6) & "'"
        Else
            SLQRY = SLQRY & ".AND.AVT->SMN_IDY='" & Trim(Label6) & "'"
            If Me.Tag = "TOP" Then
                SAL_QRY = "INL->SMN_IDY='" & Trim(Label6) & "'"
            ElseIf Me.Tag = "NOP" Then
                SAL_QRY = "ACT->SMN_IDY='" & Trim(Label6) & "'"
            Else
                SAL_QRY = "INL->SMN_IDY='" & Trim(Label6) & "'"
            End If
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
            If Me.Tag = "TOP" Then
                SAL_QRY = SAL_QRY & ".AND.INL->ROU_IDY='" & Trim(List2) & "'"
            ElseIf Me.Tag = "NOP" Then
                SAL_QRY = "ACT->ROU_IDY='" & Trim(List2) & "'"
            Else
                SAL_QRY = SAL_QRY & ".AND.INL->ROU_IDY='" & Trim(List2) & "'"
            End If
            SLQRY = SLQRY & ".AND.RDY->ROU_IDY='" & Trim(List2) & "'"
        Else
            SLQRY = SLQRY & ".AND.RDY->ROU_IDY='" & Trim(List2) & "'"
            If Me.Tag = "TOP" Then
                SAL_QRY = "INL->ROU_IDY='" & Trim(List2) & "'"
            ElseIf Me.Tag = "NOP" Then
                SAL_QRY = "ACT->ROU_IDY='" & Trim(List2) & "'"
            Else
                SAL_QRY = "INL->ROU_IDY='" & Trim(List2) & "'"
            End If
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
            SLQRY = SLQRY & ".AND.AVT->CST_IDY='" & Trim(Label7) & "'"
        Else
            SLQRY = SLQRY & ".AND.AVT->CST_IDY='" & Trim(Label7) & "'"
            SAL_QRY = "ACT->ACT_IDY='" & Trim(Label7) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
If Me.Tag = "SOT" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL')"
    OpenReport CURDIR & "\mir.RP1", "Secondary Offtake", 2, SAL_QRY, 6
    Unload M_REP_STR
ElseIf M_REP_STR.Tag = "DRL" Then
    Dim DAT As Date
    DAT = DateAdd("d", 1, DTPicker2)
    TRAIL_BALANCE Format(DAT, "MM/DD/YYYY")
    If Len(Trim(SAL_QRY)) = 0 Then
        SAL_QRY = "OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
    Else
        SAL_QRY = SAL_QRY & ".AND.OPA->DB_CR='Db'.AND.(ACT->GRP_CLF = 'Debtor'.OR.ACT->GRP_CLF = 'Creditor')"
    End If
    OpenReport CURDIR & "\Finance.RP1", "Debtors List", 2, Trim(SAL_QRY), 6
    Unload Me
ElseIf Me.Tag = "TOP" Then
    con.Execute "DELETE FROM TOP"
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "SELECT CST_IDY,SUM(GRS_AMT),SUM(NET_AMT) FROM INL WHERE INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV_DAT <= CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND TXN_TYP IN('LSL','SAL') GROUP BY CST_IDY", con, adOpenKeyset, adLockPessimistic
    While REC_SET.EOF = False
        con.Execute "INSERT INTO TOP VALUES('" & Trim(REC_SET(0) & "") & "'," & Val(REC_SET(1)) & "," & Val(REC_SET(2)) & ")"
        REC_SET.MoveNext
    Wend
    OpenReport CURDIR & "\MIR.RP1", "Top Outlets", 2, SAL_QRY, 6
    Unload M_REP_STR
ElseIf Me.Tag = "TOP_SKU" Then
    con.Execute "DELETE FROM TOP"
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "SELECT CST_IDY,SUM(PRD_QTY) FROM INV, INL WHERE INV.INV_IDY LIKE INL.INV_IDY AND INV.INV_DAT >= CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV.INV_DAT <= CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND INV.TXN_TYP IN('LSL','SAL') GROUP BY INL.CST_IDY", con, adOpenKeyset, adLockPessimistic
    While REC_SET.EOF = False
        con.Execute "INSERT INTO TOP VALUES('" & Trim(REC_SET(0) & "") & "'," & Val(REC_SET(1)) & ", NULL) "
        REC_SET.MoveNext
    Wend
    OpenReport CURDIR & "\MIR.RP1", "Top Outlets - SKU", 2, "ACT->CST_TYP <> 'Van'", 6
    Unload M_REP_STR
ElseIf Me.Tag = "SMR" Then
    Me.MousePointer = vbHourglass
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    OpenReport CURDIR & "\stock.RP1", "Stock Movement", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "NOP" Then
    Me.MousePointer = vbHourglass
    con.Execute "DELETE FROM NPO"
    If REC_SET1.State = 1 Then REC_SET1.Close
    REC_SET1.Open "SELECT ACT_IDY, CST_TYP, DBT_TYP FROM ACT WHERE ACT_TYP IN('BOT','SAL') AND MAR_IDY LIKE '" & Trim(Label5) & "%' AND LOC_IDY LIKE '" & Trim(Label3) & "%' AND CAT_IDY LIKE '" & Trim(Label4) & "%' AND SMN_IDY LIKE '" & Trim(Label6) & "%' AND DBT_TYP = 'Y' AND ROU_IDY LIKE '" & Trim(List2) & "%' AND ACT_IDY NOT IN(SELECT CST_IDY FROM INL WHERE INV_DAT >=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV_DAT <=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "'))", con, adOpenKeyset, adLockPessimistic
    While REC_SET1.EOF = False
        If Trim(REC_SET1(1)) <> "Van" Then
            If Trim(REC_SET1(2)) <> "N" Then
                con.Execute "INSERT INTO NPO VALUES('" & Trim(REC_SET1(0) & "") & "')"
            End If
        End If
        REC_SET1.MoveNext
    Wend
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\MIR.RP1", "Non Performing Outlets", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "VAN" Then
    OpenReport CURDIR & "\report.RP1", "Sales Report(VAN)", 2, Trim(SAL_QRY) & ".AND.(INV1->TXN_TYP='SAL')", 6
    Unload Me
ElseIf Me.Tag = "RAT" Then
On Error GoTo lab
    Me.MousePointer = vbHourglass
    con.Execute "DELETE FROM tpd"
    If REC_SET1.State = 1 Then REC_SET1.Close
    REC_SET1.Open "SELECT INV.PRD_IDY,INT(SUM(((PRD_QTY-INT(PRD_QTY))*1000)+(INT(PRD_QTY)*SUB_UNT))/SUB_UNT)+(MOD(SUM(((PRD_QTY-INT(PRD_QTY))*1000)+(INT(PRD_QTY)*SUB_UNT)),SUB_UNT)/1000) TOTAL,INV.PRD_RAT FROM INV,INL,PRD,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV.INV_IDY LIKE INL.INV_IDY  AND PRD.PRD_IDY LIKE INV.PRD_IDY AND INV.TXN_TYP IN('SAL') AND PRD.SAL_RAT <> INV.PRD_RAT AND ACT.MAR_IDY LIKE '" & Trim(Label5) & "%' AND ACT.LOC_IDY LIKE '" & Trim(Label3) & "%' AND ACT.CAT_IDY LIKE '" & Trim(Label4) & "%' AND INL.SMN_IDY LIKE '" & Trim(Label6) & "%' AND INL.ROU_IDY LIKE '" & Trim(List2) & "%' AND ACT.ACT_IDY LIKE '" & Trim(Label7) & "%' AND INV.INV_DAT >=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV.INV_DAT <=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') GROUP BY INV.PRD_IDY,INV.PRD_RAT", con, adOpenKeyset, adLockPessimistic
    While REC_SET1.EOF = False
        con.Execute "INSERT INTO tpd VALUES('" & Trim(REC_SET1(0) & "") & "'," & Val(REC_SET1(1) & "") & "," & Val(REC_SET1(2) & "") & ",0)"
        REC_SET1.MoveNext
    Wend
    Me.MousePointer = vbDefault
    OpenReport CURDIR & "\MIR.RP1", "Rate Changes Report", 2, "TPD->PRD_QTY > 0", 6
    Unload Me
Exit Sub
lab:
MsgBox "No records Found", vbExclamation
Unload Me
ElseIf Me.Tag = "CFR" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\MIR.RP1", "Coverage Frequency Report", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "STR" Then
    SAL_QRY = SAL_QRY & ".AND.(INL->TXN_TYP='LSL'.OR.INL->TXN_TYP='SAL')"
    OpenReport CURDIR & "\MIR.RP1", "Salesman Turnover Report", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "RAR" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    OpenReport CURDIR & "\MIR.RP1", "Retailers Addition Report", 2, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "INW_REG" Then
'    MsgBox SLQRY
    OpenReport CURDIR & "\TCM.RP1", "Mobile Inward Register", 0, Trim(SLQRY), 6
    Unload Me
'ElseIf Me.Tag = "MBL_STK" Then
'    Openreport CURDIR & "\tcm.RP1", "Mobile Stock on Hand", 2, "(AVT->SAL_IDY='')", 6
'    Unload Me
ElseIf Me.Tag = "OUT_REG" Then
    If Len(Trim(SLQRY)) > 0 Then
        SLQRY = SLQRY & ".AND.(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT'.OR.AVT->SAL_TYP='MRS')"
    Else
        SLQRY = SLQRY & "(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT'.OR.AVT->SAL_TYP='MRS')"
    End If
    OpenReport CURDIR & "\TCM.RP1", "Mobile Outward Register", 2, Trim(SLQRY), 6
    Unload Me
ElseIf Me.Tag = "AVT_RPT" Then
    If Len(Trim(SLQRY)) > 0 Then
        SLQRY = SLQRY & ".AND.(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT'.OR.AVT->SAL_TYP='MRS').AND.(AVT->AVT_STS='Yes')"
        
    Else
        SLQRY = SLQRY & "(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT'.OR.AVT->SAL_TYP='MRS').AND.(AVT->AVT_STS='Yes')"
    End If
    OpenReport CURDIR & "\TCM.RP1", "Activations Report", 2, Trim(SLQRY), 6
    Unload Me
ElseIf Me.Tag = "NAV_REP" Then
    If Len(Trim(SLQRY)) > 0 Then
        SLQRY = SLQRY & ".AND.(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT'.OR.AVT->SAL_TYP='MRS').AND.(AVT->AVT_STS<>'Yes')"
    Else
        SLQRY = SLQRY & "(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT'.OR.AVT->SAL_TYP='MRS').AND.(AVT->AVT_STS<>'Yes')"
    End If
    OpenReport CURDIR & "\TCM.RP1", "Mobiles Not Activated", 2, Trim(SLQRY), 6
    Unload Me
ElseIf Me.Tag = "CAF_REP" Then
    If Len(Trim(SLQRY)) > 0 Then
        SLQRY = SLQRY & ".AND.(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT').AND.AVT->AVT_STS='Yes'.AND.(AVT->CAF_STS='Yes')"
    Else
        SLQRY = SLQRY & "(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT').AND.AVT->AVT_STS='Yes'.AND.(AVT->CAF_STS='Yes')"
    End If
    OpenReport CURDIR & "\TCM.RP1", "CAF Received Report", 2, Trim(SLQRY), 6
    Unload Me
ElseIf Me.Tag = "CPN_REP" Then
    If Len(Trim(SLQRY)) > 0 Then
        SLQRY = SLQRY & ".AND.(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT').AND.AVT->AVT_STS='Yes'.AND.(AVT->CAF_STS<>'Yes')"
    Else
        SLQRY = SLQRY & "(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT').AND.AVT->AVT_STS='Yes'.AND.(AVT->CAF_STS<>'Yes')"
    End If
    OpenReport CURDIR & "\TCM.RP1", "CAF Pending Report", 2, Trim(SLQRY), 6
    Unload Me
ElseIf Me.Tag = "CHN_REP" Then
    If Len(Trim(SLQRY)) > 0 Then
        SLQRY = SLQRY & ".AND.(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT').AND.AVT->AVT_STS='Yes'.AND.(AVT->CHR_STS='Yes')"
    Else
        SLQRY = SLQRY & "(AVT->SAL_TYP='SAL'.OR.AVT->SAL_TYP='STO'.OR.AVT->SAL_TYP='PRT').AND.AVT->AVT_STS='Yes'.AND.(AVT->CHR_STS='Yes')"
    End If
    OpenReport CURDIR & "\TCM.RP1", "Churn Report", 2, Trim(SLQRY), 6
    Unload Me
ElseIf Me.Tag = "MBL_SUM" Then
    OpenReport CURDIR & "\TCM.RP1", "Mobile Summary Report", 2, Trim(SLQRY), 6
    Unload Me
Else
    M_REP_IQY.Tag = M_REP_STR.Tag
    Unload M_REP_STR
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
    If Check7.Enabled = True Then
        Check7.SetFocus
    Else
        Check1.SetFocus
    End If
End If
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
SETYEAROPEN M_FRM_SAQ
SSTab1.TabEnabled(1) = False

DTPicker1.MinDate = "01/01/1999"
DTPicker1.MaxDate = "12/31/2015"
DTPicker2.MinDate = "01/01/1999"
DTPicker2.MaxDate = "12/31/2015"

If Me.Tag <> "DRL" Then
    DTPicker1.SetFocus
Else
    DTPicker1.Visible = False
    Label1.Visible = False
    DTPicker2.SetFocus
End If
Check7.Value = 1
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check8.SetFocus
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
AddData List2, Me, "ROU_IDY", "RDY", "", "GROUP BY ROU_IDY"
DTPicker2.Value = Date
DTPicker1.Value = Date
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
        End If
    End If
End If
End Sub


Private Sub List2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check6.SetFocus
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

Public Sub clean()
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


