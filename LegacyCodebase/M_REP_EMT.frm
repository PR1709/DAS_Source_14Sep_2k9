VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_EMT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "M_REP_EMT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5609
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Report &Filters"
      TabPicture(0)   =   "M_REP_EMT.frx":0442
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
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DTPicker2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Check7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_EMT.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   2175
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   4095
         Begin VB.ComboBox list2 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   255
            Width           =   2295
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CheckBox Check8 
            Caption         =   "S&upplier"
            Height          =   375
            Left            =   480
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1680
            Width           =   975
         End
         Begin VB.CheckBox Check4 
            Caption         =   "&Salesman"
            Height          =   375
            Left            =   480
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "&Route"
            Height          =   375
            Left            =   480
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "R&etailer"
            Height          =   375
            Left            =   480
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   5400
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   840
         Value           =   1  'Checked
         Width           =   495
      End
      Begin MSComctlLib.ListView list1 
         Height          =   1935
         Left            =   -74760
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   3413
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
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1935
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1455
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Location"
         Height          =   375
         Left            =   960
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Ca&tegory"
         Height          =   375
         Left            =   960
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "&Market"
         Height          =   375
         Left            =   960
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   3600
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74186755
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1320
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74186755
         CurrentDate     =   37258
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   3240
         TabIndex        =   31
         Top             =   555
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   840
         TabIndex        =   30
         Top             =   555
         Width           =   345
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   7
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
         TabIndex        =   5
         Top             =   600
         Width           =   6225
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "M_REP_EMT"
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
        Check4.SetFocus
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
Else
    Text5.Enabled = False
    Text5 = ""
    Command1.SetFocus
End If
End If
End Sub

Private Sub Command1_Click()
SAL_QRY = ""
Command1.Enabled = False
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
'CODE CHANGED BY PKB FOR DATE FILTER
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
    If Me.Tag = "EMP_REG" Then
    SAL_QRY = "SNV->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "') .and. SNV->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "') "
    End If
End If

'END OF CODE FOR DATE FILTER
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'SMN' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text4) <> "" Then
        rs("DIR_ALS") = Trim(Label6) & ""
        rs("DIR_MSG") = Trim(Text4) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND. SNV->SMN_IDY='" & Trim(Label6) & "'"
        Else
            SAL_QRY = "SNV->SMN_IDY='" & Trim(Label6) & "'"
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
            SAL_QRY = SAL_QRY & ".AND.SNV->ROU_IDY='" & Trim(List2) & "'"
        Else
            SAL_QRY = "SNV->ROU_IDY='" & Trim(List2) & "'"
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
            SAL_QRY = SAL_QRY & ".AND.SNV->CST_IDY='" & Trim(Label7) & "'"
        Else
            SAL_QRY = "SNV->CST_IDY='" & Trim(Label7) & "'"
        End If
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
rs.Update

If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'SUP' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(Text5) <> "" Then
        Text6 = ""
        Check6.Value = False
        rs("DIR_ALS") = Trim(Label1) & ""
        rs("DIR_MSG") = Trim(Text5) & ""
        SAL_QRY = "SNV->CST_IDY='" & Trim(Label1) & "'"
    Else
        rs("DIR_ALS") = "*"
        rs("DIR_MSG") = "All"
    End If
    rs.Update
End If
rs.Update

If Len(Trim(Text5)) = 0 And Me.Tag <> "EMP_REG" Then
    If Trim(Len(SAL_QRY)) <> 0 Then
        SAL_QRY = SAL_QRY & ".AND.(ACT->ACT_TYP='SAL'.OR.ACT->ACT_TYP='BOT').AND.(ACT->CST_TYP <> 'Van')"
    Else
  '      SAL_QRY = "(ACT->ACT_TYP='SAL'.OR.ACT->ACT_TYP='BOT').AND.(ACT->CST_TYP <> 'Van')"
        SAL_QRY = SAL_QRY & "(ACT->ACT_TYP='PUR'.OR.ACT->ACT_TYP='SAL'.OR.ACT->ACT_TYP='BOT').AND.(ACT->CST_TYP <> 'Van')"
        'Code changed after PC feedback by PNR
    End If
End If


'If Me.Tag = "SOH" Then
'Openreport CURDIR & "\SNN.RP1", "Shells SOH", 2, "", 6
' 'Openreport CURDIR & "\SNN.RP1", "Shells SOH", 2, SAL_QRY, 6
'ElseIf Me.Tag = "ROLCOL" Then
''Openreport CURDIR & "\SNN.RP1", "Shell ROL COL", 2, "(ACT->CST_TYP <> 'Van')", 6
'Openreport CURDIR & "\SNN.RP1", "Shell ROL COL", 2, SAL_QRY, 6
If Me.Tag = "PRC" Then
    OpenReport CURDIR & "\SNN.RP1", "Party ROL COL", 2, SAL_QRY, 6 '"(ACT->CST_TYP <> 'Van')", 6
ElseIf Me.Tag = "PRCA" Then
    'Openreport CURDIR & "\SNN.RP1", "Party ROL COL Account", 2, "(ACT->CST_TYP <> 'Van')", 6 'MODIFIED BY Parvathi
    OpenReport CURDIR & "\SNN.RP1", "Party ROL COL Account", 2, SAL_QRY, 6 '"(ACT->CST_TYP <> 'Van')", 6
ElseIf Me.Tag = "SDA" Then
    'Openreport CURDIR & "\SNN.RP1", "Deposits Account", 2, "(ACT->CST_TYP <> 'Van')", 6
    'Openreport CURDIR & "\SNN.RP1", "Deposits Account", 2, "(ACT->CST_TYP <> 'Van')", 6  'MODIFIED BY Parvathi
    OpenReport CURDIR & "\SNN.RP1", "Deposits Account", 2, SAL_QRY, 6 '"(ACT->CST_TYP <> 'Van')", 6
ElseIf Me.Tag = "EMP_REG" Then
   ' Openreport CURDIR & "\SNN.RP1", "Empties Register", 2, "SNV->INV_DAT >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "') .and. SNV->INV_DAT <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "') ", 6 'CHANGED BY PKB
    'Openreport CURDIR & "\SNN.RP1", "Empties Register", 2, SAL_QRY, "(ACT->CST_TYP <> 'Van')" , 6
    OpenReport CURDIR & "\SNN.RP1", "Empties Register", 2, SAL_QRY, 6
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

Private Sub DTPicker2_LostFocus()
    If DTPicker2 < DTPicker1 Then
        MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
        DTPicker1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
SETYEAROPEN M_REP_EMT
SSTab1.TabEnabled(1) = False
Check7.Value = 1
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
If Me.Tag = "EMP_REG" Then
    Label2.Visible = True
    Label8.Visible = True
    DTPicker1.Visible = True
    DTPicker2.Visible = True
    DTPicker1.SetFocus
Else
    Check7.SetFocus
    Label2.Visible = False
    Label8.Visible = False
    DTPicker1.Visible = False
    DTPicker2.Visible = False
End If
End Sub
Private Sub Form_Load()
If Me.Tag = "EMP_REG" Then
DTPicker1.Value = Date
DTPicker2.Value = Date
End If
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
        ElseIf Label20 = "Select Supplier" Then
            Label1 = list1.SelectedItem.ListSubItems.Item(1)
            Text5 = Trim(list1.SelectedItem)
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
If rs.State = 1 Then rs.Close
rs.Open "select * from smn where smn_nme like '" & Trim(Text4) & "%'", con, adOpenStatic
If Not rs.RecordCount > 0 Then
If rs.State = 1 Then rs.Close
While Len(Trim(Text4)) < 10 And Len(Trim(Text4)) > 0
Text4 = "0" & Text4
Wend
rs.Open "select * from smn where smn_IDY like '" & Trim(Text4) & "%'", con, adOpenStatic
Text4 = ""
If rs.RecordCount > 0 Then
Text4 = rs("SMN_NME")
Label6 = rs("SMN_IDY")
Check5.SetFocus
Exit Sub
End If
End If
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

'If K = 13 Then
'Dim ls As ListItem
'Dim rs As New ADODB.Recordset
'If check("SMN", "SMN_IDY", Trim(Label6)) = True Then
'Check5.SetFocus
'Exit Sub
'End If
'rs.Open "select * from smn where smn_nme like '" & Trim(Text4) & "%'", CON, adOpenStatic
'list1.ListItems.CLEAR
'If rs.RecordCount > 0 Then
'    For A = 1 To rs.RecordCount
'        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
'        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
'        rs.MoveNext
'    Next
'    SSTab1.Tab = 1
'    Label20 = "Select Salesman"
'    list1.ColumnHeaders(1).Text = "Salesman Name"
'    list1.ColumnHeaders(2).Text = "ID"
'    list1.Visible = True
'    list1.SetFocus
'Else
'    If SSTab1.Tab = 0 Then
'        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
'        SSTab1.Tab = 0
'        Text4.Enabled = True
'        Text4.SetFocus
'    End If
'End If
'End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text5_Change()
If Text5 = "" Then Label1 = ""
End Sub

Private Sub TEXT5_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
If check("ACT", "ACT_IDY", Trim(Label1)) = True Then
Command1.SetFocus
Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select * from act where act_nme like '" & Trim(Text5) & "%' and (act_typ like 'PUR')", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy") & "")
        rs.MoveNext
    Next
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
        Text5.Enabled = True
        Text5.SetFocus
    End If
End If
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
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
