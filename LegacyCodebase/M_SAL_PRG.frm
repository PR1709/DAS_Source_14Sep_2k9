VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SAL_PRG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Growth Report"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "M_SAL_PRG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Details >>"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7800
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Report Filters"
      TabPicture(0)   =   "M_SAL_PRG.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ProgressBar1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_SAL_PRG.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   6240
         TabIndex        =   41
         Top             =   1320
         Width           =   2535
         Begin VB.ListBox PTY_TYP 
            Height          =   285
            ItemData        =   "M_SAL_PRG.frx":047A
            Left            =   120
            List            =   "M_SAL_PRG.frx":0493
            Style           =   1  'Checkbox
            TabIndex        =   42
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Party Type"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Select Dates"
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   6015
         Begin VB.CheckBox Check12 
            Caption         =   "&FMCG Calendar"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   280
            Width           =   1455
         End
         Begin VB.ComboBox cmbCaldates 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   250
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   4560
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   255
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
            Left            =   2880
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   255
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74317827
            CurrentDate     =   37258
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   4230
            TabIndex        =   40
            Top             =   318
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   2400
            TabIndex        =   39
            Top             =   318
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   6015
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2760
            MaxLength       =   40
            TabIndex        =   31
            Top             =   250
            Width           =   2895
         End
         Begin VB.CheckBox Check11 
            Caption         =   "All S&uppliers"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   228
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Supplier"
            Height          =   195
            Left            =   1920
            TabIndex        =   33
            Top             =   318
            Width           =   570
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5640
            TabIndex        =   32
            Top             =   250
            Visible         =   0   'False
            Width           =   165
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   8775
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5640
            MaxLength       =   40
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5640
            MaxLength       =   40
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5640
            MaxLength       =   40
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   960
            Width           =   2655
         End
         Begin VB.OptionButton optPks 
            Caption         =   "Pack &Size"
            Height          =   255
            Left            =   4440
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1005
            Width           =   1095
         End
         Begin VB.OptionButton optPkt 
            Caption         =   "Pack &Type"
            Height          =   255
            Left            =   4440
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   645
            Width           =   1095
         End
         Begin VB.OptionButton optBrd 
            Caption         =   "&Brand"
            Height          =   255
            Left            =   4440
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   285
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   960
            Width           =   2655
         End
         Begin VB.OptionButton optPrd 
            Caption         =   "&Product"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1005
            Width           =   855
         End
         Begin VB.OptionButton optSeg 
            Caption         =   "Se&gment"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   645
            Width           =   1095
         End
         Begin VB.OptionButton optCat 
            Caption         =   "Categor&y"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   4680
            TabIndex        =   14
            Top             =   1560
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All Products"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   650
         Left            =   6720
         TabIndex        =   2
         Top             =   4200
         Width           =   2175
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   180
            Width           =   855
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   6840
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   8385
         _ExtentX        =   14790
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
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
         TabIndex        =   21
         Top             =   960
         Width           =   8385
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   5295
         TabIndex        =   20
         Top             =   7320
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Processed"
         Height          =   195
         Left            =   4920
         TabIndex        =   19
         Top             =   6960
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5880
         TabIndex        =   18
         Top             =   7920
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   7080
         Width           =   45
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   855
         Left            =   4680
         Top             =   7560
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Processing Report"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   7320
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   3720
         Visible         =   0   'False
         Width           =   2295
      End
   End
End
Attribute VB_Name = "M_SAL_PRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSWEK As New ADODB.Recordset
Private Sub Check11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check11.Value = 1 Then
        PTY_TYP.Enabled = True
        PTY_TYP.SetFocus
    Else
        Text9.SetFocus
    End If
End If
End Sub

Private Sub Check12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check12.Value = 1 Then
        If RSWEK.RecordCount > 0 Then
            cmbCaldates.Enabled = True
            cmbCaldates.SetFocus
        End If
    Else
        DTPicker1.Enabled = True
        DTPicker1.SetFocus
    End If
End If
End Sub

Private Sub Check7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check7.Value = 1 Then
        Command1.SetFocus
        optSeg.Value = False
        optCat.Value = False
        optPrd.Value = False
        optBrd.Value = False
        optPkt.Value = False
        optPks.Value = False
        Label7 = ""
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        con.Execute "UPDATE DIR SET DIR_MSG='All' where DIR_IDY='PRD' AND DIR_IDX='IDY'"
        con.Execute "UPDATE DIR SET DIR_MSG='All' where DIR_IDY='SEG' AND DIR_IDX='IDY'"
        con.Execute "UPDATE DIR SET DIR_MSG='All' where DIR_IDY='PGR' AND DIR_IDX='IDY'"
        con.Execute "UPDATE DIR SET DIR_MSG='All' where DIR_IDY='BRD' AND DIR_IDX='IDY'"
        con.Execute "UPDATE DIR SET DIR_MSG='All' where DIR_IDY='PKT' AND DIR_IDX='IDY'"
        con.Execute "UPDATE DIR SET DIR_MSG='All' where DIR_IDY='PKS' AND DIR_IDX='IDY'"
    Else
        Frame2.Enabled = True
        optCat.SetFocus
    End If
End If
End Sub

Private Sub cmbCaldates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If RSWEK.State = 1 Then RSWEK.Close
    RSWEK.Open "SELECT FRM_DAT,TO_DAT FROM WEK WHERE WEK_IDY=" & Val(cmbCaldates.Text), con, adOpenKeyset, adLockPessimistic
    If Not RSWEK.EOF Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
        DTPicker1.Value = RSWEK("FRM_DAT")
        DTPicker2.Value = RSWEK("TO_DAT")
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
        Check11.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
Command1.Enabled = False
'Command2.Enabled = False
'dates
rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    rs("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    rs.Update
End If
con.Execute "UPDATE DIR SET DIR_MSG='All' WHERE DIR_IDY LIKE 'PGR' AND DIR_IDX LIKE 'IDY'"
con.Execute "UPDATE DIR SET DIR_MSG='All' WHERE DIR_IDY LIKE 'PRD' AND DIR_IDX LIKE 'IDY'"
con.Execute "UPDATE DIR SET DIR_MSG='All' WHERE DIR_IDY LIKE 'SEG' AND DIR_IDX LIKE 'IDY'"

con.Execute "UPDATE DIR SET DIR_MSG='All' WHERE DIR_IDY LIKE 'BRD' AND DIR_IDX LIKE 'IDY'"
con.Execute "UPDATE DIR SET DIR_MSG='All' WHERE DIR_IDY LIKE 'PKT' AND DIR_IDX LIKE 'IDY'"
con.Execute "UPDATE DIR SET DIR_MSG='All' WHERE DIR_IDY LIKE 'PKS' AND DIR_IDX LIKE 'IDY'"

'Products
SAL_QRY = "inV->inv_dat >= ctod('" & Format(DTPicker1.Value, "MM/DD/YYYY") & "').and.inv->inv_dat <= ctod('" & Format(DTPicker2.Value, "MM/DD/YYYY") & "').AND.INV->TXN_TYP='SAL'.AND.(INV->STK_TYP='Salable'.OR.INV->STK_TYP='Free')"
If Check11.Value = 0 Then
    If Len(Trim(Label4)) <> 0 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'SUP' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Label4)) > 0 Then
            rs("DIR_ALS") = Trim(Label4) & ""
            rs("DIR_MSG") = Trim(Text9) & ""
            rs.Update
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.prd->cmp_idy='" & Trim(Label4) & "'"
            Else
                SAL_QRY = "prd->cmp_IDY='" & Trim(Label4) & "'"
            End If
        End If
    End If
    End If
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'SUP' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
    rs("DIR_ALS") = "*"
    rs("DIR_MSG") = "All"
    rs.Update
    End If
End If
If optPrd.Value = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'PRD' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Text2)) > 0 Then
            rs("DIR_MSG") = Trim(Text2) & ""
        Else
            rs("DIR_MSG") = "All"
        End If
        rs.Update
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.INV->prd_idy='" & Trim(Label7) & "'"
        End If
    End If
ElseIf optSeg.Value = True Then
'Segment
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'SEG' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Text4)) > 0 Then
            rs("DIR_MSG") = Trim(Text4) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.PGR->GRP_CLS='" & Trim(Text4) & "'"
            End If
        Else
            rs("dir_msg") = "All"
        End If
        rs.Update
    End If
ElseIf optCat.Value = True Then
'Category
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir where dir_idy like 'PGR' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(Trim(Text3)) > 0 Then
            rs("DIR_MSG") = Trim(Text3) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.pgr->grp_cat='" & Trim(Text3) & "'"
            End If
        Else
            rs("dir_msg") = "All"
        End If
        rs.Update
    End If
ElseIf optBrd.Value = True Then
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'BRD' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
        If Len(Trim(Text5)) > 0 Then
            rs("DIR_MSG") = Trim(Text5) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.PRD->BRD_NME='" & Trim(Text5) & "'"
            End If
        Else
            rs("dir_msg") = "All"
        End If
    rs.Update
End If
ElseIf optPkt.Value = True Then
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'PKT' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
        If Len(Trim(Text6)) > 0 Then
            rs("DIR_MSG") = Trim(Text6) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.PRD->PRD_PKT='" & Trim(Text6) & "'"
            End If
        Else
            rs("dir_msg") = "All"
        End If
    rs.Update
End If
ElseIf optPks.Value = True Then
If rs.State = 1 Then rs.Close
rs.Open "select * from dir where dir_idy like 'PKS' AND dir_idx like 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
        If Len(Trim(Text1)) > 0 Then
                rs("DIR_MSG") = Trim(Text1) & ""
                If Trim(Len(SAL_QRY)) <> 0 Then
                    SAL_QRY = SAL_QRY & ".AND.PRD->PRD_PCK='" & Trim(Text1) & "'"
                End If
        Else
            rs("dir_msg") = "All"
        End If
    rs.Update
End If
End If
' TO SOLVE THE ERROR OF SALES GROWTH REPORT IDENTIFIED BY MR SRINIVASA AND SOLVED BY PROMODA BY REPLACE 4 IN PLACE OF 3 IN FOR LOOP
If PTY_TYP.Selected(0) = False Then
    For A = 1 To 3
        If PTY_TYP.Selected(A) = True Then
            If SALQRY = "" Then
                SALQRY = "ACT->cst_typ='" & PTY_TYP.List(A) & "'"
            Else
                SALQRY = SALQRY & ".or.ACT->cst_typ='" & PTY_TYP.List(A) & "'"
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
If Trim(sprg) = "Category" Then
    If Month(DTPicker1.Value) >= 4 And Month(DTPicker1.Value) <= 9 Then
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Channel SP H1", 2, Trim(SAL_QRY), 6
    Else
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Channel SP H2", 2, Trim(SAL_QRY), 6
    End If
ElseIf Trim(sprg) = "Location" Then
    If Month(DTPicker1.Value) >= 4 And Month(DTPicker1.Value) <= 9 Then
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Location SP H1", 2, Trim(SAL_QRY), 6
    Else
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Location SP H2", 2, Trim(SAL_QRY), 6
    End If
ElseIf Trim(sprg) = "Market" Then
    If Month(DTPicker1.Value) >= 4 And Month(DTPicker1.Value) <= 9 Then
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Market SP H1", 2, Trim(SAL_QRY), 6
    Else
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Market SP H2", 2, Trim(SAL_QRY), 6
    End If
ElseIf Trim(sprg) = "Salesman" Then
    If Month(DTPicker1.Value) >= 4 And Month(DTPicker1.Value) <= 9 Then
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Salesman SP H1", 2, Trim(SAL_QRY), 6
    Else
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Salesman SP H2", 2, Trim(SAL_QRY), 6
    End If
ElseIf Trim(sprg) = "Route" Then
    If Month(DTPicker1.Value) >= 4 And Month(DTPicker1.Value) <= 9 Then
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Route SP H1", 2, Trim(SAL_QRY), 6
    Else
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Route SP H2", 2, Trim(SAL_QRY), 6
    End If
ElseIf Trim(sprg) = "Retailer" Then
    If Month(DTPicker1.Value) >= 4 And Month(DTPicker1.Value) <= 9 Then
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Retailer SP H1", 2, Trim(SAL_QRY), 6
    Else
        OpenReport CURDIR & "\Sal_prg.RP1", "Retailer Retailer SP H2", 2, Trim(SAL_QRY), 6
    End If
End If
Command1.Enabled = True
Check12.SetFocus
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPicker2.SetFocus
End If
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If DTPicker2.Value >= DTPicker1.Value Then
        Check11.SetFocus
    Else
        MsgBox "Invalid Date!", vbExclamation, "DAS Version 4.0"
        DTPicker2.SetFocus
    End If
End If
End Sub

Private Sub Form_Activate()
Check12.SetFocus
If RSWEK.State = 1 Then RSWEK.Close
RSWEK.Open "SELECT WEK_IDY FROM WEK", con, adOpenKeyset, adLockPessimistic
While Not RSWEK.EOF
    cmbCaldates.AddItem RSWEK("WEK_IDY")
    RSWEK.MoveNext
Wend
PTY_TYP.Selected(0) = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
SETYEAROPEN Me
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
            Label7 = list1.SelectedItem.ListSubItems.Item(1)
            Text2 = Trim(list1.SelectedItem)
            Text3 = ""
            Text4 = ""
            Check7.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Brand" Then
            'Label5 = list1.SelectedItem
            Text5 = Trim(list1.SelectedItem)
            Text6 = ""
            Text1 = ""
            Text2 = ""
            Text3 = ""
            Text4 = ""
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Pack Type" Then
           'Label5 = list1.SelectedItem
            Text6 = Trim(list1.SelectedItem)
            Text5 = ""
            Text1 = ""
            Text2 = ""
            Text3 = ""
            Text4 = ""
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Pack Size" Then
            'Label5 = list1.SelectedItem
            Text1 = Trim(list1.SelectedItem)
            Text5 = ""
            Text6 = ""
            Text2 = ""
            Text3 = ""
            Text4 = ""
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Salesman" Then
            Label7 = list1.SelectedItem.ListSubItems.Item(1)
            Text5 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check10.SetFocus
        ElseIf Label20 = "Select Retailer" Then
            Label9 = list1.SelectedItem.ListSubItems.Item(1)
            Text6 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Supplier" Then
            Label4 = list1.SelectedItem.ListSubItems.Item(1)
            Text9 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            Check11.Value = False
            list1.Visible = False
            PTY_TYP.SetFocus
        ElseIf Label20 = "Select Retailers Location" Then
            Label10 = list1.SelectedItem.ListSubItems.Item(1)
            Text8 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check5.SetFocus
        ElseIf Label20 = "Select Retailers Category" Then
            Label8 = list1.SelectedItem.ListSubItems.Item(1)
            Text1 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check9.SetFocus
        ElseIf Label20 = "Select Market" Then
            Label11 = list1.SelectedItem.ListSubItems.Item(1)
            Text7 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Check4.SetFocus
        End If
    End If
End If
End Sub


Private Sub optBrd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text5.Enabled = True
    Text5.SetFocus
End If
End Sub

Private Sub optCat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text3.Enabled = True
    Text3.SetFocus
End If
End Sub

Private Sub optPks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub optPkt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text6.Enabled = True
    Text6.SetFocus
End If
End Sub

Private Sub optPrd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text2.Enabled = True
    Text2.SetFocus
End If
End Sub

Private Sub optSeg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text4.Enabled = True
    Text4.SetFocus
End If
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
'            CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(1) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
'            CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(2) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
 '           CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(3) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
  '          CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(4) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
   '         CHECK1.SetFocus
        End If
    Else
        
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
rs.Open "select * from prd,pgr where prd_nme like '" & Trim(Text2) & "%' and prd.grp_idy like pgr.grp_idy and cmp_idy like '" & Trim(Label4) & "%'", con, adOpenStatic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        'list1.ColumnHeaders.Clear
        If Len(Trim(rs("PRD_NME"))) > 0 Then
            Set ls = list1.ListItems.Add(, , Trim(rs("prd_nme")))
            ls.ListSubItems.Add , , Trim(rs("prd_idy") & "")
            ls.ListSubItems.Add , , Trim(rs("grp_cls") & "")
        End If
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Product"
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Product Name"
    list1.ColumnHeaders(1).Width = 2000
    list1.ColumnHeaders(2).Text = "ID"
    list1.ColumnHeaders(3).Text = "Product Segment"
    list1.ColumnHeaders(3).Width = 2000
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
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select * from pgr where GRP_cat like '" & Trim(Text3) & "%' GROUP BY GRP_cat", con, adOpenStatic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
If rs.RecordCount > 0 Then
    Label20 = "Select Product Category"
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , , list1.Width - 100
    list1.ColumnHeaders(1).Text = "Product Category"
    For A = 1 To rs.RecordCount
        If Len(Trim(rs("GRP_CAT"))) > 0 Then Set ls = list1.ListItems.Add(, , Trim(rs("grp_cat")))
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
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text4_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select * from pgr where grp_cls like '" & Trim(Text4) & "%' group by grp_cls", con, adOpenStatic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
If rs.RecordCount > 0 Then
    Label20 = "Select Product Segment"
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , , list1.Width - 100
    list1.ColumnHeaders(1).Text = "Product Segment"
    For A = 1 To rs.RecordCount
        If Len(Trim(rs("GRP_CLS"))) > 0 Then list1.ListItems.Add , , Trim(rs("grp_cls"))
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
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text9_Change()
If Text9 = "" Then Label4 = ""
End Sub

Private Sub Text9_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("act", "act_idy", Trim(Label4)) = True Then
    Command1.SetFocus
Else
    rs.Open "select * from act  where act_nme like '" & Trim(Text9) & "%' and (act_typ like 'PUR' or act_typ like 'BOT')", con, adOpenStatic
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
    list1.ColumnHeaders(1).Width = 3500
    list1.ColumnHeaders(1).Text = "Supplier Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text9.Enabled = True
        Text9.SetFocus
    End If
End If
End If
End If
End Sub
'Brand
Private Sub TEXT5_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If check("prd", "prd_idy", Trim(Label5)) = True Then
Check4.SetFocus
Else
rs.Open "select brd_nme from prd where brd_nme like '" & Trim(Text2) & "%' group by brd_nme", con, adOpenStatic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
If rs.RecordCount > 0 Then
    SSTab1.Tab = 1
    Label20 = "Select Brand"
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Brand Name"
    list1.ColumnHeaders(1).Width = 2000
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        'list1.ColumnHeaders.Clear
        If Len(Trim(rs("BRD_NME"))) > 0 Then Set ls = list1.ListItems.Add(, , Trim(rs("brd_nme")))
        rs.MoveNext
    Next
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "SKU Brand Not Found, Define SKU Brand!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text5.Enabled = True
        Text5.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub
'Pack Type
Private Sub TEXT6_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select prd_idy,prd_pkt from prd where prd_pkt like '" & Trim(Text6) & "%' group by prd_pkt", con, adOpenStatic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
If rs.RecordCount > 0 Then
    SSTab1.Tab = 1
    Label20 = "Select Pack Type"
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Pack Type"
    list1.ColumnHeaders(1).Width = 2000
    For A = 1 To rs.RecordCount
        'list1.ColumnHeaders.Clear
        If Len(Trim(rs("prd_pkt"))) > 0 Then Set ls = list1.ListItems.Add(, , Trim(rs("prd_pkt")))
        rs.MoveNext
    Next
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Pack Type Not Found, Define Pack!", vbExclamation, "DAS Version 4.0"
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
'Pack Size
Private Sub TEXT1_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select prd_idy,prd_pck from prd where prd_pck like '" & Trim(Text1) & "%' group by prd_pck", con, adOpenStatic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
If rs.RecordCount > 0 Then
    SSTab1.Tab = 1
    Label20 = "Select Pack Size"
    list1.ColumnHeaders.Add
    list1.ColumnHeaders(1).Text = "Pack Size"
    list1.ColumnHeaders(1).Width = 2000
    For A = 1 To rs.RecordCount
        'list1.ColumnHeaders.Clear
        If Len(Trim(rs("prd_pck"))) > 0 Then Set ls = list1.ListItems.Add(, , Trim(rs("prd_pck")))
        rs.MoveNext
    Next
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Pack Size Not Found, Define Pack!", vbExclamation, "DAS Version 4.0"
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

