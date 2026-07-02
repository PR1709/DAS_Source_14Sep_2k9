VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_PQY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   ClipControls    =   0   'False
   Icon            =   "M_REP_PRG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Details >>"
      Height          =   375
      Left            =   0
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   7800
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
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
      TabPicture(0)   =   "M_REP_PRG.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Shape1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label15"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label16"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPicker1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPicker2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ProgressBar1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Check11"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text9"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check7"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "PTY_TYP"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbCaldates"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Check12"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_PRG.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox Check12 
         Caption         =   "&FMCG Calendar"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   460
         Width           =   1455
      End
      Begin VB.ComboBox cmbCaldates 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   54
         Top             =   430
         Width           =   615
      End
      Begin VB.ListBox PTY_TYP 
         Height          =   735
         ItemData        =   "M_REP_PRG.frx":047A
         Left            =   1560
         List            =   "M_REP_PRG.frx":048D
         Style           =   1  'Checkbox
         TabIndex        =   52
         Top             =   840
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Caption         =   "Scheme Report Options"
         Height          =   1455
         Left            =   5040
         TabIndex        =   46
         Top             =   1560
         Visible         =   0   'False
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "&2 Value - Consolidated"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&3 SKU - Consolidated"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&1 Invoice - Detailed"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   7680
         TabIndex        =   34
         Top             =   3600
         Width           =   1095
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   7335
         Begin VB.ComboBox LIST2 
            Height          =   315
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Rou&te ID"
            Height          =   195
            Left            =   4320
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   773
            Width           =   1095
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5400
            MaxLength       =   40
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1155
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5400
            MaxLength       =   40
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox Check6 
            Caption         =   "&Retailer"
            Height          =   195
            Left            =   4320
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1223
            Width           =   855
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Salesma&n"
            Height          =   195
            Left            =   4320
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   308
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Retailers Cate&gory"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   773
            Width           =   1695
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Retailer &Location"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   308
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   705
            Width           =   1695
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1155
            Width           =   1695
         End
         Begin VB.CheckBox Check9 
            Caption         =   "&Market"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1223
            Width           =   855
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7200
            TabIndex        =   43
            Top             =   1155
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7200
            TabIndex        =   42
            Top             =   240
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3600
            TabIndex        =   33
            Top             =   705
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3600
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3600
            TabIndex        =   31
            Top             =   1155
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All Products"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4560
         MaxLength       =   40
         TabIndex        =   16
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox Check11 
         Caption         =   "All S&uppliers"
         Height          =   375
         Left            =   6310
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   6840
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   4815
         Begin VB.OptionButton CHECK3 
            Caption         =   "Categor&y"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   285
            Width           =   1095
         End
         Begin VB.OptionButton CHECK2 
            Caption         =   "&Segment"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   645
            Width           =   1095
         End
         Begin VB.OptionButton CHECK1 
            Caption         =   "&Product"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1005
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   4680
            TabIndex        =   22
            Top             =   960
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   4560
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   405
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24903683
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   2880
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   405
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   24903683
         CurrentDate     =   37258
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Type"
         Height          =   195
         Left            =   525
         TabIndex        =   53
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Supplier"
         Height          =   195
         Left            =   3720
         TabIndex        =   45
         Top             =   885
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Processing Report"
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Top             =   7320
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   855
         Left            =   4680
         Top             =   7560
         Width           =   2775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5880
         TabIndex        =   21
         Top             =   7080
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5880
         TabIndex        =   20
         Top             =   7920
         Width           =   45
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   5295
         TabIndex        =   18
         Top             =   7320
         Width           =   360
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7560
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   4230
         TabIndex        =   6
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   345
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
         TabIndex        =   4
         Top             =   600
         Width           =   8385
      End
   End
End
Attribute VB_Name = "M_REP_PQY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rswek As New ADODB.Recordset
Dim CTR As Long
Dim COMCTR As Long
Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then
    Text2.Enabled = True
    Text2.SetFocus
End If
End Sub

Private Sub Check10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check10.Value = 1 Then
        LIST2.Enabled = True
        LIST2.SetFocus
    Else
        Check6.SetFocus
    End If
End If
End Sub

Private Sub Check11_Click()
Check11_KeyPress 13
End Sub

Private Sub Check11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check11.Value = 0 Then
        Check11.Enabled = False
        Text9.SetFocus
    Else
        Text9 = ""
        PTY_TYP.SetFocus
    End If
End If
End Sub

Private Sub CHECK2_KeyPress(K As Integer)
If K = 13 Then
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

Private Sub Check4_keypress(K As Integer)
If K = 13 Then
    If Check4.Value = 1 Then
        Text5.Enabled = True
        Text5.SetFocus
    Else
        Check10.SetFocus
    End If
End If
End Sub
Private Sub Check5_keypress(K As Integer)
If K = 13 Then
    If Check5.Value = 1 Then
        Text1.Enabled = True
        Text1.SetFocus
    Else
        Check9.SetFocus
    End If
End If
End Sub
Private Sub Check6_keypress(K As Integer)
If K = 13 Then
    If Check6.Value = 1 Then
        Text6.Enabled = True
        Text6.SetFocus
    Else
        Command1.SetFocus
    End If
End If
End Sub
Private Sub Check8_KeyPress(K As Integer)
If K = 13 Then
    If Check8.Value = 1 Then
        Text8.Enabled = True
        Text8.SetFocus
    Else
        Check5.SetFocus
    End If
End If
End Sub
Private Sub Check7_KeyPress(K As Integer)
If K = 13 Then
    If Check7.Value = 1 Then
        Frame2.Enabled = False
        CHECK1.Value = 0
        CHECK2.Value = 0
        CHECK3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = 0
        Check8.SetFocus
    Else
        Check7.Enabled = False
        Frame2.Enabled = True
        CHECK3.SetFocus
    End If
End If
End Sub

Private Sub Check9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check9.Value = 1 Then
        Text7.Enabled = True
        Text7.SetFocus
    Else
        Check4.SetFocus
    End If
End If
End Sub

Private Sub cmbCaldates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If rswek.State = 1 Then rswek.Close
rswek.Open "select frm_dat,to_dat from wek where wek_idy=" & Val(cmbCaldates.Text), CON, adOpenKeyset, adLockOptimistic
If Not rswek.EOF Then
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    DTPicker1.Value = Format(rswek("frm_dat"), "MM/DD/YYYY")
    DTPicker2.Value = Format(rswek("to_dat"), "MM/DD/YYYY")
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    PTY_TYP.SetFocus
End If
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
Command1.Enabled = False
Command2.Enabled = False
'dates
RS.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    RS("DIR_ALS") = Format(DTPicker1, "DD/MM/YYYY")
    RS("DIR_MSG") = Format(DTPicker2, "DD/MM/YYYY")
    RS.Update
    If Me.Tag = "Scheme Value" Then
        SAL_QRY = "inl->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inl->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    Else
        SAL_QRY = "inv->inv_dat>=ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inv->inv_dat<=ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    End If
End If
'Products
If Check11.Value = 0 Then
    If Len(Trim(Label4)) <> 0 Then
    If RS.State = 1 Then RS.Close
    RS.Open "select * from dir where dir_idy like 'SUP' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        If Len(Trim(Label4)) > 0 Then
            RS("DIR_ALS") = Trim(Label4) & ""
            RS("DIR_MSG") = Trim(Text9) & ""
            RS.Update
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.prd->cmp_idy='" & Trim(Label4) & "'"
            Else
                SAL_QRY = "prd->cmp_IDY='" & Trim(Label4) & "'"
            End If
        End If
    End If
    End If
Else
    If RS.State = 1 Then RS.Close
    RS.Open "select * from dir where dir_idy like 'SUP' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
    RS("DIR_ALS") = "*"
    RS("DIR_MSG") = "All"
    RS.Update
    End If
End If
If CHECK1.Value = True Then
    If RS.State = 1 Then RS.Close
    RS.Open "select * from dir where dir_idy like 'PRD' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        If Len(Trim(Label5)) > 0 Then
            RS("DIR_ALS") = Trim(Label5) & ""
            RS("DIR_MSG") = Trim(Text2) & ""
            RS.Update
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.prd->prd_idy='" & Trim(Label5) & "'"
            Else
                SAL_QRY = "prd->prd_IDY='" & Trim(Label5) & "'"
            End If
        End If
    End If
ElseIf CHECK2.Value = True Then
'Segment
    If RS.State = 1 Then RS.Close
    RS.Open "select * from dir where dir_idy like 'SEG' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        If Len(Trim(Text4)) > 0 Then
            RS("DIR_MSG") = Trim(Text4) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.pgr->grP_cls='" & Trim(Text4) & "'"
            Else
                SAL_QRY = "pgr->grp_cls='" & Trim(Text4) & "'"
            End If
            RS.Update
        End If
    End If
ElseIf CHECK3.Value = True Then
'Category
    If RS.State = 1 Then RS.Close
    RS.Open "select * from dir where dir_idy like 'PGR' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        If Len(Trim(Text3)) > 0 Then
            RS("DIR_MSG") = Trim(Text3) & ""
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.pgr->grp_cat='" & Trim(Text3) & "'"
            Else
                SAL_QRY = "PGR->grp_cat='" & Trim(Text3) & "'"
            End If
            RS.Update
        End If
        RS.Update
    End If
End If
'Location
If RS.State = 1 Then RS.Close
RS.Open "select * from dir where dir_idy like 'LOC' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    If Trim(Label10) <> "" Then
        RS("DIR_ALS") = Trim(Label10) & ""
        RS("DIR_MSG") = Trim(Text8) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->LOC_IDY='" & Trim(Label10) & "'"
        Else
            SAL_QRY = "ACT->LOC_IDY='" & Trim(Label10) & "'"
        End If
    Else
        RS("DIR_ALS") = "*"
        RS("DIR_MSG") = "All"
    End If
    RS.Update
End If
'Category Retailers
If RS.State = 1 Then RS.Close
RS.Open "select * from dir where dir_idy like 'CAT' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    If Trim(Label8) <> "" Then
        RS("DIR_ALS") = Trim(Label8) & ""
        RS("DIR_MSG") = Trim(Text1) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->CAT_IDY='" & Trim(Label8) & "'"
        Else
            SAL_QRY = "ACT->CAT_IDY='" & Trim(Label8) & "'"
        End If
    Else
        RS("DIR_ALS") = "*"
        RS("DIR_MSG") = "All"
    End If
    RS.Update
End If
'Markets
If RS.State = 1 Then RS.Close
RS.Open "select * from dir where dir_idy like 'MAR' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
        If Trim(Label11) <> "" Then
        RS("DIR_ALS") = Trim(Label11) & ""
        RS("DIR_MSG") = Trim(Text7) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->MAR_IDY='" & Trim(Label11) & "'"
        Else
            SAL_QRY = "ACT->MAR_IDY='" & Trim(Label11) & "'"
        End If
    Else
        RS("DIR_ALS") = "*"
        RS("DIR_MSG") = "All"
    End If
    RS.Update
End If
'Salesman
If RS.State = 1 Then RS.Close
RS.Open "select * from dir where dir_idy like 'SMN' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    If Trim(Label7) <> "" Then
        RS("DIR_ALS") = Trim(Label7) & ""
        RS("DIR_MSG") = Trim(Text5) & ""
        If Me.Tag <> "SPD" And Me.Tag = "SEG_PRT" Then
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label7) & "'"
            Else
                SAL_QRY = "INL->SMN_IDY='" & Trim(Label7) & "'"
            End If
        ElseIf Me.Tag = "Sales Scheme" Or Me.Tag = "Sales Scheme Date" Then
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label7) & "'"
            Else
                SAL_QRY = "INL->SMN_IDY='" & Trim(Label7) & "'"
            End If
        Else
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.INL->SMN_IDY='" & Trim(Label7) & "'"
            Else
                SAL_QRY = "INL->SMN_IDY='" & Trim(Label7) & "'"
            End If
        End If
    Else
        RS("DIR_ALS") = "*"
        RS("DIR_MSG") = "All"
    End If
    RS.Update
End If
'RouteId
If RS.State = 1 Then RS.Close
RS.Open "select * from dir where dir_idy like 'ROU' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    If Check10.Value = 1 And Trim(LIST2.Text) <> "" Then
        RS("DIR_ALS") = Trim(LIST2) & ""
        RS("DIR_MSG") = Trim(LIST2) & ""
        If Me.Tag <> "SPD" And Me.Tag = "SEG_PRT" Then
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.INL->ROU_IDY='" & Trim(LIST2) & "'"
            Else
                SAL_QRY = "INL->ROU_IDY='" & Trim(LIST2) & "'"
            End If
        Else
            If Trim(Len(SAL_QRY)) <> 0 Then
                SAL_QRY = SAL_QRY & ".AND.INL->ROU_IDY='" & Trim(LIST2) & "'"
            Else
                SAL_QRY = "INL->ROU_IDY='" & Trim(LIST2) & "'"
            End If
        End If
    Else
        RS("DIR_ALS") = "*"
        RS("DIR_MSG") = "All"
    End If
    RS.Update
End If
If PTY_TYP.Selected(0) = False Then
    For A = 1 To 4
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

'Customer ID
If RS.State = 1 Then RS.Close
RS.Open "select * from dir where dir_idy like 'DLR' AND dir_idx like 'IDY'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    If Trim(Label9) <> "" Then
        RS("DIR_ALS") = Trim(Label9) & ""
        RS("DIR_MSG") = Trim(Text6) & ""
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.ACT->ACT_IDY='" & Trim(Label9) & "'"
        Else
            SAL_QRY = "ACT->ACT_IDY='" & Trim(Label9) & "'"
        End If
    Else
        RS("DIR_ALS") = "*"
        RS("DIR_MSG") = "All"
    End If
    RS.Update
End If
RS.Update
If Me.Tag = "SPD" Then
   SAL_QRY = SAL_QRY & ".and.INV->TXN_TYP = 'SAL'"
   Openreport CURDIR & "\MIR.RP1", "Product Pickup Retailer", 2, SAL_QRY, 6
ElseIf Me.Tag = "Scheme Value" Then
    SAL_QRY = SAL_QRY & ".and. lnv->lnv_amt>0"
    Openreport CURDIR & "\REPORT.RP1", "Sales Scheme Value SKU", 2, SAL_QRY, 6
   ' SET_SPD
ElseIf Me.Tag = "NOP" Then
    Me.MousePointer = vbHourglass
    CON.Execute "DELETE FROM NPO"
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CST_IDY FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD,PGR WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND INV_DAT >=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INV_DAT <=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND INV.PRD_IDY LIKE '" & Trim(Label5) & "%' AND GRP_CLS LIKE '" & Trim(Text4) & "%' AND GRP_CAT LIKE '" & Trim(Text2) & "%') AND SMN_IDY LIKE'" & Trim(Label7) & "%' and ROU_IDY LIKE '" & Trim(LIST2) & "%'", CON, adOpenKeyset, adLockOptimistic
    
    If RS.RecordCount > 0 Then RS.MoveFirst
    While RS.EOF = False
        CON.Execute "INSERT INTO NPO VALUES('" & Trim(RS(0) & "") & "')"
        RS.MoveNext
    Wend
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM ACT WHERE ACT_TYP IN('BOT','SAL') AND DBT_TYP = 'Y' AND ACT_IDY NOT IN(SELECT CST_IDY FROM NPO)", CON, adOpenKeyset, adLockOptimistic
    CON.Execute "DELETE FROM NPO"
    While RS.EOF = False
        CON.Execute "INSERT INTO NPO VALUES('" & Trim(RS(0) & "") & "')"
        RS.MoveNext
    Wend
    
    Me.MousePointer = vbDefault
    Openreport CURDIR & "\MIR.rp1", "Non Performing Outlets", 0, SAL_QRY, 6
    Unload Me
ElseIf Me.Tag = "PPD" Then
    SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP = 'SAL'"
    Openreport CURDIR & "\MIR.RP1", "Product Pickup Product", 2, SAL_QRY, 6
ElseIf Me.Tag = "Stock Type" Then
    SAL_QRY = SAL_QRY & ".AND. INV->TXN_TYP ='SAL'"
    Openreport CURDIR & "\REPORT.RP1", "Sales Product Type", 2, SAL_QRY, 6
ElseIf Me.Tag = "Sales Product Type" Then
    SAL_QRY = SAL_QRY & ".AND. INV->TXN_TYP ='SAL'"
    Openreport CURDIR & "\REPORT.RP1", "Sales Product Group", 2, SAL_QRY, 6
ElseIf Me.Tag = "Invoice Returns" Then
    SAL_QRY = SAL_QRY & ".AND. INV->TXN_TYP ='SAL' .AND. (INV->STK_TYP = 'Damage' .OR. INV->STK_TYP = 'TakeBack'.OR. INV->STK_TYP = 'Breakage'.OR. INV->STK_TYP = 'Leakage')"
    Openreport CURDIR & "\REPORT.RP1", "Sales Takeback", 2, SAL_QRY, 6
ElseIf Me.Tag = "Sales Scheme" Then
    SAL_QRY = SAL_QRY & ".AND. INV->TXN_TYP ='SAL' .AND. INV->STK_TYP = 'Free'"
    If Option1.Value = True Then
        Openreport CURDIR & "\REPORT.RP1", "Sales Scheme", 2, SAL_QRY, 6
    ElseIf Option3.Value = True Then
        Openreport CURDIR & "\REPORT.RP1", "Sales Scheme Consolidated", 2, SAL_QRY, 6
    Else
        Openreport CURDIR & "\REPORT.RP1", "Sales Scheme SKU", 2, SAL_QRY, 6
    End If
ElseIf Me.Tag = "Sales Scheme Date" Then
    SAL_QRY = SAL_QRY & ".AND. INV->TXN_TYP ='SAL' .AND. INV->STK_TYP = 'Free'"
    Openreport CURDIR & "\REPORT.RP1", "Sales Scheme Date", 2, SAL_QRY, 6
ElseIf Me.Tag = "PPV" Then
    SAL_QRY = SAL_QRY & ".and. INV->TXN_TYP = 'SAL'"
    Openreport CURDIR & "\MIR.RP1", "Product Pickup Value", 2, SAL_QRY, 6
ElseIf Me.Tag = "Free" Then
    If Trim(Len(SAL_QRY)) <> 0 Then
        SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    Else
        SAL_QRY = "INV->TXN_TYP='FRS'"
    End If
    Openreport CURDIR & "\report.rp1", "Free Stock Issues", 2, SAL_QRY, 6
ElseIf Me.Tag = "sup" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL')"
    Openreport CURDIR & "\REP_LIB.RP1", "Supplier wise", 2, SAL_QRY, 6
ElseIf Me.Tag = "sup_b" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL')"
    Openreport CURDIR & "\REP_LIB.RP1", "Supplier wise - True Sale", 2, SAL_QRY, 6
ElseIf Me.Tag = "stk" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL')"
    Openreport CURDIR & "\REP_LIB.RP1", "Supplier wise - Stock", 2, SAL_QRY, 6
ElseIf Me.Tag = "RET_DIS" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL').AND.(inv->add_amt>0.or.inv->dis_amt>0)"
    Openreport CURDIR & "\REPORT.RP1", "Product Discount Retailer Base", 2, SAL_QRY, 6
ElseIf Me.Tag = "SUP_DIS" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL').AND.(inv->add_amt>0.or.inv->dis_amt>0)"
    Openreport CURDIR & "\REPORT.RP1", "Product Discount", 2, SAL_QRY, 6
ElseIf Me.Tag = "SEG_DIS" Then
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL').AND.(inv->add_amt>0.or.inv->dis_amt>0)"
    Openreport CURDIR & "\REPORT.RP1", "Product Discount Retailer", 2, SAL_QRY, 6
Else
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL')"
    Openreport CURDIR & "\report.rp1", "Sales Product", 2, Trim(SAL_QRY), 6
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Me.Height = 4800 Then
    Me.Height = 5970
    Command3.Caption = "&Details <<"
Else
    Me.Height = 4800
    Command3.Caption = "&Details >>"
End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Check11.Enabled = True
Check11.SetFocus
End If
If KeyCode = 38 And Shift = 1 Then DTPicker1.SetFocus
End Sub

Private Sub DTPicker2_LostFocus()
If DTPicker1 > DTPicker2 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "Effmcg"
    DTPicker1.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
SETYEAR M_REP_PQY
DTPicker1.SetFocus
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
Check7.Value = 1
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check8.Value = 0
If Me.Tag = "Sales Scheme" Or Me.Tag = "Scheme Value" Then
    Frame4.Visible = True
End If
PTY_TYP.Selected(0) = True
Check12.SetFocus
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
        If Label20.Caption = "Select Retailers Category" Then Text1.SetFocus
        If Label20.Caption = "Select Market" Then Text7.SetFocus
        If Label20.Caption = "Select Supplier" Then Text9.SetFocus
    Else
    
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
AddData LIST2, Me, "ROU_IDY", "RDY", "", "GROUP BY ROU_IDY"
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
            Check7.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Check8.SetFocus
        ElseIf Label20 = "Select Product Segment" Then
            Text4 = list1.SelectedItem
            Text2 = ""
            Text3 = ""
            Label5 = ""
            Check7.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Check8.SetFocus
        ElseIf Label20 = "Select Product" Then
            Label5 = list1.SelectedItem.ListSubItems.Item(1)
            Text2 = Trim(list1.SelectedItem)
            Text3 = ""
            Text4 = ""
            Check7.Value = False
            SSTab1.Tab = 0
            list1.Visible = False
            Check8.SetFocus
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

Private Sub list2_keypress(K As Integer)
If K = 13 Then
    Check6.SetFocus
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
            CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(1) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(2) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(3) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            CHECK1.SetFocus
        End If
    ElseIf PTY_TYP.Selected(4) = True Then
        If Check7.Enabled = True Then
            Check7.SetFocus
        Else
            CHECK1.SetFocus
        End If
    Else
        
    End If
End If
End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ProgressBar1.Value > 1 Then
    Me.MousePointer = vbHourglass
End If
End Sub

Private Sub Text1_Change()
If Text1 = "" Then Label8 = ""
End Sub

Private Sub Text1_Keydown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("cat", "cat_idy", Trim(Text1)) = True Then
Check9.SetFocus
Else
RS.Open "select * from cat where cat_nme like '" & Trim(Text1) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("cat_nme")))
        ls.ListSubItems.Add , , Trim(RS("cat_idy") & "")
        RS.MoveNext
    Next
    list1.ColumnHeaders.Clear
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    SSTab1.Tab = 1
    Label20 = "Select Retailers Category"
    list1.ColumnHeaders(1).Text = "Retailers Category"
    list1.ColumnHeaders(1).Width = 3500
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Retailers Category Not Found, Define Category !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
End If
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text2_Change()
If Text2 = "" Then Label5 = ""
End Sub

Private Sub TEXT2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("prd", "prd_idy", Trim(Label5)) = True Then
Check4.SetFocus
Else
RS.Open "select * from prd,pgr where prd_nme like '" & Trim(Text2) & "%' and prd.grp_idy like pgr.grp_idy and cmp_idy like '" & Trim(Label4) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        'list1.ColumnHeaders.Clear
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
    list1.ColumnHeaders(1).Width = 2000
    list1.ColumnHeaders(2).Text = "ID"
    list1.ColumnHeaders(3).Text = "Product Segment"
    list1.ColumnHeaders(3).Width = 2000
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
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
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
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text4_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("pgr", "grp_cls", Trim(Text4)) = True Then
Check4.SetFocus
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
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub


Private Sub Text5_Change()
If Text5 = "" Then Label7 = ""
End Sub

Private Sub Text5_keydown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("smn", "smn_idy", Trim(Label7)) = True Then
    Check10.SetFocus
Else
RS.Open "select * from smn where smn_nme like '" & Trim(Text5) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("smn_nme")))
        ls.ListSubItems.Add , , Trim(RS("smn_idy") & "")
        RS.MoveNext
    Next
    list1.ColumnHeaders.Clear
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    SSTab1.Tab = 1
    Label20 = "Select Salesman"
    list1.ColumnHeaders(1).Text = "Salesman Name"
    list1.ColumnHeaders(1).Width = 2500
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text5.Enabled = True
        Text5.SetFocus
    End If
End If
End If
End If
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Text6_Change()
If Text6 = "" Then Label9 = ""
End Sub

Private Sub Text6_keydown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("act", "act_idy", Trim(Label9)) = True Then
    Command1.SetFocus
Else
    RS.Open "select * from act  where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT')", CON, adOpenStatic
    list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("act_nme")))
        ls.ListSubItems.Add , , Trim(RS("act_idy") & "")
        RS.MoveNext
    Next
    list1.ColumnHeaders.Clear
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    SSTab1.Tab = 1
    Label20 = "Select Retailer"
    list1.ColumnHeaders(1).Width = 2000
    list1.ColumnHeaders(1).Text = "Retailer Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Retailer Not Found, Define Retailer !", vbExclamation, "Effmcg"
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

Private Sub Text7_Change()
If Text7 = "" Then Label11 = ""
End Sub

Private Sub Text7_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
RS.Open "select * from mar where mar_nme like '" & Trim(Text7) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("mar_nme")))
        ls.ListSubItems.Add , , Trim(RS("mar_idy") & "")
        RS.MoveNext
    Next
    list1.ColumnHeaders.Clear
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    SSTab1.Tab = 1
    Label20 = "Select Market"
    list1.ColumnHeaders(1).Width = 2000
    list1.ColumnHeaders(1).Text = "Market Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Market Not Found, Define Market !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text8.Enabled = True
        Text8.SetFocus
    End If
End If
End If
End Sub

Private Sub Text7_keypress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Text8_Change()
If Text8 = "" Then Label10 = ""
End Sub

Private Sub Text8_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("loc", "loc_idy", Sincrement(Label10)) = True Then
CHECK3.SetFocus
Else
RS.Open "select * from loc where loc_nme like '" & Trim(Text8) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("loc_nme")))
        ls.ListSubItems.Add , , Trim(RS("loc_idy") & "")
        RS.MoveNext
    Next
    list1.ColumnHeaders.Clear
    list1.ColumnHeaders.Add
    list1.ColumnHeaders.Add
    SSTab1.Tab = 1
    Label20 = "Select Retailers Location"
    list1.ColumnHeaders(1).Text = "Retailers Location"
    list1.ColumnHeaders(1).Width = 2000
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Location Not Found, Define Location !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text8.Enabled = True
        Text8.SetFocus
    End If
End If
End If
End If
End Sub

Public Sub SET_SPD()
Me.MousePointer = vbHourglass
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Label3.Visible = True
Label13 = 0
Label14 = 0
COMCTR = 0
ProgressBar1.Visible = True
CON.Execute "DELETE FROM PPC"
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
RS.Open "SELECT SUM(TOT_PCS),PRD_IDY,CST_IDY,SUB_UNT FROM INV,INL WHERE INV.INV_IDY LIKE INL.INV_IDY AND INL.INV_DAT>=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INL.INV_DAT<=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND INL.TXN_TYP ='SAL' AND INL.SMN_IDY LIKE '" & Trim(Label7) & "%' AND INL.ROU_IDY LIKE '" & Trim(LIST2) & "%' AND STK_TYP IN('Free','Salable') GROUP BY INL.CST_IDY,PRD_IDY", CON, adOpenKeyset, adLockOptimistic
rs2.Open "SELECT SUM(TOT_PCS),PRD_IDY,CST_IDY,SUB_UNT FROM INV,INL WHERE INV.INV_IDY LIKE INL.INV_IDY AND INL.INV_DAT>=CTOD('" & Format(DTPicker1, "MM/DD/YYYY") & "') AND INL.INV_DAT<=CTOD('" & Format(DTPicker2, "MM/DD/YYYY") & "') AND INL.TXN_TYP ='SAL' AND INL.SMN_IDY LIKE '" & Trim(Label7) & "%' AND INL.ROU_IDY LIKE '" & Trim(LIST2) & "%' AND STK_TYP IN('Damage','TakeBack','Breakage','Leakage') GROUP BY INL.CST_IDY,PRD_IDY", CON, adOpenKeyset, adLockOptimistic
RS1.Open "SELECT * FROM PPC", CON, adOpenKeyset, adLockOptimistic
Label14 = RS.RecordCount + rs2.RecordCount
CTR = RS.RecordCount
For A = 1 To RS.RecordCount
    RS1.AddNew
    RS1!CST_IDY = Trim(RS(2))
    RS1!PRD_IDY = Trim(RS(1))
    RS1!prd_qty = Val(RS(0) & "")
    RS1!inv_idy = Val(RS(3) & "")
    RS1.Update
    COMCTR = COMCTR + 1
    inc_pgb
    RS.MoveNext
Next
For A = 1 To rs2.RecordCount
    CON.Execute "UPDATE PPC SET PRD_QTY = PRD_QTY - " & Val(rs2(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs2!PRD_IDY) & "'", d
    If d = 0 Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM PPC", CON, adOpenKeyset, adLockOptimistic
        RS1.AddNew
        RS1!CST_IDY = Trim(rs2(2))
        RS1!PRD_IDY = Trim(rs2(1))
        RS1!prd_qty = Val(rs2(0) & "") * -1
        RS1!inv_idy = Val(rs2(3) & "")
        RS1.Update
    End If
    COMCTR = COMCTR + 1
    inc_pgb
    rs2.MoveNext
Next

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM PPC", CON, adOpenKeyset
For A = 1 To RS.RecordCount
    RS!prd_qty = STM(Val(RS!prd_qty & ""), Val(RS!inv_idy & ""))
    RS!inv_idy = ""
    RS.Update
    RS.MoveNext
Next
ProgressBar1.Visible = False
Label3.Visible = False
Me.MousePointer = vbDefault
Openreport CURDIR & "\MIR.RP1", "Product Pickup", 2, SAL_QRY, 6
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Text9_Change()
If Text9 = "" Then Label4 = ""
End Sub

Private Sub Text9_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("act", "act_idy", Trim(Label4)) = True Then
    Command1.SetFocus
Else
    RS.Open "select * from act  where act_nme like '" & Trim(Text9) & "%' and (act_typ like 'PUR' or act_typ like 'BOT')", CON, adOpenStatic
    list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("act_nme")))
        ls.ListSubItems.Add , , Trim(RS("act_idy") & "")
        RS.MoveNext
    Next
    list1.ColumnHeaders.Clear
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
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text9.Enabled = True
        Text9.SetFocus
    End If
End If
End If
End If
End Sub
Public Sub inc_pgb()
On Error Resume Next
DoEvents
Label13 = COMCTR
Label13.Refresh
ProgressBar1.Value = (100 / CTR) * COMCTR
End Sub
Private Sub Check12_Click()
Check12_KeyDown 13, 0
End Sub

Private Sub Check12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check12.Value = 1 Then
        cmbCaldates.Enabled = True
        If rswek.State = 1 Then rswek.Close
        rswek.Open "select wek_idy from wek", CON, adOpenKeyset, adLockOptimistic
        cmbCaldates.Clear
        If Not rswek.EOF Then
            While Not rswek.EOF
                cmbCaldates.AddItem rswek("wek_idy")
                rswek.MoveNext
            Wend
            cmbCaldates.SetFocus
        Else
            MsgBox "Sales Calendar does not exist!", vbExclamation, "Effmcg"
            Check12.SetFocus
            Check12.Value = 0
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
'cmbCaldates_KeyDown 13, 0
End Sub
