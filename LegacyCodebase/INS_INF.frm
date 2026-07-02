VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form INS_INF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Distributor Setup"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "INS_INF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Distributor &Details"
      TabPicture(0)   =   "INS_INF.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "apgst"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cst"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ad1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Cmpnme"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "zone"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Town"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ffr"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "wds"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "depot"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "unt"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "&Invoice Remarks"
      TabPicture(1)   =   "INS_INF.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Check1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Shape2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Shape1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label13"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "INS_INF.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "list1"
      Tab(2).Control(1)=   "Label20"
      Tab(2).ControlCount=   2
      Begin VB.TextBox unt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
         Height          =   1935
         Left            =   -74400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "INS_INF.frx":0496
         Top             =   2400
         Width           =   6255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "I &Accept the License Agreement"
         Height          =   375
         Left            =   -74160
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5655
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   960
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1500
         Width           =   5655
      End
      Begin VB.TextBox depot 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1215
      End
      Begin VB.TextBox wds 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4260
         Width           =   1215
      End
      Begin VB.TextBox ffr 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4785
         Width           =   2295
      End
      Begin VB.TextBox Town 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox zone 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3180
         Width           =   1215
      End
      Begin VB.TextBox Cmpnme 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   80
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   3975
      End
      Begin VB.TextBox ad1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1140
         Width           =   5655
      End
      Begin VB.TextBox cst 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2415
         Width           =   2295
      End
      Begin VB.TextBox apgst 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2295
      End
      Begin MSComctlLib.ListView list1 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   7223
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "CST"
         Height          =   195
         Left            =   1395
         TabIndex        =   38
         Top             =   2835
         Width           =   315
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Code"
         Height          =   195
         Left            =   3300
         TabIndex        =   37
         Top             =   3608
         Width           =   705
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   2655
         Left            =   -74500
         Top             =   2300
         Width           =   6495
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Left            =   -74520
         Top             =   2280
         Width           =   6495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Remarks"
         Height          =   195
         Left            =   -74700
         TabIndex        =   27
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Market"
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
         TabIndex        =   25
         Top             =   600
         Width           =   7545
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3720
         TabIndex        =   22
         Top             =   4260
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3720
         TabIndex        =   21
         Top             =   3900
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3720
         TabIndex        =   20
         Top             =   3540
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3720
         TabIndex        =   19
         Top             =   3180
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distributor Code-COLA"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Officer In Charge"
         Height          =   195
         Left            =   510
         TabIndex        =   17
         Top             =   4860
         Width           =   1200
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town - District"
         Height          =   195
         Left            =   690
         TabIndex        =   16
         Top             =   3975
         Width           =   1020
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depot - Unit"
         Height          =   195
         Left            =   855
         TabIndex        =   15
         Top             =   3615
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone - Region"
         Height          =   195
         Left            =   690
         TabIndex        =   9
         Top             =   3255
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distributor Name"
         Height          =   195
         Left            =   555
         TabIndex        =   8
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Left            =   1140
         TabIndex        =   7
         Top             =   1215
         Width           =   570
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT - TIN"
         Height          =   195
         Left            =   990
         TabIndex        =   6
         Top             =   2130
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GST"
         Height          =   195
         Left            =   1380
         TabIndex        =   5
         Top             =   2490
         Width           =   330
      End
   End
End
Attribute VB_Name = "INS_INF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub zonelist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from zon where zon_nme like '" & zone & "%'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("zon_nme")))
        ls.ListSubItems.Add , , Trim(rs("zon_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 2
    Label20 = "Select Zone"
    list1.ColumnHeaders(1).Text = "Zone Name"
    list1.ColumnHeaders(2).Text = "Zone ID"
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Zone Not Found, Define Zone...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    zone.SetFocus
End If
End Sub

Private Sub townlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from twn where twn_nme like '" & Town & "%'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("twn_nme")))
        ls.ListSubItems.Add , , Trim(rs("twn_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 2
    Label20 = "Select Town"
    list1.ColumnHeaders(1).Text = "Town Name"
    list1.ColumnHeaders(2).Text = "Town ID"
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Town Not Found, Define Town...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Town.SetFocus
End If
End Sub

Private Sub depotlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dep where dep_nme like '" & depot & "%'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("dep_nme")))
        ls.ListSubItems.Add , , Trim(rs("dep_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 2
    Label20 = "Select Depot"
    list1.ColumnHeaders(1).Text = "Depot Name"
    list1.ColumnHeaders(2).Text = "Depot ID"
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Depot Not Found, Define Depot...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    depot.SetFocus
End If
End Sub

Private Sub ffrlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from ffr where ffr_nme like '" & ffr & "%'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("ffr_nme")))
        ls.ListSubItems.Add , , Trim(rs("ffr_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 2
    Label20 = "Select FFR"
    list1.ColumnHeaders(1).Text = "FFR Name"
    list1.ColumnHeaders(2).Text = "FFR ID"
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "FFR Not Found, Define FFR...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    ffr.SetFocus
End If
End Sub

Private Sub wdlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from wds where wdI_nme like '" & wds & "%'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("wdI_nme")))
        ls.ListSubItems.Add , , Trim(rs("wdI_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 2
    Label20 = "Select WD"
    list1.ColumnHeaders(1).Text = "WD Name"
    list1.ColumnHeaders(2).Text = "WD ID"
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "WD Not Found, Define WD...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    wds.SetFocus
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Command2.Enabled = True
Else
Command2.Enabled = False
End If
End Sub

Private Sub Cmpnme_KEYDOWN(K As Integer, S As Integer)
If K = 13 Then
    If Len(Trim(Cmpnme)) = 0 Then
        MsgBox "Invalid Company Name, Enter Name !", vbExclamation
        Cmpnme.SetFocus
    Else
        ad1.SetFocus
    End If
End If
End Sub

Private Sub ad1_KEYDOWN(K As Integer, S As Integer)
If K = 13 Then
    If Len(Trim(ad1)) = 0 Then
        MsgBox "Invalid Address, Enter Address !", vbExclamation
        ad1.SetFocus
    Else
        Text1.SetFocus
    End If
ElseIf K = 38 Then
    If Len(Trim(ad1)) = 0 Then
        MsgBox "Invalid Address, Enter Address !", vbExclamation
        ad1.SetFocus
    Else
        Cmpnme.SetFocus
    End If
End If
End Sub
Private Sub apgst_KEYDOWN(K As Integer, S As Integer)
If K = 13 Then
    CST.SetFocus
ElseIf K = 38 Then
    Text1.SetFocus
End If
End Sub
Private Sub Command1_Click()
Unload Me
Set Form = Nothing
End Sub

Private Sub Command2_Click()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT1'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(Cmpnme) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT2'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(ad1) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT3'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(Text1) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT4'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(Text6) & ""
rs.Update

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM1'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(Text2) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM2'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(Text3) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM3'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(Text4) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM4'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(Text5) & ""
rs.Update

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'CST'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(CST) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'GST'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs("DIR_MSG") = Trim(apgst) & ""
rs.Update
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'FFR'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_MSG") = Trim(ffr) & ""
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'DEP'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Trim(depot) & ""
    rs("DIR_MSG") = Trim(Label10)
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'TWN'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Trim(Town) & ""
    rs("DIR_MSG") = Trim(Label11)
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'ZON'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Trim(zone) & ""
    rs("DIR_MSG") = Trim(Label29)
    rs.Update
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'UNT'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Trim(unt) & ""
    rs("DIR_MSG") = Trim(Label29)
    rs.Update
Else
    rs.AddNew
    rs("DIR_IDY") = "UNT"
    rs("DIR_IDX") = "IDY"
    rs!DIR_ALS = Trim(unt)
    rs!DIR_MSG = ""
    rs.Update
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'WDS'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs("DIR_ALS") = Trim(wds) & ""
    rs("DIR_MSG") = Trim(Label12)
    rs.Update
End If
If Text7.Visible = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dir WHERE dir_idY like 'SYS' AND DIR_IDX LIKE 'PSW'", CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        rs("DIR_ALS") = "Yes"
        rs.Update
    End If
End If

Unload Me
Set Form = Nothing
End Sub

Private Sub Cst_KEYDOWN(K As Integer, S As Integer)
If K = 13 Then
    Text6.SetFocus
ElseIf K = 38 Then
    apgst.SetFocus
End If
End Sub

Private Sub DISPDETAILS()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT1'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then Cmpnme = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT2'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then ad1 = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT3'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then Text1 = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'TT4'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then Text6 = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'CST'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then CST = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'GST'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then apgst = Trim(rs("DIR_MSG") & "")

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'FFR'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    ffr = Trim(rs("DIR_MSG") & "")
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM1'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then Text2 = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM2'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then Text3 = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM3'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then Text4 = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idX like 'RM4'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then Text5 = Trim(rs("DIR_MSG") & "")
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'DEP'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    depot = Trim(rs("DIR_ALS") & "")
    Label10 = Trim(rs("DIR_MSG") & "")
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'TWN'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    Town = Trim(rs("DIR_ALS") & "")
    Label11 = Trim(rs("DIR_MSG") & "")
End If


If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'ZON'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    zone = Trim(rs("DIR_ALS") & "")
    Label29 = Trim(rs("DIR_MSG") & "")
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'UNT'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    unt = Trim(rs("DIR_ALS") & "")
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'WDS'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    wds = Trim(rs("DIR_ALS") & "")
    Label12 = Trim(rs("DIR_MSG") & "")
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from dir WHERE dir_idY like 'SYS' AND DIR_IDX LIKE 'PSW'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Trim(rs("dir_als")) = "Yes" Then
        Text7.Visible = False
        Check1.Visible = False
        Shape1.Visible = False
        Shape2.Visible = False
        Command2.Enabled = True
    Else
        Text7.Visible = True
        Check1.Visible = True
        Shape1.Visible = True
        Shape2.Visible = True
        Command2.Enabled = False
    End If
End If

End Sub

Private Sub depot_keydown(K As Integer, S As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
'If RS.State = 1 Then RS.Close
'        RS.Open "select * from dep where dep_idy like '" & Trim(depot) & "'", con, adOpenStatic
'        If RS.RecordCount > 0 Then
'            depot = Trim(depot)
'            Label10 = RS("dep_nme")
            If K = 38 Then
                zone.SetFocus
            Else
                unt.SetFocus
            End If
'        Else
'            depotlist
'        End If
End If
End Sub

Private Sub Form_Activate()
Cmpnme.SetFocus
SSTab1.TabEnabled(2) = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        If Label20.Caption = "Select Zone" Then zone.SetFocus
        If Label20.Caption = "Select Depot" Then depot.SetFocus
        If Label20.Caption = "Select Town" Then Town.SetFocus
        If Label20.Caption = "Select WD" Then wds.SetFocus
        If Label20.Caption = "Select FFR" Then ffr.SetFocus
    Else
        Unload Me
        Set Form = Nothing
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
List1_keypress 13
End Sub

Private Sub List1_keypress(K As Integer)
If K = 13 Then
    If list1.ListItems.COUNT > 0 Then
        If Label20 = "Select Zone" Then
            zone = Trim(list1.SelectedItem.ListSubItems.Item(1))
            Label29 = Trim(list1.SelectedItem)
            list1.Visible = False
            SSTab1.Tab = 0
            depot.SetFocus
        ElseIf Label20 = "Select Depot" Then
            depot = Trim(list1.SelectedItem.ListSubItems.Item(1))
            Label10 = Trim(list1.SelectedItem)
            list1.Visible = False
            SSTab1.Tab = 0
            Town.SetFocus
        ElseIf Label20 = "Select Town" Then
            Town = Trim(list1.SelectedItem.ListSubItems.Item(1))
            Label11 = Trim(list1.SelectedItem)
            list1.Visible = False
            SSTab1.Tab = 0
            wds.SetFocus
        ElseIf Label20 = "Select WD" Then
            wds = Trim(list1.SelectedItem.ListSubItems.Item(1))
            Label12 = Trim(list1.SelectedItem)
            list1.Visible = False
            SSTab1.Tab = 0
            ffr.SetFocus
        ElseIf Label20 = "Select FFR" Then
            ffr = Trim(list1.SelectedItem.ListSubItems.Item(1))
            Label13 = Trim(list1.SelectedItem)
            list1.Visible = False
            SSTab1.Tab = 0
            Command2.SetFocus
        End If
    End If
End If
End Sub

Private Sub list1_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then
    Command1.Visible = False
    Command2.Visible = False
Else
    Command1.Visible = True
    Command2.Visible = True
End If
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    apgst.SetFocus
ElseIf KeyCode = 38 Then
    ad1.SetFocus
End If
End Sub

Private Sub TEXT6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    zone.SetFocus
ElseIf KeyCode = 38 Then
    CST.SetFocus
End If
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text3.SetFocus
ElseIf KeyCode = 38 Then
    SSTab1.Tab = 0
    ffr.SetFocus
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text4.SetFocus
ElseIf KeyCode = 38 Then
    Text2.SetFocus
End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text5.SetFocus
ElseIf KeyCode = 38 Then
    Text3.SetFocus
End If
End Sub
Private Sub TEXT5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Command2.Enabled = True Then
    Command2.SetFocus
    Else
    Check1.SetFocus
    End If
ElseIf KeyCode = 38 Then
    Text4.SetFocus
End If
End Sub

Private Sub town_keydown(K As Integer, S As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
'If RS.State = 1 Then RS.Close
'        RS.Open "select * from twn where twn_idy like '" & Trim(Town) & "'", con, adOpenStatic
'        If RS.RecordCount > 0 Then
'            Town = Trim(Town)
'            Label11 = RS("twn_nme")
            If K = 38 Then
                unt.SetFocus
            Else
                wds.SetFocus
            End If
'        Else
'            townlist
'        End If
End If
End Sub
Private Sub wds_keydown(K As Integer, S As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
'If RS.State = 1 Then RS.Close
'        RS.Open "select * from wds where wdI_idy like '" & Trim(wds) & "'", con, adOpenStatic
'        If RS.RecordCount > 0 Then
'            wds = Trim(wds)
'            Label12 = RS("wdI_nme")
            If K = 38 Then
                Town.SetFocus
            Else
                ffr.SetFocus
            End If
'        Else
'            wdlist
'        End If
End If
End Sub
Private Sub unt_keydown(K As Integer, S As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
'If RS.State = 1 Then RS.Close
'        RS.Open "select * from wds where wdI_idy like '" & Trim(wds) & "'", con, adOpenStatic
'        If RS.RecordCount > 0 Then
'            wds = Trim(wds)
'            Label12 = RS("wdI_nme")
            If K = 38 Then
                depot.SetFocus
            Else
                Town.SetFocus
            End If
'        Else
'            wdlist
'        End If
End If
End Sub
Private Sub ffr_keydown(K As Integer, S As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    ffr = Trim(ffr)
    If K = 38 Then
        wds.SetFocus
    Else
        SSTab1.Tab = 1
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
DISPDETAILS
End Sub

Private Sub zone_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Zone Code"
End Sub

Private Sub town_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Town Code"
End Sub

Private Sub ffr_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Supplier Representative Code"
End Sub

Private Sub depot_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Depot Code"
End Sub

Private Sub wds_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Your Reference Code as per Supplier Records"
End Sub
Private Sub unt_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Your Reference Code as per Supplier Records"
End Sub

Private Sub zone_keydown(K As Integer, S As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
'If RS.State = 1 Then RS.Close
'        RS.Open "select * from zon where zon_idy like '" & Trim(zone) & "" & "'", con, adOpenStatic
'        If RS.RecordCount > 0 Then
'            zone = Trim(zone)
'            Label29 = RS("zon_nme")
            If K = 38 Then
                Text6.SetFocus
            Else
                depot.SetFocus
            End If
'        Else
'            zonelist
'        End If
End If
End Sub

Private Sub zone_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub town_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub ffr_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub wds_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub depot_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub
Private Sub unt_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

