VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_ROU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retailer Visit Sequence"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_ROU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5150
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Visit Sequence "
      TabPicture(0)   =   "M_FRM_ROU.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "PB"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DT"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtVar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSort"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Combo2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Combo3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtsmn"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_ROU.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LST_CAP"
      Tab(1).Control(1)=   "list4"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtsmn 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4320
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "M_FRM_ROU.frx":047A
         Left            =   6840
         List            =   "M_FRM_ROU.frx":047C
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Combo3"
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "M_FRM_ROU.frx":047E
         Left            =   3360
         List            =   "M_FRM_ROU.frx":0497
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   6615
         Begin VB.TextBox SMN_IDY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2880
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   292
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "M_FRM_ROU.frx":04DB
            Left            =   5160
            List            =   "M_FRM_ROU.frx":04F4
            Style           =   2  'Dropdown List
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   300
            Width           =   1335
         End
         Begin VB.ComboBox cmbRoute 
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Salesman"
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Day"
            Height          =   195
            Index           =   2
            Left            =   4440
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Route"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "&Order Retailers"
         Enabled         =   0   'False
         Height          =   330
         Left            =   7185
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtVar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         MaxLength       =   4
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Load Retailers"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24707074
         CurrentDate     =   37874
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   100
         TabIndex        =   4
         Top             =   4800
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   3435
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1320
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   6059
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Retailers List....Please Wait....."
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   4800
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label LST_CAP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   16
         Top             =   720
         Width           =   8295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":0538
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":0A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":0E44
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":1198
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":17D0
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":1B24
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":1F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":2278
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":2684
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":2790
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":2AE4
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ROU.frx":2EF8
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl+S)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACT_LST"
                  Text            =   "&1 All Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_CAT"
                  Text            =   "&2 Retailer Category"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_LOC"
                  Text            =   "&3 Retailer Location"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_MAR"
                  Text            =   "&4 Retailer Market"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ROU"
                  Text            =   "&5 Retailer Beat"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_SMN"
                  Text            =   "&6 Retailer Salesman"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SUP_LST"
                  Text            =   "&7 Supplier List"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ID_LST"
                  Text            =   "&8 Retailer ID List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "act"
                  Text            =   "&Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dlr"
                  Text            =   "&Dealers List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sup"
                  Text            =   "&Supplier List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Account ID "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "m_frm_rou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsact As New ADODB.Recordset
Dim rsrou As New ADODB.Recordset
Dim RSMAR As New ADODB.Recordset
Dim RSCAT As New ADODB.Recordset
Dim R As Integer, C As Integer
Dim MD As Boolean

Private Sub cmbRoute_Change()
If Len(Trim(cmbRoute)) = 3 Then SLT
End Sub

Private Sub cmbRoute_KeyDown(k As Integer, Shift As Integer)
If k = 13 Then
    Command1.Enabled = True
    Command1.SetFocus
    cmdSort.Enabled = True
End If
End Sub

Private Sub cmbRoute_keypress(k As Integer)
If k = 39 Or k = 41 Or k = 37 Then k = 0
If k > 47 And k < 58 Or k = 8 Or k = 13 Then
ElseIf k > 96 And k < 123 Then
    k = k - 32
ElseIf k > 64 And k < 91 Then
Else
    k = 0
End If
End Sub

Private Sub cmdSort_Click()
MS.Col = 0
MS.ColSel = 0
If MS.Rows = 1 Then Exit Sub
Me.MousePointer = vbHourglass
MS.Sort = 2
R = 1
Do While Val(MS.TextMatrix(R, 0)) <> 0
    C = R + 1
    If C <> MS.Rows Then
    Do While Val(MS.TextMatrix(C, 0)) <> 0
        If Val(MS.TextMatrix(R, 0)) > Val(MS.TextMatrix(C, 0)) Then
            SNUM = Val(MS.TextMatrix(R, 0))
            RNAME = MS.TextMatrix(R, 1)
            MNAME = MS.TextMatrix(R, 2)
            ROUT = Trim(MS.TextMatrix(R, 3))
            SMN = Trim(MS.TextMatrix(R, 4))
            VDAY = Trim(MS.TextMatrix(R, 5))
            DIST = Val(MS.TextMatrix(R, 6))
            ARRT = Trim(MS.TextMatrix(R, 7))
            DEPT = Trim(MS.TextMatrix(R, 8))
            CNAME = Trim(MS.TextMatrix(R, 9))
            LNAME = Trim(MS.TextMatrix(R, 10))
            SMNID = Trim(MS.TextMatrix(R, 11))
            ACTID = Trim(MS.TextMatrix(R, 12))
            MS.TextMatrix(R, 0) = MS.TextMatrix(C, 0)
            MS.TextMatrix(R, 1) = MS.TextMatrix(C, 1)
            MS.TextMatrix(R, 2) = MS.TextMatrix(C, 2)
            MS.TextMatrix(R, 3) = MS.TextMatrix(C, 3)
            MS.TextMatrix(R, 4) = MS.TextMatrix(C, 4)
            MS.TextMatrix(R, 5) = MS.TextMatrix(C, 5)
            MS.TextMatrix(R, 6) = MS.TextMatrix(C, 6)
            MS.TextMatrix(R, 7) = MS.TextMatrix(C, 7)
            MS.TextMatrix(R, 8) = MS.TextMatrix(C, 8)
            MS.TextMatrix(R, 9) = MS.TextMatrix(C, 9)
            MS.TextMatrix(R, 10) = MS.TextMatrix(C, 10)
            MS.TextMatrix(R, 11) = MS.TextMatrix(C, 11)
            MS.TextMatrix(R, 12) = MS.TextMatrix(C, 12)
            MS.TextMatrix(C, 0) = SNUM
            MS.TextMatrix(C, 1) = RNAME
            MS.TextMatrix(C, 2) = MNAME
            MS.TextMatrix(C, 3) = ROUT
            MS.TextMatrix(C, 4) = SMN
            MS.TextMatrix(C, 5) = VDAY
            MS.TextMatrix(C, 6) = DIST
            MS.TextMatrix(C, 7) = ARRT
            MS.TextMatrix(C, 8) = DEPT
            MS.TextMatrix(C, 9) = CNAME
            MS.TextMatrix(C, 10) = LNAME
            MS.TextMatrix(C, 11) = SMNID
            MS.TextMatrix(C, 12) = ACTID
        End If
        C = C + 1
        If C = MS.Rows Then
            Exit Do
        End If
    Loop
    End If
    R = R + 1
    If R = MS.Rows Then
        Exit Do
    End If
Loop
MS.SetFocus
Me.MousePointer = vbDefault
End Sub

Private Sub Combo1_KeyPress(k As Integer)
If k = 13 Then
    Command1.Enabled = True
    cmdSort.Enabled = True
    Command1.SetFocus
End If
End Sub

Private Sub Combo2_keypress(L As Integer)
If L = 13 Then
    MS.Text = Combo2
    Combo2.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Len(Trim(Combo3)) = 3 Then SLT
End Sub

Private Sub Combo3_KeyDown(L As Integer, Shift As Integer)
If L = 13 Then
    C = 0
    STMP = MS.TextMatrix(MS.Row, 0)
    SNUM = MS.Row
    For R = 1 To MS.Rows - 1
        If Val(MS.TextMatrix(R, 0)) = STMP And Trim(MS.TextMatrix(R, 3)) = Trim(Combo3.Text) Then
            MS.TextMatrix(SNUM, 0) = 0
            Exit For
        End If
    Next
    MS.Text = Combo3
    Combo3.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Combo3_keypress(k As Integer)
If k = 39 Or k = 41 Or k = 37 Then k = 0
If k > 47 And k < 58 Or k = 8 Or k = 13 Then
ElseIf k > 96 And k < 123 Then
k = k - 32
ElseIf k > 64 And k < 91 Then
Else
k = 0
End If
End Sub

Private Sub Command1_Click()
    Label2.Visible = True
    Label2.Caption = "Processing Query..."
    Me.MousePointer = vbHourglass
    DoEvents
    Command1.Enabled = False
    If RSMAR.State = 1 Then RSMAR.Close
    With MS
        .Rows = 2
        R = 1
        If RS.State = 1 Then RS.Close
        If Option1(0).Value = True Then
            MS.FixedCols = 0
            RS.Open "SELECT SEQ_NUM,ACT_NME,CAT_NME,ROU_DST,TIM_INN,TIM_OUT,MAR_NME,SMN_NME,DAY_IDY,RDY.ROU_IDY RDY_IDY,RDY.SMN_IDY RDY_SMN_IDY,ACT.ACT_IDY,LOC.LOC_NME FROM smn,mar,RDY,LOC,ACT left join cat on act.cat_idy =cat.cat_idy WHERE ACT.ACT_IDY LIKE RDY.ACT_IDY AND rdy.smn_idy like smn.smn_idy AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.LOC_IDY LIKE LOC.LOC_IDY AND RDY.ROU_IDY LIKE '" & Trim(cmbRoute) & "' ORDER BY ACT_NME", CON, adOpenKeyset, adLockOptimistic
            'GROUP BY RDY.ACT_IDY
        ElseIf Option1(1).Value = True Then
            MS.FixedCols = 1
            RS.Open "SELECT SEQ_NUM,ACT_NME,CAT_NME,ROU_DST,TIM_INN,TIM_OUT,MAR_NME,SMN_NME,DAY_IDY,RDY.ROU_IDY RDY_IDY,RDY.SMN_IDY RDY_SMN_IDY,ACT.ACT_IDY,LOC.LOC_NME FROM smn,mar,RDY,LOC,ACT left join cat on act.cat_idy =cat.cat_idy WHERE ACT.ACT_IDY LIKE RDY.ACT_IDY AND rdy.smn_idy like smn.smn_idy AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.LOC_IDY LIKE LOC.LOC_IDY AND RDY.SMN_IDY LIKE '" & Trim(SMN_IDY.Tag) & "' ORDER BY ACT_NME", CON, adOpenKeyset, adLockOptimistic
        ElseIf Option1(2).Value = True Then
            MS.FixedCols = 1
            RS.Open "SELECT SEQ_NUM,ACT_NME,CAT_NME,ROU_DST,TIM_INN,TIM_OUT,MAR_NME,SMN_NME,DAY_IDY,RDY.ROU_IDY RDY_IDY,RDY.SMN_IDY RDY_SMN_IDY,ACT.ACT_IDY,LOC.LOC_NME FROM smn,mar,RDY,LOC,ACT left join cat on act.cat_idy =cat.cat_idy WHERE ACT.ACT_IDY LIKE RDY.ACT_IDY AND rdy.smn_idy like smn.smn_idy AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.LOC_IDY LIKE LOC.LOC_IDY AND RDY.DAY_IDY = " & Val(Combo1.ListIndex + 1) & " GROUP BY RDY.ACT_IDY ORDER BY ACT_NME", CON, adOpenKeyset, adLockOptimistic
        End If
        'DoEvents
        
        If RS.RecordCount > 0 Then
        PB.Visible = True
        Label2.Left = 4560
        Label2.Caption = "Loading Retailers... Please Wait..."
        Label2.Refresh
        PB.Max = RS.RecordCount + 1
        MS.Rows = RS.RecordCount + 1
        For A = 1 To RS.RecordCount
            MS.TextMatrix(A, 0) = Val(RS(0) & "")
            MS.TextMatrix(A, 1) = Trim(RS(1))
            MS.TextMatrix(A, 2) = Trim(RS(6) & "")
            MS.TextMatrix(A, 3) = Trim(RS(9) & "")
            MS.TextMatrix(A, 4) = Trim(RS(7))
            MS.TextMatrix(A, 5) = Combo1.list(Val(RS(8) & "") - 1)
            MS.TextMatrix(A, 6) = Trim(RS(3))
            MS.TextMatrix(A, 7) = Trim(RS(4))
            MS.TextMatrix(A, 8) = Trim(RS(5))
            MS.TextMatrix(A, 9) = Trim(RS(2) & "")
            MS.TextMatrix(A, 10) = Trim(RS(12) & "")
            MS.TextMatrix(A, 11) = Trim(RS(10) & "")
            MS.TextMatrix(A, 12) = Trim(RS(11) & "")
            RS.MoveNext
            PB.Value = A
        Next
        PB.Visible = False
        Label2.Visible = False
        MS.SetFocus
        If Option1(0).Value = True Then
            MS.Col = 0
        Else
            MS.Col = 1
        End If
        Else
            MsgBox "Retailers Not Found", vbExclamation
            cmdSort.Enabled = False
            Label2.Visible = False
            PB.Visible = False
            cancelrecord
        End If
    End With
    Me.MousePointer = vbNormal
End Sub

Private Sub DT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MS.Col = 7 Then
        MS.Text = Format(DT, "hh:mm:ss")
        DT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 8 Then
        If Format(DT, "hh:mm:ss") > Format(MS.TextMatrix(MS.Row, 7), "hh:mm:ss") Then
            MS.Text = Format(DT, "hh:mm:ss")
            DT.Visible = False
            MS.SetFocus
        Else
            MsgBox "Invalid Departure Time ", vbExclamation
            DT.Visible = True
            DT.SetFocus
        End If
    End If
End If
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
Option1.Item(0).SetFocus
SSTab1.TabEnabled(1) = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then Shell "c:\windows\calc.exe"
If KeyCode = 83 And Shift = 2 Then saverecord
If KeyCode = 88 And Shift = 2 Then Unload Me
If KeyCode = 81 And Shift = 2 Then cancelrecord
If KeyCode = 112 Then M_HELP
If KeyCode = 69 And Shift = 2 Then If m_frm_rou.ActiveControl.name = "SMN_IDY" Then V_SMN_LST
If KeyCode > 111 And KeyCode < 120 Then If TypeOf m_frm_rou.ActiveControl Is TextBox Then m_frm_rou.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "Option1" Then
        Unload Me
    ElseIf SSTab1.Tab = 1 And list4.Visible = True Then
        list4.Visible = False
        SSTab1.Tab = 0
        If SMN_IDY.Enabled = True Then SMN_IDY.SetFocus
    Else
        saverecord
    End If
End If
End Sub

Private Sub Form_Load()
AddData cmbRoute, Me, "rou_idy", "rdy", "", "group by rou_idy"
AddData Combo3, Me, "rou_idy", "rdy", "", "group by rou_idy"
LOADMS
End Sub

Private Sub list4_DblClick()
list4_KeyPress 13
End Sub

Private Sub list4_KeyPress(k As Integer)
If k <> 13 Then Exit Sub
list4.Visible = False
SSTab1.Tab = 0
If list4.Tag = "MS" Then
    txtsmn.Visible = False
    MS.Text = list4.SelectedItem.ListSubItems.Item(1)
    MS.TextMatrix(MS.Row, 10) = list4.SelectedItem
    MS.SetFocus
Else
    SMN_IDY = list4.SelectedItem.ListSubItems.Item(1)
    SMN_IDY.Tag = list4.SelectedItem
    Command1.Enabled = True
    cmdSort.Enabled = True
    Command1.SetFocus
End If
End Sub

Private Sub MS_KeyPress(k As Integer)
If MS.Row = 0 Then Exit Sub
If k > 96 And k < 123 Then
    k = k - 32
ElseIf k > 47 And k < 58 Or k = 46 Then
End If
If k = 39 Or k = 41 Then k = 0
If k = 27 Then
Else
    If MS.Col = 0 Or MS.Col = 6 Then
        If k <> 13 Then
            MS.TextMatrix(MS.Row, MS.Col) = ""
            txtVar.Visible = True
            txtVar.Left = MS.Left + MS.CellLeft
            txtVar.Top = MS.Top + MS.CellTop
            txtVar.Height = MS.CellHeight
            txtVar.Width = MS.CellWidth
            txtVar.Visible = True
            txtVar.SetFocus
            If k = 13 Then
                k = 0
            Else
                txtVar = Chr(k)
            End If
            txtVar.SelStart = Len(Trim(txtVar))
        End If
    ElseIf MS.Col = 7 Or MS.Col = 8 Then
        If IsDate(MS.Text) Then DT = MS.Text
        DT.Visible = True
        DT.Top = MS.Top + MS.CellTop
        DT.Left = MS.CellLeft + MS.Left
        DT.Width = MS.CellWidth
        DT.Height = MS.CellHeight
        DT.Visible = True
        DT.SetFocus
    ElseIf MS.Col = 5 Then
        If k <> 13 Then
            Combo2.Visible = True
            Combo2.Enabled = True
            Combo2.Top = MS.Top + MS.CellTop
            Combo2.Left = MS.CellLeft + MS.Left
            Combo2.Width = MS.CellWidth
            Combo2.Visible = True
            Combo2.SetFocus
        End If
    ElseIf MS.Col = 3 Then
        If k <> 13 Then
            Combo3.Visible = True
            Combo3.Enabled = True
            Combo3.Top = MS.Top + MS.CellTop
            Combo3.Left = MS.CellLeft + MS.Left
            Combo3.Width = MS.CellWidth
            Combo3.Visible = True
            Combo3.SetFocus
        End If
    ElseIf MS.Col = 4 Then
        If k <> 0 And k <> 27 And k <> 13 Then
            txtsmn = ""
            txtsmn.Visible = True
            txtsmn.Left = MS.Left + MS.CellLeft
            txtsmn.Top = MS.Top + MS.CellTop
            txtsmn.Height = MS.CellHeight
            txtsmn.Width = MS.CellWidth
            txtsmn.Visible = True
            txtsmn.SetFocus
            If k = 13 Then
                k = 0
            Else
                txtsmn = Chr(k)
            End If
            txtsmn.SelStart = Len(Trim(txtsmn))
        End If
    End If
End If
End Sub

Private Sub MS_LeaveCell()
Combo2.Visible = False
Combo3.Visible = False
txtVar.Visible = False
txtsmn.Visible = False
DT.Visible = False
End Sub

Private Sub MS_Scroll()
If txtVar.Visible = True Then
    txtVar = ""
    txtVar.Visible = False
    MS.SetFocus
ElseIf DT.Visible = True Then
    DT.Visible = False
    MS.SetFocus
ElseIf Combo2.Visible = True Then
    Combo2.Visible = False
    MS.SetFocus
ElseIf Combo3.Visible = True Then
    Combo3.Visible = False
    MS.SetFocus
ElseIf txtsmn.Visible = True Then
    txtsmn = ""
    txtsmn.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Option1_KeyPress(I As Integer, k As Integer)
On Error Resume Next
If k = 13 Then
    If I = 0 Then
        SMN_IDY.Enabled = False
        SMN_IDY = ""
        Combo1.Enabled = False
        Combo1.ListIndex = 0
        cmbRoute.Enabled = True
        cmbRoute.SetFocus
    ElseIf I = 1 Then
        cmbRoute.Enabled = False
        cmbRoute.ListIndex = 0
        Combo1.Enabled = False
        Combo1.ListIndex = 0
        SMN_IDY.Enabled = True
        SMN_IDY.SetFocus
    ElseIf I = 2 Then
        SMN_IDY.Enabled = False
        SMN_IDY = ""
        cmbRoute.Enabled = False
        cmbRoute.ListIndex = 0
        Combo1.Enabled = True
        Combo1.SetFocus
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "s" Then
    saverecord
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "p" Then
    
ElseIf Button.KEY = "c" Then
    cancelrecord
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If Me.ActiveControl.name = "SMN_IDY" Then V_SMN_LST
End If
End Sub

Private Sub Txtsmn_KeyPress(k As Integer)
If k = 39 Or k = 41 Or k = 37 Then k = 0
If k = 13 Then
    If CHECKTB("smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancelrecord
        Exit Sub
    End If
    If check("SMN", "SMN_IDY", Sincrement(Trim(txtsmn))) = True Then
        MS.Text = GetValue("SMN_NME", "SMN WHERE SMN_IDY LIKE '" & Sincrement(Trim(txtsmn)) & "'")
        MS.TextMatrix(MS.Row, 10) = Sincrement(Trim(txtsmn))
        txtsmn.Visible = False
        MS.SetFocus
    Else
        list4.Tag = "MS"
        V_SMN_LST
    End If
End If
End Sub

Private Sub txtvar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MS.Col = 0 Then
        For R = 1 To MS.Rows - 1
            If Val(txtVar) = 0 Then
            ElseIf Trim(MS.TextMatrix(R, MS.Col)) = Trim(txtVar) Then
                MsgBox "Sequence should not contain duplicate values", vbExclamation
                txtVar.Visible = True
                txtVar.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        Next
        If R = MS.Rows Then
            MS.Col = 0
            txtVar.Visible = False
            MS.Text = txtVar.Text
            txtVar = ""
            If MS.Row < MS.Rows - 1 Then
                MS.Row = MS.Row + 1
            End If
            MS.SetFocus
        End If
    Else
        txtVar.Visible = False
        MS.Text = txtVar.Text
        txtVar = ""
        MS.Col = MS.Col + 1
        MS.SetFocus
    End If
End If
End Sub
Private Sub saverecord()
If MsgBox("Save Record?", vbYesNo + vbExclamation) = vbNo Then
    cancelrecord
    Exit Sub
End If
If Option1(1).Value = True Then
    If Trim(SMN_IDY) = "" Then
        MsgBox "Select Salesman ID", vbExclamation
        SSTab1.Tab = 0
        SMN_IDY.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
End If
If MS.Rows = 1 Then
    MsgBox "Load Retailers ", vbExclamation
    Command1.Enabled = True
    cmdSort.Enabled = True
    Command1.SetFocus
    Exit Sub
End If
If rsrou.State = 1 Then rsrou.Close
If Option1(0).Value = True Then
    CON.Execute "DELE FROM RDY WHERE ROU_IDY LIKE '" & Trim(cmbRoute) & "'"
    SAL_QRY = "RDY->ROU_IDY='" & Trim(cmbRoute) & "'"
ElseIf Option1(1).Value = True Then
    CON.Execute "DELE FROM RDY WHERE SMN_IDY LIKE '" & Trim(SMN_IDY.Tag) & "'"
    SAL_QRY = "RDY->SMN_IDY='" & Trim(SMN_IDY.Tag) & "'"
Else
    CON.Execute "DELE FROM RDY WHERE DAY_IDY =" & Trim(Combo1.ListIndex + 1)
    SAL_QRY = "RDY->DAY_IDY=" & Trim(Combo1.ListIndex + 1)
End If
rsrou.Open "select * from rdy", CON, adOpenKeyset, adLockOptimistic
R = 1
For R = 1 To MS.Rows - 1
    rsrou.AddNew
    rsrou!ROU_IDY = Trim(cmbRoute)
    rsrou("Act_idy") = Trim(MS.TextMatrix(R, 12)) & ""
    rsrou("rou_dst") = Val(MS.TextMatrix(R, 6))
    rsrou("tim_inn") = Trim(MS.TextMatrix(R, 7)) & ""
    rsrou("tim_out") = Trim(MS.TextMatrix(R, 8)) & ""
    rsrou("seq_num") = Val(MS.TextMatrix(R, 0))
    If Trim(MS.TextMatrix(R, 5)) = "Monday" Then
        dayid = 1
    ElseIf Trim(MS.TextMatrix(R, 5)) = "Tuesday" Then
        dayid = 2
    ElseIf Trim(MS.TextMatrix(R, 5)) = "Wednesday" Then
        dayid = 3
    ElseIf Trim(MS.TextMatrix(R, 5)) = "Thursday" Then
        dayid = 4
    ElseIf Trim(MS.TextMatrix(R, 5)) = "Friday" Then
        dayid = 5
    ElseIf Trim(MS.TextMatrix(R, 5)) = "Saturday" Then
        dayid = 6
    ElseIf Trim(MS.TextMatrix(R, 5)) = "Sunday" Then
        dayid = 7
    End If
    rsrou("DAY_IDY") = dayid
    rsrou("SMN_IDY") = Trim(MS.TextMatrix(R, 11))
    rsrou("ROU_IDY") = Trim(MS.TextMatrix(R, 3))
    rsrou.Update
    rsrou.MoveNext
Next
Openreport CURDIR & "\LISTS.RP1", "Route Sequence by Route", 2, SAL_QRY, 6
cancelrecord
End Sub

Private Sub txtVar_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 0
End Sub
Private Sub SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesman ID or Press <Enter> To Select From The List"
End Sub

Private Sub smn_idy_kEYPRESS(k As Integer)
If k > 96 And k < 123 Then k = k - 32
If k = 39 Or k = 41 Or k = 37 Then k = 0
If k = 13 Then
    If CHECKTB("smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        'cancel
        Exit Sub
    End If
    If check("SMN", "SMN_IDY", Sincrement(Trim(SMN_IDY))) = True Then
        SMN_IDY.Tag = Sincrement(Trim(SMN_IDY))
        SMN_IDY = GetValue("SMN_NME", "SMN WHERE SMN_IDY LIKE '" & Trim(SMN_IDY.Tag) & "'")
        Command1.Enabled = True
        cmdSort.Enabled = True
        Command1.SetFocus
    Else
        list4.Tag = ""
        V_SMN_LST
    End If
End If
End Sub
Private Sub V_SMN_LST()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    'cancelrecord
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim lst As ListItem
LST_CAP = "Select Salesman"
If REC_SET.State = 1 Then REC_SET.Close
If list4.Tag = "MS" Then
REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(txtsmn) & "%'", CON, adOpenKeyset, adLockOptimistic
Else
REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(SMN_IDY) & "%'", CON, adOpenKeyset, adLockOptimistic
End If
list4.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    list4.ColumnHeaders.Item(1).Text = "Salesman ID"
    list4.ColumnHeaders.Item(2).Text = "Salesman Name"
    For A = 1 To REC_SET.RecordCount
        Set lst = list4.ListItems.Add(, , Trim(REC_SET("smn_IDY")))
        lst.ListSubItems.Add , , Trim(REC_SET("smn_NME"))
        REC_SET.MoveNext
    Next
    REC_SET.Close
    SSTab1.Tab = 1
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesmen Not Found, Define Salesmen !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        If list4.Tag = "MS" Then
            txtsmn.Visible = True
            txtsmn.SetFocus
        Else
            SMN_IDY.SetFocus
        End If
    End If
End If
End Sub

Public Sub cancelrecord()
On Error Resume Next
clrctr Me
list4.Tag = ""
SMN_IDY.Tag = ""
cmbRoute.ListIndex = 0
Combo1.ListIndex = 0
MS.Clear
MS.Rows = 1
Combo2.Visible = False
txtVar.Visible = False
txtsmn.Visible = False
Combo3.Visible = False
Command1.Enabled = False
cmdSort.Enabled = False
Option1(0).SetFocus
AddData cmbRoute, Me, "rou_idy", "rdy", "", "group by rou_idy"
AddData Combo3, Me, "rou_idy", "rdy", "", "group by rou_idy"
LOADMS
Label2.Left = 3360
SSTab1.Tab = 0
End Sub
Private Sub LOADMS()
With MS
.Cols = 13
.Rows = 1
.FixedCols = 0
.TextMatrix(0, 0) = "Order #"
.ColAlignment(0) = 7
.ColWidth(0) = 700
.Col = 0
.Row = 0
.CellForeColor = vbBlue
.TextMatrix(0, 1) = "Retailer Name"
.ColAlignment(1) = 1
.ColWidth(1) = 1900
.TextMatrix(0, 2) = "Market"
.ColWidth(2) = 2300
.ColAlignment(2) = 1
.TextMatrix(0, 3) = "Route"
.ColAlignment(3) = 1
.Col = 3
.Row = 0
.CellForeColor = vbBlue
.ColWidth(3) = 800
.TextMatrix(0, 4) = "Salesman"
.ColAlignment(4) = 1
.Col = 4
.Row = 0
.CellForeColor = vbBlue
.ColWidth(4) = 2000
.TextMatrix(0, 5) = "Visit day"
.ColAlignment(5) = 1
.Col = 5
.Row = 0
.CellForeColor = vbBlue
.ColWidth(5) = 1200
.TextMatrix(0, 6) = "Dist.Kms"
.ColWidth(6) = 1200
.Col = 6
.Row = 0
.CellForeColor = vbBlue
.ColAlignment(6) = 7
.ColAlignment(7) = 1
.TextMatrix(0, 7) = "Arr.Time"
.ColWidth(7) = 1300
.Col = 7
.Row = 0
.CellForeColor = vbBlue
.TextMatrix(0, 8) = "Dep.Time"
.ColWidth(8) = 1300
.ColAlignment(8) = 1
.Row = 0
.Col = 8
.CellForeColor = vbBlue
.TextMatrix(0, 9) = "Category"
.ColWidth(9) = 1900
.ColAlignment(9) = 1
.TextMatrix(0, 10) = "Location"
.ColWidth(10) = 2000
.ColAlignment(10) = 1
.ColWidth(11) = 0
.ColWidth(12) = 0
.ColWidth(11) = 0
End With
End Sub
