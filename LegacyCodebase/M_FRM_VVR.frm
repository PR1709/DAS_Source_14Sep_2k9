VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_VVR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Van Variance Check"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_VVR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&Van Variance"
      TabPicture(0)   =   "M_FRM_VVR.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ACT_NME"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "NET_AMT"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TO_DAT"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FRM_DAT"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ALL_CHK"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "List2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "LDS_IDY"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ls"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command5"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ProgressBar1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "VAN_CHK"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "VAN_IDY"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "&Products - SKU Summary"
      TabPicture(1)   =   "M_FRM_VVR.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command7"
      Tab(1).Control(1)=   "MS"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "Label7"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Invoice Details"
      TabPicture(2)   =   "M_FRM_VVR.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command6"
      Tab(2).Control(1)=   "MS1"
      Tab(2).Control(2)=   "Label8"
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(4)=   "Tot_lbl"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "View"
      TabPicture(3)   =   "M_FRM_VVR.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LST_CAP"
      Tab(3).Control(1)=   "list4"
      Tab(3).ControlCount=   2
      Begin VB.TextBox VAN_IDY 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7575
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox VAN_CHK 
         Caption         =   "&Van Issue ID"
         Height          =   195
         Left            =   6240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   675
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   4750
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton Command7 
         Caption         =   "G&enerate"
         Height          =   375
         Left            =   -67800
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ge&nerate"
         Height          =   375
         Left            =   -67920
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   7560
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin MSComctlLib.ListView ls 
         Height          =   2205
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2520
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3889
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
            Text            =   "Invoice ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Retailer Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Net Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   700
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   8415
         Begin VB.OptionButton smn_chk 
            Appearance      =   0  'Flat
            Caption         =   "&Salesman"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   268
            Width           =   1095
         End
         Begin VB.ListBox Route 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            ItemData        =   "M_FRM_VVR.frx":04B2
            Left            =   7080
            List            =   "M_FRM_VVR.frx":04B4
            Style           =   1  'Checkbox
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   160
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            Caption         =   "R&oute"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   268
            Width           =   855
         End
         Begin VB.OptionButton mar_chk 
            Appearance      =   0  'Flat
            Caption         =   "&Market"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   268
            Width           =   855
         End
         Begin VB.TextBox SMN_IDY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1440
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   230
            Width           =   1455
         End
         Begin VB.TextBox MAR_IDY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   4320
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   230
            Width           =   1455
         End
      End
      Begin VB.TextBox LDS_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">>"
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Alt+L"
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<"
         Height          =   375
         Left            =   6480
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Alt+U"
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Alt+D"
         Top             =   4350
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Alt+S"
         Top             =   2520
         Width           =   375
      End
      Begin VB.ListBox List2 
         Height          =   2205
         ItemData        =   "M_FRM_VVR.frx":04B6
         Left            =   6960
         List            =   "M_FRM_VVR.frx":04B8
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox ALL_CHK 
         Caption         =   "&All"
         Height          =   255
         Left            =   6240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1245
         Width           =   675
      End
      Begin MSComCtl2.DTPicker FRM_DAT 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16580611
         CurrentDate     =   37628
      End
      Begin MSComCtl2.DTPicker TO_DAT 
         Height          =   330
         Left            =   4560
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16580611
         CurrentDate     =   37628
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   24
         Top             =   960
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4000
         Left            =   -74940
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   840
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   3
         Cols            =   12
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid MS1 
         Height          =   3555
         Left            =   -74950
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1080
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   6271
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
         ScrollBars      =   2
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   4560
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16580611
         CurrentDate     =   37628
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Double Click On SKU Name to DISPLAY Invoice wise SKU Quantity in Units (Pieces)"
         Height          =   195
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Loading SKU's Please Wait..."
         Height          =   195
         Left            =   -74880
         TabIndex        =   40
         Top             =   4800
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Loading SKU's Please Wait..."
         Height          =   195
         Left            =   -74880
         TabIndex        =   39
         Top             =   4875
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Processing Query Please Wait...."
         Height          =   195
         Left            =   3000
         TabIndex        =   38
         Top             =   4800
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   3890
         TabIndex        =   33
         Top             =   675
         Width           =   345
      End
      Begin VB.Label Label3 
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69960
         TabIndex        =   30
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Tot_lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -67905
         TabIndex        =   29
         Top             =   4800
         Width           =   390
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
         TabIndex        =   25
         Top             =   600
         Width           =   8295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Variance ID"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   675
         Width           =   840
      End
      Begin VB.Label NET_AMT 
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label ACT_NME 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   4035
         TabIndex        =   6
         Top             =   1275
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   1275
         Width           =   345
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   600
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
            Picture         =   "M_FRM_VVR.frx":04BA
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":09FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":0DC6
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":111A
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":1752
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":1AA6
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":1F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":21FA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":2606
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":2712
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":2A66
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_VVR.frx":2E7A
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LDS_BAT"
                  Text            =   "Variance Sheet - Batchwise"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LDS_PRD"
                  Text            =   "Variance Sheet - Productwise"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LDS_INV"
                  Text            =   "Variance Sheet - Inv.Track"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LDS_PTR"
                  Text            =   "Variance Sheet - Prd.Track"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TRP_SHT"
                  Text            =   "Trip Sheet"
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
Attribute VB_Name = "M_FRM_VVR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, CTR, COM_CTR

Private Sub ALL_CHK_KeyPress(K As Integer)
If K = 13 Then
    If ALL_CHK.Value = 0 Then
        Frame1.Enabled = True
        smn_chk.SetFocus
    Else
        Frame1.Enabled = False
        smn_chk.Value = False
        mar_chk.Value = False
        Option3.Value = False
        Label6.Visible = True
        LoadInvoices
        Label6.Visible = False
    End If
End If
End Sub

Private Sub Command1_Click()
For A = 1 To ls.ListItems.count
    If ls.ListItems.Item(A).Selected = True Then
    List2.AddItem ls.ListItems.Item(A)
    End If
Next
A = 1
b = ls.ListItems.count
While A <= b
    If ls.ListItems.Item(A).Selected = True Then
    ls.ListItems.Remove A
        A = A - 1
        b = b - 1
    End If
    A = A + 1
Wend
ls.SetFocus
End Sub

Private Sub Command2_Click()
If CHECKTB("inl.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
GETINVOICE
Dim lst As ListItem
Dim rs As New ADODB.Recordset
If List2.ListCount > 0 Then
rs.Open "SELECT * FROM INL,ACT WHERE ACT.ACT_IDY LIKE INL.CST_IDY AND INV_IDY in (SELECT CST_IDY FROM NPO)", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    Set lst = ls.ListItems.Add(, , rs("inv_idy"))
    lst.ListSubItems.Add , , Trim(rs!ACT_NME)
    lst.ListSubItems.Add , , convert(rs!NET_AMT)
    rs.MoveNext
Next
List2.CLEAR
End If
End Sub

Private Sub Command3_Click()
If CHECKTB("act.dbf,inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim lst As ListItem
Dim rs As New ADODB.Recordset
For A = 1 To List2.ListCount
    If List2.Selected(A - 1) = True Then
        Set lst = ls.ListItems.Add(, , List2.List(A - 1))
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM INL,ACT WHERE ACT.ACT_IDY LIKE INL.CST_IDY AND INV_IDY LIKE '" & Trim(List2.List(A - 1)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            lst.ListSubItems.Add , , Trim(rs!ACT_NME)
            lst.ListSubItems.Add , , convert(rs!NET_AMT)
        End If
    End If
Next
A = 1
b = List2.ListCount
While A <= b
    If List2.Selected(A - 1) = True Then
    List2.RemoveItem A - 1
        A = A - 1
        b = b - 1
    End If
    A = A + 1
Wend
If List2.ListCount > 0 Then
    List2.Selected(0) = True
End If
List2.SetFocus
End Sub

Private Sub Command4_Click()
For A = 1 To ls.ListItems.count
    List2.AddItem ls.ListItems.Item(A)
Next
ls.ListItems.CLEAR
List2.SetFocus
End Sub

Private Sub Command5_Click()
If CHECKTB("inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
ls.ListItems.CLEAR
Label6.Visible = True
Label6.Refresh
Me.MousePointer = vbHourglass
If MOD_REC = True Then
    Dim RSINL As New ADODB.Recordset
    RSINL.Open "select * from inl where dch_idy like '" & Trim(LDS_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    List2.CLEAR
    MOD_REC = True
    For A = 1 To RSINL.RecordCount
    List2.AddItem RSINL!INV_IDY
    RSINL.MoveNext
    Next
    LoadInvoices
    Label6.Visible = False
    If RSINL.RecordCount > 0 Then   ' edited by radhika
        Command5.Enabled = False
    End If                          'end
    RSINL.Close
Else
    LoadInvoices
    Label6.Visible = False
End If
Me.MousePointer = vbNormal
End Sub
Private Sub Command5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Click To Refresh And Load Invoices"
End Sub

Private Sub Command6_Click()
If CHECKTB("inl.dbf,act.dbf,mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Label8.Visible = True
Label8.Refresh
Me.MousePointer = vbHourglass
Dim RSINL As New ADODB.Recordset
MS1.Rows = 1
Tot_lbl = "0.00"
GETINVOICE
RSINL.Open "select * from inl,act,mar where inl.cst_idy like act.act_idy and act.mar_idy like mar.mar_idy AND TXN_TYP LIKE 'SAL' and inv_idy in (SELECT CST_IDY FROM NPO)", con, adOpenKeyset, adLockPessimistic
MS1.Rows = RSINL.RecordCount + 1
For A = 1 To RSINL.RecordCount
    MS1.TextMatrix(A, 0) = Trim(RSINL!INV_IDY)
    MS1.TextMatrix(A, 1) = Trim(RSINL!ACT_NME)
    MS1.TextMatrix(A, 2) = Trim(RSINL!mar_nme)
    MS1.TextMatrix(A, 3) = convert(RSINL!NET_AMT)
    Tot_lbl = convert(Val(Tot_lbl) + RSINL!NET_AMT)
    RSINL.MoveNext
Next
MS1.ColWidth(4) = 0
MS1.ColWidth(3) = 1200
RSINL.Close
Me.MousePointer = vbNormal
Label8.Visible = False
If MS1.Rows > 1 Then
    MS1.SetFocus
    MS1.Row = 1
    MS1.Col = 0
End If
End Sub

Private Sub Command7_Click()
If CHECKTB("prd.dbf,inl.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Label7.Visible = True
Label7.Refresh
Me.MousePointer = vbHourglass
Dim RSINL As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
rsprd.Open "select * from prd", con, adOpenKeyset, adLockPessimistic
MS.Rows = 1
GETINVOICE
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TOT_PCS),BAT_IDY,INV.STK_TYP,INV.SUB_UNT,PRD_IDY,PRD_RAT FROM INL,INV WHERE INL.INV_IDY LIKE INV.INV_IDY AND INL.TXN_TYP LIKE 'SAL' AND INV.INV_IDY IN (SELECT CST_IDY FROM NPO) GROUP BY BAT_IDY,INV.STK_TYP", con, adOpenKeyset, adLockPessimistic
b = 1
chk = True
If RSINL.RecordCount > 0 Then  ' edited by radhika
    Command7.Enabled = False
End If  'end
For A = 1 To RSINL.RecordCount
    If Trim(RSINL!stk_typ) = "Salable" Then
        For C = 1 To MS.Rows - 1
            If Trim(MS.TextMatrix(C, 8)) = Trim(RSINL!BAT_IDY) Then
            MS.TextMatrix(C, 4) = Val(RSINL(0))
            chk = False
            End If
        Next C
        If chk = True Then
            MS.AddItem ""
            MS.TextMatrix(b, 0) = ""
            MS.TextMatrix(b, 1) = Int(Val(RSINL(0)) / Val(RSINL(3)))
            MS.TextMatrix(b, 2) = Val(RSINL(0)) Mod Val(RSINL(3))
            MS.TextMatrix(b, 3) = Val(RSINL(0))
            MS.TextMatrix(b, 4) = 0
            MS.TextMatrix(b, 5) = 0
            MS.TextMatrix(b, 6) = convert(RSINL!PRD_rat)
            MS.TextMatrix(b, 8) = RSINL!BAT_IDY
            MS.TextMatrix(b, 9) = RSINL!SUB_UNT
            MS.TextMatrix(b, 10) = RSINL!prd_idy
            b = b + 1
        End If
    End If
    RSINL.MoveNext
Next
chk = True
If RSINL.RecordCount > 0 Then RSINL.MoveFirst
For A = 1 To RSINL.RecordCount
    If Trim(RSINL!stk_typ) = "Free" Then
        For C = 1 To MS.Rows - 1
            If Trim(MS.TextMatrix(C, 8)) = Trim(RSINL!BAT_IDY) Then
            MS.TextMatrix(C, 4) = Val(RSINL(0))
            chk = False
            End If
        Next C
        If chk = True Then
            MS.AddItem ""
            MS.TextMatrix(MS.Rows - 1, 0) = ""
            MS.TextMatrix(MS.Rows - 1, 1) = 0
            MS.TextMatrix(MS.Rows - 1, 2) = 0
            MS.TextMatrix(MS.Rows - 1, 4) = Val(RSINL(0))
            MS.TextMatrix(MS.Rows - 1, 3) = 0
            MS.TextMatrix(MS.Rows - 1, 5) = 0
            MS.TextMatrix(MS.Rows - 1, 6) = convert(RSINL!PRD_rat)
            MS.TextMatrix(MS.Rows - 1, 8) = RSINL!BAT_IDY
            MS.TextMatrix(MS.Rows - 1, 9) = RSINL!SUB_UNT
            MS.TextMatrix(MS.Rows - 1, 10) = RSINL!prd_idy
        End If
        chk = True
    End If
    RSINL.MoveNext
Next A

chk = True
If RSINL.RecordCount > 0 Then RSINL.MoveFirst
For A = 1 To RSINL.RecordCount
    If Trim(RSINL!stk_typ) = "Replc." Then
        For C = 1 To MS.Rows - 1
            If Trim(MS.TextMatrix(C, 8)) = Trim(RSINL!BAT_IDY) Then
            MS.TextMatrix(C, 5) = Val(RSINL(0))
            chk = False
            End If
        Next C
        If chk = True Then
            MS.AddItem ""
            MS.TextMatrix(MS.Rows - 1, 0) = ""
            MS.TextMatrix(MS.Rows - 1, 1) = ""
            MS.TextMatrix(MS.Rows - 1, 2) = ""
            MS.TextMatrix(MS.Rows - 1, 5) = Val(RSINL(0))
            MS.TextMatrix(MS.Rows - 1, 3) = 0
            MS.TextMatrix(MS.Rows - 1, 4) = 0
            MS.TextMatrix(MS.Rows - 1, 6) = convert(RSINL!PRD_rat)
            MS.TextMatrix(MS.Rows - 1, 8) = RSINL!BAT_IDY
            MS.TextMatrix(MS.Rows - 1, 9) = RSINL!SUB_UNT
            MS.TextMatrix(MS.Rows - 1, 10) = RSINL!prd_idy
        End If
        chk = False
    End If
    RSINL.MoveNext
Next A
RSINL.filter = "stk_typ='TakeBack' or stk_typ='Damage' or stk_typ='Breakage' or stk_typ='Leakage'"
chk = True
If RSINL.RecordCount > 0 Then RSINL.MoveFirst
For A = 1 To RSINL.RecordCount
    For C = 1 To MS.Rows - 1
        If Trim(MS.TextMatrix(C, 8)) = Trim(RSINL!BAT_IDY) Then
        MS.TextMatrix(C, 11) = Val(RSINL(0)) + Val(MS.TextMatrix(C, 11))
        chk = False
        End If
    Next C
    If chk = True Then
        MS.AddItem ""
        MS.TextMatrix(MS.Rows - 1, 0) = ""
        MS.TextMatrix(MS.Rows - 1, 1) = ""
        MS.TextMatrix(MS.Rows - 1, 2) = ""
        MS.TextMatrix(MS.Rows - 1, 11) = Val(RSINL(0))
        MS.TextMatrix(MS.Rows - 1, 5) = 0
        MS.TextMatrix(MS.Rows - 1, 3) = 0
        MS.TextMatrix(MS.Rows - 1, 4) = 0
        MS.TextMatrix(MS.Rows - 1, 6) = convert(RSINL!PRD_rat)
        MS.TextMatrix(MS.Rows - 1, 8) = RSINL!BAT_IDY
        MS.TextMatrix(MS.Rows - 1, 9) = RSINL!SUB_UNT
        MS.TextMatrix(MS.Rows - 1, 10) = RSINL!prd_idy
    End If
    chk = False
    RSINL.MoveNext
Next A

For A = 1 To MS.Rows - 1
    If rsprd.BOF = False Then rsprd.MoveFirst
    rsprd.Find "prd_idy like '" & Trim(MS.TextMatrix(A, 10)) & "'"
    If rsprd.EOF = False Then MS.TextMatrix(A, 0) = rsprd!PRD_NME
    If MS.TextMatrix(A, 0) <> "" Then
        MS.TextMatrix(A, 1) = QTY(Int(Val(STM(Val(MS.TextMatrix(A, 3)) + Val(MS.TextMatrix(A, 4)) + Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 9))))))
        MS.TextMatrix(A, 2) = (Val(MS.TextMatrix(A, 3)) + Val(MS.TextMatrix(A, 4)) + Val(MS.TextMatrix(A, 5))) Mod Val(MS.TextMatrix(A, 9))
    End If
    MS.TextMatrix(A, 3) = QTY(STM(Val(MS.TextMatrix(A, 3)), Val(MS.TextMatrix(A, 9))))
    MS.TextMatrix(A, 4) = QTY(STM(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 9))))
    MS.TextMatrix(A, 5) = QTY(STM(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 9))))
    MS.TextMatrix(A, 11) = QTY(STM(Val(MS.TextMatrix(A, 11)), Val(MS.TextMatrix(A, 9))))
Next
If MS.Rows > 1 Then
MS.SetFocus
MS.Col = 0
MS.Row = 1
End If
Label7.Visible = False
Me.MousePointer = vbNormal
End Sub

Private Sub DTPicker1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date"
End Sub

Private Sub DTPicker1_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
VAN_CHK.SetFocus
End If
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If list4.Visible = True And LST_CAP = "Select Van Variance ID" Then
        list4.Visible = False
        SSTab1.Tab = 0
        LDS_IDY.SetFocus
    ElseIf list4.Visible = True And LST_CAP = "Select Salesman" Then
        list4.Visible = False
        SSTab1.Tab = 0
        SMN_IDY.SetFocus
    ElseIf list4.Visible = True And LST_CAP = "Select Market" Then
        list4.Visible = False
        SSTab1.Tab = 0
        MAR_IDY.SetFocus
    Else
        If M_FRM_VVR.ActiveControl.name = "LDS_IDY" Then
            Unload Me
        ElseIf List2.ListCount = 0 Then
            MsgBox "No Invoices Found", vbExclamation
            CANCELRECORD
        Else
            saverecord
        End If
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

'If List1.Visible = False And List2.Visible = False And list3.Visible = False And list4.Visible = False And list5.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    'If KeyCode = 80 And Shift = 2 Then Toolbar1_ButtonClick
    If KeyCode = 83 And Shift = 2 Then saverecord
    'If KeyCode = 88 And Shift = 2 Then EXT_REC
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_VVR.ActiveControl.name = "LDS_IDY" Then
            V_LDS_LST
        ElseIf M_FRM_VVR.ActiveControl.name = "M_INV_IDY" Then
            'V_INV_LST
        ElseIf M_FRM_VVR.ActiveControl.name = "M_CST_IDY" Then
            'V_DLR_LST
        ElseIf M_FRM_VVR.ActiveControl.name = "M_SMN_IDY" Then
            'V_SMN_LST
        ElseIf M_FRM_VVR.ActiveControl.name = "M_ACT_IDY" Then
            'V_ACT_LST
        ElseIf M_FRM_VVR.ActiveControl.name = "Text3" Then
            'V_PRD_LST
        End If
    End If
    If KeyCode = 83 And Shift = 4 Then
        Command1_Click
    ElseIf KeyCode = 76 And Shift = 4 Then
        Command4_Click
    ElseIf KeyCode = 85 And Shift = 4 Then
        Command3_Click
    ElseIf KeyCode = 68 And Shift = 4 Then
        Command2_Click
    End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_VVR.ActiveControl Is TextBox Then M_FRM_VVR.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Form_Load()
SETYEAR M_FRM_VVR
SSTab1.TabEnabled(3) = False
SSTab1.Tab = 0
MS.ColWidth(0) = 2400
MS.ColWidth(1) = 1000
MS.ColWidth(2) = 1000
MS.ColWidth(3) = 1000
MS.ColWidth(4) = 1020
MS.ColWidth(5) = 1000
MS.ColWidth(6) = 0
MS.ColWidth(7) = 0
MS.ColWidth(8) = 0
MS.ColWidth(9) = 0
MS.ColWidth(10) = 0
MS.ColWidth(11) = 1000
MS.TextMatrix(0, 0) = "Product Name"
MS.ColAlignment(1) = 7
MS.ColAlignment(2) = 7
MS.ColAlignment(3) = 7
MS.ColAlignment(4) = 7
MS.ColAlignment(5) = 7
MS.ColAlignment(6) = 7
MS.ColAlignment(7) = 7

MS.TextMatrix(0, 1) = "Cases"
MS.TextMatrix(0, 2) = "Pieces"
MS.TextMatrix(0, 3) = "Sale"
MS.TextMatrix(0, 4) = "Free"
MS.TextMatrix(0, 5) = "Replc."
MS.TextMatrix(0, 6) = "MRP"
MS.TextMatrix(0, 7) = "Amount"
MS.TextMatrix(0, 11) = "TakeBack"
MS1.ColWidth(0) = 1500
MS1.ColWidth(1) = 3500
MS1.ColWidth(2) = 2200
MS1.ColWidth(3) = 1500
MS1.ColWidth(4) = 0
MS1.ColAlignment(0) = 1
MS1.ColAlignment(3) = 7
MS1.ColAlignment(4) = 7
MS1.TextMatrix(0, 0) = "Invoice ID"
MS1.TextMatrix(0, 1) = "Retailer Name"
MS1.TextMatrix(0, 2) = "Market"
MS1.TextMatrix(0, 3) = "Amount"
MS1.TextMatrix(0, 4) = "Qty - Pieces"
AddData Route, Me, "rou_idy", "rdy", "", "group by rou_idy"
End Sub

Public Sub newrecord()
SSTab1.Tab = 0
CANCELRECORD
LDS_IDY.Enabled = False
clrctr M_FRM_VVR
MOD_REC = False
'CHANGED
DTPicker1.SetFocus
End Sub
Public Sub CANCELRECORD()
SSTab1.Tab = 0
clrctr M_FRM_VVR
SETYEAR M_FRM_VVR
ls.ListItems.CLEAR
ROUTECLEAR
List2.CLEAR
MS.Rows = 1
MS1.Rows = 1
VAN_CHK.Value = 0
ALL_CHK.Value = 1
smn_chk.Value = 0
mar_chk.Value = 0
list4.ColumnHeaders.Item(3).Text = ""
list4.ColumnHeaders.Item(4).Text = ""
list4.ColumnHeaders.Item(5).Text = ""
LDS_IDY.Enabled = True
LDS_IDY.SetFocus
MOD_REC = False
Frame1.Enabled = False
Command5.Enabled = True ' EDITED BY RADHIKA
Command7.Enabled = True     'END
End Sub

Private Sub FRM_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select From Date"
End Sub

Private Sub FRM_DAT_KeyDOWN(KeyCode As Integer, s As Integer)
If KeyCode = 13 Then TO_DAT.SetFocus
End Sub

Private Sub LDS_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Loading Sheet ID or Press <Enter> To Select ID's From Then List"
End Sub

Private Sub LDS_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
    If CHECKTB("lds.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If Len(Trim(LDS_IDY)) = 0 Then
        V_LDS_LST
    Else
        Dim RSlds As New ADODB.Recordset
        If Len(Trim(LDS_IDY)) <> 10 Then
            RSlds.Open "SELECT * FROM LDS WHER LDS_IDY LIKE 'LDS" & Tincrement(Trim(LDS_IDY)) & "'", con, adOpenKeyset, adLockPessimistic
        Else
            RSlds.Open "SELECT * FROM LDS WHER LDS_IDY LIKE '" & Trim(LDS_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        End If
        If RSlds.RecordCount > 0 Then
            If Len(Trim(LDS_IDY)) <> 10 Then
                LDS_IDY = "LDS" & Tincrement(LDS_IDY)
            Else
                LDS_IDY = Trim(LDS_IDY)
            End If
            LDS_IDY = Trim(LDS_IDY)
            LDS_IDY.Enabled = False
            FRM_DAT = RSlds!lst_Dat
            TO_DAT = RSlds!len_Dat
            If Len(Trim(RSlds!MAR_IDY)) > 0 Then
                Frame1.Enabled = True
                mar_chk.SetFocus
                MAR_IDY = Trim(RSlds!MAR_IDY)
                ALL_CHK.Value = 0
            ElseIf Len(Trim(RSlds!SMN_IDY)) > 0 Then
                Frame1.Enabled = True
                smn_chk.SetFocus
                SMN_IDY = Trim(RSlds!SMN_IDY)
                ALL_CHK.Value = 0
            ElseIf Len(Trim(RSlds!rou_idy)) > 0 Then
                Frame1.Enabled = True
                SETROUTE Trim(RSlds!rou_idy & "")
                Route.Enabled = True
                ALL_CHK.Value = 0
            Else
                ALL_CHK.Value = 1
            End If
            RSlds.Close
            LDS_IDY.Enabled = False
            displaydetails
        Else
            MsgBox "Invalid Delivery Sheet Number ", vbExclamation
            LDS_IDY.SetFocus
            SendKeys "{home}+{end}"
        End If
    End If
End If
End Sub

Private Sub List2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Click '<' To Deselect A Particular Invoice Or '<<' To Deselect All Invoices"
End Sub

Private Sub list4_DblClick()
list4_KeyPress 13
End Sub

Private Sub list4_KeyPress(K As Integer)
If K <> 13 Then Exit Sub
If Trim(LST_CAP) = "Select Market" Then
    SMN_IDY.Text = ""
    MAR_IDY = list4.SelectedItem
    list4.Visible = False
    SSTab1.Tab = 0
    Command5.SetFocus
ElseIf LST_CAP = "Select Van Variance ID" Then
    LDS_IDY = list4.SelectedItem
    list4.Visible = False
    SSTab1.Tab = 0
    LDS_IDY_KeyPress 13
ElseIf LST_CAP = "Select Van Reference" Then
    VAN_IDY = list4.SelectedItem
    VAN_IDY.Enabled = False
    list4.Visible = False
    SSTab1.Tab = 0
    Command5.SetFocus
Else
    MAR_IDY.Text = ""
    SMN_IDY = list4.SelectedItem
    list4.Visible = False
    SSTab1.Tab = 0
    Command5.SetFocus
End If

End Sub

Private Sub ls_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If ls.SortOrder = lvwAscending Then
ls.SortOrder = lvwDescending
ElseIf ls.SortOrder = lvwDescending Then
ls.SortOrder = lvwAscending
End If
ls.SortKey = CH.Index - 1
End Sub

Private Sub ls_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Click '>' To Select A Particular Invoice And '>>' To Select All Invoices"
End Sub

Private Sub MAR_CHK_keypress(K As Integer)
If K = 13 Then
    If mar_chk.Value = True Then
        MAR_IDY.Enabled = True
        MAR_IDY.SetFocus
    End If
End If
End Sub

Private Sub MAR_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Market ID or Press <Enter> To Select From The List"
End Sub

Private Sub MS_DblClick()
If CHECKTB("act.dbf,mar.dbf,inl.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSINV As New ADODB.Recordset
If MS.TextMatrix(MS.Row, 8) <> "" Then
    RSINV.Open "SELECT SUM(TOT_PCS),INL.INV_IDY,INL.CST_IDY FROM INL,INV WHERE INL.INV_IDY LIKE INV.INV_IDY AND INV.TXN_TYP = 'SAL' AND INV.BAT_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 8)) & "' AND INV.STK_TYP IN('Free','Salable','Replc.') and INV.INV_IDY IN(SELECT CST_IDY FROM NPO) GROUP BY INV.INV_IDY", con, adOpenKeyset, adLockPessimistic
    MS1.Rows = 1
    MS1.ColWidth(3) = 0
    MS1.ColWidth(4) = 1200
    For A = 1 To RSINV.RecordCount
        MS1.AddItem ""
        MS1.TextMatrix(A, 0) = Trim(RSINV(1))
        MS1.TextMatrix(A, 3) = Trim(RSINV(2))
        MS1.TextMatrix(A, 4) = Trim(RSINV(0))
        RSINV.MoveNext
    Next
End If

If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT ACT_IDY,ACT_NME,MAR_NME FROM ACT,MAR WHER ACT.MAR_IDY LIKE MAR.MAR_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS1.Rows - 1
    If RSINV.BOF = False Then RSINV.MoveFirst
    RSINV.Find "ACT_IDY like '" & Trim(MS1.TextMatrix(A, 3)) & "'"
    If RSINV.EOF = False Then
        MS1.TextMatrix(A, 1) = Trim(RSINV!ACT_NME)
        MS1.TextMatrix(A, 2) = Trim(RSINV!mar_nme)
    End If
Next
Tot_lbl = "0.00"
SSTab1.Tab = 2
MS1.SetFocus
End Sub

Private Sub Option3_keypress(K As Integer)
If K = 13 Then
SMN_IDY = ""
MAR_IDY = ""
Route.Enabled = True
Route.SetFocus
End If
End Sub



Private Sub Route_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Routes"
End Sub

Private Sub Route_kEYPRESS(K As Integer)
If K = 13 Then
Route.Enabled = False
Command5.SetFocus
End If
End Sub

Private Sub SMN_CHK_KeyPress(K As Integer)
If K = 13 Then
    If smn_chk.Value = True Then
        SMN_IDY.Enabled = True
        SMN_IDY.SetFocus
    Else
        mar_chk.SetFocus
    End If
End If
End Sub

Private Sub Mar_IDY_KEYPRESS(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
    If CHECKTB("mar.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If check("mar", "mar_IDY", Sincrement(Trim(MAR_IDY))) = True Then
        MAR_IDY = Sincrement(Trim(MAR_IDY))
        SMN_IDY.Text = ""
        Command1.SetFocus
    Else
        V_MAR_LST
    End If
End If
End Sub

Private Sub SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesman ID or Press <Enter> To Select From The List"
End Sub

Private Sub smn_idy_kEYPRESS(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
    If CHECKTB("smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If check("SMN", "SMN_IDY", Sincrement(Trim(SMN_IDY))) = True Then
        SMN_IDY = Sincrement(Trim(SMN_IDY))
        MAR_IDY.Text = ""
        Command1.SetFocus
    Else
        V_SMN_LST
    End If
End If
End Sub
Private Sub V_SMN_LST()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim lst As ListItem
LST_CAP = "Select Salesman"
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    list4.ColumnHeaders.Item(1).Text = "Salesman ID"
    list4.ColumnHeaders.Item(2).Text = "Salesman Name"
    For A = 1 To REC_SET.RecordCount
        Set lst = list4.ListItems.Add(, , Trim(REC_SET("smn_IDY")))
        lst.ListSubItems.Add , , Trim(REC_SET("smn_NME"))
        REC_SET.MoveNext
    Next
    REC_SET.Close
    SSTab1.Tab = 3
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Salesmen Not Found, Define Salesmen !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    SMN_IDY.SetFocus
    End If
End If
End Sub
Private Sub V_MAR_LST()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim lst As ListItem
LST_CAP = "Select Market"
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from mar WHERE mar_nme like '" & Trim(MAR_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    list4.ColumnHeaders.Item(1).Text = "Market ID"
    list4.ColumnHeaders.Item(2).Text = "Market Name"
    For A = 1 To REC_SET.RecordCount
        Set lst = list4.ListItems.Add(, , Trim(REC_SET("mar_IDY")))
        lst.ListSubItems.Add , , Trim(REC_SET("mar_NME"))
        REC_SET.MoveNext
    Next
    REC_SET.Close
    SSTab1.Tab = 3
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    MAR_IDY.SetFocus
    End If
End If
End Sub
Private Sub V_LDS_LST()
If CHECKTB("lds.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim lst As ListItem
LST_CAP = "Select Van Variance ID"
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from lds ", con, adOpenKeyset, adLockPessimistic
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    list4.ColumnHeaders.Item(1).Text = "ID"
    list4.ColumnHeaders.Item(2).Text = "Date"
    list4.ColumnHeaders.Item(1).Width = 2000
    list4.ColumnHeaders.Item(2).Width = 1500
    For A = 1 To REC_SET.RecordCount
        Set lst = list4.ListItems.Add(, , Trim(REC_SET("LDS_IDY")))
        lst.ListSubItems.Add , , Format(REC_SET("LDS_DAT"), "dd/mmm/yyyy")
        REC_SET.MoveNext
    Next
    REC_SET.Close
    SSTab1.Tab = 3
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Van Variance ID Not Found!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    LDS_IDY.SetFocus
    End If
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    If LDS_IDY.Enabled = True Then
        LDS_IDY.SetFocus
    Else
        DTPicker1.SetFocus
    End If
ElseIf SSTab1.Tab = 1 Then
    If MS.Rows > 1 Then
        MS.Row = 1
        MS.Col = 0
        MS.SetFocus
    End If
ElseIf SSTab1.Tab = 2 Then
    If MS1.Rows > 1 Then
        MS1.Row = 1
        MS1.Col = 0
        MS1.SetFocus
    End If
End If
End Sub

Private Sub TO_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select To Date"
End Sub

Private Sub TO_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
ALL_CHK.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
FRM_DAT.SetFocus
End If
End Sub

Public Sub LoadInvoices()
If CHECKTB("inl.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Label6.Visible = True
Label6.Refresh
Dim lst As ListItem
Dim RSINL As New ADODB.Recordset
M_FRM_VVR.MousePointer = vbHourglass
    If VAN_CHK.Value = 1 Then
        RSINL.Open "SELECT inv_idy,ACT_NME,NET_AMT FROM INL,ACT WHERE INL.ORD_IDY LIKE '" & Trim(VAN_IDY) & "' AND INL.CST_IDY LIKE ACT.ACT_IDY", con, adOpenKeyset, adLockPessimistic
    Else
    If ALL_CHK.Value = 1 Then
        RSINL.Open "SELECT inv_idy,act_nme,net_amt FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' and DCH_IDY = ''" & CTR1, con, adOpenKeyset, adLockPessimistic
    Else
        If smn_chk.Value = True Then
            RSINL.Open "SELECT inv_idy,act_nme,net_amt FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND INL.SMN_IDY LIKE '" & Trim(SMN_IDY) & "%'" & CTR1 & " AND RTRIM(LTRIM(DCH_IDY))=''", con, adOpenKeyset, adLockPessimistic
        ElseIf mar_chk.Value = True Then
            RSINL.Open "SELECT inv_idy,act_nme,net_amt FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND ACT.MAR_IDY LIKE '" & Trim(MAR_IDY) & "%'" & CTR1 & " AND RTRIM(LTRIM(DCH_IDY))=''", con, adOpenKeyset, adLockPessimistic
        ElseIf Option3.Value = True Then
            QRY1 = GETROUTE
            If Trim(QRY1) = "" Then QRY1 = "''"
            RSINL.Open "SELECT inv_idy,act_nme,net_amt FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND INL.ROU_IDY IN(" & QRY1 & ")" & CTR1 & "AND RTRIM(LTRIM(DCH_IDY))=''", con, adOpenKeyset, adLockPessimistic
        Else
            Exit Sub
        End If
     End If
    End If
M_FRM_VVR.MousePointer = vbNormal
If RSINL.EOF = True Then
   MsgBox "No Invoices Found For this Van Variance ID Combination", vbExclamation
   If LDS_IDY.Enabled = False Then
'        FRM_DAT.SetFocus
        DTPicker1.SetFocus
    Else
        LDS_IDY.SetFocus
   End If
   Exit Sub
End If
Label6.Visible = False
DoEvents
ls.ListItems.CLEAR
CTR = RSINL.RecordCount
COM_CTR = 0
ProgressBar1.Visible = True
Do While True
    If RSINL.EOF = True Then
       Exit Do
    End If
    Set lst = ls.ListItems.Add(, , Trim(RSINL!INV_IDY))
    lst.ListSubItems.Add , , Trim(RSINL!ACT_NME)
    lst.ListSubItems.Add , , convert(RSINL!NET_AMT)
    RSINL.MoveNext
    incpcb
Loop
If RSINL.RecordCount > 0 Then ' edited by radhika
    Command5.Enabled = False
End If                          'end
ProgressBar1.Visible = False
RSINL.Close
ls.SetFocus
Label6.Visible = False
End Sub
Public Sub incpcb()
COM_CTR = COM_CTR + 1
ProgressBar1.Value = (100 / CTR) * COM_CTR
End Sub

Public Sub saverecord()
Dim RSlds As New ADODB.Recordset
Dim RSINL As New ADODB.Recordset
If List2.ListCount = 0 Then Exit Sub

If MsgBox("Confirm Transaction Save?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
CANCELRECORD
Exit Sub
End If
If CHECKTB("inl.dbf,pnl.dbf,lds.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Me.MousePointer = vbHourglass
Me.Refresh
GETINVOICE
If MOD_REC = False Then
    INC = "LDS" & T7increment("IDY_LDS")
    GETINVOICE
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT CST_IDY FROM NPO)", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        For A = 1 To RSINL.RecordCount
            RSINL!DCH_IDY = Trim(INC)
            RSINL.MoveNext
        Next
        RSINL.UpdateBatch adAffectAllChapters
    End If
    For A = 0 To List2.ListCount - 1
        PNL.Execute "UPDATE PNL SET DCH_IDY ='" & INC & "' WHERE INV_IDY LIKE '" & Trim(List2.List(A)) & "'"
    Next
    If RSlds.State = 1 Then RSlds.Close
    RSlds.Open "SELECT * FROM LDS", con, adOpenKeyset, adLockPessimistic
    RSlds.AddNew
    RSlds!LDS_IDY = Trim(INC)
    MsgBox "New Delivery Sheet No. " & Trim(INC), vbExclamation
Else
    INC = Trim(LDS_IDY)
    con.Execute "UPDATE INL SET DCH_IDY ='' WHERE DCH_IDY ='" & Trim(LDS_IDY) & "'", d
    PNL.Execute "UPDATE PNL SET DCH_IDY ='' WHERE DCH_IDY ='" & Trim(LDS_IDY) & "'", d
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT CST_IDY FROM NPO)", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        For A = 1 To RSINL.RecordCount
            RSINL!DCH_IDY = Trim(INC)
            RSINL.Update
            RSINL.MoveNext
        Next
    End If
    For A = 0 To List2.ListCount - 1
        PNL.Execute "UPDATE PNL SET DCH_IDY ='" & INC & "' WHERE INV_IDY LIKE '" & Trim(List2.List(A)) & "'"
    Next
    RSlds.Open "SELECT * FROM LDS WHERE LDS_IDY LIKE '" & Trim(LDS_IDY) & "'", con, adOpenKeyset, adLockPessimistic
End If
    RSlds!LDS_DAT = Date
    RSlds!LDS_TIM = Format(Time, "h:m:s")
    RSlds!lst_Dat = FRM_DAT
    RSlds!len_Dat = TO_DAT
    If ALL_CHK.Value = 0 Then
        If mar_chk.Value = True Then
            RSlds!MAR_IDY = Trim(MAR_IDY)
            RSlds!SMN_IDY = ""
            RSlds!rou_idy = ""
        ElseIf smn_chk.Value = True Then
            RSlds!SMN_IDY = Trim(SMN_IDY)
            RSlds!MAR_IDY = ""
            RSlds!rou_idy = ""
        ElseIf Option3.Value = True Then
            RSlds!rou_idy = GETROUTE
            RSlds!MAR_IDY = ""
            RSlds!SMN_IDY = ""
        End If
    Else
        RSlds!rou_idy = ""
        RSlds!MAR_IDY = ""
        RSlds!SMN_IDY = ""
    End If
    RSlds.Update
    RSlds.Close
Me.MousePointer = vbDefault
CANCELRECORD
End Sub

Public Function GETINVOICE()
If CHECKTB("npo.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Function
End If
con.Execute "DELETE FROM NPO"
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM NPO", con, adOpenKeyset, adLockPessimistic
For A = 0 To List2.ListCount - 1
    rs.AddNew
        rs(0) = Trim(List2.List(A))
    rs.Update
Next
End Function

Public Sub displaydetails()
If CHECKTB("inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSINL As New ADODB.Recordset
RSINL.Open "select * from inl where dch_idy like '" & Trim(LDS_IDY) & "'", con, adOpenKeyset, adLockPessimistic
List2.CLEAR
MOD_REC = True
For A = 1 To RSINL.RecordCount
    List2.AddItem RSINL!INV_IDY
    RSINL.MoveNext
Next
RSINL.Close
FRM_DAT.SetFocus
End Sub

Public Sub Mod_LoadInvoices()
Dim lst As ListItem
Dim RSINL As New ADODB.Recordset
'    If Len(Trim(GETINVOICE)) = 0 Then
'        CTR = ""
'    Else
'        CTR = "and INV_IDY not in (" & GETINVOICE & ")"
'    End If
'    If ALL_CHK.Value = 1 Then
'        RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL'" & CTR, CON, adOpenKeyset, adlockpessimistic
'    Else
'        If smn_chk.Value = True Then
'            RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND INL.SMN_IDY LIKE '" & Trim(SMN_IDY) & "'" & CTR, CON, adOpenKeyset, adlockpessimistic
'        ElseIf mar_chk.Value = True Then
'            RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND ACT.MAR_IDY LIKE '" & Trim(MAR_IDY) & "'" & CTR, CON, adOpenKeyset, adlockpessimistic
'        ElseIf Option3.Value = True Then
'            For A = 1 To Route.ListCount - 1
'                If Route.Selected(A - 1) = True Then
'                    If Len(Trim(query)) = 0 Then
'                        query = "'" & Route.List(A - 1) & "'"
'                    Else
'                        query = query & ",'" & Route.List(A - 1) & "'"
'                    End If
'                End If
'            Next
'            RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND INL.ROU_IDY IN(" & query & ") AND SUBSTR(INV_IDY,1,3)='SAL' " & CTR, CON, adOpenKeyset, adlockpessimistic
'        End If
'    End If

For A = 1 To RSINL.RecordCount
    Set lst = ls.ListItems.Add(, , Trim(RSINL!INV_IDY))
    lst.ListSubItems.Add , , Trim(RSINL!ACT_NME)
    lst.ListSubItems.Add , , convert(RSINL!NET_AMT)
    RSINL.MoveNext
Next
RSINL.Close
FRM_DAT.SetFocus
ls.SetFocus
End Sub
Public Function GETROUTE() As String
For A = 1 To Route.ListCount - 1
    If Route.Selected(A - 1) = True Then
        If Len(Trim(Query)) = 0 Then
            Query = "'" & Route.List(A - 1) & "'"
        Else
            Query = Query & ",'" & Route.List(A - 1) & "'"
        End If
    End If
Next
GETROUTE = Query
End Function

Public Sub SETROUTE(str As String)
For A = 1 To Len(Trim(str))
    For b = 0 To Route.ListCount - 1
    If Trim(Route.List(A - 1)) = Trim(Mid(Trim(str), b + 1, 1)) Then Route.Selected(A - 1) = True
    Next
Next
End Sub

Public Sub ROUTECLEAR()
For A = 0 To Route.ListCount - 1
    Route.Selected(A) = False
Next
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
newrecord
ElseIf Button.KEY = "s" Then
saverecord
ElseIf Button.KEY = "e" Then
Unload Me
ElseIf Button.KEY = "p" Then
'PRINTRECORD
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "smn" Then
V_SMN_LST
ElseIf Button.KEY = "ed" Then
    If M_FRM_VVR.ActiveControl.name = "LDS_IDY" Then
        LST_CAP = "Select Van Variance ID"
        V_LDS_LST
    ElseIf M_FRM_VVR.ActiveControl.name = "M_CST_IDY" Then
        V_LST_CAP = "Select Retailer"
        'V_DLR_LST
    ElseIf M_FRM_VVR.ActiveControl.name = "M_SMN_IDY" Then
        V_LST_CAP = "Select Salesman"
        'V_SMN_LST
    ElseIf M_FRM_VVR.ActiveControl.name = "M_ACT_IDY" Then
        V_LST_CAP = "Select Account"
        'V_ACT_LST
    ElseIf M_FRM_VVR.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        'V_PRD_LST
    End If
ElseIf Button.KEY = "DAY" Then
'    If Button.Value = tbrPressed Then
'        Text1.Visible = True
'        Text1 = Format(Date, "dddd")
'        Text1.BackColor = vbBlue
'        Text1.ForeColor = vbWhite
'    Else
'        Text1 = ""
'        Text1.Visible = False
'    End If
ElseIf Button.KEY = "smn1" Then
    If Button.Value = tbrPressed Then
'        Text2.Visible = True
        list1.Tag = "Salesman"
        V_SMN_LST
'        Text2.BackColor = vbBlue
'        Text2.ForeColor = vbWhite
    Else
'        Text2 = ""
'        Text2.Visible = False
    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
If check("prt", "prt_nme", "M_FRM_VVR") = False Then
Check3.Value = 0
Else
Check3.Value = 1
End If
Check3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Public Sub MODLOADINVOICES()
If CHECKTB("act.dbf,inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSINL As New ADODB.Recordset
If ALL_CHK.Value = 1 Then
    RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND DCH_IDY IN('','" & Trim(LDS_IDY) & "')", con, adOpenKeyset, adLockPessimistic
Else
    If smn_chk.Value = True Then
        RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND INL.SMN_IDY LIKE '" & Trim(SMN_IDY) & "%'" & CTR & " AND DCH_IDY IN('','" & Trim(LDS_IDY) & "')", con, adOpenKeyset, adLockPessimistic
    ElseIf mar_chk.Value = True Then
        RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND ACT.MAR_IDY LIKE '" & Trim(MAR_IDY) & "%'" & CTR & " AND DCH_IDY IN('','" & Trim(LDS_IDY) & "')", con, adOpenKeyset, adLockPessimistic
    ElseIf Option3.Value = True Then
        RSINL.Open "SELECT * FROM INL,ACT WHERE INL.CST_IDY LIKE ACT.ACT_IDY AND INV_DAT >=CTOD('" & FRM_DAT & "') AND INV_DAT <=CTOD('" & TO_DAT & "') AND TXN_TYP LIKE 'SAL' AND INL.ROU_IDY IN(" & GETROUTE & ")" & CTR & " AND DCH_IDY IN('','" & Trim(LDS_IDY) & "')", con, adOpenKeyset, adLockPessimistic
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If CHECKTB("inl.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
With MDIForm1
    If Len(Trim(LDS_IDY)) = 0 Then
        MsgBox "Select Van Variance ID To Print", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        LDS_IDY.Enabled = True
        LDS_IDY.SetFocus
        Exit Sub
    End If
    SAL_QRY = "INL->DCH_IDY = '" & Trim(LDS_IDY) & "' .and. INV->STK_TYP <> 'Shortage' .and. INV->STK_TYP <> 'Damage' .and. INV->STK_TYP <> 'TakeBack' .and. INV->STK_TYP <> 'Breakage' .and. inv->stk_typ <> 'Leakage'"
If ButtonMenu.KEY = "LDS_BAT" Then
    OpenReport CURDIR & "\VVR_RPT.RP1", "Variance Sheet", 2, SAL_QRY, 6
ElseIf ButtonMenu.KEY = "LDS_PRD" Then
    OpenReport CURDIR & "\VVR_RPT.RP1", "Variance Sheet Product", 2, SAL_QRY, 6
ElseIf ButtonMenu.KEY = "LDS_INV" Then
    OpenReport CURDIR & "\VVR_RPT.RP1", "Variance Sheet Invoice Track", 2, SAL_QRY, 6
ElseIf ButtonMenu.KEY = "LDS_PTR" Then
    SAL_QRY = "INL->DCH_IDY = '" & Trim(LDS_IDY) & "'"
    OpenReport CURDIR & "\VVR_RPT.RP1", "Variance Sheet Product Track", 2, SAL_QRY, 6
ElseIf ButtonMenu.KEY = "TRP_SHT" Then
    SAL_QRY = "INL->DCH_IDY = '" & Trim(LDS_IDY) & "' .and. INV->STK_TYP <> 'Shortage'"
    OpenReport CURDIR & "\VVR_RPT.RP1", "Variance Trip Sheet", 2, SAL_QRY, 6
End If
End With
End Sub

Private Sub VAN_CHK_Click()
If VAN_CHK.Value = 1 Then
CUC (False)
Else
CUC (True)
End If
End Sub

Private Sub VAN_CHK_KeyPress(K As Integer)
If K = 13 Then
If VAN_CHK.Value = 1 Then
VAN_IDY.Enabled = True
VAN_IDY.SetFocus
Else
FRM_DAT.SetFocus
End If
End If
End Sub
Public Sub CUC(b As Boolean)
FRM_DAT.Enabled = b
TO_DAT.Enabled = b
ALL_CHK.Enabled = b
smn_chk.Enabled = b
SMN_IDY.Enabled = b
mar_chk.Enabled = b
MAR_IDY.Enabled = b
Option3.Enabled = b
Route.Enabled = b
End Sub
Private Sub VAN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Van Ref. ID or Press <Enter> To Select From The List"
End Sub
Private Sub VAN_LST()
If CHECKTB("van.dbf,smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rsVan As New ADODB.Recordset
Dim ls As ListItem
rsVan.Open "SELECT * FROM VAN,smn WHERE van.smn_idy like smn.smn_idy and VAN_CLS='S' AND VAN_CHK = 'V' GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
If rsVan.RecordCount > 0 Then
    list4.ListItems.CLEAR
    For A = 1 To rsVan.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rsVan!INV_IDY))
        ls.ListSubItems.Add , , Format(rsVan!INV_DAT, "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(rsVan("SMN_NME") & "")
        'ls.ListSubItems.Add , , Trim(rsVan("SMN_AD4") & "")
        ls.ListSubItems.Add , , Trim(rsVan("ROU_IDY") & "")
        ls.ListSubItems.Add , , Trim(rsVan("MNL_REF") & "")
        rsVan.MoveNext
    Next
    SSTab1.Tab = 3
    LST_CAP = "Select Van Reference"
    list4.ColumnHeaders.Item(3).Text = "Salesman"
    list4.ColumnHeaders.Item(4).Text = "Route"
    list4.ColumnHeaders.Item(5).Text = "Reference"
    list4.Visible = True
    list4.SetFocus
Else
    MsgBox "Van Reference Not found!", vbExclamation
    SSTab1.Tab = 0
    SendKeys "{home}+{end}"
    VAN_IDY.SetFocus
End If
End Sub

Private Sub VAN_IDY_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii = 13 Then
    If CHECKTB("van.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If

    If Len(Trim(VAN_IDY)) = 0 Then
        VAN_LST
    Else
        If querycheck("VAN", "INV_IDY", "RSI" & Tincrement(VAN_IDY), "AND VAN_CLS='S' AND VAN_CHK='V'") = True Then
            VAN_IDY = "RSI" & Tincrement(VAN_IDY)
            VAN_IDY.Enabled = False
            'M_CST_IDY.SetFocus
        Else
            VAN_LST
        End If
    End If
End If
End Sub
