VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_RSI 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van Stock Issue"
   ClientHeight    =   5610
   ClientLeft      =   1425
   ClientTop       =   1845
   ClientWidth     =   8820
   ForeColor       =   &H00404040&
   Icon            =   "M_FRM_RSI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   2280
      ScaleHeight     =   1875
      ScaleWidth      =   4035
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   375
         Left            =   720
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1365
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2280
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1365
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   2040
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24444931
         CurrentDate     =   37315
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "As On Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   73
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         Caption         =   "Select Date to Bring Forward Opening Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   69
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   2040
      ScaleHeight     =   3555
      ScaleWidth      =   4635
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   0
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Sales Account View"
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
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SAL_REP"
                  Object.Tag             =   "SAL_REP"
                  Text            =   "&a Sales Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CAS_REC"
                  Text            =   "&b Cases Reconcile"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UNT_REC"
                  Text            =   "&c Units Reconcile"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_STL"
                  Text            =   "&d Sale Estimate"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Add Opening Stock"
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
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bill"
            Object.ToolTipText     =   "Convert to Bill"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RS"
            Object.ToolTipText     =   "Return Stocks"
            ImageIndex      =   13
         EndProperty
      EndProperty
      Begin VB.CheckBox Pieces 
         Caption         =   "Check to do in Cases"
         Height          =   195
         Left            =   4800
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483637
      ForeColor       =   8388608
      TabCaption(0)   =   "&Voucher Details"
      TabPicture(0)   =   "M_FRM_RSI.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TRP_NOS"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "M_CST_IDY"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "M_SMN_IDY"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "M_INV_IDY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "M_INV_DAT"
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(7)=   "M_CST_NME"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(11)=   "Label16"
      Tab(0).Control(12)=   "M_GRS_AMT"
      Tab(0).Control(13)=   "Label40"
      Tab(0).Control(14)=   "M_SMN_NME"
      Tab(0).Control(15)=   "Label23"
      Tab(0).Control(16)=   "Label1"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "&Products"
      TabPicture(1)   =   "M_FRM_RSI.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "amount"
      Tab(1).Control(3)=   "MS"
      Tab(1).Control(4)=   "M_Prd_Nme"
      Tab(1).Control(5)=   "Total"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Collection &Details"
      TabPicture(2)   =   "M_FRM_RSI.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "View"
      TabPicture(3)   =   "M_FRM_RSI.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label49"
      Tab(3).Control(1)=   "list4"
      Tab(3).Control(2)=   "list1"
      Tab(3).Control(3)=   "list5"
      Tab(3).Control(4)=   "Frame1"
      Tab(3).Control(5)=   "LIST3"
      Tab(3).Control(6)=   "list2"
      Tab(3).Control(7)=   "LIST7"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "&Settlement"
      TabPicture(4)   =   "M_FRM_RSI.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "MSS"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "&Empty Shell Details"
      TabPicture(5)   =   "M_FRM_RSI.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SHL_SET"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "txtempshl"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "ms2"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Empt&y Shells Settlement"
      TabPicture(6)   =   "M_FRM_RSI.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command6"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "ms1"
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Cash Currency Details"
         Height          =   3975
         Left            =   4560
         TabIndex        =   81
         Top             =   600
         Width           =   3855
         Begin VB.TextBox TXTTOTAL 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   102
            Text            =   " "
            Top             =   3550
            Width           =   1095
         End
         Begin VB.TextBox txt500 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   101
            Text            =   " "
            Top             =   661
            Width           =   1095
         End
         Begin VB.TextBox txt100 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   100
            Text            =   " "
            Top             =   1022
            Width           =   1095
         End
         Begin VB.TextBox txt50 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   99
            Text            =   " "
            Top             =   1383
            Width           =   1095
         End
         Begin VB.TextBox txt20 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   98
            Text            =   " "
            Top             =   1744
            Width           =   1095
         End
         Begin VB.TextBox txt10 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   97
            Text            =   " "
            Top             =   2105
            Width           =   1095
         End
         Begin VB.TextBox txt5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   96
            Text            =   " "
            Top             =   2466
            Width           =   1095
         End
         Begin VB.TextBox txt2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   95
            Text            =   " "
            Top             =   2827
            Width           =   1095
         End
         Begin VB.TextBox txt1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   94
            Text            =   " "
            Top             =   3188
            Width           =   1095
         End
         Begin VB.TextBox txtchange 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   93
            Text            =   " "
            Top             =   3550
            Width           =   1095
         End
         Begin VB.TextBox txt1000 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   92
            Text            =   " "
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label LBLCHG 
            AutoSize        =   -1  'True
            Caption         =   "Change"
            Height          =   195
            Left            =   240
            TabIndex        =   91
            Top             =   3600
            Width           =   555
         End
         Begin VB.Label LBL1 
            AutoSize        =   -1  'True
            Caption         =   "1s"
            Height          =   195
            Left            =   240
            TabIndex        =   90
            Top             =   3240
            Width           =   165
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "2s"
            Height          =   195
            Left            =   240
            TabIndex        =   89
            Top             =   2880
            Width           =   165
         End
         Begin VB.Label LBL5 
            AutoSize        =   -1  'True
            Caption         =   "5s"
            Height          =   195
            Left            =   240
            TabIndex        =   88
            Top             =   2520
            Width           =   165
         End
         Begin VB.Label LBL10 
            AutoSize        =   -1  'True
            Caption         =   "10s"
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label LBL20 
            AutoSize        =   -1  'True
            Caption         =   "20s"
            Height          =   195
            Left            =   240
            TabIndex        =   86
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label LBL50 
            AutoSize        =   -1  'True
            Caption         =   "50s"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label LBL100 
            AutoSize        =   -1  'True
            Caption         =   "100s"
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label LBL500 
            AutoSize        =   -1  'True
            Caption         =   "500s"
            Height          =   195
            Left            =   240
            TabIndex        =   83
            Top             =   720
            Width           =   345
         End
         Begin VB.Label LBL1000 
            AutoSize        =   -1  'True
            Caption         =   "1000s"
            Height          =   195
            Left            =   240
            TabIndex        =   82
            Top             =   345
            Width           =   435
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ge&nerate"
         Height          =   375
         Left            =   -67560
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton SHL_SET 
         Caption         =   "Upd&ate Shells"
         Height          =   375
         Left            =   -67560
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtempshl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74520
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox TRP_NOS 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "M_FRM_RSI.frx":0506
         Left            =   -72360
         List            =   "M_FRM_RSI.frx":050D
         TabIndex        =   60
         Top             =   2350
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "P&rint"
         Height          =   375
         Left            =   -67440
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   -67440
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Collection Details"
         Height          =   3975
         Left            =   360
         TabIndex        =   37
         Top             =   600
         Width           =   3735
         Begin VB.TextBox M_COM_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            TabIndex        =   79
            Text            =   "0.00"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox M_BRK_QTY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   65
            Text            =   "0.00"
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox M_LKG_QTY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "0.00"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox M_TAX_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox M_CAS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            TabIndex        =   54
            Text            =   "0.00"
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox M_EXP_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox M_CRT_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "0.00"
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox M_CHQ_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            TabIndex        =   48
            Text            =   "0.00"
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox M_DIS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox M_NET_SAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox M_FRE_QTY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox M_GRS_SAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Commission Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   75
            TabIndex        =   80
            Top             =   1748
            Width           =   1110
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Breakage"
            Height          =   195
            Left            =   4815
            TabIndex        =   67
            Top             =   1035
            Width           =   810
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Leakage"
            Height          =   195
            Left            =   4875
            TabIndex        =   66
            Top             =   1395
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Taxes && Others"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   105
            TabIndex        =   52
            Top             =   2468
            Width           =   1080
         End
         Begin VB.Label LLabel12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Expense Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   285
            TabIndex        =   51
            Top             =   2108
            Width           =   900
         End
         Begin VB.Label Label11Y 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Credit"
            Height          =   195
            Left            =   780
            TabIndex        =   49
            Top             =   3555
            Width           =   405
         End
         Begin VB.Label Label10H 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cheque Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   345
            TabIndex        =   47
            Top             =   3195
            Width           =   840
         End
         Begin VB.Label Label9L 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cash Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   540
            TabIndex        =   46
            Top             =   2828
            Width           =   645
         End
         Begin VB.Label LBLLabel8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Discount Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   270
            TabIndex        =   44
            Top             =   1388
            Width           =   915
         End
         Begin VB.Label LBLLabel7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Net Sales"
            Height          =   195
            Left            =   495
            TabIndex        =   42
            Top             =   1028
            Width           =   690
         End
         Begin VB.Label LBLLabel5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Free Quantity"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   668
            Width           =   945
         End
         Begin VB.Label LBLLabel4 
            AutoSize        =   -1  'True
            Caption         =   "Gross Sales"
            Height          =   195
            Left            =   345
            TabIndex        =   38
            Top             =   308
            Width           =   840
         End
      End
      Begin MSComctlLib.ListView LIST7 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
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
            Text            =   "Account Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   1080
         Left            =   -68400
         TabIndex        =   32
         Top             =   480
         Width           =   2055
         Begin VB.CheckBox Check3 
            Caption         =   "&Complete Returns"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   300
            Width           =   1815
         End
         Begin VB.CheckBox Check4 
            Caption         =   "C&hallan Close"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   700
            Width           =   1815
         End
         Begin VB.CheckBox Trolley 
            Caption         =   "P&ush Cart"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   180
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
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
            Text            =   "Account Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView LIST3 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vehicle Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Salesman"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Market"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Route"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox M_CST_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "M_FRM_RSI.frx":0516
         Left            =   -72960
         List            =   "M_FRM_RSI.frx":0523
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74880
         TabIndex        =   4
         Top             =   4500
         Width           =   8295
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73800
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox M_SMN_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   -72360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24444931
         CurrentDate     =   37099
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   375
         Left            =   -69240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         PromptInclude   =   0   'False
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4275
         Left            =   -74940
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   19
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
      End
      Begin MSComctlLib.ListView list5 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Challan ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Salesman"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Market"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Route"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product (SKU) Name"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "MRP"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Sale Rate"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Short Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Batch ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Batch Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Pur.Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Mfg. Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Exp. Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSS 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   16711680
         WordWrap        =   -1  'True
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid ms1 
         Height          =   4035
         Left            =   -74910
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   960
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   7117
         _Version        =   393216
         Cols            =   19
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid ms2 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   960
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   7117
         _Version        =   393216
         Cols            =   19
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   2
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trip No"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73065
         TabIndex        =   59
         Top             =   2410
         Width           =   525
      End
      Begin VB.Label M_Prd_Nme 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74880
         TabIndex        =   29
         Top             =   4800
         Width           =   45
      End
      Begin VB.Label M_CST_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   -70920
         TabIndex        =   26
         Top             =   1800
         Width           =   2460
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73290
         TabIndex        =   25
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   195
         Left            =   -73095
         TabIndex        =   22
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   -72900
         TabIndex        =   21
         Top             =   1275
         Width           =   345
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73455
         TabIndex        =   20
         Top             =   2955
         Width           =   900
      End
      Begin VB.Label M_GRS_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   -72360
         TabIndex        =   19
         Top             =   3405
         Width           =   1215
      End
      Begin VB.Label Label49 
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
         TabIndex        =   18
         Top             =   540
         Width           =   8295
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   -69960
         TabIndex        =   17
         Top             =   4350
         Width           =   60
      End
      Begin VB.Label M_SMN_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   -70920
         TabIndex        =   16
         Top             =   2880
         Width           =   2460
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -66525
         TabIndex        =   15
         Top             =   4785
         Width           =   45
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   -74040
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Challan Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73680
         TabIndex        =   13
         Top             =   668
         Width           =   1125
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":0530
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":0A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":0E3C
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":1190
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":17C8
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":1B1C
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":1F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":2270
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":267C
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":2788
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":2ADC
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":2EF0
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI.frx":343C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_RSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, bslab, REC_SET1 As New ADODB.Recordset, rs4 As New ADODB.Recordset, REC_SET5 As New ADODB.Recordset, MSROW, INVIDY, INC, dlr_idy
Dim RET_STK As Boolean, save_rec, ACTIDY
Dim RSCUR As New ADODB.Recordset
Dim invid As String, FLD_CHK As String, IDN As Boolean
Dim BMAKEBILL As Boolean
Dim RSVAN As New ADODB.Recordset
Dim AR(5, 3)
Dim RSCOM As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset

Private Sub CREATE()
MOD_REC = False
M_INV_IDY = ""
If Check3.Value = 1 And Trolley.Value = 1 Then
MakeBill
Else
saverecord
End If
End Sub

Public Sub GRS_CAL()
On Error Resume Next
M_GRS_AMT = convert(AMT(Val(Total)))
M_NET_AMT = convert(AMT(Round(Val(M_GRS_AMT) - Val(billdisamt), 2)))
M_NET_AMT = convert(AMT(Round(Val(M_NET_AMT) - Val(addisamt), 2)))
M_NET_AMT = convert(AMT(Round(Val(M_NET_AMT) + Val(taxamt), 2)))
M_NET_AMT = convert(AMT(Val(M_NET_AMT) - Val(damageamt)))
End Sub

Private Sub saverecord()
On Error Resume Next
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "Effmcg") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("smn.dbf,act.dbf,van.dbf,prd.dbf,brd.dbf,inl.dbf,snv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("smn", "smn_idy", Trim(M_SMN_IDY)) = False Then
    MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
    M_SMN_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If
If check("act", "act_idy", Trim(M_CST_IDY)) = False Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "Effmcg"
    M_CST_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If
If Check3.Enabled = False And Check4.Enabled = False Then
    MsgBox "Changes not saved", vbInformation
    INC = Trim(M_INV_IDY)
    GoTo VAN_PRT_CHK:
    CANCELRECORD
    Exit Sub
End If
If Trolley.Value = 1 Then
Else
    If TRP_NOS = "All" And MOD_REC = False And Check3.Value = 0 Then
        MsgBox "Invalid Trip Number", vbExclamation
        SSTab1.Tab = 0
        TRP_NOS.Enabled = True
        TRP_NOS.SetFocus
        Exit Sub
    End If
End If
For A = 0 To TRP_NOS.ListCount
    If TRP_NOS.List(A) = "All" Then
        b = 0
    Else
        b = 1
        Exit For
    End If
Next
If Check3.Value = 1 And MS.TextMatrix(1, 1) = "All" And b = 0 Then
    If MsgBox("Trip Wise Data will be Merged ?", vbQuestion + vbYesNo) = vbNo Then
        SSTab1.Tab = 0
        Exit Sub
    End If
End If
If Check3.Value = 1 And MS.TextMatrix(1, 1) = "All" And Check3.Enabled = True Then
    If MsgBox("Trip Wise Data will be Merged ?", vbQuestion + vbYesNo) = vbNo Then
        SSTab1.Tab = 0
        Exit Sub
    End If
End If
If CHECKSTOCK = True Then Exit Sub
If REC_SET1.State = 1 Then REC_SET1.Close
If REC_SET.State = 1 Then REC_SET.Close
If MOD_REC = True Then
    If Check3.Value = 1 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from van,PRD where VAN.PRD_IDY LIKE PRD.PRD_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset
        For A = 1 To REC_SET.RecordCount
            retqty = MTS(Val(REC_SET!ret_qty & ""), Val(REC_SET!sub_unt))
            brkqty = MTS(Val(REC_SET!brk_qty & ""), Val(REC_SET!sub_unt))
            lkgqty = MTS(Val(REC_SET!lkg_qty & ""), Val(REC_SET!sub_unt))
            ret_qty = STM(Val(retqty) + Val(brkqty) + Val(lkgqty), REC_SET!sub_unt)
            CON.Execute "update prd where prd_idy like '" & REC_SET("prd_idy") & "' set cur_lev = " & CONQTY(Val(ret_qty), REC_SET("prd_idy"))
            CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & REC_SET("prd_idy") & "' set PRD_PCS = PRD_PCS - " & MTS(Val(ret_qty), Val(REC_SET("SUB_UNT")))
            CON.Execute "update Brd where BAT_idy like '" & REC_SET("BAT_idy") & "' set PRD_QTY = " & CONBRDQTY(Val(ret_qty), REC_SET("BAT_idy"))
            REC_SET.MoveNext
        Next

    ElseIf Check3.Value = 0 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from van,PRD where VAN.PRD_IDY LIKE PRD.PRD_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset
        For A = 1 To REC_SET.RecordCount
            CON.Execute "update prd where prd_idy like '" & REC_SET("prd_idy") & "' set cur_lev = " & ADDQTY(Val(REC_SET("TOT_qty")), REC_SET("prd_idy"))
            CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & REC_SET("prd_idy") & "' set PRD_PCS = PRD_PCS + " & MTS(Val(REC_SET("TOT_qty")), Val(REC_SET("SUB_UNT")))
            CON.Execute "update Brd where BAT_idy like '" & REC_SET("BAT_idy") & "' set PRD_QTY = " & ADDBRDQTY(Val(REC_SET("TOT_qty")), REC_SET("BAT_idy"))
            REC_SET.MoveNext
        Next
    End If
    CON.Execute "delete from van where inv_idy like '" & Trim(M_INV_IDY) & "'", d
    INC = Trim(M_INV_IDY)
Else
    If Trim(MS.TextMatrix(1, 0)) <> "" Then
        INC = "RSI" & T7increment("IDY_RSI")
        INVIDY = INC
        If Trim(M_INV_IDY) = "" And Trolley.Value = 1 Then
            CON.Execute "UPDATE INL SET ORD_IDY = '" & Trim(INC) & "' WHERE INV_IDY = '" & Trim(invid) & "'"
        End If
    Else
         CANCELRECORD
        Exit Sub
    End If
End If
If REC_SET.State = 1 Then REC_SET.Close
If RSCUR.State = 1 Then RSCUR.Close
RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY = ''", CON, adOpenKeyset, adLockOptimistic
REC_SET.Open "select * from van where inv_idy = ''", CON, adOpenKeyset, adLockOptimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
        REC_SET.AddNew
        REC_SET("inv_idy") = Trim(INC)
        REC_SET("inv_dat") = Trim(M_INV_DAT)
        If MS.TextMatrix(A, 1) = "All" Then
            REC_SET("TRP_IDY") = 0
        Else
            REC_SET("TRP_IDY") = Val(MS.TextMatrix(A, 1))
        End If
        REC_SET("prd_idy") = Trim(MS.TextMatrix(A, 10))
        REC_SET("ACT_idy") = Trim(M_ACT_IDY)
        REC_SET("CST_idy") = Trim(M_CST_IDY)
        REC_SET("SMN_idy") = Trim(M_SMN_IDY)
        If Pieces.Value = 0 Then
            MS.TextMatrix(A, 2) = STM(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 3) = STM(Val(MS.TextMatrix(A, 3)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 4) = STM(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 6) = STM(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 5) = STM(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 15) = Val(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 16) = STM(Val(MS.TextMatrix(A, 16)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 17) = STM(Val(MS.TextMatrix(A, 17)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 18) = STM(Val(MS.TextMatrix(A, 18)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 8) = convert(Val(MS.TextMatrix(A, 8)) * Val(MS.TextMatrix(A, 12)))
        End If
        REC_SET("OPN_QTY") = Val(MS.TextMatrix(A, 2))
        REC_SET("ISS_QTY") = Val(MS.TextMatrix(A, 3))
        REC_SET("tot_qty") = Val(MS.TextMatrix(A, 4))
        REC_SET("fre_qty") = Val(MS.TextMatrix(A, 16))
        REC_SET("brk_qty") = Val(MS.TextMatrix(A, 17))
        REC_SET("lkg_qty") = Val(MS.TextMatrix(A, 18))
        REC_SET("BIL_QTY") = Val(MS.TextMatrix(A, 6))
        REC_SET("RET_QTY") = Val(MS.TextMatrix(A, 5))
        REC_SET("BAT_idy") = Trim(MS.TextMatrix(A, 13))
        REC_SET!PRD_rat = Val(MS.TextMatrix(A, 15))
        If Check3.Value = 1 Then
            REC_SET!van_cls = "S"
        Else
            REC_SET!van_cls = "N"
        End If
        If Check4.Value = 1 Or save_rec = "Save" Then
        save_rec = ""
        REC_SET!van_cls = "C"
        End If
        REC_SET!COL_AMT = Val(MS.TextMatrix(A, 9))
        REC_SET!CAS_AMT = Val(M_CAS_AMT)
        REC_SET!CHQ_AMT = Val(M_CHQ_AMT)
        REC_SET!DIS_AMT = Val(M_DIS_AMT)
        REC_SET!EXP_AMT = Val(M_EXP_AMT)
        REC_SET!ADD_AMT = Val(M_TAX_AMT)
        REC_SET!COM_AMT = Val(M_COM_AMT)
        If Trolley.Value = 1 Then
            REC_SET!VAN_CHK = "P"
        Else
            REC_SET!VAN_CHK = "V"
        End If
        REC_SET.Update
        If Check3.Value = 1 Then
            retqty = MTS(Val(Trim(MS.TextMatrix(A, 5) & "")), Val(MS.TextMatrix(A, 12)))
            brkqty = MTS(Val(Trim(MS.TextMatrix(A, 17) & "")), Val(MS.TextMatrix(A, 12)))
            lkgqty = MTS(Val(Trim(MS.TextMatrix(A, 18) & "")), Val(MS.TextMatrix(A, 12)))
            ret_qty = STM(Val(retqty) + Val(brkqty) + Val(lkgqty), Val(MS.TextMatrix(A, 12)))
            CON.Execute "update prd where prd_idy like '" & Trim(MS.TextMatrix(A, 10)) & "' set cur_lev = " & ADDQTY(QTY(Trim(ret_qty & "")), Trim(MS.TextMatrix(A, 10)))
            CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS + " & MTS(Val(ret_qty), Val(MS.TextMatrix(A, 12)))
            CON.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & ADDBRDQTY(QTY(Val(ret_qty)), Trim(MS.TextMatrix(A, 13)))
        ElseIf Check3.Value = 0 Then
            CON.Execute "update prd where prd_idy like '" & Trim(MS.TextMatrix(A, 10)) & "' set cur_lev = " & CONQTY(QTY(Trim(MS.TextMatrix(A, 4) & "")), Trim(MS.TextMatrix(A, 10)))
            CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
            CON.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(QTY(Val(MS.TextMatrix(A, 4))), Trim(MS.TextMatrix(A, 13)))
        End If
    End If
Next
RSCUR.AddNew    ' EDITED BY RADHIKA
    RSCUR(0) = Trim(INC)
    RSCUR(1) = convert(Val(Trim(txt1000) & ""))
    RSCUR(2) = convert(Val(Trim(txt500) & ""))
    RSCUR(3) = convert(Val(Trim(txt100) & ""))
    RSCUR(4) = convert(Val(Trim(txt50) & ""))
    RSCUR(5) = convert(Val(Trim(txt20) & ""))
    RSCUR(6) = convert(Val(Trim(txt10) & ""))
    RSCUR(7) = convert(Val(Trim(txt5) & ""))
    RSCUR(8) = convert(Val(Trim(txt2) & ""))
    RSCUR(9) = convert(Val(Trim(txt1) & ""))
    RSCUR(10) = convert(Val(Trim(txtchange) & ""))
RSCUR.Update    ' END
SAVESHELL Trim(INC)
If REC_SET1.State = 1 Then REC_SET1.Close
If MOD_REC = False Then MsgBox "New Challan ID is " & INC, vbExclamation, "Effmcg"
Label23 = labdisp("idy_rsi")
VAN_PRT_CHK:
If CHECKTB("prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_RSI") = True Then
    CON.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
    CON.Execute "update dir where dir_idy like 'REM' set dir_als = 'RSI'"
    If Trolley.Value = 1 Then
        If Pieces.Value = 0 Then
            Openreport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Pieces)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        Else
            Openreport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Cases)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        End If
    Else
        Openreport CURDIR & "\DOS_VCH.RP1", "Van Challan", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
    End If
    CANCELRECORD
End If
CANCELRECORD
End Sub

Private Sub CANCELRECORD()
clrctr M_FRM_RSI
setval
Form_Load
MS.Rows = 1
ms1.Rows = 2
ms2.Rows = 2
MS.AddItem ""
SSTab1.Tab = 0
Toolbar1.Buttons(12).Enabled = False
M_INV_IDY.Enabled = True
M_INV_IDY.SetFocus
Trolley.Enabled = False
For A = 1 To MSS.Rows - 1
    MSS.TextMatrix(A, 1) = "0.00"
    MSS.TextMatrix(A, 2) = "0.00"
    MSS.TextMatrix(A, 3) = "0.00"
Next
If Trolley.Value = 0 Then TRP_NOS.Enabled = True
TRP_NOS.Clear
TRP_NOS.AddItem "All"
TRP_NOS = ""
invid = ""
MOD_REC = False
End Sub

Private Sub newrecord()
CANCELRECORD
If Trolley.Value = 1 Then
MS.ColWidth(7) = 0
MS.ColWidth(6) = 1050
End If
'CHECK3.Enabled = False
'Check4.Enabled = False
M_INV_IDY.Enabled = False
M_INV_DAT.SetFocus
SSTab1.Tab = 0
MS.FixedCols = 0
idnew = True
IDN = True
End Sub

Private Sub SALESMANLIST()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Label49 = "Select Salesman"
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(M_SMN_IDY) & "%'", CON, adOpenKeyset, adLockOptimistic
ElseIf Check2.Value = 1 Then
    REC_SET.Open "select * from SMN WHERE SMN_AD4 like '" & Trim(M_SMN_IDY) & "%'", CON, adOpenKeyset, adLockOptimistic
End If
List1.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = List1.ListItems.Add(, , Trim(REC_SET("smn_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    List1.Visible = True
    List1.SetFocus
    
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Salesmen Not Found, Define Salesmen !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    M_SMN_IDY.SetFocus
    End If
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    TRP_NOS.AddItem "All"
    Check4.Enabled = False
    M_INV_DAT.SetFocus
ElseIf Check3.Value = 0 Then
    M_INV_DAT.SetFocus
End If
End Sub

Private Sub Check3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check or Uncheck To Confirm Returns"
End Sub

Private Sub CHECK3_KeyPress(K As Integer)
If K = 13 Then
    M_INV_DAT.SetFocus
End If
End Sub

Private Sub Check4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check or Uncheck To Close Challan"
End Sub

Private Sub Check4_keypress(K As Integer)
If K = 13 Then
    M_INV_DAT.SetFocus
End If
End Sub

Private Sub Command1_Click()
Picture2.Visible = False
SSTab1.Enabled = True
Toolbar1.Enabled = True
If SSTab1.Tab = 1 Then
    MS.SetFocus
End If
End Sub

Private Sub Command2_Click()
If CHECKTB("inl.dbf,inv.dbf,lnv.dbf,prd.dbf,brd.dbf,vch.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSINL As New ADODB.Recordset
If Len(Trim(M_INV_IDY)) = 0 Then
    MsgBox "Select Van Challan ", vbExclamation
    CANCELRECORD
    Exit Sub
End If
RSINL.Open "SELECT SUM(SPD_AMT),SUM(CSD_AMT),SUM(TAX_AMT),SUM(EXD_AMT),SUM(PKG_CST) FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then
    inldiscount = Val(RSINL(0) & "") + Val(RSINL(1) & "")
    inltax = Val(RSINL(2) & "") + Val(RSINL(3) & "") + Val(RSINL(4) & "")
End If
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(DIS_AMT),SUM(ADD_AMT),SUM(TAX_AMT) FROM INV WHERE INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') AND STK_TYP IN('Salable','Free')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then
    invdiscount = Val(RSINL(0) & "") + Val(RSINL(1) & "")
    invtax = Val(RSINL(2) & "")
End If
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(FRT_AMT) FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "' and rtrim(ltrim(TXN_RM2))<>'Scheme Amount'", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then inldiscount = Val(inldiscount) + Val(RSINL(0) & "")
MSS.TextMatrix(7, 2) = convert(Val(inldiscount) + Val(invdiscount))
MSS.TextMatrix(8, 2) = convert(Val(inltax) + Val(invtax))
'grossamt
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT sum((prd_SDR /sub_unt) * tot_pcs),stk_typ FROM INV,BRD WHERE BRD.BAT_IDY LIKE INV.BAT_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') AND INV.STK_TYP IN('Salable','TakeBack','Damage','Free','Breakage','Leakage')group by stk_typ", CON, adOpenKeyset, adLockOptimistic
RSamt = 0
If RSINL.RecordCount > 0 Then
    For A = 1 To RSINL.RecordCount
        If Trim(RSINL!stk_typ) = "Salable" Then
            RSamt = Val(RSINL(0) & "") + RSamt
        ElseIf Trim(RSINL!stk_typ) = "Free" Then
            RSamt = Val(RSINL(0) & "") + RSamt
        
        ElseIf Trim(RSINL!stk_typ) = "Damage" Then
            RSamt = RSamt - Val(RSINL(0) & "")
        ElseIf Trim(RSINL!stk_typ) = "TakeBack" Or Trim(RSINL!stk_typ) = "Breakage" Or Trim(RSINL!stk_typ) = "Leakage" Then
            RSamt = RSamt - Val(RSINL(0) & "")
        End If
        RSINL.MoveNext
    Next
    MSS.TextMatrix(1, 2) = convert(Val(RSamt))
End If

'freeqty
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT sum((prd_SDR /sub_unt) * tot_pcs) FROM INV,BRD WHERE INV.BAT_IDY LIKE BRD.BAT_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') AND STK_TYP IN('Free') and sch_idy IS NULL", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(2, 2) = convert(Val(RSINL(0)))
'schqty
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT * FROM lnv,brd,prd WHERE lnv.INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') and lev_des like brd.bat_idy and brd.prd_idy like prd.prd_idy", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then
    For A = 1 To RSINL.RecordCount
        b = Val(b) + ((Val(RSINL!prd_sdr) / Val(RSINL!sub_unt)) * Val(RSINL!INV_GRS))
        RSINL.MoveNext
    
    Next
    MSS.TextMatrix(5, 2) = convert(Val(b))
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "SELECT sum(lnv_amt) FROM lnv WHERE lnv.INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') and (len(alltrim(lev_des))=0 or inv_grs = 0) ", CON, adOpenKeyset, adLockOptimistic
    If RSINL.RecordCount > 0 Then C = Val(RSINL(0) & "")
    MSS.TextMatrix(6, 2) = convert(Val(C))
End If
'CASH
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH,ACT WHERE VCH.DBT_IDY LIKE ACT.ACT_IDY AND ACT.GRP_CLF LIKE 'Cash Book' and inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(10, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH,ACT WHERE VCH.DBT_IDY LIKE ACT.ACT_IDY AND ACT.GRP_CLF LIKE 'Bank Book' and inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(11, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH WHERE inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(12, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(net_amt) FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(12, 2) = convert(Val(RSINL(0) & "") - Val(MSS.TextMatrix(12, 2)))
SETCALC
End Sub

Private Sub SETCALC()
MSS.TextMatrix(1, 3) = convert(Val(MSS.TextMatrix(1, 1)) - Val(MSS.TextMatrix(1, 2)))
For A = 2 To 7
b = Val(b) + Val(MSS.TextMatrix(A, 1))
C = Val(C) + Val(MSS.TextMatrix(A, 2))
MSS.TextMatrix(A, 3) = convert(Val(MSS.TextMatrix(A, 1)) - Val(MSS.TextMatrix(A, 2)))
Next
MSS.TextMatrix(8, 3) = convert(Val(MSS.TextMatrix(8, 1)) - Val(MSS.TextMatrix(8, 2)))
MSS.TextMatrix(9, 1) = convert(Val(MSS.TextMatrix(1, 1)) + Val(MSS.TextMatrix(8, 1)) - Val(b))
MSS.TextMatrix(9, 2) = convert(Val(MSS.TextMatrix(1, 2)) + Val(MSS.TextMatrix(8, 2)) - Val(C))
MSS.TextMatrix(13, 1) = convert(Val(MSS.TextMatrix(9, 1)) - (Val(MSS.TextMatrix(11, 1)) + Val(MSS.TextMatrix(11, 1)) + Val(MSS.TextMatrix(12, 1))))
MSS.TextMatrix(13, 2) = convert(Val(MSS.TextMatrix(9, 2)) - (Val(MSS.TextMatrix(10, 2)) + Val(MSS.TextMatrix(11, 2)) + Val(MSS.TextMatrix(12, 2))))
MSS.TextMatrix(9, 3) = convert(Val(MSS.TextMatrix(9, 1)) - Val(MSS.TextMatrix(9, 2)))
MSS.TextMatrix(10, 3) = convert(Val(MSS.TextMatrix(10, 1)) - Val(MSS.TextMatrix(10, 2)))
MSS.TextMatrix(11, 3) = convert(Val(MSS.TextMatrix(11, 1)) - Val(MSS.TextMatrix(11, 2)))
MSS.TextMatrix(12, 3) = convert(Val(MSS.TextMatrix(12, 1)) - Val(MSS.TextMatrix(12, 2)))
End Sub

Private Sub Command3_Click()
If Len(Trim(M_INV_IDY)) > 0 Then
    Dim RS As New ADODB.Recordset
    CON.Execute "delete from stl"
    RS.Open "select * from stl", CON, adOpenKeyset, adLockOptimistic
    For A = 1 To MSS.Rows - 1
    RS.AddNew
    RS!TXN_RMK = Trim(MSS.TextMatrix(A, 0))
    RS!rct_Qty = Val(MSS.TextMatrix(A, 1))
    RS!ISS_QTY = Val(MSS.TextMatrix(A, 2))
    RS!INV_IDY = Trim(M_INV_IDY)
    RS!PRD_IDY = ""
    RS.Update
    Next
    RS.Close
    Openreport CURDIR & "\dos_vch.rp1", "Van Settlement", 0, "", 6
Else
    MsgBox "Select Van Challan", vbExclamation
    CANCELRECORD
End If
End Sub

Private Sub Command4_Click()
Picture2.Visible = False
SSTab1.Enabled = True
Toolbar1.Enabled = True
V_RSR_LST
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
If Len(Trim(M_INV_IDY)) = 0 Then
    MsgBox "Select Van Challan ", vbExclamation
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("snv.dbf,shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
With ms1
.Rows = ms2.Rows
For A = 2 To ms2.Rows - 1
    .TextMatrix(A, 0) = ms2.TextMatrix(A, 0)
    .TextMatrix(A, 1) = ms2.TextMatrix(A, 1)
    .TextMatrix(A, 2) = ms2.TextMatrix(A, 2)
    .TextMatrix(A, 3) = ms2.TextMatrix(A, 3)
    .TextMatrix(A, 4) = ms2.TextMatrix(A, 4)
    .TextMatrix(A, 9) = ms2.TextMatrix(A, 7)
Next
.SetFocus
'.Row = 1
.Col = 1
Dim RS As New ADODB.Recordset
RS.Open "SELECT SUM(R_SHL_NOS),SUM(R_PCK_NOS),SNV.SHL_IDY,prd_idy FROM SNV WHERE TXN_TYP LIKE 'SAL' AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') GROUP BY SNV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS.RecordCount
    For b = 2 To .Rows - 1
        If Trim(RS!PRD_IDY & "") = Trim(.TextMatrix(b, 9)) Then
            .TextMatrix(b, 5) = Val(RS(0) & "")
            .TextMatrix(b, 6) = Val(RS(1) & "")
            GoTo n
        End If
    Next
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = GetValue("PRD_NME", "PRD WHERE PRD_IDY LIKE '" & Trim(RS!PRD_IDY) & "'")
    .TextMatrix(.Rows - 1, 5) = Val(RS(0) & "")
    .TextMatrix(.Rows - 1, 6) = Val(RS(1) & "")
    .TextMatrix(.Rows - 1, 9) = Val(RS(2) & "")
n:
    RS.MoveNext
Next
For b = 2 To .Rows - 1
    .TextMatrix(b, 7) = Val(.TextMatrix(b, 5)) - Val(.TextMatrix(b, 3))
    .TextMatrix(b, 8) = Val(.TextMatrix(b, 6)) - Val(.TextMatrix(b, 4))
Next

End With
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command4_Click
End If
End Sub

Private Sub list3_DblClick()
List3_KeyPress 13
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If LIST3.ListItems.Count > 0 Then
        M_CST_IDY = LIST3.SelectedItem.ListSubItems(1)
        M_CST_IDY_keydown 13, 0
        LIST3.Visible = False
    End If
End If
End Sub

Private Sub exitrecord()
Unload Me
End Sub

Private Sub list2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list2.SortOrder = lvwAscending Then
list2.SortOrder = lvwDescending
ElseIf list2.SortOrder = lvwDescending Then
list2.SortOrder = lvwAscending
End If
list2.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub list7_DblClick()
List7_keypress 13
End Sub
Private Sub list7_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If LIST7.SortOrder = lvwAscending Then
LIST7.SortOrder = lvwDescending
ElseIf LIST7.SortOrder = lvwDescending Then
LIST7.SortOrder = lvwAscending
End If
LIST7.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub list2_DblClick()
list2_keypress 13
End Sub

Private Sub list3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If LIST3.SortOrder = lvwAscending Then
LIST3.SortOrder = lvwDescending
ElseIf LIST3.SortOrder = lvwDescending Then
LIST3.SortOrder = lvwAscending
End If
LIST3.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub List4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list4.SortOrder = lvwAscending Then
list4.SortOrder = lvwDescending
ElseIf list4.SortOrder = lvwDescending Then
list4.SortOrder = lvwAscending
End If
list4.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub list4_DblClick()
list4_KeyPress 13
End Sub

Private Sub list4_KeyPress(KeyAscii As Integer)
Dim LS1 As ListSubItem
If KeyAscii = 13 Then
    If list4.ListItems.Count > 0 Then
        list4.Visible = False
        If list4.ListItems.Count = 0 Then
            Option3.Visible = False
            Exit Sub
        End If
        Set LS1 = list4.SelectedItem.ListSubItems.Item(6)
        dispprd LS1.Text
    End If
End If
End Sub

Private Sub list2_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list2.ListItems.Count > 0 Then
        M_ACT_IDY = list2.SelectedItem.ListSubItems.Item(1)
        M_ACT_NME = list2.SelectedItem
        SSTab1.Tab = 0
        list2.Visible = False
        TRP_NOS.SetFocus
    End If
End If
End Sub
Private Sub List7_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If LIST7.ListItems.Count > 0 Then
    LIST7.Visible = False
    SSTab1.Tab = 1
    MS.SetFocus
    End If
End If
End Sub
Private Sub List6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    Dim REC_SET As New ADODB.Recordset
    List6.Visible = False
    MS.Text = List6
    List6.Visible = False
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from prd,BRD where BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.prd_idy like '" & MS.TextMatrix(MS.Row, 10) & "'", CON, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then MS.TextMatrix(MS.Row, 4) = AMT(REC_SET("PRD_MRP") & "")
    discalc
    MS.Col = 10
    MS.SetFocus
End If
End Sub

Private Sub V_ACT_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
TYP = ""
If TYP = "" Then
    TYP = "Stock Issues"
End If
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' and ACT_NME LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", CON, adOpenKeyset, adLockOptimistic
Else
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' and ACT_AD4 LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", CON, adOpenKeyset, adLockOptimistic
End If
list2.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list2.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Account"
    list2.Visible = True
    list2.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Stock Issues Account Not Found, Define Stock Issues Account...!", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    M_ACT_IDY.SetFocus
End If
End If
End Sub
Private Sub V_ACT_S_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
TYP = ""
If TYP = "" Then
    TYP = "Sales"
    End If
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' AND ACT_TYP LIKE '' ", CON, adOpenKeyset, adLockOptimistic
Else
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' and ACT_TYP LIKE '' ", CON, adOpenKeyset, adLockOptimistic
End If
LIST7.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = LIST7.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Account"
    LIST7.Visible = True
    LIST7.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Stock Issues Account Not Found, Define Stock Issues Account...!", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    M_ACT_IDY.SetFocus
End If
End If
End Sub

Private Sub V_DLR_LST()
If CHECKTB("act.dbf,mar.dbf,smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,RDY,smn where (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and  act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.ACT_IDY LIKE RDY.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND DBT_TYP='Y' and act.cst_typ like 'Van' GROUP BY ACT.ACT_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,RDY,smn where (act.ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.ACT_IDY LIKE RDY.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND DBT_TYP='Y' and act.cst_typ like 'Van' GROUP BY ACT.ACT_IDY", CON, adOpenKeyset, adLockOptimistic
End If
LIST3.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = LIST3.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_nme") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("MAR_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("rou_idy") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Vehicle"
    LIST3.Visible = True
    LIST3.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Vehicle Not Found, Define Vehicle !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    End If
End If
End Sub

Private Sub ListView1_DblClick()
ListView1_KeyPress 13
End Sub

Private Sub ListView1_KeyPress(K As Integer)
If K = 13 Then
    Picture1.Visible = False
    SSTab1.Enabled = True
    Toolbar1.Enabled = True
    MAKEBILLCONTD
End If
End Sub

Private Sub M_CAS_AMT_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_CRT_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt1000.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 Then
    M_CHQ_AMT.SetFocus
End If
End Sub

Private Sub M_GRS_SAL_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_FRE_QTY_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_NET_SAL_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_DIS_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_EXP_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_COM_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_CAS_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_CHQ_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_CRT_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_CHQ_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CRT_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CAS_AMT.SetFocus
    calc
End If
End Sub

Private Sub M_GRS_SAL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
M_FRE_QTY.SetFocus
End If
End Sub

Private Sub M_FRE_QTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_NET_SAL.SetFocus
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_GRS_SAL.SetFocus
End If
End Sub
Private Sub M_BRK_QTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_LKG_QTY.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_FRE_QTY.SetFocus
    calc
End If
End Sub
Private Sub M_LKG_QTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_NET_SAL.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_BRK_QTY.SetFocus
    calc
End If
End Sub
Private Sub M_NET_SAL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_DIS_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_FRE_QTY.SetFocus
    calc
End If
End Sub
Private Sub M_DIS_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_COM_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_NET_SAL.SetFocus
    calc
End If
End Sub
Private Sub M_EXP_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_TAX_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_COM_AMT.SetFocus
    calc
End If
End Sub
Private Sub M_COM_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_EXP_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_DIS_AMT.SetFocus
    calc
End If
End Sub

Private Sub M_CAS_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CHQ_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_TAX_AMT.SetFocus
    calc
End If
End Sub

Private Sub M_TAX_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TAX_AMT_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CAS_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_EXP_AMT.SetFocus
    calc
End If
End Sub

Private Sub M_TAX_AMT_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_FRE_QTY_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_NET_SAL_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_DIS_AMT_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_EXP_AMT_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_COM_AMT_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_CHQ_AMT_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_CRT_AMT_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_CST_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Vehicle ID"
End Sub

Private Sub M_CST_IDY_keydown(K As Integer, S As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If CHECKTB("act.dbf,smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_CST_IDY)) & "' AND (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') AND DBT_TYP='Y' and cst_typ like 'Van'", CON, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then
        M_CST_IDY = Sincrement(M_CST_IDY)
        If Len(Trim(M_CST_IDY)) = 0 Then Exit Sub
        M_CST_NME = Trim(REC_SET("act_nme") & "")
        If Len(Trim(REC_SET("ROU_IDY"))) <> 0 Then M_ROU_IDY = Trim(REC_SET("ROU_IDY") & "")
        If IDN = True Then
            SMN = Split(GetSalesman(M_CST_IDY, M_INV_DAT), "|")
            If UBound(SMN) > 0 Then
                M_SMN_IDY = Trim(SMN(0))
                M_SMN_NME = Trim(SMN(1))
            End If
        End If
        If REC_SET1.State = 1 Then REC_SET1.Close
        M_DBT_RAP = Trim(REC_SET("dbt_rap"))
        SSTab1.Tab = 0
        If K = 38 Then
            M_INV_DAT.SetFocus
        Else
            If TRP_NOS.Enabled = True Then
                TRP_NOS.SetFocus
            Else
                M_SMN_IDY.SetFocus
            End If
        End If
    Else
        V_DLR_LST
    End If
End If
End Sub

Private Sub M_INV_DAT_CloseUp()
M_INV_DAT.SetFocus
End Sub

Private Sub M_INV_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date Using Arrow Keys Or Press F4"
End Sub

Private Sub M_INV_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then M_CST_IDY.SetFocus
End Sub

Private Sub Form_Activate()
checkdongle
If Me.Tag = "Van" Then
list5.SetFocus
Else
SSTab1.Tab = 0
End If
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "PUSH CART") = True Then
    Trolley.Value = 1
    TRP_NOS.Enabled = False
Else
    Trolley.Value = 0
End If
If M_INV_IDY.Enabled = True Then
    M_INV_IDY.SetFocus
Else
    M_INV_DAT.SetFocus
End If
Label23 = labdisp("idy_rsi")
SSTab1.TabEnabled(3) = False
'setval
SETYEAR M_FRM_RSI
txt1000 = "0.00"
txt500 = "0.00"
txt100 = "0.00"
txt50 = "0.00"
txt20 = "0.00"
txt10 = "0.00"
txt5 = "0.00"
txt2 = "0.00"
txt1 = "0.00"
txtchange = "0.00"
End Sub

Private Sub INVOICELIST()
If CHECKTB("smn.dbf,van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Trolley.Value = 1 Then
    qry = " and vaN_chk= 'P'"
Else
    qry = " and vaN_chk= 'V'"
End If
If Check1.Value = 1 Then
    REC_SET.Open "select inv_idy,inv_dat,SMN_NME,SMN_AD4 from SMN,VAN where SMN_nme like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy " & qry & " GROUP BY INV_IDY ", CON, adOpenKeyset, adLockOptimistic
Else
    REC_SET.Open "select inv_idy,inv_dat,SMN_NME,SMN_AD4 from SMN,VAN where SMN_ad4 like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy " & qry & " GROUP BY INV_IDY", CON, adOpenKeyset, adLockOptimistic
End If
list5.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_AD4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Van Challan Document"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Van Challan Number Not Found...!", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then Shell "c:\windows\calc.exe"
If List1.Visible = False And list4.Visible = False And list5.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 82 And Shift = 2 Then RETURNSTOCKS
    If KeyCode = 83 And Shift = 2 Then
    If Check3.Value = 1 And Trolley.Value = 1 Then
        MakeBill
    Else
        saverecord
    End If
    End If
    If KeyCode = 88 And Shift = 2 Then exitrecord
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_RSI.ActiveControl.name = "M_INV_IDY" Then
            INVOICELIST
        ElseIf M_FRM_RSI.ActiveControl.name = "M_SMN_IDY" Then
            SALESMANLIST
        ElseIf M_FRM_RSI.ActiveControl.name = "Text3" Then
            PRODUCTLIST
        ElseIf M_FRM_RSI.ActiveControl.name = "M_CST_IDY" Then
            V_DLR_LST
        ElseIf M_FRM_RSI.ActiveControl.name = "M_ACT_IDY" Then
            V_ACT_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_RSI.ActiveControl Is TextBox Then M_FRM_RSI.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub List5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list5.SortOrder = lvwAscending Then
list5.SortOrder = lvwDescending
ElseIf list5.SortOrder = lvwDescending Then
list5.SortOrder = lvwAscending
End If
list5.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub list5_DblClick()
If list5.ListItems.Count > 0 Then
    list5.Visible = False
    If Label49 = "Select Van Challan" Then
        
        AddOpeningStock
    Else
        M_INV_IDY = list5.SelectedItem
        DISPDETAILS
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If List1.Visible = True Then
        List1.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_SMN_IDY.SetFocus
    ElseIf list5.Visible = True Then
        list5.Visible = False
        SSTab1.Tab = 0
        If M_INV_IDY.Enabled = True Then
            M_INV_IDY.SetFocus
        Else
            M_INV_DAT.SetFocus
        End If
    ElseIf Picture2.Visible = True Then
        Picture2.Visible = False
        Toolbar1.Enabled = True
        SSTab1.Enabled = True
        If SSTab1.Tab = 1 Then
            MS.SetFocus
        End If
    ElseIf list4.Visible = True Then
        list4.Visible = False
        If Text3.Visible = True Then
            Text3.SetFocus
        Else
            Check2.Caption = "Place"
            SSTab1.Tab = 1
            MS.SetFocus
        End If
    ElseIf list2.Visible = True Then
        list2.Visible = False
        M_ACT_IDY.SetFocus
        SSTab1.Tab = 0
    ElseIf LIST3.Visible = True Then
        LIST3.Visible = False
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        Text3.Visible = False
        amount.Visible = False
        GRS_CAL
        discalc
        SSTab1.Tab = 2
    ElseIf LIST3.Visible = True Then
        LIST3.Visible = False
        M_CST_IDY.SetFocus
        SSTab1.Tab = 0
    Else
        If M_FRM_RSI.ActiveControl.name = "M_INV_IDY" Then
            Unload Me
        Else
            If Check3.Value = 1 And Trolley.Value = 1 Then
                MakeBill
            Else
                saverecord
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
MOD_REC = False
MS.ColWidth(0) = 2200
MS.ColWidth(1) = 500
MS.ColWidth(2) = 1050
MS.ColWidth(3) = 1050
MS.ColWidth(4) = 1050
MS.ColWidth(5) = 1050
MS.ColWidth(6) = 1050
MS.ColWidth(7) = 1050
MS.ColWidth(8) = 1000
MS.ColWidth(9) = 1200
MS.ColWidth(10) = 1100
MS.ColWidth(11) = 1500
MS.ColWidth(12) = 0
MS.ColWidth(13) = 0
MS.ColWidth(14) = 1600
MS.ColWidth(15) = 1000
MS.ColWidth(16) = 1050
MS.ColWidth(17) = 1050
MS.ColWidth(18) = 1050
ms2.RowHeight(0) = 450
MS.TextMatrix(0, 0) = "Product Name"
SCl MS, 0
MS.ColAlignment(1) = 7
MS.TextMatrix(0, 1) = "Trip"
MS.TextMatrix(0, 2) = "Opg Stk"
MS.ColAlignment(2) = 7
SCl MS, 2
MS.TextMatrix(0, 3) = "New Stk"
MS.ColAlignment(3) = 7
SCl MS, 3
MS.TextMatrix(0, 4) = "Total Stk"
MS.ColAlignment(4) = 7
MS.TextMatrix(0, 5) = "Ret. Stk"
MS.ColAlignment(5) = 7
SCl MS, 5
MS.RowHeight(0) = 450
MSS.RowHeight(0) = 450
ms1.RowHeight(0) = 450
MS.TextMatrix(0, 6) = "Bill Stk"
MS.ColAlignment(6) = 7
MS.TextMatrix(0, 7) = "Variance"
MS.ColAlignment(7) = 7
MS.TextMatrix(0, 8) = "MRP"
MS.ColAlignment(8) = 7
MS.TextMatrix(0, 9) = "Amount"
MS.ColAlignment(9) = 7
MS.TextMatrix(0, 10) = "Product ID"
MS.TextMatrix(0, 11) = "Short Name"
MS.TextMatrix(0, 12) = "SUBUNTS"
MS.TextMatrix(0, 13) = "batchidy"
MS.TextMatrix(0, 14) = "Batch"
MS.TextMatrix(0, 15) = "Sale Rate"
MS.TextMatrix(0, 16) = "Free. Stk"
MS.ColAlignment(15) = 7
MS.ColAlignment(16) = 7
SCl MS, 16
MS.TextMatrix(0, 17) = "Transit Breakage"
MS.ColAlignment(17) = 7
SCl MS, 17
MS.TextMatrix(0, 18) = "Transit Leakage"
MS.ColAlignment(18) = 7
SCl MS, 18
MS.Col = 0
MS.Col = 3
If CHECKTB("dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If querycheck("DIR", "DIR_MSG", "C", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PCS'") = True Then
    Pieces.Value = 1
Else
    Pieces.Value = 0
End If
MSS.ColAlignment(1) = 7
MSS.ColAlignment(2) = 7
MSS.ColAlignment(3) = 7
MSS.TextMatrix(0, 1) = "      Settlement"
MSS.TextMatrix(0, 2) = "  Reconciliation"
MSS.TextMatrix(0, 3) = "Difference"
MSS.ColWidth(0) = 2150
MSS.ColWidth(1) = 1700
MSS.ColWidth(2) = 1700
MSS.ColWidth(3) = 1500
MSS.Row = 0
MSS.Col = 1
MSS.CellFontBold = True
MSS.Col = 2
MSS.CellFontBold = True
MSS.Col = 3
MSS.CellFontBold = True
MSS.Rows = 1
MSS.Rows = 15
MSS.TextMatrix(1, 0) = "Sales"
MSS.TextMatrix(2, 0) = "Free Qty    (-)"
MSS.TextMatrix(3, 0) = "Breakage    (-)"
MSS.TextMatrix(4, 0) = "Leakage     (-)"
MSS.RowHeight(3) = 0
MSS.RowHeight(4) = 0
MSS.TextMatrix(5, 0) = "Scheme Qty  (-)"
MSS.TextMatrix(6, 0) = "Scheme Less (-)"
MSS.TextMatrix(7, 0) = "Discount    (-)"
MSS.TextMatrix(8, 0) = "Tax/Others  (+)"
MSS.TextMatrix(9, 0) = "Net Sales      "
MSS.TextMatrix(10, 0) = "Cash          "
MSS.TextMatrix(11, 0) = "Cheques       "
MSS.TextMatrix(12, 0) = "Credits       "
MSS.TextMatrix(13, 0) = "Variance      "
MSS.TextMatrix(14, 0) = "Expenses   (-)"
MSS.Col = 0
For A = 1 To 14
MSS.Row = A
MSS.CellFontBold = True
Next
With ms1
.Cols = 10
.Rows = 3
.FixedRows = 2
.FixedCols = 1
.MergeCells = flexMergeFree
.MergeRow(0) = True
.ColWidth(0) = 1800
.ColWidth(1) = 850
.ColWidth(2) = 850
.ColWidth(3) = 850
.ColWidth(4) = 850
.ColWidth(5) = 850
.ColWidth(6) = 850
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(9) = 0
For A = 0 To .Cols - 1
.Row = 0
.Col = A
.CellFontBold = True
.Row = 1
.Col = A
.CellFontBold = True
Next
.Row = 0
.Col = 1
.CellAlignment = 4
.TextMatrix(0, 1) = "        Actuals"
.Col = 2
.CellAlignment = 4
.TextMatrix(0, 2) = "        Actuals"
.Row = 0
.Col = 3
.CellAlignment = 4
.TextMatrix(0, 1 + 2) = "Quick Settlement"
.Col = 4
.CellAlignment = 4
.TextMatrix(0, 2 + 2) = "Quick Settlement"
.Col = 5
.CellAlignment = 4
.TextMatrix(0, 3 + 2) = "Final Settlement"
.Col = 6
.CellAlignment = 4
.TextMatrix(0, 4 + 2) = "Final Settlement"
.Col = 7
.CellAlignment = 4
.TextMatrix(0, 5 + 2) = "Variance"
.Col = 8
.CellAlignment = 4
.TextMatrix(0, 6 + 2) = "Variance"
.ColAlignment(1) = 7
.ColAlignment(2) = 7
.ColAlignment(3) = 7
.ColAlignment(4) = 7
.ColAlignment(5) = 7
.ColAlignment(6) = 7
.ColAlignment(7) = 7
.ColAlignment(8) = 7

.TextMatrix(1, 0) = "Shell Name"
.TextMatrix(1, 1) = "Shells"
.TextMatrix(1, 2) = "Bottles"
.TextMatrix(1, 3) = "Shells"
.TextMatrix(1, 4) = "Bottles"
.TextMatrix(1, 5) = "Shells"
.TextMatrix(1, 6) = "Bottles"
.TextMatrix(1, 7) = "Shells"
.TextMatrix(1, 8) = "Bottles"
End With

With ms2
.Cols = 9
.Rows = 3
.FixedRows = 2
.FixedCols = 1
.MergeCells = flexMergeFree
.MergeRow(0) = True
.MergeCol(0) = True
For A = 0 To .Cols - 1
.Row = 0
.Col = A
.CellFontBold = True
.Row = 1
.Col = A
.CellFontBold = True
Next
.ColWidth(0) = 2200
.ColWidth(1) = 900
.ColWidth(2) = 900
.ColWidth(3) = 900
.ColWidth(4) = 900
.ColWidth(5) = 900
.ColWidth(6) = 900
.ColWidth(7) = 0
.ColWidth(8) = 0
.ColAlignment(1) = 7
.ColAlignment(2) = 7
.ColAlignment(3) = 7
.ColAlignment(4) = 7
.ColAlignment(5) = 7
.ColAlignment(6) = 7
.Row = 0
.Col = 1
.CellAlignment = 4
.Col = 2
.CellAlignment = 4
.Col = 3
.CellAlignment = 4
.Col = 4
.CellAlignment = 4
.Col = 5
.CellAlignment = 4
.Col = 6
.CellAlignment = 4
.TextMatrix(0, 1) = "Returnables"
.TextMatrix(0, 2) = "Returnables"
.TextMatrix(0, 3) = "Returned"
SCl ms2, 3
.TextMatrix(0, 4) = "Returned"
SCl ms2, 4
.TextMatrix(0, 5) = "Difference"
.TextMatrix(0, 6) = "Difference"
.TextMatrix(1, 0) = "Shell Name"
.TextMatrix(0, 0) = "Shell Name"
.TextMatrix(1, 1) = "Shells"
.TextMatrix(1, 2) = "Bottles"
.TextMatrix(1, 3) = "Shells"
.TextMatrix(1, 4) = "Bottles"
.TextMatrix(1, 5) = "Shells"
.TextMatrix(1, 6) = "Bottles"
End With
MSS.Row = 13
MSS.Col = 1
MSS.CellFontBold = True
MSS.Col = 2
MSS.CellFontBold = True
MSS.Col = 3
MSS.CellFontBold = True
For A = 1 To 14
MSS.TextMatrix(A, 1) = "0.00"
MSS.TextMatrix(A, 2) = "0.00"
MSS.TextMatrix(A, 3) = "0.00"
Next
MS.Col = 5
MS.Col = 0
M_INV_DAT = Date
lrdate = Date
dcdate = Date
waydate = Date
Label17 = "0.00"
Label18 = "0"
Label19 = "0.00"
Label20 = "0.00"
Label23 = labdisp("idy_rsi")
M_NET_AMT = "0.00"
M_GRS_AMT = "0.00"
crdays = 0
Total = "0.00"
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Trolley.Value = 1 Then
    CON.Execute "INSERT INTO PRT VALUES('PUSH CART')"
Else
    CON.Execute "DELETE FROM PRT WHERE PRT_NME LIKE 'PUSH CART'"
End If
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Challan Number"
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
If M_INV_IDY = "" And K = 27 Then Unload Me
End Sub

Private Sub LIST5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If list5.ListItems.Count > 0 Then
    If Label49 = "Select Van Challan" Then
        AddOpeningStock
    Else
        M_INV_IDY = list5.SelectedItem
        DISPDETAILS
    End If
End If
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
If Label49.Caption = "Select Salesman" Then
    SALESMANLIST
ElseIf Label49.Caption = "Select Invoice" Then
    INVOICELIST
ElseIf Label49.Caption = "Select Product" Then
    PRODUCTLIST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If Label49.Caption = "Select Salesman" Then
    SALESMANLIST
ElseIf Label49.Caption = "Select Invoice" Then
    INVOICELIST
ElseIf Label49.Caption = "Select Product" Then
    If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If Check2.Value = 1 Then
    CON.Execute "insert into prt values('prdnum')"
    Else
    CON.Execute "delete from prt where prt_nme like 'prdnum'"
    End If
    PRODUCTLIST
End If
'End If
End Sub


Private Sub MS_RowColChange()
M_Prd_Nme = MS.TextMatrix(MS.Row, 0)
End Sub

Private Sub MS_EnterCell()
If MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Type Product Name or Press <Enter> to get the List"
ElseIf MS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Trip No"
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Opening Stock"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "New Stock"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Stock"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Return Stock"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Bill Stock"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Variance"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Amount"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product ID"
ElseIf MS.Col = 11 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
ElseIf MS.Col = 12 Then
    M_MDI.STS_BAR.Panels(1).Text = "Batch"
ElseIf MS.Col = 16 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Free Stock"
End If
End Sub

Private Sub MS_Scroll()
If amount.Visible = True Then
    amount = ""
    amount.Visible = False
    MS.SetFocus
ElseIf List6.Visible = True Then
    List6.Visible = False
    MS.SetFocus
ElseIf Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub ms2_KeyPress(K As Integer)
If ms2.Row > 1 Then
    If Check3.Value = 1 Then
        If ms2.Col = 3 Or ms2.Col = 4 Then
            If (K > 47 And K < 58) Or K = 45 Then
                txtempshl = ""
                txtempshl.MaxLength = 8
                txtempshl.Visible = True
                txtempshl.Top = ms2.Top + ms2.CellTop
                txtempshl.Left = ms2.CellLeft + ms2.Left
                txtempshl.Width = ms2.CellWidth
                txtempshl.Height = ms2.CellHeight
                txtempshl.Visible = True
                txtempshl.SetFocus
                txtempshl = Chr(K)
                txtempshl.SelStart = Len(txtempshl)
            Else
                K = 0
            End If
        End If
    End If
End If
End Sub

Private Sub ms2_Scroll()
If txtempshl.Visible = True Then
    txtempshl = ""
    txtempshl.Visible = False
    ms2.SetFocus
End If
End Sub

Private Sub ShlSet()
SHLCAS = 0
Dim RS As New ADODB.Recordset
ms2.Rows = 2
RS.Open "select * from shl,prd where prd.prd_idy like shl.prd_idy and shl_typ= 'Shell'", CON, adOpenKeyset, adLockOptimistic
For A = 1 To MS.Rows - 1
    If MS.TextMatrix(A, 0) <> "" Then
        If RS.RecordCount > 0 Then
            RS.MoveFirst
            RS.Find "prd_idy='" & Trim(MS.TextMatrix(A, 10)) & "'"
            If RS.EOF = False Then
                For b = 1 To ms2.Rows - 1
                    If MS.TextMatrix(A, 10) = ms2.TextMatrix(b, 7) Then
                        If Pieces.Value = 0 Then
                            SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 16)) + Val(MS.TextMatrix(A, 17))), Val(MS.TextMatrix(A, 12)))
                        Else
                            Aa = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                            bb = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
                            C = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
                            d = MTS(Val(MS.TextMatrix(A, 16)), Val(MS.TextMatrix(A, 12)))
                            e = MTS(Val(MS.TextMatrix(A, 17)), Val(MS.TextMatrix(A, 12)))
                            f = MTS(Val(MS.TextMatrix(A, 18)), Val(MS.TextMatrix(A, 12)))
                            SHLCAS = QTY(STM(Val(Aa) - (Val(bb) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 12))))
                        End If
                        ms2.TextMatrix(b, 1) = Int(SHLCAS) + Val(ms2.TextMatrix(b, 1))
                        ms2.TextMatrix(b, 2) = Val(Aa) - (Val(bb) + Val(e) + Val(f)) + ms2.TextMatrix(b, 1)
                        GoTo n
                    End If
                Next
                ms2.AddItem ""
                ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(RS!PRD_NME)
                ms2.TextMatrix(ms2.Rows - 1, 7) = Trim(RS!PRD_IDY & "")
                ms2.TextMatrix(ms2.Rows - 1, 8) = Trim(RS!shl_idy & "")
                If Pieces.Value = 0 Then
                    SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 16)) + Val(MS.TextMatrix(A, 17))), Val(MS.TextMatrix(A, 12)))
                Else
                    Aa = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                    bb = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
                    C = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
                    d = MTS(Val(MS.TextMatrix(A, 16)), Val(MS.TextMatrix(A, 12)))
                    e = MTS(Val(MS.TextMatrix(A, 17)), Val(MS.TextMatrix(A, 12)))
                    f = MTS(Val(MS.TextMatrix(A, 18)), Val(MS.TextMatrix(A, 12)))
                    SHLCAS = QTY(STM(Val(Aa) - (Val(bb) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 12))))
                End If
                ms2.TextMatrix(ms2.Rows - 1, 1) = Int(SHLCAS) + Val(ms2.TextMatrix(ms2.Rows - 1, 1))
                ms2.TextMatrix(ms2.Rows - 1, 2) = Val(Aa) - (Val(bb) + Val(e) + Val(f))
            End If
n:
        End If
    End If
Next
If RS.State = 1 Then RS.Close
RS.Open "select * from snv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'VAN'", CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS.RecordCount
    For b = 2 To ms2.Rows - 1
        If Trim(RS!PRD_IDY) = Trim(ms2.TextMatrix(b, 7)) Then
            ms2.TextMatrix(b, 3) = Val(RS!R_shl_nos & "")
            ms2.TextMatrix(b, 4) = Val(RS!R_pck_nos & "")
            Exit For
        End If
    Next
    RS.MoveNext
Next
For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 5) = Val(ms2.TextMatrix(b, 1)) - Val(ms2.TextMatrix(b, 3))
    ms2.TextMatrix(b, 6) = Val(ms2.TextMatrix(b, 2)) - Val(ms2.TextMatrix(b, 4))
Next
End Sub

Private Sub Shl_Set_Click()
If Len(Trim(M_INV_IDY)) = 0 And MOD_REC = True Then
    MsgBox "Select Van Challan ", vbExclamation
    CANCELRECORD
    Exit Sub
End If
SHLCAS = 0
Dim RS As New ADODB.Recordset
ms2.Rows = 2
RS.Open "select * from shl,prd where prd.prd_idy like shl.prd_idy and shl_typ= 'Shell'", CON, adOpenKeyset, adLockOptimistic
For A = 1 To MS.Rows - 1
    If MS.TextMatrix(A, 0) <> "" Then
        If RS.RecordCount > 0 Then
            RS.MoveFirst
            RS.Find "prd_idy='" & Trim(MS.TextMatrix(A, 10)) & "'"
            If RS.EOF = False Then
                For b = 1 To ms2.Rows - 1
                    If MS.TextMatrix(A, 10) = ms2.TextMatrix(b, 7) Then
                        If Pieces.Value = 0 Then
                            SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 16)) + Val(MS.TextMatrix(A, 17))), Val(MS.TextMatrix(A, 12)))
                        Else
                            Aa = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                            bb = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
                            C = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
                            d = MTS(Val(MS.TextMatrix(A, 16)), Val(MS.TextMatrix(A, 12)))
                            e = MTS(Val(MS.TextMatrix(A, 17)), Val(MS.TextMatrix(A, 12)))
                            f = MTS(Val(MS.TextMatrix(A, 18)), Val(MS.TextMatrix(A, 12)))
                            SHLCAS = QTY(STM(Val(Aa) - (Val(bb) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 12))))
                        End If
                        ms2.TextMatrix(b, 1) = Int(SHLCAS) + Val(ms2.TextMatrix(b, 1))
                        ms2.TextMatrix(b, 2) = Val(Aa) - (Val(bb) + Val(e) + Val(f)) + ms2.TextMatrix(b, 1)
                        GoTo n
                    End If
                Next
                ms2.AddItem ""
                ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(RS!PRD_NME)
                ms2.TextMatrix(ms2.Rows - 1, 7) = Trim(RS!PRD_IDY & "")
                ms2.TextMatrix(ms2.Rows - 1, 8) = Trim(RS!shl_idy & "")
                If Pieces.Value = 0 Then
                    SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 16)) + Val(MS.TextMatrix(A, 17))), Val(MS.TextMatrix(A, 12)))
                Else
                    Aa = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                    bb = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
                    C = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
                    d = MTS(Val(MS.TextMatrix(A, 16)), Val(MS.TextMatrix(A, 12)))
                    e = MTS(Val(MS.TextMatrix(A, 17)), Val(MS.TextMatrix(A, 12)))
                    f = MTS(Val(MS.TextMatrix(A, 18)), Val(MS.TextMatrix(A, 12)))
                    SHLCAS = QTY(STM(Val(Aa) - (Val(bb) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 12))))
                End If
                ms2.TextMatrix(ms2.Rows - 1, 1) = Int(SHLCAS) + Val(ms2.TextMatrix(ms2.Rows - 1, 1))
                ms2.TextMatrix(ms2.Rows - 1, 2) = Val(Aa) - (Val(bb) + Val(e) + Val(f))
            End If
n:
        End If
    End If
Next
If RS.State = 1 Then RS.Close
RS.Open "select * from snv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'VAN'", CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS.RecordCount
    For b = 2 To ms2.Rows - 1
        If Trim(RS!PRD_IDY) = Trim(ms2.TextMatrix(b, 7)) Then
            ms2.TextMatrix(b, 3) = Val(RS!R_shl_nos & "")
            ms2.TextMatrix(b, 4) = Val(RS!R_pck_nos & "")
            Exit For
        End If
    Next
    RS.MoveNext
Next
For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 5) = Val(ms2.TextMatrix(b, 1)) - Val(ms2.TextMatrix(b, 3))
    ms2.TextMatrix(b, 6) = Val(ms2.TextMatrix(b, 2)) - Val(ms2.TextMatrix(b, 4))
Next
If SSTab1.Tab = 5 Then
    ms2.SetFocus
    'ms2.Row = 2
    ms2.Col = 1
End If
'If RS.State = 1 Then RS.Close
'RS.Open "SELECT * FROM SNV WHERE INV_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
'For A = 1 To RS.RecordCount
'    For b = 2 To ms2.Rows - 1
'        If Trim(ms2.TextMatrix(b, 7)) = Trim(RS!shl_idy) Then
'            ms2.TextMatrix(b, 3) = Val(RS!i_shl_nos & "")
'            ms2.TextMatrix(b, 4) = Val(RS!i_pck_nos & "")
'            Exit For
'        End If
'    Next
'    RS.MoveNext
'Next
'For b = 2 To ms2.Rows - 1
'    ms2.TextMatrix(b, 5) = Val(ms2.TextMatrix(b, 1)) - Val(ms2.TextMatrix(b, 3))
'    ms2.TextMatrix(b, 6) = Val(ms2.TextMatrix(b, 2)) - Val(ms2.TextMatrix(b, 4))
'Next
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
newrecord
ElseIf Button.KEY = "s" Then
If Check3.Value = 1 And Trolley.Value = 1 Then
MakeBill
Else
saverecord
End If

ElseIf Button.KEY = "e" Then
exitrecord
ElseIf Button.KEY = "p" Then
Printrecord
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "RS" Then
RETURNSTOCKS
ElseIf Button.KEY = "crt" Then
    Picture2.Visible = True
    DTPicker1.SetFocus
    SSTab1.Enabled = False
    Toolbar1.Enabled = False
ElseIf Button.KEY = "Bill" Then
If Check3.Value = 1 And Trolley.Value = 1 Then
MakeBill
Else
saverecord
End If
ElseIf Button.KEY = "ed" Then
    If M_FRM_RSI.ActiveControl.name = "M_INV_IDY" Then
        Label49 = "Select Invoice"
        INVOICELIST
    ElseIf M_FRM_RSI.ActiveControl.name = "M_SMN_IDY" Then
        Label49 = "Select Salesman"
        SALESMANLIST
    ElseIf M_FRM_RSI.ActiveControl.name = "M_CST_IDY" Then
        Label49 = "Select Vehicle"
        V_DLR_LST
    ElseIf M_FRM_RSI.ActiveControl.name = "M_ACT_IDY" Then
        Label49 = "Select Account"
        V_ACT_LST
    ElseIf M_FRM_RSI.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        PRODUCTLIST
    End If
ElseIf Button.KEY = "DAY" Then
    If Button.Value = tbrPressed Then
        Text1.Visible = True
        Text1 = Format(Date, "dd/mm/yyyy") & " " & Format(Date, "dddd")
        Text1.BackColor = vbBlue
        Text1.ForeColor = vbWhite
    Else
        Text1.Visible = False
    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
If check("prt", "prt_nme", "M_FRM_RSI") = False Then
Check3.Value = 0
Else
Check3.Value = 1
End If
Check3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Private Sub M_INV_IDY_KeyDown(K As Integer, S As Integer)
If K = 13 Or K = 40 Then
If Len(Trim(M_INV_IDY)) = 0 Then
    INVOICELIST
Else
    inV = M_INV_IDY
    clrctr M_FRM_RSI
    txt1000 = "0.00"
    txt500 = "0.00"
    txt100 = "0.00"
    txt50 = "0.00"
    txt20 = "0.00"
    txt10 = "0.00"
    txt5 = "0.00"
    txt2 = "0.00"
    txt1 = "0.00"
    txtchange = "0.00"
    Label9 = ""
    Label10 = ""
    Label11 = ""
    Label17 = "0.00"
    Label18 = "0"
    Label19 = "0.00"
    Label20 = "0.00"
    M_GRS_AMT = "0.00"
    M_NET_AMT = "0.00"
    M_SMN_NME = ""
    Label40 = ""
    M_INV_IDY = inV
    DISPDETAILS
End If
End If
End Sub
Public Sub DISPDETAILS()
On Error GoTo LBL
If CHECKTB("van.dbf,prd.dbf,brd.dbf,act.dbf,smn.dbf,shl.dbf,snv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSVAN As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
If RSCUR.State = 1 Then RSCUR.Close
If Trolley.Value = 1 Then
    If Len(Trim(M_INV_IDY)) = 10 Then
        RSVAN.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "' AND VAN_CHK='P'", CON, adOpenKeyset, adLockOptimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
    Else
        RSVAN.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like 'RSI" & Tincrement(M_INV_IDY) & "' AND VAN_CHK='P'", CON, adOpenKeyset, adLockOptimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY LIKE 'RSI" & Tincrement(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
    End If
Else
    If Len(Trim(M_INV_IDY)) = 10 Then
        RSVAN.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "' AND VAN_CHK='V'", CON, adOpenKeyset, adLockOptimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
    Else
        RSVAN.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like 'RSI" & Tincrement(M_INV_IDY) & "' AND VAN_CHK='V' ", CON, adOpenKeyset, adLockOptimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY LIKE 'RSI" & Tincrement(Trim(M_INV_IDY)) & "'", CON, adOpenKeyset, adLockOptimistic
    End If
End If
If RSVAN.RecordCount > 0 Then
    If Len(Trim(M_INV_IDY)) = 10 Then
        M_INV_IDY = Trim(M_INV_IDY)
    Else
        M_INV_IDY = "RSI" & Tincrement(M_INV_IDY)
    End If
    TRP_NOS.Clear
    TRP_NOS.AddItem "All"
    If rsact.State = 1 Then rsact.Close
    rsact.Open "SELECT * FROM VAN WHERE INV_IDY LIKE '" & Trim(M_INV_IDY) & "' GROUP BY TRP_IDY", CON, adOpenKeyset, adLockOptimistic
    For A = 1 To rsact.RecordCount
        If rsact.RecordCount = 1 And Val(rsact!trp_idy & "") = 0 And Trolley.Value = 0 Then
            TRP_NOS.AddItem "All"
            Exit For
        ElseIf rsact.RecordCount = 1 And Val(rsact!trp_idy & "") = 0 And Trolley.Value = 1 Then
            TRP_NOS.AddItem "0"
        End If
        TRP_NOS.AddItem Val(rsact!trp_idy & "")
        rsact.MoveNext
    Next
    TRP_NOS = TRP_NOS.List(1)
    M_INV_IDY.Enabled = False
    M_INV_DAT = RSVAN("INV_DAT")
    Trolley.Enabled = False
    M_CST_IDY = Trim(RSVAN!cst_idy)
    M_CST_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(RSVAN!cst_idy) & "'")
    M_ACT_IDY = Trim(RSVAN!ACT_IDY)
    M_ACT_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(RSVAN!ACT_IDY) & "'")
    M_SMN_IDY = Trim(RSVAN!SMN_IDY)
    M_SMN_NME = GetValue("SMN_NME", "SMN WHERE SMN_IDY LIKE '" & Trim(RSVAN!SMN_IDY) & "'")
    IDN = False
    If RSVAN!VAN_CHK = "V" Then
        Trolley.Value = 0
        MS.ColWidth(7) = 1050
    Else
        Trolley.Value = 1
        MS.ColWidth(7) = 0
    End If
    
    If RSVAN!van_cls = "S" Then
        Check3.Value = 1
        Check3.Enabled = False
        Check4.Enabled = True
    ElseIf RSVAN!van_cls = "C" Then
        Check3.Value = 1
        Check3.Enabled = False
        Check4.Value = 1
        Check4.Enabled = False
    ElseIf RSVAN!van_cls = "N" Then
        Check3.Value = 0
        Check3.Enabled = True
        Check4.Value = 0
        Check4.Enabled = False
    Else
        Check3.Value = 0
    End If
    
    MS.Rows = RSVAN.RecordCount + 1
    'MS.FixedCols = 1
    For A = 1 To RSVAN.RecordCount
        MS.TextMatrix(A, 0) = Trim(RSVAN!PRD_NME)
        If Val(RSVAN!trp_idy & "") = 0 And Trolley.Value = 0 Then
            MS.TextMatrix(A, 1) = "All"
        ElseIf Val(RSVAN!trp_idy & "") = 0 And Trolley.Value = 1 Then
            MS.TextMatrix(A, 1) = "0"
        Else
            MS.TextMatrix(A, 1) = Trim(Val(RSVAN!trp_idy & ""))
        End If
        If Pieces.Value = 0 Then
            MS.TextMatrix(A, 2) = MTS(Val(RSVAN!OPN_QTY), RSVAN!sub_unt)
            MS.TextMatrix(A, 3) = MTS(Val(RSVAN!ISS_QTY), RSVAN!sub_unt)
            MS.TextMatrix(A, 4) = MTS(Val(RSVAN!tot_qty), RSVAN!sub_unt)
            MS.TextMatrix(A, 6) = MTS(Val(RSVAN!BIL_QTY), RSVAN!sub_unt)
            MS.TextMatrix(A, 5) = MTS(Val(RSVAN!ret_qty), RSVAN!sub_unt)
            MS.TextMatrix(A, 16) = MTS(Val(RSVAN!fre_qty & ""), RSVAN!sub_unt)
            MS.TextMatrix(A, 17) = MTS(Val(RSVAN!brk_qty & ""), RSVAN!sub_unt)
            MS.TextMatrix(A, 18) = MTS(Val(RSVAN!lkg_qty & ""), RSVAN!sub_unt)
            MS.TextMatrix(A, 8) = convert(Val(RSVAN!PRD_MRP) / Val(RSVAN!sub_unt))
            MS.TextMatrix(A, 15) = convert(Val(RSVAN!PRD_rat) / Val(RSVAN!sub_unt))
        Else
            MS.TextMatrix(A, 2) = QTY(RSVAN!OPN_QTY)
            MS.TextMatrix(A, 3) = QTY(RSVAN!ISS_QTY)
            MS.TextMatrix(A, 4) = QTY(RSVAN!tot_qty)
            MS.TextMatrix(A, 6) = QTY(RSVAN!BIL_QTY)
            MS.TextMatrix(A, 5) = QTY(RSVAN!ret_qty)
            MS.TextMatrix(A, 16) = QTY(RSVAN!fre_qty & "")
            MS.TextMatrix(A, 17) = QTY(RSVAN!brk_qty & "")
            MS.TextMatrix(A, 18) = QTY(RSVAN!lkg_qty & "")
            MS.TextMatrix(A, 8) = convert(RSVAN!PRD_MRP & "")
            MS.TextMatrix(A, 15) = convert(RSVAN!PRD_rat)
        End If
        M_CAS_AMT = convert(RSVAN!CAS_AMT & "")
        M_CHQ_AMT = convert(RSVAN!CHQ_AMT & "")
        M_DIS_AMT = convert(RSVAN!DIS_AMT & "")
        M_EXP_AMT = convert(RSVAN!EXP_AMT & "")
        M_TAX_AMT = convert(RSVAN!ADD_AMT & "")
        M_COM_AMT = convert(RSVAN!COM_AMT & "")
        MS.TextMatrix(A, 9) = convert(RSVAN!COL_AMT & "")
        'CHANGED
        MS.TextMatrix(A, 10) = Trim(RSVAN!PRD_IDY)
        MS.TextMatrix(A, 11) = Trim(RSVAN!prd_prn)
        MS.TextMatrix(A, 12) = Val(RSVAN!sub_unt)
        MS.TextMatrix(A, 13) = Trim(RSVAN!BAT_IDY)
        MS.TextMatrix(A, 14) = Trim(RSVAN!Bat_nme & "")
        RSVAN.MoveNext
    Next
    If Not RSCUR.EOF Then
        txt1000 = convert(Val(Trim(RSCUR(1)) & ""))
        txt500 = convert(Val(Trim(RSCUR(2)) & ""))
        txt100 = convert(Val(Trim(RSCUR(3)) & ""))
        txt50 = convert(Val(Trim(RSCUR(4)) & ""))
        txt20 = convert(Val(Trim(RSCUR(5)) & ""))
        txt10 = convert(Val(Trim(RSCUR(6)) & ""))
        txt5 = convert(Val(Trim(RSCUR(7)) & ""))
        txt2 = convert(Val(Trim(RSCUR(8)) & ""))
        txt1 = convert(Val(Trim(RSCUR(9)) & ""))
        txtchange = convert(Val(Trim(RSCUR(10)) & ""))
    End If
    ShlSet
    SETDISPDETAILS
    discalc
    calc
    list5.Visible = False
    MOD_REC = True
    SSTab1.Tab = 0
    M_INV_DAT.SetFocus
Else
    INVOICELIST
End If
Exit Sub
LBL:
If err.Number = 35773 Then
    MsgBox "Invalid Voucher Selected for this Financial Year", vbExclamation
    CANCELRECORD
Else
    MsgBox err.Description
    CANCELRECORD
End If
End Sub
Private Sub SETDISPDETAILS()
If CHECKTB("van.dbf,prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSVAN As New ADODB.Recordset
If Len(Trim(M_INV_IDY)) = 10 Then
    RSVAN.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
Else
    RSVAN.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like 'RSI" & Tincrement(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
End If
If RSVAN.RecordCount > 0 Then
    For A = 1 To RSVAN.RecordCount
        TOTVAL = Val(TOTVAL) + Val(RATCAL(QTY(Val(RSVAN!tot_qty & "")), Val(RSVAN!PRD_rat), Val(RSVAN!sub_unt)))
        RetVal = Val(RetVal) + Val(RATCAL(QTY(Val(RSVAN!ret_qty & "")), Val(RSVAN!PRD_rat), Val(RSVAN!sub_unt)))
        freval = Val(freval) + Val(RATCAL(QTY(Val(RSVAN!fre_qty & "")), Val(RSVAN!PRD_rat), Val(RSVAN!sub_unt)))
        brkval = Val(brkval) + Val(RATCAL(QTY(Val(RSVAN!brk_qty & "")), Val(RSVAN!PRD_rat), Val(RSVAN!sub_unt)))
        lkgval = Val(lkgval) + Val(RATCAL(QTY(Val(RSVAN!lkg_qty & "")), Val(RSVAN!PRD_rat), Val(RSVAN!sub_unt)))
        RSVAN.MoveNext
    Next
    RSVAN.MoveFirst
    MSS.TextMatrix(1, 1) = convert(Val(TOTVAL) - (Val(RetVal) + Val(brkval) + Val(lkgval)))
    
    MSS.TextMatrix(2, 1) = convert(Val(freval))
    'MSS.TextMatrix(3, 1) = convert(Val(brkval))
    'MSS.TextMatrix(4, 1) = convert(Val(lkgval))
    
    MSS.TextMatrix(7, 1) = convert(Val(RSVAN!DIS_AMT & ""))
    MSS.TextMatrix(14, 1) = convert(Val(RSVAN!EXP_AMT & ""))
    MSS.TextMatrix(8, 1) = convert(Val(RSVAN!ADD_AMT & ""))
    MSS.TextMatrix(10, 1) = convert(Val(RSVAN!CAS_AMT & ""))
    MSS.TextMatrix(11, 1) = convert(Val(RSVAN!CHQ_AMT & ""))
End If
SETCALC
End Sub
Private Sub M_SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Salesmen"
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If List1.SortOrder = lvwAscending Then
List1.SortOrder = lvwDescending
ElseIf List1.SortOrder = lvwDescending Then
List1.SortOrder = lvwAscending
End If
List1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub List1_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If List1.ListItems.Count > 0 Then
        M_SMN_IDY = List1.SelectedItem.ListSubItems.Item(1)
        M_SMN_IDY_KeyDown 13, 0
        List1.Visible = False
    End If
End If
End Sub

Private Sub List1_DblClick()
List1_keypress 13
End Sub

Private Sub M_SMN_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    If M_INV_IDY.Enabled = True Then
        M_INV_IDY.SetFocus
    Else
        M_INV_DAT.SetFocus
    End If
ElseIf SSTab1.Tab = 1 Then
    If MS.Rows > 1 Then
        MS.Row = 1
        MS.Col = 0
        MS.SetFocus
    End If
ElseIf SSTab1.Tab = 5 Then
    If RSVAN.State = 1 Then RSVAN.Close
    If Len(Trim(M_INV_IDY)) = 10 Then
        RSVAN.Open "SELECT * FROM VAN WHERE INV_IDY='" & Trim(M_INV_IDY) & "'", CON
    Else
        RSVAN.Open "SELECT * FROM VAN WHERE INV_IDY='RSI" & Tincrement(M_INV_IDY) & "'", CON
    End If
    'If rsvan.EOF Then SHL_SET.Enabled = False
    SHL_SET.SetFocus
'    If ms2.Rows > 1 Then ms2.Row = 1
'    ms2.SetFocus
ElseIf SSTab1.Tab = 6 Then
    ms1.SetFocus
    If RSVAN.State = 1 Then RSVAN.Close
    If Len(Trim(M_INV_IDY)) = 10 Then
        RSVAN.Open "SELECT * FROM VAN WHERE INV_IDY='" & Trim(M_INV_IDY) & "'", CON
    Else
        RSVAN.Open "SELECT * FROM VAN WHERE INV_IDY='RSI" & Tincrement(M_INV_IDY) & "'", CON
    End If
    If RSVAN.EOF Then Command6.Enabled = False
    If ms1.Row > 1 Then ms1.Row = 2
    MS.Col = 1
    If Command6.Enabled = True Then Command6.SetFocus
ElseIf SSTab1.Tab = 2 Then
    M_GRS_SAL.SetFocus
End If
If List1.Visible = True Then List1.Visible = False
If list4.Visible = True Then list4.Visible = False
If list5.Visible = True Then list5.Visible = False

If SSTab1.Tab <> 1 Then
'    tax.Visible = False
'    units.Visible = False
    Total.Visible = False
Else
'    tax.Visible = True
'    units.Visible = True
    Total.Visible = True
End If
Text3.Visible = False
amount.Visible = False
If SSTab1.Tab = 3 Then
Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub M_SMN_IDY_KeyDown(K As Integer, S As Integer)
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
Dim REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("smn.dbf,inl.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from smn where smn_idy like '" & Sincrement(Trim(M_SMN_IDY)) & "'", CON, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        If REC_SET1.State = 1 Then REC_SET1.Close
        REC_SET1.Open "SELECT * FROM INL WHERE SMN_IDY LIKE '" & Sincrement(Trim(M_SMN_IDY)) & "' AND TXN_TYP LIKE 'RSI'", CON, adOpenKeyset, adLockOptimistic
        If REC_SET1.RecordCount > 0 Then
            For A = 1 To REC_SET1.RecordCount
                If REC_SET2.State = 1 Then REC_SET2.Close
                REC_SET2.Open "SELECT * FROM INL WHERE ORD_IDY LIKE '" & Trim(REC_SET1("INV_IDY")) & "' AND TXN_TYP LIKE 'RSR'", CON, adOpenKeyset, adLockOptimistic
                If REC_SET2.RecordCount = 0 Then
                    If MOD_REC = False Then
                        MsgBox "Return Stock Should be posted for Salesman Mr." & Trim(REC_SET("smn_nme"))
                        M_SMN_IDY.SetFocus
                        M_SMN_IDY = ""
                        SSTab1.Tab = 0
                    Else
                        SSTab1.Tab = 1
                        If MOD_REC = False Then Toolbar1.Buttons(12).Enabled = True
                        MS.SetFocus
                    End If
                    Exit Sub
                Else
                    M_SMN_IDY = Sincrement(M_SMN_IDY)
                    M_SMN_NME = Trim(REC_SET("smn_nme"))
                    If K <> 38 Then
                        SSTab1.Tab = 1
                        If MOD_REC = False Then Toolbar1.Buttons(12).Enabled = True
                        MS.SetFocus
                    Else
                        TRP_NOS.SetFocus
                    End If
                End If
                REC_SET1.MoveNext
            Next
        Else
            If K = 38 Then
                If TRP_NOS.Enabled = True Then
                    TRP_NOS.SetFocus
                Else
                    M_CST_IDY.SetFocus
                End If
            Else
                M_SMN_IDY = Sincrement(M_SMN_IDY)
                M_SMN_NME = Trim(REC_SET("smn_nme"))
                If MOD_REC = False Then Toolbar1.Buttons(12).Enabled = True
                SSTab1.Tab = 1
                MS.SetFocus
            End If
        End If
    Else
        SALESMANLIST
    End If
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_LeaveCell()
If Text3.Visible = True Then Text3.Visible = False
If amount.Visible = True Then amount.Visible = False
If List6.Visible = True Then List6.Visible = False
End Sub

Private Sub MS_KeyPress(K As Integer)
If MOD_REC = False And TRP_NOS.Enabled = True Then
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K > 47 And K < 58 Or K = 46 Then
End If
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If MS.Col = 0 And MS.TextMatrix(MS.Row, 0) = "" Then
        Text3 = ""
        PRODUCTLIST
    ElseIf MS.Col = 3 And MS.TextMatrix(MS.Rows - 1, 0) <> "" And Check3.Value = 0 And Check3.Enabled = True Then
        MS.AddItem ""
        MS.Col = 0
        MS.Row = MS.Rows - 1
        Exit Sub
    Else
        If MS.Cols - 1 <> MS.Col Then
        MS.Col = MS.Col + 1
        End If
    End If
End If
If MS.Row <> 0 Then
    If MS.Col = 0 Then
    Text3 = ""
        If K = 13 Then
            If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                MS.Col = 2
                MS.SetFocus
            Else
                Text3 = ""
                PRODUCTLIST
            End If
        ElseIf K <> 27 Then
            If K = vbKeySpace Then
            Else
                If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                Else
                Text3.Visible = True
                Text3.Top = MS.Top + MS.CellTop
                Text3.Left = MS.CellLeft + MS.Left
                Text3.Width = MS.CellWidth
                Text3.Height = MS.CellHeight
                Text3.Visible = True
                Text3.SetFocus
                If K <> 13 Then Text3 = Chr(K)
                Text3.SelStart = Len(Text3)
                End If
            End If
        End If
    ElseIf MS.Col = 3 Then
        If MS.TextMatrix(MS.Row, 0) <> "" And Check3.Value = 0 Then
               If K > 47 And K < 58 Or K = 46 Then
                    amount = ""
                    amount.MaxLength = 10
                    amount.Visible = True
                    amount.Top = MS.Top + MS.CellTop
                    amount.Left = MS.CellLeft + MS.Left
                    amount.Width = MS.CellWidth
                    amount.Height = MS.CellHeight
                    amount.Visible = True
                    amount.SetFocus
                    amount = Chr(K)
                    amount.SelStart = Len(amount)
               ElseIf K = 13 And MOD_REC = False Then
                    If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                        MS.AddItem ""
                        MS.Col = 0
                        MS.Row = MS.Rows - 1
                    Else
                        MS.Col = MS.Col + 1
                    End If
            Else
                K = 0
            End If
        End If
        ElseIf MS.Col = 2 Then
        If MS.TextMatrix(MS.Row, 0) <> "" And Check3.Value = 0 Then
               If K > 47 And K < 58 Or K = 46 Then
                    amount = ""
                    amount.MaxLength = 10
                    amount.Visible = True
                    amount.Top = MS.Top + MS.CellTop
                    amount.Left = MS.CellLeft + MS.Left
                    amount.Width = MS.CellWidth
                    amount.Height = MS.CellHeight
                    amount.Visible = True
                    amount.SetFocus
                    amount = Chr(K)
                    amount.SelStart = Len(amount)
                ElseIf K = 13 And MOD_REC = False Then
                    If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                        MS.AddItem ""
                        MS.Col = 0
                        MS.Row = MS.Rows - 1
                    Else
                        MS.Col = MS.Col + 1
                    End If
            Else
                K = 0
            End If
        End If
    ElseIf MS.Col = 5 Or MS.Col = 16 Or MS.Col = 17 Or MS.Col = 18 Then
        If MS.TextMatrix(MS.Row, 0) <> "" And Check3.Value = 1 And Val(MS.TextMatrix(MS.Row, 1)) = Val(TRP_NOS) Then
           If K > 47 And K < 58 Or K = 46 Then
                amount = ""
                amount.MaxLength = 10
                amount.Visible = True
                amount.Top = MS.Top + MS.CellTop
                amount.Left = MS.CellLeft + MS.Left
                amount.Width = MS.CellWidth
                amount.Height = MS.CellHeight
                amount.Visible = True
                amount.SetFocus
                amount = Chr(K)
                amount.SelStart = Len(amount)
            ElseIf K = 13 And MOD_REC = False Then
                If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                    MS.AddItem ""
                    MS.Col = 0
                    MS.Row = MS.Rows - 1
                Else
                    MS.Col = MS.Col + 1
                End If
            Else
                K = 0
            End If
        End If
    End If
End If
End Sub

Private Sub Text3_KeyDown(K As Integer, S As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If MS.Col = 0 Then
        If CHECKTB("prd.dbf,brd.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM prd,brd WHERE brd.prd_idy like prd.prd_idy and brd.prd_IDY LIKE '" & Sincrement(Text3) & "' AND PRD_QTY >0 AND PRD.STK_TYP IN('Salable','Free')", CON, adOpenKeyset
        If REC_SET.RecordCount = 1 Then
            dispprd Trim(REC_SET!BAT_IDY)
        ElseIf REC_SET.RecordCount > 1 Then
            Text3 = Sincrement(Text3)
            PRODUCTIDLIST
        Else
            PRODUCTLIST
       End If
       End If
End If
End Sub

Public Sub calc()
M_GRS_AMT = 0
For A = 1 To MS.Rows - 1
M_GRS_AMT = Val(M_GRS_AMT) + Val(MS.TextMatrix(A, 9))
M_GRS_AMT = convert(AMT(M_GRS_AMT & ""))
Next
M_NET_AMT = convert(Val(M_GRS_AMT))
M_NET_SAL = convert(Val(M_GRS_SAL) - Val(M_FRE_QTY))
M_CRT_AMT = convert((Val(M_NET_SAL) + Val(M_TAX_AMT)) - (Val(M_CAS_AMT) + Val(M_CHQ_AMT) + Val(M_DIS_AMT) + Val(M_COM_AMT) + Val(M_EXP_AMT)))
End Sub

Private Sub amount_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
ElseIf K = 13 Then
If MS.Col = 3 Then
        If CHECKTB("van.dbf,prd.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
        amount = Val(amount)
        If REC_SET1.State = 1 Then REC_SET1.Close
        If REC_SET2.State = 1 Then REC_SET2.Close
        If REC_SET3.State = 1 Then REC_SET3.Close
        If Trim(M_INV_IDY) <> "" Then
            REC_SET2.Open "select TOT_QTY,SUB_UNT from VAN,PRD where VAN.PRD_IDY LIKE PRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 13)) & "' and inv_idy like '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
            If REC_SET2.RecordCount > 0 Then
                For A = 1 To REC_SET2.RecordCount
                    INVTOT = Val(INVTOT) + MTS(Val(REC_SET2(0)), Val(REC_SET2!sub_unt))
                    REC_SET2.MoveNext
                Next
                REC_SET2.MoveFirst
                INVTOT = QTY(STM(Val(INVTOT), Val(REC_SET2!sub_unt)))
            End If
        End If
        If Pieces.Value = 1 Then
        STK = Val(MTS(Val(MS.TextMatrix(MS.Row, 2)), Val(MS.TextMatrix(MS.Row, 12))))
        STK = STK + Val(MTS(Val(amount), Val(MS.TextMatrix(MS.Row, 12))))
        ElseIf Pieces.Value = 0 Then
            amount = Int(AMT(Val(amount)))
            STK = Val(MS.TextMatrix(MS.Row, 2)) + Int(Val(amount))
            pcs = Int((Val(STK)))
            GoTo PiecesTrack
        End If
        b = InStr(amount, ".")
        If b <> 0 Then M = Mid(amount, b + 1, Len(amount) - b + 1)
        If Len(M) = 1 Then
            amount = amount & "00"
        ElseIf Len(M) = 2 Then
            amount = amount & "0"
        ElseIf Len(M) = 0 Then
            amount = amount & ".000"
        End If
        If b > 0 And Val(MS.TextMatrix(MS.Row, 12)) > 0 Then
            If InStr(b + 1, amount, ".") > 0 Then
                MS.TextMatrix(MS.Row, 4) = "0.000"
                Exit Sub
            End If
            M = Mid(amount, b + 1, 3)
            If M >= Val(MS.TextMatrix(MS.Row, 12)) Then
                MsgBox "Product Sub Units Are " & Val(MS.TextMatrix(MS.Row, 12)), vbExclamation, "Effmcg"
                amount.Visible = True
                amount.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                M_QTY = Val(STK)
            End If
        Else
            M_QTY = Val(STK)
        End If
PiecesTrack:
        If CHECKTB("dir.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
        If Pieces.Value = 0 Then M_QTY = pcs
            REMIND = MTS(Val(ADDBRDQTY(QTY(Val(INVTOT)), Trim(MS.TextMatrix(MS.Row, 13)))), Val(MS.TextMatrix(MS.Row, 12)))
            For A = 1 To MS.Rows - 1
                If MS.Row <> A Then
                    If Trim(MS.TextMatrix(MS.Row, 13)) = Trim(MS.TextMatrix(A, 13)) Then
                        If Pieces = 0 Then
                            ENTQTY = Val(ENTQTY) + Val(MS.TextMatrix(A, 4))
                        Else
                            ENTQTY = Val(ENTQTY) + MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                        End If
                    End If
                End If
            Next
            CURQTY = Val(M_QTY) + Val(ENTQTY)
            If querycheck("DIR", "DIR_MSG", "U", " and dir_idy like 'SYS' AND DIR_IDX LIKE 'STK'") = False Then
                If Val(REMIND) - Val(CURQTY) < 0 Then
                    MsgBox "Stock Out Of Range, Check Stock !", vbExclamation, "Effmcg"
                    If amount.Visible = True Then
                        amount.SetFocus
                    Else
                        MS.SetFocus
                    End If
                    SendKeys "{home}+{end}"
                    Exit Sub
                 End If
            End If
            MS.Text = amount
            If Pieces.Value = 1 Then
                MS.TextMatrix(MS.Row, 4) = QTY(STM(Val(STK), MS.TextMatrix(MS.Row, 12)))
            Else
                MS.TextMatrix(MS.Row, 4) = Val(STK)
            End If
            amount.Visible = False
            If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                MS.AddItem ""
                MS.Col = 0
                MS.Row = MS.Rows - 1
            End If
    MS.SetFocus
    discalc
ElseIf MS.Col = 2 Then
        If CHECKTB("van.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
        amount = Val(amount)
        If REC_SET1.State = 1 Then REC_SET1.Close
        If REC_SET2.State = 1 Then REC_SET2.Close
        If REC_SET3.State = 1 Then REC_SET3.Close
        If Trim(M_INV_IDY) <> "" Then
             REC_SET2.Open "select TOT_QTY,SUB_UNT from VAN where BAT_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 13)) & "' and inv_idy like '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
            For A = 1 To REC_SET2.RecordCount
                INVTOT = Val(INVTOT) + MTS(Val(REC_SET2(0)), Val(REC_SET2!sub_unt))
                REC_SET2.MoveNext
            Next
            REC_SET2.MoveFirst
            INVTOT = QTY(STM(Val(INVTOT), Val(REC_SET2!sub_unt)))
        End If
        If Pieces.Value = 1 Then
        STK = Val(MTS(Val(MS.TextMatrix(MS.Row, 3)), Val(MS.TextMatrix(MS.Row, 12))))
        STK = STK + Val(MTS(Val(amount), Val(MS.TextMatrix(MS.Row, 12))))
        ElseIf Pieces.Value = 0 Then
            amount = Int(AMT(Val(amount)))
            STK = Val(MS.TextMatrix(MS.Row, 3)) + Int(Val(amount))
            pcs = Int((Val(STK)))
            GoTo PiecesTrack1
        End If
        b = InStr(amount, ".")
        If b <> 0 Then M = Mid(amount, b + 1, Len(amount) - b + 1)
        If Len(M) = 1 Then
            amount = amount & "00"
        ElseIf Len(M) = 2 Then
            amount = amount & "0"
        ElseIf Len(M) = 0 Then
            amount = amount & ".000"
        End If
        If b > 0 And Val(MS.TextMatrix(MS.Row, 12)) > 0 Then
            If InStr(b + 1, amount, ".") > 0 Then
                MS.TextMatrix(MS.Row, 4) = "0.000"
                Exit Sub
            End If
            M = Mid(amount, b + 1, 3)
            If M >= Val(MS.TextMatrix(MS.Row, 12)) Then
                MsgBox "Product Sub Units Are " & Val(MS.TextMatrix(MS.Row, 12)), vbExclamation, "Effmcg"
                amount.Visible = True
                amount.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                M_QTY = Val(STK)
            End If
        Else
            M_QTY = Val(STK)
        End If
PiecesTrack1:
        If CHECKTB("dir.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
        If Pieces.Value = 0 Then M_QTY = pcs
            REMIND = MTS(Val(ADDBRDQTY(QTY(Val(INVTOT)), Trim(MS.TextMatrix(MS.Row, 13)))), Val(MS.TextMatrix(MS.Row, 12)))
            For A = 1 To MS.Rows - 1
                If MS.Row <> A Then
                    If Trim(MS.TextMatrix(MS.Row, 13)) = Trim(MS.TextMatrix(A, 13)) Then
                        If Pieces = 0 Then
                            ENTQTY = Val(ENTQTY) + Val(MS.TextMatrix(A, 4))
                        Else
                            ENTQTY = Val(ENTQTY) + MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                        End If
                    End If
                End If
            Next
            CURQTY = Val(M_QTY)
            If querycheck("DIR", "DIR_MSG", "U", " and dir_idy like 'SYS' AND DIR_IDX LIKE 'STK'") = False Then
                If Val(REMIND) - Val(CURQTY) < 0 Then
                    MsgBox "Stock Out Of Range, Check Stock !", vbExclamation, "Effmcg"
                    If amount.Visible = True Then
                        amount.SetFocus
                    Else
                        MS.SetFocus
                    End If
                    SendKeys "{home}+{end}"
                    Exit Sub
                 End If
            End If
            MS.Text = amount
            If Pieces.Value = 1 Then
                MS.TextMatrix(MS.Row, 4) = QTY(STM(Val(STK), Val(MS.TextMatrix(MS.Row, 12))))
            Else
                MS.TextMatrix(MS.Row, 4) = Val(STK)
            End If
            amount.Visible = False
            If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                MS.AddItem ""
                MS.Col = 0
                MS.Row = MS.Rows - 1
            End If
    MS.SetFocus
    discalc
ElseIf MS.Col = 5 Or MS.Col = 16 Or MS.Col = 17 Or MS.Col = 18 Then
        If Pieces.Value = 1 Then
            amount = QTY(amount)
            b = InStr(amount, ".")
            If b > 0 And Val(MS.TextMatrix(MS.Row, 12)) > 0 Then
                If InStr(b + 1, amount, ".") > 0 Then
                    MS.TextMatrix(MS.Row, 4) = "0.000"
                    Exit Sub
                End If
                M = Mid(amount, b + 1, 4)
                If M >= Val(MS.TextMatrix(MS.Row, 12)) Then
                    MsgBox "Product Sub Units Are " & Val(MS.TextMatrix(MS.Row, 12)), vbExclamation, "Effmcg"
                    If amount.Visible = True Then
                        amount.SetFocus
                        SendKeys "{home}+{end}"
                        Exit Sub
                    Else
                        amount.Visible = False
                        MS.SetFocus
                    End If
                Else
                    amount.Visible = False
                    MS.Text = QTY(amount)
                    P_BIL_QTY
                    discalc
                    If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                    MS.SetFocus
                End If
            Else
                MS.Text = QTY(amount)
                If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                P_BIL_QTY
                discalc
                amount.Visible = False
                If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                MS.SetFocus
            End If
        Else
                MS.Text = amount
                amount.Visible = False
                If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                P_BIL_QTY
                discalc
            MS.SetFocus
        End If
    Else
        K = 0
    End If
End If
End Sub

Public Sub dispprd(prdid As String)
If CHECKTB("prd.dbf,brd.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
MSROW = MS.Row
Dim REC_SET As New ADODB.Recordset
For A = 1 To MS.Rows - 1
    If Trim(prdid) = MS.TextMatrix(A, 13) And Val(TRP_NOS) = Val(MS.TextMatrix(A, 1)) Then
        SSTab1.Tab = 1
        list4.Visible = False
        MsgBox "Product Already Exists In Invoice !", vbExclamation, "Effmcg"
        Text3.Visible = False
        MS.SetFocus
        MS.Col = 0
        MS.Row = A
        Exit Sub
    End If
Next
REC_SET.Open "select * from PRD,BRD,ACT where BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.BAT_idy like '" & Trim(prdid) & "' AND ACT.ACT_IDY LIKE '" & Trim(M_CST_IDY) & "%'", CON, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    'MS.FixedCols = 1
    list4.Visible = False
    MS.TextMatrix(MS.Row, 0) = REC_SET("PRD_NME")
    A = "0.000"
    If Pieces.Value = 0 Then A = 0
    MS.TextMatrix(MS.Row, 1) = Val(TRP_NOS)
    MS.TextMatrix(MS.Row, 2) = A
    MS.TextMatrix(MS.Row, 2) = A
    MS.TextMatrix(MS.Row, 3) = A
    MS.TextMatrix(MS.Row, 4) = A
    MS.TextMatrix(MS.Row, 5) = A
    MS.TextMatrix(MS.Row, 6) = A
    MS.TextMatrix(MS.Row, 7) = A
    If Pieces.Value = 1 Then
        MS.TextMatrix(MS.Row, 8) = convert(REC_SET!PRD_MRP)
        If REC_SET("DBT_RAP") = "A" Then
            MS.TextMatrix(MS.Row, 15) = convert(REC_SET!BAS_RAT)
        ElseIf REC_SET("DBT_RAP") = "B" Then
            MS.TextMatrix(MS.Row, 15) = convert(REC_SET!prd_sdr)
        ElseIf REC_SET("DBT_RAP") = "C" Then
            MS.TextMatrix(MS.Row, 15) = convert(REC_SET!PRD_MRP)
        End If
    Else
        MS.TextMatrix(MS.Row, 8) = convert(Val(REC_SET!PRD_MRP) / Val(REC_SET!sub_unt))
        If REC_SET("DBT_RAP") = "A" Then
            MS.TextMatrix(MS.Row, 15) = convert(Val(REC_SET!BAS_RAT) / Val(REC_SET!sub_unt))
        ElseIf REC_SET("DBT_RAP") = "B" Then
            MS.TextMatrix(MS.Row, 15) = convert(Val(REC_SET!prd_sdr) / Val(REC_SET!sub_unt))
        ElseIf REC_SET("DBT_RAP") = "C" Then
            MS.TextMatrix(MS.Row, 15) = convert(Val(REC_SET!PRD_MRP) / Val(REC_SET!sub_unt))
        End If
    End If
    MS.TextMatrix(MS.Row, 9) = ""
    MS.TextMatrix(MS.Row, 10) = Trim(REC_SET("prd_idy"))
    MS.TextMatrix(MS.Row, 11) = Trim(REC_SET("prd_prn"))
    MS.TextMatrix(MS.Row, 12) = Val(REC_SET!sub_unt)
    MS.TextMatrix(MS.Row, 13) = Trim(REC_SET!BAT_IDY)
    MS.TextMatrix(MS.Row, 14) = Trim(REC_SET!Bat_nme)
    Text3.Visible = False
    calc
    SSTab1.Tab = 1
    MS.Row = MSROW
    MS.Col = 3
    MS.SetFocus
End If
End Sub

Public Sub discalc()
units = 0
tax = 0
Total = 0
M_GRS_SAL = 0
M_FRE_QTY = 0
M_BRK_QTY = 0
M_LKG_QTY = 0
For A = 1 To MS.Rows - 1
    'Customer Discount calculating on units and rate
    If MS.TextMatrix(A, 0) <> "" Then
        If Pieces.Value = 0 Then
            MS.TextMatrix(A, 9) = convert(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 6)))
            If Trolley.Value = 1 Then MS.TextMatrix(A, 6) = Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 16)) + Val(MS.TextMatrix(A, 17)) + Val(MS.TextMatrix(A, 18)))
            MS.TextMatrix(A, 7) = Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 17)) + Val(MS.TextMatrix(A, 18)))
            M_GRS_SAL = convert(Val(M_GRS_SAL) + (Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 5)))) * Val(MS.TextMatrix(A, 15)))
            M_FRE_QTY = convert(Val(M_FRE_QTY) + Val(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 16))))
            'M_BRK_QTY = convert(Val(M_BRK_QTY) + Val(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 17))))
            'M_LKG_QTY = convert(Val(M_LKG_QTY) + Val(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 18))))
        Else
            MS.TextMatrix(A, 9) = convert(RATCAL(QTY(Val(MS.TextMatrix(A, 6) & "")), Val(MS.TextMatrix(A, 15) & ""), Val(MS.TextMatrix(A, 12))))
            Aa = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
            b = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
            C = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
            d = MTS(Val(MS.TextMatrix(A, 16)), Val(MS.TextMatrix(A, 12)))
            e = MTS(Val(MS.TextMatrix(A, 17)), Val(MS.TextMatrix(A, 12)))
            f = MTS(Val(MS.TextMatrix(A, 18)), Val(MS.TextMatrix(A, 12)))
            ' + VAL(C)) AFTER VAL(B)
            M_GRS_SAL = convert(Val(M_GRS_SAL) + (RATCAL(STM(Val(Aa) - (Val(b) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 12))), Val(MS.TextMatrix(A, 15)), Val(MS.TextMatrix(A, 12)))))
            M_FRE_QTY = convert(Val(M_FRE_QTY) + Val(RATCAL(STM(Val(d), Val(MS.TextMatrix(A, 12))), Val(MS.TextMatrix(A, 15)), Val(MS.TextMatrix(A, 12)))))
            'M_BRK_QTY = convert(Val(M_BRK_QTY) + Val(RATCAL(STM(Val(e), Val(MS.TextMatrix(A, 12))), Val(MS.TextMatrix(A, 15)), Val(MS.TextMatrix(A, 12)))))
            'M_LKG_QTY = convert(Val(M_LKG_QTY) + Val(RATCAL(STM(Val(f), Val(MS.TextMatrix(A, 12))), Val(MS.TextMatrix(A, 15)), Val(MS.TextMatrix(A, 12)))))
            MS.TextMatrix(A, 7) = QTY(STM(Val(Aa) - (Val(b) + Val(C) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 12))))
            If Trolley.Value = 1 Then MS.TextMatrix(A, 6) = QTY(STM(Val(Aa) - (Val(b) + Val(d) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 12))))
        End If
        Total = convert(AMT(Val(Total) + Val(MS.TextMatrix(A, 9))))
    End If
Next
If Check3.Value = 1 And Check3.Enabled = True Then COMMISSION
calc
M_GRS_AMT = convert(AMT(Total))
Total = convert(AMT(Total))
M_NET_AMT = convert(AMT(Val(M_GRS_AMT)))
GRS_CAL
End Sub

Public Sub PRODUCTLIST()
If CHECKTB("prt.dbf,prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("prt", "prt_nme", "p_prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') AND PRD_QTY > 0 ORDER BY BRW_SEQ", CON, adOpenKeyset
Else
    REC_SET.Open "SELECT * FROM brd,prd WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD_PRN LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') AND PRD_QTY > 0 ORDER BY BRW_SEQ", CON, adOpenKeyset
End If
list4.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(REC_SET!PRD_MRP)
        ls.ListSubItems.Add , , convert(REC_SET!prd_sdr)
        If Val(REC_SET("PRD_QTY") & "") <> 0 Then
            ls.ListSubItems.Add , , Trim(QTY(REC_SET("PRD_QTY") & ""))
        Else
            ls.ListSubItems.Add , , "0.000"
        End If
        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_IDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_NME") & "")
        ls.ListSubItems.Add , , convert(REC_SET("PRD_PDR") & "")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_dom"), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_doe"), "dd/MM/yyyy")
        REC_SET.MoveNext
    Next
    Label49 = "Select Product"
    SSTab1.Tab = 3
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    SSTab1.Tab = 1
    If Text3.Visible = True Then
        Text3.SetFocus
    Else
        Check2.Caption = "Place"
        MS.SetFocus
    End If
    End If
End If
End Sub
Public Sub PRODUCTIDLIST()
If CHECKTB("prt.dbf,prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("prt", "prt_nme", "p_prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND prd.PRD_idy LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') AND PRD_QTY >0 ORDER BY BRW_SEQ", CON, adOpenKeyset
list4.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(REC_SET!PRD_MRP)
        ls.ListSubItems.Add , , convert(REC_SET!prd_sdr)
        If Val(REC_SET("PRD_QTY") & "") <> 0 Then
            ls.ListSubItems.Add , , Trim(QTY(REC_SET("PRD_QTY") & ""))
        Else
            ls.ListSubItems.Add , , "0.000"
        End If
        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_IDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_NME") & "")
        ls.ListSubItems.Add , , convert(REC_SET("PRD_PDR") & "")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_dom"), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_doe"), "dd/MM/yyyy")
        REC_SET.MoveNext
    Next
    Label49 = "Select Product"
    SSTab1.Tab = 3
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    SSTab1.Tab = 1
    If Text3.Visible = True Then
        Text3.SetFocus
    Else
        Check2.Caption = "Place"
        MS.SetFocus
    End If
    End If
End If
End Sub
Public Sub Printrecord()
If check("PRT", "PRT_NME", "M_FRM_RSI") = True Then
M_FRM_PRV.Check3.Value = 1
End If
M_FRM_PRV.Tag = "RSI"
M_FRM_PRV.Show 1
End Sub

Public Sub setval()
If list5.Visible = False Then clrctr M_FRM_RSI
Check4.Enabled = True
Check3.Enabled = True
Check3.Value = 0
Check4.Value = 0
RET_STK = False
Label17 = "0.00"
Label18 = "0"
Label19 = "0.00"
Label20 = "0.00"
M_GRS_SAL = "0.00"
M_FRE_QTY = "0.00"
M_NET_SAL = "0.00"
M_GRS_SAL = "0.00"
M_DIS_AMT = "0.00"
M_EXP_AMT = "0.00"
M_CAS_AMT = "0.00"
M_CHQ_AMT = "0.00"
M_CRT_AMT = "0.00"
M_NET_AMT = "0.00"
txt1000 = "0.00"
txt500 = "0.00"
txt100 = "0.00"
txt50 = "0.00"
txt20 = "0.00"
txt10 = "0.00"
txt5 = "0.00"
txt2 = "0.00"
txt1 = "0.00"
txtchange = "0.00"
Label41 = "0.00"
M_SMN_NME = ""
M_ACT_NME = ""
M_CST_NME = ""
Label40 = ""
ad = "0.00"
cd = "0.00"
units = "0.00"
tax = "0.00"
Total = "0.00"
M_INV_DAT = Date
End Sub

Public Sub AddOpeningStock()
If CHECKTB("van.dbf,prd.dbf,brd.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET  As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
SSTab1.Tab = 1
list5.Visible = False
MS.SetFocus
REC_SET.Open "select SUM((INT(RET_QTY)*SUB_UNT)+((RET_QTY-INT(RET_QTY))*1000)),* from VAN,PRD,BRD where inv_idy like '" & Trim(list5.SelectedItem) & "' and VAN.BAT_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY GROUP BY VAN.PRD_IDY ", CON, adOpenKeyset, adLockOptimistic
REC_SET1.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
MS.Rows = REC_SET.RecordCount + 1
For A = 1 To REC_SET.RecordCount
    MS.TextMatrix(A, 0) = REC_SET("prd_nme")
    MS.TextMatrix(A, 1) = Val(TRP_NOS)
    If Pieces.Value = 1 Then
        MS.TextMatrix(A, 2) = QTY(STM(Val(REC_SET(0) & ""), Val(REC_SET!sub_unt)))
        MS.TextMatrix(A, 4) = QTY(STM(Val(REC_SET(0) & ""), Val(REC_SET!sub_unt)))
        MS.TextMatrix(A, 8) = convert(REC_SET("prd_mrp"))
        If REC_SET1.RecordCount > 0 Then
            If REC_SET1("DBT_RAP") = "A" Then
                MS.TextMatrix(A, 15) = convert(REC_SET!BAS_RAT)
            ElseIf REC_SET1("DBT_RAP") = "B" Then
                MS.TextMatrix(A, 15) = convert(REC_SET!prd_sdr)
            ElseIf REC_SET1("DBT_RAP") = "C" Then
                MS.TextMatrix(A, 15) = convert(REC_SET!PRD_MRP)
            End If
        Else
            MS.TextMatrix(A, 15) = convert(REC_SET!prd_sdr)
        End If
    Else
        MS.TextMatrix(A, 2) = Val(REC_SET(0))
        MS.TextMatrix(A, 4) = MS.TextMatrix(A, 2)
        MS.TextMatrix(A, 8) = convert(REC_SET("prd_mrp") / Val(REC_SET!sub_unt))
        If REC_SET1.RecordCount > 0 Then
            If REC_SET1("DBT_RAP") = "A" Then
                MS.TextMatrix(A, 15) = convert(Val(REC_SET!BAS_RAT) / Val(REC_SET!sub_unt))
            ElseIf REC_SET1("DBT_RAP") = "B" Then
                MS.TextMatrix(A, 15) = convert(Val(REC_SET!prd_sdr) / Val(REC_SET!sub_unt))
            ElseIf REC_SET1("DBT_RAP") = "C" Then
                MS.TextMatrix(A, 15) = convert(Val(REC_SET!PRD_MRP) / Val(REC_SET!sub_unt))
            End If
        Else
            MS.TextMatrix(A, 15) = convert(Val(REC_SET!prd_sdr) / Val(REC_SET!sub_unt))
        End If
    End If
    MS.TextMatrix(A, 10) = Trim(REC_SET("PRD_IDY"))
    MS.TextMatrix(A, 11) = Trim(REC_SET("PRD_PRN"))
    MS.TextMatrix(A, 12) = Trim(REC_SET("SUB_UNT"))
    MS.TextMatrix(A, 13) = Trim(REC_SET("BAT_IDY"))
    MS.TextMatrix(A, 14) = Trim(REC_SET("BAT_NME"))
    REC_SET.MoveNext
Next
discalc
End Sub

Public Sub V_RSR_LST()
If CHECKTB("smn.dbf,van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Trolley.Value = 1 Then
    REC_SET.Open "select * from SMN,VAN where VAN.smn_idy like smn.smn_idy AND INV_DAT = CTOD('" & Trim(DTPicker1) & "') AND VAN_CHK= 'P' AND VAN_CLS='C' GROUP BY INV_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    REC_SET.Open "select * from SMN,VAN where VAN.smn_idy like smn.smn_idy AND INV_DAT = CTOD('" & Trim(DTPicker1) & "') AND VAN_CHK= 'V' AND VAN_CLS='C' GROUP BY INV_IDY", CON, adOpenKeyset, adLockOptimistic
End If
list5.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET!INV_IDY))
        ls.ListSubItems.Add , , Format(REC_SET!INV_DAT, "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_nme") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Van Challan"
    list5.Visible = True
    list5.SetFocus
Else
    MsgBox "Van Challan Number Not Found...!", vbExclamation, "Effmcg"
    MS.SetFocus
End If
End Sub

Public Function CHECKSTOCK() As Boolean
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Function
End If
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
For A = 1 To MS.Rows - 1
    If MS.TextMatrix(A, 0) <> "" Then
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM PRD,brd WHERE prd.PRD_IDY LIKE brd.prd_idy and brd.BAT_idy like '" & Trim(MS.TextMatrix(A, 12)) & "'", CON, adOpenKeyset, adLockOptimistic
        If RS.RecordCount > 0 Then
            If MOD_REC = True Then
            Else
            If Pieces.Value = 1 Then
                For d = 1 To MS.Rows - 1
                    If A <> d Then
                        If Trim(MS.TextMatrix(A, 12)) = Trim(MS.TextMatrix(d, 12)) Then
                            b = Val(b) + MTS(Val(MS.TextMatrix(d, 4)), Val(MS.TextMatrix(d, 12)))
                        End If
                    End If
                Next
                b = Val(b) + MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                C = MTS(Val(RS!prd_qty), Val(MS.TextMatrix(A, 12)))
            Else
                For d = 1 To MS.Rows - 1
                    If A <> d Then
                        If Trim(MS.TextMatrix(A, 12)) = Trim(MS.TextMatrix(d, 12)) Then
                            b = Val(b) + Val(MS.TextMatrix(d, 4))
                        End If
                    End If
                Next
                b = Val(b) + Val(MS.TextMatrix(A, 4))
                C = MTS(Val(RS!prd_qty), Val(MS.TextMatrix(A, 12)))
            End If
                If Val(b) > Val(C) Then
                    MsgBox "Stock out of Range, Check Stock", vbExclamation, "Effmcg"
                    SSTab1.Tab = 1
                    MS.SetFocus
                    MS.Row = A
                    CHECKSTOCK = True
                End If
            End If
        End If
    End If
Next
End Function

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If CHECKTB("inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RS As New ADODB.Recordset
With MDIForm1
If Trim(M_INV_IDY) <> "" Then
    SAL_QRY = "INL->ORD_IDY = '" & Trim(M_INV_IDY) & "' .AND. INL->TXN_TYP = 'SAL'"
    If ButtonMenu.KEY = "SAL_REP" Then
        Openreport CURDIR & "\REPORT.RP1", "Sales Report (Van)", 2, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "CAS_REC" Then
        updatevan
        Openreport CURDIR & "\REP_LIB.RP1", "Van Reconcile Cases", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "UNT_REC" Then
        updatevan
        Openreport CURDIR & "\REP_LIB.RP1", "Van Reconcile", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "VAN_STL" Then
        Openreport CURDIR & "\DOS_VCH.RP1", "Sale Estimate", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
    End If
Else
    MsgBox "Select Van Challan for Reports Print", vbExclamation
    CANCELRECORD
End If
End With
End Sub
Public Sub updatevan()
If CHECKTB("cld.dbf,dir.dbf,van.dbf,prd.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSCLD As New ADODB.Recordset
Dim RSVAN As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
CON.Execute "DELETE FROM CLD"
CON.Execute "update dir set dir_msg = '" & M_ROU_IDY & "' where dir_idy = 'ROU' and dir_idx = 'IDY'"
CON.Execute "update dir set dir_msg = '" & M_SMN_NME & "' where dir_idy = 'SMN' and dir_idx = 'IDY'"
CON.Execute "update dir set dir_msg = '" & M_CST_NME & "' where dir_idy = 'DLR' and dir_idx = 'IDY'"
CON.Execute "update dir set dir_msg = '" & M_INV_IDY & "' where dir_idy = 'INV' and dir_idx = 'NUM'"
CON.Execute "update dir set dir_msg = '" & Format(M_INV_DAT, "dd/MM/yyyy") & "' where dir_idy = 'HDR' and dir_idx = 'DAT'"
If RSVAN.State = 1 Then RSVAN.Close
RSVAN.Open "SELECT SUM(((tot_qty- int(tot_qty))*1000)+ (int(tot_qty)*sub_unt)) p_tot_qty,SUM(((ret_qty- int(ret_qty))*1000)+ (int(ret_qty)*sub_unt)+((brk_qty- int(brk_qty))*1000)+ (int(brk_qty)*sub_unt)+((lkg_qty- int(lkg_qty))*1000)+ (int(lkg_qty)*sub_unt)) p_ret_qty,*  FROM VAN,PRD WHERE PRD.PRD_IDY LIKE VAN.PRD_IDY AND INV_IDY LIKE '" & Trim(M_INV_IDY) & "' group by VAN.prd_idy", CON, adOpenKeyset, adLockOptimistic
For A = 1 To RSVAN.RecordCount
    CON.Execute "INSERT INTO CLD VALUES('" & Trim(RSVAN!BAT_IDY) & "'," & Val(RSVAN!p_tot_qty) & "," & Val(RSVAN!p_ret_qty) & ",0,0,0,0,0,0)"
    RSVAN.MoveNext
Next
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT SUM(TOT_PCS),BAT_IDY,STK_TYP FROM INV WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "' AND STK_TYP IN('Salable','Free','Damage','TakeBack','Breakage','Leakage') GROUP BY BAT_IDY,STK_TYP", CON, adOpenKeyset, adLockOptimistic
For A = 1 To RSINV.RecordCount
    If Trim(RSINV(2)) = "Salable" Or Trim(RSINV(2)) = "Free" Then
        CON.Execute "UPDATE CLD SET PRT_STK= PRT_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!BAT_IDY) & "'", d
        If d = 0 Then CON.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!BAT_IDY) & "',0,0," & Val(RSINV(0)) & ",0,0,0,0,0)"
    Else
        CON.Execute "UPDATE CLD SET SAL_STK= SAL_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!BAT_IDY) & "'", d
        If d = 0 Then CON.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!BAT_IDY) & "',0,0,0," & Val(RSINV(0)) & ",0,0,0,0)"
    End If
    RSINV.MoveNext
Next
End Sub

Private Sub Trolley_Click()
If Trolley.Value = 1 Then
   Check4.Enabled = False
   TRP_NOS.Enabled = False
   MS.ColWidth(7) = 0
   MS.ColWidth(6) = 1050
Else
    TRP_NOS.Enabled = True
    MS.ColWidth(7) = 1050
    Check4.Enabled = True
End If
If M_INV_IDY.Enabled = True Then
    M_INV_IDY.SetFocus
    SSTab1.Tab = 0
Else
    SSTab1.Tab = 0
    M_INV_DAT.SetFocus
End If

End Sub

Private Sub Trolley_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check or Uncheck Push Cart"
End Sub

Private Sub Trolley_kEYPRESS(K As Integer)
If K = 13 Then
    If M_INV_IDY.Enabled = True Then
        M_INV_IDY.SetFocus
        SSTab1.Tab = 0
    Else
        SSTab1.Tab = 0
        M_INV_DAT.SetFocus
    End If
End If
End Sub

Private Sub MakeBill()
On Error Resume Next
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "Effmcg") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("smn.dbf,act.dbf,prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("smn", "smn_idy", Trim(M_SMN_IDY)) = False Then
    MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
    M_SMN_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If check("act", "act_idy", Trim(M_CST_IDY)) = False Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "Effmcg"
    M_CST_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If Check3.Enabled = False And Check4.Enabled = False Then
    MsgBox "Changes not saved", vbInformation
    INC = Trim(M_INV_IDY)
    If check("PRT", "PRT_NME", "M_FRM_RSI") = True Then
        CON.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
        CON.Execute "update dir where dir_idy like 'REM' set dir_als = 'RSI'"
        If Pieces.Value = 0 Then
            Openreport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Pieces)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        Else
            Openreport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Cases)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        End If
    End If
    CANCELRECORD
    Exit Sub
End If

Dim RSINL As New ADODB.Recordset
Dim REC_PNL As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
Dim REC_IRT As New ADODB.Recordset
If MsgBox("Convert to Bill", vbYesNo + vbQuestion) = vbNo Then
    BMAKEBILL = False
    Exit Sub
End If
'If check("INL", "ORD_IDY", Trim(M_INV_IDY)) = True Then
'    MakeBill = True
'    Exit Function
'End If
'M_SEL_ACT.Tag = "Sales"
'M_SEL_ACT.Show 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where grp_clf like 'Sales' and ACT_TYP LIKE '' ", CON, adOpenKeyset, adLockOptimistic
ListView1.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = ListView1.ListItems.Add(, , Trim(REC_SET("act_idy")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_nme"))
        REC_SET.MoveNext
    Next
Else
End If
Toolbar1.Enabled = False
SSTab1.Enabled = False
Picture1.Visible = True
ListView1.SetFocus
End Sub
Private Sub MAKEBILLCONTD()
If CHECKTB("act.dbf,irt.dbf,inl.dbf,inv.dbf,prd.dbf,brd.dbf,pnl.dbf,pnv.dbf,dir.dbf,mar.dbf,smn.dbf,cat.dbf,loc.dbf,pgr.dbf,van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSINL As New ADODB.Recordset
Dim REC_PNL As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
Dim REC_IRT As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If Trolley.Value <> 1 And Check3.Value <> 1 Then
    Exit Sub
ElseIf Trolley.Value = 1 And Check3.Value = 1 And Check4.Value = 0 Then
        RSINL.Open "select * from inl where inv_idy = ''", CON, adOpenKeyset, adLockOptimistic
        RSINL.AddNew
        RSINL!INV_IDY = "SAL" & T7increment("idy_sal")
        INC = RSINL!INV_IDY
        invid = Trim(INC)
        RSINL("inv_dat") = M_INV_DAT
        RSINL("ord_idy") = M_INV_IDY
        RSINL("SMN_IDY") = M_SMN_IDY
        RSINL("cst_idy") = Trim(M_CST_IDY & "")
        RSINL("act_idy") = ListView1.SelectedItem
        RSINL("txn_typ") = "SAL"
        RSINL("grs_amt") = Val(Total)
        RSINL("net_amt") = Val(Total)
        RSINL("ROU_IDY") = GetValue("ROU_IDY", "ACT WHERE ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'")
        RATEVAL = GetValue("DBT_RAP", "ACT WHERE ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'")
        RSINL!USR_IDY = Trim(USERNAME)
        RSINL!sys_dat = Date
        RSINL!sys_tim = Format(Time, "HH:MM:SS")
        RSINL!wbl_dat = Date
        RSINL!lrn_dat = Date
        RSINL!DCH_IDY = ""
        RSINL!vch_idy = ""
        RSINL!Sch_Idy = ""
        RSINL.Update
        'IRT UPDATION
        If check("IRT", "INV_IDY", Trim(INC)) = False Then
          If Val(Total) > 0 Then
            REC_IRT.Open "select * from IRT", CON, adOpenKeyset, adLockOptimistic
            REC_IRT.AddNew
            REC_IRT("inv_idy") = INC
            REC_IRT("inv_dat") = M_INV_DAT
            REC_IRT("PTY_idy") = Trim(M_CST_IDY & "")
            REC_IRT("ACT_idy") = ""
            REC_IRT("txn_typ") = "SAL"
            REC_IRT("smn_idy") = Trim(M_SMN_IDY & "")
            REC_IRT("net_amt") = Val(Total)
            REC_IRT("due_dat") = Date
            REC_IRT.Update
          End If
        End If
        'END OF IRT.DBF UPDATION
    If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "select * from inv where inv_idy = ''", CON, adOpenKeyset, adLockOptimistic
    For A = 1 To MS.Rows - 1
        If Trim(MS.TextMatrix(A, 0)) <> "" And Val(MS.TextMatrix(A, 5)) > 0 Then
            RSINV.AddNew
            RSINV("inv_idy") = Trim(INC)
            RSINV("inv_dat") = Trim(M_INV_DAT)
            RSINV("prd_idy") = MS.TextMatrix(A, 10)
            RSINV("RAT_APP") = Trim(RATEVAL)
            RSINV("STK_TYP") = "Salable"
            If Pieces.Value = 1 Then
                RSINV("prd_qty") = Val(MS.TextMatrix(A, 6))
                RSINV("prd_rat") = Val(MS.TextMatrix(A, 15))
            Else
                RSINV("PRD_QTY") = STM(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
                RSINV("prd_rat") = convert(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 12)))
            End If
            RSINV("txn_typ") = "SAL"
            RSINV!BAT_IDY = Trim(MS.TextMatrix(A, 13))
            RSINV!sub_unt = Val(MS.TextMatrix(A, 12))
            If Pieces.Value = 1 Then
                RSINV!PRD_UNT = Int(Val(MS.TextMatrix(A, 6)))
                RSINV!PRD_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12))) Mod Val(MS.TextMatrix(A, 12))
                RSINV!TOT_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
            Else
                RSINV!PRD_UNT = Int(Val(MS.TextMatrix(A, 6) / MS.TextMatrix(A, 12)))
                RSINV!PRD_PCS = Val(MS.TextMatrix(A, 6)) Mod Val(MS.TextMatrix(A, 12))
                RSINV!TOT_PCS = Val(MS.TextMatrix(A, 6))
            End If
            RSINV("COL_AMT") = Val((MS.TextMatrix(A, 9)))
            RSINV!stk_upd = "Y"
            RSINV!ORD_IDY = M_INV_IDY
            RSINV.Update
            If Pieces.Value = 1 Then
                CON.Execute "update PRD where prd_idy like '" & Trim(MS.TextMatrix(A, 10)) & "' set CUR_LEV = " & CONQTY(Trim(MS.TextMatrix(A, 6)), Trim(MS.TextMatrix(A, 10))), d
                CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12))), d
                CON.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(Trim(MS.TextMatrix(A, 6)), Trim(MS.TextMatrix(A, 13))), d
            Else
                CON.Execute "update PRD where prd_idy like '" & Trim(MS.TextMatrix(A, 10)) & "' set CUR_LEV = " & CONQTY(QTY(STM(MS.TextMatrix(A, 6), Val(MS.TextMatrix(A, 12)))), Trim(MS.TextMatrix(A, 10)))
                CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS - " & Val(MS.TextMatrix(A, 6))
                CON.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(STM(MS.TextMatrix(A, 6), Val(MS.TextMatrix(A, 12))), Trim(MS.TextMatrix(A, 13)))
            End If
        End If
    Next
    'PRINT DATABASE LOGIC BEGIN
    If MOD_REC = True Then
        PNL.Execute "delete from pnv where inv_idy like '" & Trim(INC) & "'"
        PNL.Execute "delete from pnl where inv_idy like '" & Trim(INC) & "'"
    End If
    REC_PNL.Open "SELECT * FROM PNL", PNL, adOpenKeyset, adLockOptimistic
    REC_PNL.AddNew
    REC_PNL!INV_IDY = INC
    REC_PNL("inv_dat") = M_INV_DAT
    REC_PNL("ord_idy") = Trim(M_INV_IDY & "")
    REC_PNL("cst_idy") = Trim(M_CST_IDY & "")
    REC_PNL("cst_nme") = Trim(M_CST_NME)
    REC_PNL("cst_ad1") = GetValue("ACT_AD1", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL("cst_ad2") = GetValue("ACT_AD2", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL("cst_ad3") = GetValue("ACT_AD3", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL("cst_ad4") = GetValue("ACT_AD4", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL("cst_zip") = GetValue("ACT_ZIP", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL("Act_idy") = Trim(M_ACT_IDY & "")
    REC_PNL("dch_idy") = ""
    REC_PNL("csd_pcg") = 0
    REC_PNL("csd_amt") = 0
    REC_PNL("spd_pcg") = 0
    REC_PNL("spd_amt") = 0
    REC_PNL("tax_pcg") = 0
    REC_PNL("tax_amt") = 0
    REC_PNL("exd_pcg") = 0
    REC_PNL("EXD_AMT") = 0
    REC_PNL("pkg_cst") = 0
    REC_PNL("FRT_AMT") = 0
    REC_PNL("LRN_IDY") = 0
    REC_PNL("LRN_AMT") = 0
    REC_PNL("LRN_DAT") = 0
    REC_PNL("WBL_IDY") = 0
    REC_PNL("WBL_DAT") = 0
    REC_PNL("TXN_RMK") = 0
    REC_PNL("TTL_CAS") = 0
    REC_PNL("txn_typ") = "SAL"
    REC_PNL("grs_amt") = Val(Total)
    REC_PNL("net_amt") = Val(Total)
    REC_PNL("crt_day") = 0
    REC_PNL("due_dat") = Date
    REC_PNL("rou_off") = 0
    REC_PNL("ROU_IDY") = GetValue("ROU_IDY", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL("TXN_TPT") = 0
    REC_PNL("TXN_RM1") = 0
    REC_PNL("TXN_RM2") = 0
    REC_PNL!usr_nme = Trim(USERNAME)
    REC_PNL!sys_dat = Date
    REC_PNL!sys_tim = Format(Time, "HH:MM:SS")
    REC_PNL("SCH_CHK") = "Y"
    REC_PNL!SRN_AMT = 0
    REC_PNL!vch_idy = ""
    REC_PNL!Sch_Idy = ""
    REC_PNL!cst_gst = GetValue("ACT_CST", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL!cst_cst = GetValue("ACT_GST", "ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'")
    REC_PNL!MAR_IDY = GetValue("mar_idy", "act where act_idy = '" & Trim(M_CST_IDY) & "'")
    REC_PNL!mar_nme = GetValue("mar_nme", "mar where mar_idy = '" & Trim(REC_PNL("mar_idy")) & "'")
    REC_PNL!smn_nme = GetValue("smn_nme", "smn where smn_idy = '" & Trim(M_SMN_IDY) & "'")
    REC_PNL!loc_IDY = GetValue("loc_idy", "act where act_idy = '" & Trim(M_CST_IDY) & "'")
    REC_PNL!loc_nme = GetValue("loc_nme", "loc where loc_idy = '" & Trim(REC_PNL("loc_idy")) & "'")
    REC_PNL!cat_IDY = GetValue("cat_idy", "act where act_idy = '" & Trim(M_CST_IDY) & "'")
    REC_PNL!cat_nme = GetValue("cat_nme", "cat where cat_idy = '" & Trim(REC_PNL("cat_idy")) & "'")
    REC_PNL!exd_cal = ""
    REC_PNL!SMN_IDY = M_SMN_IDY
    REC_PNL!frt_typ = ""
    REC_PNL!tax_act = ""
    REC_PNL!frt_act = ""
    REC_PNL!des_mod = ""
    REC_PNL!sto_rmk = ""
    REC_PNL!vch_dat = Date
    REC_PNL!ins_nme = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'TT1'")
    REC_PNL!ins_adr = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'TT2'")
    REC_PNL!sys_rm1 = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'RM1'")
    REC_PNL!sys_rm2 = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'RM2'")
    REC_PNL!sys_rm3 = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'RM3'")
    REC_PNL!sys_rm4 = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'RM4'")
    REC_PNL!ins_gst = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'GST'")
    REC_PNL!ins_cst = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'CST'")
    REC_PNL!FRM_DAT = Date
    REC_PNL!TO_DAT = Date
    REC_PNL!rou_act = ""
    REC_PNL!dch_dat = Date
    REC_PNL!DIS_CHK = ""
    REC_PNL.Update
    REC_PNL.Close
    If REC_PNL.State = 1 Then REC_PNL.Close
    REC_PNL.Open "select * from pnv where inv_idy = ''", PNL, adOpenKeyset, adLockOptimistic
    For A = 1 To MS.Rows - 1
        If Trim(MS.TextMatrix(A, 0)) <> "" And Val(MS.TextMatrix(A, 3)) > 0 And MS.TextMatrix(A, 2) <> "" Then
            REC_PNL.AddNew
            REC_PNL("inv_idy") = Trim(INC)
            REC_PNL("inv_dat") = Trim(M_INV_DAT)
            REC_PNL("prd_idy") = MS.TextMatrix(A, 10)
            REC_PNL("grp_idy") = GetValue("grp_idy", "prd where prd_idy = '" & Trim(MS.TextMatrix(A, 10)) & "'")
            REC_PNL("grp_cat") = GetValue("grp_cat", "pgr where grp_idy = '" & Trim(REC_PNL("grp_idy")) & "'")
            REC_PNL("grp_cls") = GetValue("grp_cls", "pgr where grp_idy = '" & Trim(REC_PNL("grp_idy")) & "'")
            REC_PNL("RAT_APP") = ""
            If Pieces.Value = 1 Then
                REC_PNL("prd_qty") = Val(MS.TextMatrix(A, 6))
                REC_PNL("prd_rat") = Val(MS.TextMatrix(A, 15))
                REC_PNL!MRP_RAT = Val(MS.TextMatrix(A, 8))
            Else
                REC_PNL("PRD_QTY") = STM(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
                REC_PNL("prd_rat") = Val(Val(MS.TextMatrix(A, 15))) * Val(MS.TextMatrix(A, 12))
                REC_PNL!MRP_RAT = Val(MS.TextMatrix(A, 8)) * Val(MS.TextMatrix(A, 12))
            End If
            REC_PNL("ADD_AMT") = 0
            REC_PNL("trd_dis") = 0
            REC_PNL("trd_tax") = 0
            REC_PNL("DIS_AMT") = 0
            REC_PNL("TAX_AMT") = 0
            REC_PNL("txn_typ") = "SAL"
            REC_PNL!BAT_IDY = Trim(MS.TextMatrix(A, 13))
            REC_PNL!PRD_UOM = Val((MS.TextMatrix(A, 4)))
            REC_PNL!sub_unt = Val(MS.TextMatrix(A, 12))
            If Pieces.Value = 1 Then
                REC_PNL!PRD_UNT = Int(Val(MS.TextMatrix(A, 6)))
                REC_PNL!PRD_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12))) Mod Val(MS.TextMatrix(A, 12))
                REC_PNL!TOT_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
            Else
                REC_PNL!PRD_UNT = Int(Val(MS.TextMatrix(A, 6) / MS.TextMatrix(A, 12)))
                REC_PNL!PRD_PCS = Val(MS.TextMatrix(A, 6)) Mod Val(MS.TextMatrix(A, 12))
                REC_PNL!TOT_PCS = Val(MS.TextMatrix(A, 6))
            End If
            REC_PNL("COL_AMT") = Val(MS.TextMatrix(A, 9))
            REC_PNL!stk_upd = "Y"
            REC_PNL("prd_nme") = Trim(MS.TextMatrix(A, 10))
            REC_PNL("lst_rat") = 0
            REC_PNL("lst_dis") = 0
            REC_PNL("lst_rat") = 0
            REC_PNL("ret_qty") = 0
            REC_PNL("iss_qty") = 0
            REC_PNL("tot_qty") = 0
            REC_PNL("rat_app") = ""
            REC_PNL("cpn_rat") = 0
            REC_PNL("cpn_amt") = 0
            REC_PNL("add_dis") = 0
            REC_PNL("fst_flg") = ""
            REC_PNL("opn_qty") = 0
            REC_PNL("sch_amt") = 0
            REC_PNL("ord_idy") = 0
            REC_PNL("bat_nme") = Trim(MS.TextMatrix(A, 14))
            REC_PNL("stk_typ") = "Salable"
            REC_PNL("sch_idy") = ""
            REC_PNL.Update
        End If
    Next
    SAVESHELL Trim(INC)
    'pnv salable update close
    save_rec = "Save"
    MsgBox "New Invoice ID " & INC, vbExclamation
    BMAKEBILL = True
Else
End If
If REC_SET1.State = 1 Then REC_SET1.Close
If REC_SET.State = 1 Then REC_SET.Close
If MOD_REC = True Then
    If Check3.Value = 1 And Check3.Enabled = False Then
    Else
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from van,PRD where VAN.PRD_IDY LIKE PRD.PRD_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset
    For A = 1 To REC_SET.RecordCount
    CON.Execute "update prd where prd_idy like '" & REC_SET("prd_idy") & "' set cur_lev = " & ADDQTY(Val(REC_SET("TOT_qty")), Trim(REC_SET("prd_idy")))
    CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & REC_SET("prd_idy") & "' set PRD_PCS = PRD_PCS + " & MTS(Val(REC_SET("TOT_qty")), Val(REC_SET("SUB_UNT")))
    CON.Execute "update Brd where BAT_idy like '" & REC_SET("BAT_idy") & "' set PRD_QTY = " & ADDBRDQTY(Val(REC_SET("TOT_qty")), Trim(REC_SET("BAT_idy")))
    REC_SET.MoveNext
    Next
    End If
    CON.Execute "delete from van where inv_idy like '" & Trim(M_INV_IDY) & "'", d
    INC = Trim(M_INV_IDY)
Else
    If Trim(MS.TextMatrix(1, 0)) <> "" Then
        INC = "RSI" & T7increment("IDY_RSI")
        INVIDY = INC
        If Trim(M_INV_IDY) = "" And Trolley.Value = 1 Then
            CON.Execute "UPDATE INL SET ORD_IDY = '" & Trim(INC) & "' WHERE INV_IDY = '" & Trim(invid) & "'"
        End If
    Else
         CANCELRECORD
        Exit Sub
    End If
End If
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from van where inv_idy = ''", CON, adOpenKeyset, adLockOptimistic
If RSCUR.State = 1 Then RSCUR.Close
RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY=''", CON, adOpenKeyset, adLockOptimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
        REC_SET.AddNew
        REC_SET("inv_idy") = Trim(INC)
        REC_SET("inv_dat") = Trim(M_INV_DAT)
        If MS.TextMatrix(A, 1) = "All" Then
            REC_SET("TRP_IDY") = 0
        Else
            REC_SET("TRP_IDY") = Val(MS.TextMatrix(A, 1))
        End If
        REC_SET("prd_idy") = Trim(MS.TextMatrix(A, 10))
        REC_SET("ACT_idy") = Trim(M_ACT_IDY)
        REC_SET("CST_idy") = Trim(M_CST_IDY)
        REC_SET("SMN_idy") = Trim(M_SMN_IDY)
        If Pieces.Value = 0 Then
            MS.TextMatrix(A, 1) = STM(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 2) = STM(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 3) = STM(Val(MS.TextMatrix(A, 3)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 5) = STM(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 6) = STM(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 4) = STM(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 15) = Val(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 16) = STM(Val(MS.TextMatrix(A, 16)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 17) = STM(Val(MS.TextMatrix(A, 17)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 18) = STM(Val(MS.TextMatrix(A, 18)), Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 8) = convert(Val(MS.TextMatrix(A, 8)) * Val(MS.TextMatrix(A, 12)))
        End If
        REC_SET("OPN_QTY") = Val(MS.TextMatrix(A, 2))
        REC_SET("ISS_QTY") = Val(MS.TextMatrix(A, 3))
        REC_SET("tot_qty") = Val(MS.TextMatrix(A, 4))
        REC_SET("fre_qty") = Val(MS.TextMatrix(A, 16))
        REC_SET("BRK_qty") = Val(MS.TextMatrix(A, 17))
        REC_SET("LKG_qty") = Val(MS.TextMatrix(A, 18))
        REC_SET("BIL_QTY") = Val(MS.TextMatrix(A, 6))
        REC_SET("RET_QTY") = Val(MS.TextMatrix(A, 5))
        REC_SET("BAT_idy") = Trim(MS.TextMatrix(A, 13))
        REC_SET!PRD_rat = Val(MS.TextMatrix(A, 15))
        If Check3.Value = 1 Then
            REC_SET!van_cls = "S"
        Else
            REC_SET!van_cls = "N"
        End If
        If Check4.Value = 1 Or save_rec = "Save" Then
        save_rec = ""
        REC_SET!van_cls = "C"
        End If
        REC_SET!COL_AMT = Val(MS.TextMatrix(A, 9))
        REC_SET!CAS_AMT = Val(M_CAS_AMT)
        REC_SET!CHQ_AMT = Val(M_CHQ_AMT)
        REC_SET!DIS_AMT = Val(M_DIS_AMT)
        REC_SET!EXP_AMT = Val(M_EXP_AMT)
        REC_SET!ADD_AMT = Val(M_TAX_AMT)
        If Trolley.Value = 1 Then
            REC_SET!VAN_CHK = "P"
        Else
            REC_SET!VAN_CHK = "V"
        End If
        REC_SET.Update
        If Check3.Value = 1 And Check3.Enabled = False Then
        ElseIf Check3.Value = 0 And Check3.Enabled = True Then
            CON.Execute "update prd where prd_idy like '" & Trim(MS.TextMatrix(A, 10)) & "' set cur_lev = " & CONQTY(QTY(Trim(MS.TextMatrix(A, 4) & "")), Trim(MS.TextMatrix(A, 10)))
            CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
            CON.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(QTY(Val(MS.TextMatrix(A, 4))), Trim(MS.TextMatrix(A, 13)))
        End If
    End If
Next
RSCUR.AddNew    ' EDITED BY RADHIKA
RSCUR(0) = Trim(INC)
RSCUR(1) = Val(Trim(txt1000))
RSCUR(2) = Val(Trim(txt500))
RSCUR(3) = Val(Trim(txt100))
RSCUR(4) = Val(Trim(txt50))
RSCUR(5) = Val(Trim(txt20))
RSCUR(6) = Val(Trim(txt10))
RSCUR(7) = Val(Trim(txt5))
RSCUR(8) = Val(Trim(txt2))
RSCUR(9) = Val(Trim(txt1))
RSCUR(10) = Val(Trim(txtchange))
RSCUR.Update    'END
If REC_SET1.State = 1 Then REC_SET1.Close
If MOD_REC = False Then MsgBox "New Challan ID is " & INC, vbExclamation, "Effmcg"
Label23 = labdisp("idy_rsi")
VAN_PRT_CHK:
If CHECKTB("prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_RSI") = True Then
    CON.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
    CON.Execute "update dir where dir_idy like 'REM' set dir_als = 'RSI'"
    If Trolley.Value = 1 Then
        If Pieces.Value = 0 Then
            Openreport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Pieces)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        Else
            Openreport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Cases)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        End If
    Else
        Openreport CURDIR & "\DOS_VCH.RP1", "Van Challan", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
    End If
    CANCELRECORD
End If
CANCELRECORD
End Sub
Public Sub RETURNSTOCKS()
Dim RSINV As New ADODB.Recordset
If MOD_REC = True Then
    If CHECKTB("inv.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    RSINV.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "' AND STK_TYP IN('Salable','Free') GROUP BY BAT_IDY", CON, adOpenKeyset, adLockOptimistic
    If RSINV.RecordCount > 0 Then
        For A = 1 To RSINV.RecordCount
            For b = 1 To MS.Rows - 1
                If MS.TextMatrix(b, 13) = Trim(RSINV(1)) Then
                    If Pieces.Value = 0 Then
                        MS.TextMatrix(b, 6) = Val(RSINV(0))
                    Else
                        MS.TextMatrix(b, 6) = QTY(STM(Val(RSINV(0)), Val(RSINV(2))))
                    End If
                End If
            Next
            RSINV.MoveNext
        Next
        RET_STK = True
        MS.SetFocus
        discalc
    End If
    If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "' AND STK_TYP IN('Damage','TakeBack','Breakage','Leakage') GROUP BY BAT_IDY", CON, adOpenKeyset, adLockOptimistic
    If RSINV.RecordCount > 0 Then
        For A = 1 To RSINV.RecordCount
            For b = 1 To MS.Rows - 1
                If MS.TextMatrix(b, 13) = Trim(RSINV(1)) Then
                    If Pieces.Value = 0 Then
                        MS.TextMatrix(b, 6) = Val(MS.TextMatrix(b, 6)) - Val(RSINV(0))
                    Else
                        MS.TextMatrix(b, 6) = QTY(STM(MTS(Val(MS.TextMatrix(b, 6)), Val(RSINV(2))) - Val(RSINV(0)), Val(RSINV(2))))
                    End If
                End If
            Next
            RSINV.MoveNext
        Next
        RET_STK = True
        MS.SetFocus
        discalc
    End If
End If
End Sub

Private Sub Trolley_Validate(cancel As Boolean)
If Trolley.Value = 1 Then
    Check4.Value = 0
    Check4.Enabled = False
ElseIf Trolley.Value = 0 Then
    Check4.Value = 0
    Check4.Enabled = True
End If
End Sub

Public Sub P_BIL_QTY()
If Trolley.Value = 1 And Check3.Value = 1 Then
    For A = 1 To MS.Rows - 1
        If MS.TextMatrix(A, 0) <> "" Then
            If Pieces.Value = 1 Then
                MS3 = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
                MS4 = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 12)))
                MS5 = Val(MS3) - Val(MS4)
                MS.TextMatrix(A, 6) = QTY(STM(Val(MS5), Val(MS.TextMatrix(A, 12))))
            Else
                MS.TextMatrix(A, 6) = Val(MS.TextMatrix(A, 4)) - Val(MS.TextMatrix(A, 5))
            End If
        End If
    Next
End If
End Sub

Public Sub BILCALC()
For A = 1 To MS.Rows - 1
    'Customer Discount calculating on units and rate
    If MS.TextMatrix(A, 0) <> "" Then
        If Pieces.Value = 0 Then
            MS.TextMatrix(A, 9) = convert(Val(MS.TextMatrix(A, 6)) * Val(MS.TextMatrix(A, 8)))
        Else
            MS.TextMatrix(A, 9) = convert(RATCAL(QTY(Val(MS.TextMatrix(A, 6) & "")), Val(MS.TextMatrix(A, 8) & ""), Val(MS.TextMatrix(A, 12))))
        End If
        Total = convert(AMT(Val(Total) + Val(MS.TextMatrix(A, 9))))
    End If
Next
M_GRS_AMT = convert(AMT(Total))
Total = convert(AMT(Total))
End Sub

Private Sub TRP_NOS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Dim RS As New ADODB.Recordset
    If Check3.Value = 1 And TRP_NOS = "All" And MOD_REC = True Then
        If CHECKTB("van.dbf,brd.dbf,prd.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
        If RSCUR.State = 1 Then RSCUR.Close
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
        RS.Open "SELECT VAN.*,BRD.*,SUM(INT(FRE_QTY) * SUB_UNT)+(FRE_QTY-INT(FRE_QTY))*1000 P_FRE_QTY,SUM(INT(TOT_QTY) * SUB_UNT)+(OPN_QTY-INT(OPN_QTY))*1000 P_TOT_QTY,SUM(INT(OPN_QTY) * SUB_UNT)+(OPN_QTY-INT(OPN_QTY))*1000 P_OPN_QTY,SUM(INT(ISS_QTY) * SUB_UNT)+(ISS_QTY-INT(ISS_QTY))*1000 P_ISS_QTY,SUM(INT(RET_QTY) * SUB_UNT)+(RET_QTY-INT(RET_QTY))*1000 P_RET_QTY,SUM(INT(BIL_QTY) * SUB_UNT)+(BIL_QTY-INT(BIL_QTY))*1000 P_BIL_QTY,SUB_UNT,PRD_NME,PRD_PRN FROM VAN,BRD,PRD WHERE VAN.BAT_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND INV_IDY LIKE '" & Trim(M_INV_IDY) & "' GROUP BY BRD.BAT_IDY", CON, adOpenKeyset, adLockOptimistic
        MS.Rows = 1
        For A = 1 To RS.RecordCount
            MS.AddItem ""
            MS.TextMatrix(A, 0) = Trim(RS!PRD_NME)
            MS.TextMatrix(A, 1) = "All"
            If Pieces.Value = 0 Then
                MS.TextMatrix(A, 2) = Val(RS!P_OPN_QTY & "")
                MS.TextMatrix(A, 3) = Val(RS!P_ISS_QTY & "")
                MS.TextMatrix(A, 4) = Val(RS!p_tot_qty & "")
                MS.TextMatrix(A, 6) = Val(RS!P_BIL_QTY & "")
                MS.TextMatrix(A, 5) = Val(RS!p_ret_qty & "")
                MS.TextMatrix(A, 16) = Val(RS!P_FRE_QTY & "")
                MS.TextMatrix(A, 8) = convert(Val(RS!PRD_MRP) / Val(RS!sub_unt))
                MS.TextMatrix(A, 15) = convert(Val(RS!PRD_rat) / Val(RS!sub_unt))
            Else
                MS.TextMatrix(A, 2) = QTY(STM(Val(RS!P_OPN_QTY), Val(RS!sub_unt)))
                MS.TextMatrix(A, 3) = QTY(STM(Val(RS!P_ISS_QTY), Val(RS!sub_unt)))
                MS.TextMatrix(A, 4) = QTY(STM(Val(RS!p_tot_qty), Val(RS!sub_unt)))
                MS.TextMatrix(A, 6) = QTY(STM(Val(RS!P_BIL_QTY), Val(RS!sub_unt)))
                MS.TextMatrix(A, 5) = QTY(STM(Val(RS!p_ret_qty), Val(RS!sub_unt)))
                MS.TextMatrix(A, 16) = QTY(STM(Val(RS!P_FRE_QTY), Val(RS!sub_unt)))
                MS.TextMatrix(A, 8) = convert(RS!PRD_MRP & "")
                MS.TextMatrix(A, 15) = convert(RS!PRD_rat)
            End If
            M_CAS_AMT = convert(RS!CAS_AMT & "")
            M_CHQ_AMT = convert(RS!CHQ_AMT & "")
            M_DIS_AMT = convert(RS!DIS_AMT & "")
            M_EXP_AMT = convert(RS!EXP_AMT & "")
            M_TAX_AMT = convert(RS!ADD_AMT & "")
            MS.TextMatrix(A, 9) = convert(RS!COL_AMT & "")
            MS.TextMatrix(A, 10) = Trim(RS!PRD_IDY)
            MS.TextMatrix(A, 11) = Trim(RS!prd_prn)
            MS.TextMatrix(A, 12) = Val(RS!sub_unt)
            MS.TextMatrix(A, 13) = Trim(RS!BAT_IDY)
            MS.TextMatrix(A, 14) = Trim(RS!Bat_nme)
            RS.MoveNext
        Next
        If RSCUR.RecordCount > 0 Then
            txt1000 = convert(Val(RSCUR(1) & ""))
            txt500 = convert(Val(RSCUR(2) & ""))
            txt100 = convert(Val(RSCUR(3) & ""))
            txt50 = convert(Val(RSCUR(4) & ""))
            txt20 = convert(Val(RSCUR(5) & ""))
            txt10 = convert(Val(RSCUR(6) & ""))
            txt5 = convert(Val(RSCUR(7) & ""))
            txt2 = convert(Val(RSCUR(8) & ""))
            txt1 = convert(Val(RSCUR(9) & ""))
            txtchange = convert(Val(RSCUR(10) & ""))
        End If
        TRP_NOS.Enabled = False
                P_BIL_QTY
        discalc
        M_SMN_IDY.SetFocus
    Else
        If MOD_REC = False Then
            If TRP_NOS = "All" Then
                MsgBox "Enter Trip number", vbExclamation
                SendKeys "{home}+{end}"
                TRP_NOS.SetFocus
            Else
                If Len(Trim(TRP_NOS)) = 0 Or Val(TRP_NOS) = 0 Then
                    MsgBox "Enter Trip number", vbExclamation
                    TRP_NOS.SetFocus
                        SendKeys "{home}+{end}"
                Else
                    TRP_NOS.Enabled = False
                    M_SMN_IDY.SetFocus
                End If
            End If
        ElseIf MOD_REC = True Then
            If CHECKTB("van.dbf") = 1 Then
                MsgBox UCase(TNAME) & " Missing in " & Me.Caption
                CANCELRECORD
                Exit Sub
            End If
            If RS.State = 1 Then RS.Close
            RS.Open "select * from van where inv_idy like '" & Trim(M_INV_IDY) & "' and trp_idy = " & Val(TRP_NOS), CON, adOpenKeyset, adLockOptimistic
            If RS.RecordCount > 0 Then
                TRP_NOS.Enabled = False
                M_SMN_IDY.SetFocus
            Else
                If TRP_NOS = "All" Or Val(TRP_NOS) = 0 Then
                    MsgBox "Enter Trip Number", vbExclamation
                    TRP_NOS.SetFocus
                    SendKeys "{home}+{end}"
                ElseIf Check3.Value = 0 Then
                    If MsgBox("Create New Trip ?", vbQuestion + vbYesNo) = vbYes Then
                        TRP_NOS.Enabled = False
                        M_SMN_IDY.SetFocus
                    Else
                        SendKeys "{home}+{end}"
                        TRP_NOS.SetFocus
                    End If
                End If
            End If
            RS.Close
        End If
    End If
End If
End Sub

Private Sub TRP_NOS_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub txt1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt1 = convert(txt1)
    ADDTOTCUR
    txtchange.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt2.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt10 = convert(txt10)
    ADDTOTCUR
    txt5.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt20.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt100_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt100 = convert(txt100)
    ADDTOTCUR
    txt50.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt500.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt1000_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt1000 = convert(txt1000)
    ADDTOTCUR
    txt500.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    M_CRT_AMT.SetFocus
End If
End Sub

Private Sub txt2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt2 = convert(txt2)
    ADDTOTCUR
    txt1.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt5.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt20_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt20 = convert(txt20)
    ADDTOTCUR
    txt10.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt50.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt5 = convert(txt5)
    ADDTOTCUR
    txt2.SetFocus
    SendKeys "{home}+{end}"
Else
    txt10.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt50_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt50 = convert(txt50)
    ADDTOTCUR
    txt20.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt100.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt500_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt500 = convert(txt500)
    ADDTOTCUR
    txt100.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt1000.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtchange_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtchange = convert(txtchange)
    ADDTOTCUR
    If Val(M_CRT_AMT) = Val(TXTTOTAL) Then
        If Check3.Value = 1 And Trolley.Value = 1 Then
            MakeBill
        Else
            saverecord
        End If
    Else
        MsgBox "Invalid Amount", vbExclamation, "Effmcg"
        txtchange.SetFocus
        SendKeys "{home}+{end}"
    End If
ElseIf KeyCode = 38 And Shift = 1 Then
    txt1.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtempshl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
With ms2
    .Text = Int(Val(txtempshl))
    txtempshl.Visible = False
    .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 1)) - Val(.TextMatrix(.Row, 3))
    .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 2)) - Val(.TextMatrix(.Row, 4))
    If .Col = 3 Then
        .Col = 4
    ElseIf .Col = 4 Then
        If .Rows - 1 <> .Row Then .Row = .Row + 1
        .Col = 3
    End If
    .SetFocus
End With
End If
End Sub

Private Sub txtempshl_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
Else
K = 0
End If
End Sub

Public Sub SAVESHELL(IDY As String)
Dim RS As New ADODB.Recordset
CON.Execute "DELETE FROM SNV WHERE INV_IDY LIKE '" & Trim(IDY) & "' AND TXN_TYP='VAN'"
RS.Open "SELECT * FROM SNV WHERE INV_IDY LIKE ''", CON, adOpenKeyset, adLockOptimistic
For A = 2 To ms2.Rows - 1
    RS.AddNew
    RS!INV_IDY = Trim(IDY)
    RS!INV_DAT = M_INV_DAT
    RS!shl_idy = ms2.TextMatrix(A, 8)
    RS!PRD_IDY = ms2.TextMatrix(A, 7)
    RS!i_shl_nos = Val(ms2.TextMatrix(A, 1))
    RS!i_pck_nos = Val(ms2.TextMatrix(A, 2))
    RS!R_shl_nos = Val(ms2.TextMatrix(A, 3))
    RS!R_pck_nos = Val(ms2.TextMatrix(A, 4))
    RS!d_shl_nos = Val(ms2.TextMatrix(A, 5))
    RS!d_pck_nos = Val(ms2.TextMatrix(A, 6))
    RS!cst_idy = M_CST_IDY
    RS!TXN_TYP = "VAN"
    RS!txn_dat = M_INV_DAT
    RS!stk_upd = "Y"
    RS.Update
Next
End Sub

Public Sub COMMISSION()
Dim TOTCOM As Single, TOTQ As Double, TOTVAL As Single, TTAMT As Integer
If RSCOM.State = 1 Then RSCOM.Close
'RSCOM.Open "select * from sic WHERE act_IDY='" & Trim(M_SMN_IDY) & "' AND com_inc ='C' and FRM_DAT<=CTOD('" & Format(M_INV_DAT.Value, "dd/mm/yy") & "') AND end_DAT>=CTOD('" & Format(M_INV_DAT.Value, "dd/mm/yy") & "')", CON, adOpenKeyset, adLockOptimistic
RSCOM.Open "select * from sic WHERE act_IDY='" & Trim(M_SMN_IDY) & "' AND com_inc ='C' and FRM_DAT<=CTOD('" & M_INV_DAT & "') AND end_DAT>=CTOD('" & M_INV_DAT & "')", CON, adOpenKeyset, adLockOptimistic
If RSCOM.EOF Then
    M_COM_AMT = "0.00"
Else
    TOTCOM = 0
    While Not RSCOM.EOF
        For R = 1 To MS.Rows - 1
            If Trim(RSCOM("prd_idy")) = Trim(MS.TextMatrix(R, 10)) Then
                If Pieces.Value = 0 Then
                    TOTQ = Val(MS.TextMatrix(R, 4)) - (Val(MS.TextMatrix(R, 5)) + Val(MS.TextMatrix(R, 18)) + Val(MS.TextMatrix(R, 17)))
                    TOTQ = STM(Val(TOTQ), Val(MS.TextMatrix(R, 12)))
                Else
                    TOTQ = Val(MTS(Val(MS.TextMatrix(R, 4)), Val(MS.TextMatrix(R, 12)))) - (Val(MTS(Val(MS.TextMatrix(R, 5)), Val(MS.TextMatrix(R, 12)))) + Val(MTS(Val(MS.TextMatrix(R, 18)), Val(MS.TextMatrix(R, 12)))) + Val(MTS(Val(MS.TextMatrix(R, 17)), Val(MS.TextMatrix(R, 12)))))
                    TOTQ = STM(Val(TOTQ), Val(MS.TextMatrix(R, 12)))
                End If
                If MTS(Val(TOTQ), MS.TextMatrix(R, 12)) >= MTS(RSCOM("min_qty"), MS.TextMatrix(R, 12)) Then
                    TOTCOM = TOTCOM + Val(TOTQ) * RSCOM("min_val")
                End If
                Exit For
            End If
        Next
        RSCOM.MoveNext
    Wend
    M_COM_AMT = convert(AMT(Trim(TOTCOM)))
End If




'If RSCOM.State = 1 Then RSCOM.Close
'If RSPRD.State = 1 Then RSPRD.Close
'RSCOM.Open "SELECT SMN_IDY,CHK_FLG,COM_AMT,PRD_IDY,BEG_RNG,END_RNG,COM_PER FROM COM WHERE SMN_IDY='" & Trim(M_SMN_IDY) & "' AND FRM_DAT<=CTOD('" & M_INV_DAT.Value & "') AND TO_DAT>=CTOD('" & M_INV_DAT.Value & "')", CON, adOpenKeyset, adLockOptimistic
'If RSCOM.EOF Then
'    M_COM_AMT = "0.00"
'Else
'    TOTCOM = 0
'    If RSCOM("CHK_FLG") = "A" Then
'        For R = 1 To ms.Rows - 1
'            If Trim(ms.TextMatrix(R, 12)) <> "" Then
'                TOTQ = Val(ms.TextMatrix(R, 4)) - (Val(ms.TextMatrix(R, 5)) + Val(ms.TextMatrix(R, 18)) + Val(ms.TextMatrix(R, 17)))
'                TOTCOM = TOTCOM + (MTS(TOTQ, ms.TextMatrix(R, 12)) * Val(RSCOM("COM_AMT")) / Val(ms.TextMatrix(R, 12)))
'            End If
'        Next
'        M_COM_AMT = convert(STR(TOTCOM))
'    ElseIf RSCOM("CHK_FLG") = "S" Then
'        While Not RSCOM.EOF
'            For R = 1 To ms.Rows - 1
'                If Trim(ms.TextMatrix(R, 12)) <> "" Then
'                    If Trim(ms.TextMatrix(R, 10)) = Trim(RSCOM("PRD_IDY")) Then
'                        TOTQ = Val(ms.TextMatrix(R, 4)) - (Val(ms.TextMatrix(R, 5)) + Val(ms.TextMatrix(R, 18)) + Val(ms.TextMatrix(R, 17)))
'                        TOTCOM = TOTCOM + (MTS(TOTQ, ms.TextMatrix(R, 12)) * Val(RSCOM("COM_AMT")) / Val(ms.TextMatrix(R, 12)))
'                        Exit For
'                    End If
'                End If
'            Next
'            RSCOM.MoveNext
'        Wend
'        M_COM_AMT = convert(STR(TOTCOM))
'    ElseIf RSCOM("CHK_FLG") = "P" Then
'        For R = 1 To ms.Rows - 1
'            If RSPRD.State = 1 Then RSPRD.Close
'            RSPRD.Open "SELECT SAL_RAT FROM PRD WHERE PRD_IDY='" & Trim(ms.TextMatrix(R, 10)) & "'", CON
'            TOTQ = Val(ms.TextMatrix(R, 4)) - (Val(ms.TextMatrix(R, 5)) + Val(ms.TextMatrix(R, 18)) + Val(ms.TextMatrix(R, 17)))
'            TOTVAL = TOTVAL + CASEAMT(STR(TOTQ), ms.TextMatrix(R, 12), RSPRD("SAL_RAT"))
'        Next
'        M_COM_AMT = convert(STR(TOTVAL)) * RSCOM("COM_AMT")
'    ElseIf RSCOM("CHK_FLG") = "R" Then
'        For R = 1 To ms.Rows - 1
'            If RSPRD.State = 1 Then RSPRD.Close
''            RSPRD.Open "SELECT BRD.PRD_SDR FROM PRD,BRD WHERE PRD.PRD_IDY='" & Trim(MS.TextMatrix(R, 10)) & "' AND PRD.PRD_IDY LIKE BRD.BAT_IDY GROUP BY PRD_IDY", CON
'            RSPRD.Open "SELECT BRD.PRD_SDR FROM PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND PRD.PRD_IDY='" & Trim(ms.TextMatrix(R, 10)) & "'", CON
'            TOTQ = Val(ms.TextMatrix(R, 4)) - (Val(ms.TextMatrix(R, 5)) + Val(ms.TextMatrix(R, 18)) + Val(ms.TextMatrix(R, 17)))
'            TOTVAL = TOTVAL + CASEAMT(STR(TOTQ), ms.TextMatrix(R, 12), RSPRD("PRD_SDR"))
'        Next
'        K = 0
'        While Not RSCOM.EOF
'            AR(K, 0) = RSCOM("BEG_RNG")
'            AR(K, 1) = RSCOM("END_RNG")
'            AR(K, 2) = RSCOM("COM_PER")
'            RSCOM.MoveNext
'            K = K + 1
'        Wend
'        If TOTVAL >= AR(0, 0) And TOTVAL <= AR(0, 1) Then
'            TOTCOM = TOTVAL * (Val(AR(0, 2)) / 100)
'        ElseIf TOTVAL >= AR(1, 0) And TOTVAL <= AR(1, 1) Then
'            TOTCOM = AR(0, 1) * (Val(AR(0, 2)) / 100) + (TOTVAL - AR(0, 1)) * (Val(AR(1, 2)) / 100)
'        ElseIf TOTVAL >= AR(2, 0) And TOTVAL <= AR(2, 1) Then
'            TOTCOM = (AR(0, 1) * (Val(AR(0, 2)) / 100)) + ((AR(1, 1) - AR(1, 0)) * (Val(AR(1, 2)) / 100)) + ((TOTVAL - AR(1, 1)) * (Val(AR(2, 2)) / 100))
'        ElseIf TOTVAL >= AR(3, 0) And TOTVAL <= AR(3, 1) Then
'            TOTCOM = (AR(0, 1) * (Val(AR(0, 2)) / 100)) + ((AR(1, 1) - AR(1, 0)) * (Val(AR(1, 2)) / 100)) + ((AR(2, 1) - AR(2, 0)) * (Val(AR(2, 2)) / 100)) + ((TOTVAL - AR(2, 1)) * (Val(AR(3, 2)) / 100))
'        ElseIf TOTVAL >= AR(4, 0) And TOTVAL <= AR(4, 1) Then
'            TOTCOM = (AR(0, 1) * (Val(AR(0, 2)) / 100)) + ((AR(1, 1) - AR(1, 0)) * (Val(AR(1, 2)) / 100)) + ((AR(2, 1) - AR(2, 0)) * (Val(AR(2, 2)) / 100)) + ((AR(3, 1) - AR(3, 0)) * (Val(AR(3, 2)) / 100)) + ((TOTVAL - AR(3, 1)) * (Val(AR(4, 2)) / 100))
'        ElseIf TOTVAL > AR(4, 1) Then
'            TOTCOM = (AR(0, 1) * (Val(AR(0, 2)) / 100)) + ((AR(1, 1) - AR(1, 0)) * (Val(AR(1, 2)) / 100)) + ((AR(2, 1) - AR(2, 0)) * (Val(AR(2, 2)) / 100)) + ((AR(3, 1) - AR(3, 0)) * (Val(AR(3, 2)) / 100)) + ((AR(4, 1) - AR(4, 0)) * (Val(AR(4, 2)) / 100)) + ((TOTVAL - AR(4, 1)) * (4 / 100))
'        End If
'        M_COM_AMT = convert(STR(TOTCOM))
'    End If
'End If
End Sub
Private Sub ADDTOTCUR()
TXTTOTAL = convert(Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txt2) * 2 + Val(txt1) * 1 + Val(txtchange))
End Sub
