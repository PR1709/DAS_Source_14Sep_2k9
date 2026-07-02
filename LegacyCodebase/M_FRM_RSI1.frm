VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_FRM_RSI1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van Stock Issue"
   ClientHeight    =   5610
   ClientLeft      =   1425
   ClientTop       =   1845
   ClientWidth     =   8820
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00404040&
   Icon            =   "M_FRM_RSI1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   22
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
         NumButtons      =   21
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
               NumButtonMenus  =   18
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SAL_REP"
                  Object.Tag             =   "SAL_REP"
                  Text            =   "&a Sales Report - Value"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SAL_REP_SKU"
                  Object.Tag             =   "SAL_REP_SKU"
                  Text            =   "&b Sales Report - SKU"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CAS_REC"
                  Text            =   "&c Cases Reconcile"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UNT_REC"
                  Text            =   "&d Units Reconcile"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_STL"
                  Text            =   "&e Sale Estimate"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "stkStmt"
                  Text            =   "&f Net Load Stock Movement"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_VRE"
                  Text            =   "&g Van Empties Quick Settlement"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_FRE"
                  Text            =   "&h Van Empties Final Settlement"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_CCS"
                  Text            =   "&i Currency Check Sheet"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_SCR"
                  Text            =   "&j Salesman Commission Report"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SHL_SOH_NLD"
                  Text            =   "&k Empties SOH Net Load"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_CNQ"
                  Text            =   "&l Van Receipts Cash n Chq"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_MTF"
                  Text            =   "&m Gate Pass"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SCB_CUM"
                  Text            =   "&n Salesman Cmltve Balance"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SCB_SMY"
                  Text            =   "&o Salesman Cmltve Summary"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_CHL_ALT"
                  Text            =   "&p Van Challan"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SHT_CSH"
                  Text            =   "&q ShortCash Report"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EXP_RPT"
                  Text            =   "&r Expense Report"
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
            Enabled         =   0   'False
            Key             =   "Bill"
            Object.ToolTipText     =   "Quick Load Interface (Ctrl+A)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RS"
            Object.ToolTipText     =   "Generate Bill Stk for this Challan"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CAS_CRT"
            Object.ToolTipText     =   "Cascade Challan ID (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
      EndProperty
      Begin VB.CheckBox Pieces 
         Caption         =   "Check to do in Cases"
         Height          =   195
         Left            =   4800
         TabIndex        =   30
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
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&Voucher Details"
      TabPicture(0)   =   "M_FRM_RSI1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "M_SMN_NME"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label40"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "M_GRS_AMT"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "M_CST_NME"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label16"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label100"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LBLROUTE"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "M_INV_DAT"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "M_INV_IDY"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "M_SMN_IDY"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "M_CST_IDY"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TRP_NOS"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmbRoute"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "MNL_REF"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Picture2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "&Products"
      TabPicture(1)   =   "M_FRM_RSI1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Total"
      Tab(1).Control(1)=   "M_Prd_Nme"
      Tab(1).Control(2)=   "MS"
      Tab(1).Control(3)=   "amount"
      Tab(1).Control(4)=   "Text3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "List6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Picture3"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Collection &Details"
      TabPicture(2)   =   "M_FRM_RSI1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "LBLMSG"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "View"
      TabPicture(3)   =   "M_FRM_RSI1.frx":0496
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
      TabPicture(4)   =   "M_FRM_RSI1.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TTL_CSH_VAR"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command9"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command8"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command7"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Command5"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Command3"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Command2"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "MSS"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Labelttlvariance"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "Glass &Entries"
      TabPicture(5)   =   "M_FRM_RSI1.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ms2"
      Tab(5).Control(1)=   "txtempshl"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "SHL_SET"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Verif&Y Glass Variance"
      TabPicture(6)   =   "M_FRM_RSI1.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ms1"
      Tab(6).Control(1)=   "Command6"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.TextBox TTL_CSH_VAR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -68800
         TabIndex        =   135
         Text            =   " 0.00"
         Top             =   4820
         Width           =   1210
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&4 Sch Log"
         Height          =   375
         Left            =   -67440
         TabIndex        =   133
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&3 Estimate"
         Height          =   375
         Left            =   -67440
         TabIndex        =   117
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&2 Reconcile"
         Height          =   375
         Left            =   -67440
         TabIndex        =   116
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&1 Gate Pass"
         Height          =   375
         Left            =   -67440
         TabIndex        =   115
         Top             =   1440
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         Height          =   2175
         Left            =   2640
         ScaleHeight     =   2115
         ScaleWidth      =   4035
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   2280
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1365
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&OK"
            Height          =   375
            Left            =   720
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   1365
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   2040
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16646147
            CurrentDate     =   37315
         End
         Begin VB.Label Label9 
            BackColor       =   &H00800000&
            Caption         =   "  Select Date to Bring Forward Opening Stock"
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
            TabIndex        =   113
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Challan Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   480
            TabIndex        =   112
            Top             =   668
            Width           =   1260
         End
      End
      Begin VB.TextBox MNL_REF 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   1700
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   -74950
         ScaleHeight     =   195
         ScaleWidth      =   8640
         TabIndex        =   89
         Top             =   4800
         Width           =   8700
         Begin VB.Label Tot_fre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4440
            TabIndex        =   101
            Top             =   0
            Width           =   45
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Free"
            Height          =   195
            Left            =   3480
            TabIndex        =   100
            Top             =   0
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Var"
            Height          =   195
            Left            =   6840
            TabIndex        =   99
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Net"
            Height          =   195
            Left            =   5160
            TabIndex        =   98
            Top             =   0
            Width           =   660
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total LI"
            Height          =   195
            Left            =   1680
            TabIndex        =   97
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total LO"
            Height          =   195
            Left            =   15
            TabIndex        =   96
            Top             =   0
            Width           =   615
         End
         Begin VB.Label tot_var 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   7680
            TabIndex        =   93
            Top             =   0
            Width           =   45
         End
         Begin VB.Label tot_net 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   6000
            TabIndex        =   92
            Top             =   0
            Width           =   45
         End
         Begin VB.Label tot_lou 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   720
            TabIndex        =   91
            Top             =   0
            Width           =   45
         End
         Begin VB.Label tot_lin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   2400
            TabIndex        =   90
            Top             =   0
            Width           =   45
         End
      End
      Begin VB.ComboBox cmbRoute 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   3900
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cash Currency Details"
         Height          =   3975
         Left            =   -70200
         TabIndex        =   64
         Top             =   960
         Width           =   3855
         Begin VB.TextBox TXTTOTAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Enabled         =   0   'False
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   2640
            TabIndex        =   85
            Text            =   " "
            Top             =   3550
            Width           =   1095
         End
         Begin VB.TextBox txt500 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   84
            Text            =   " "
            Top             =   661
            Width           =   1095
         End
         Begin VB.TextBox txt100 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   83
            Text            =   " "
            Top             =   1022
            Width           =   1095
         End
         Begin VB.TextBox txt50 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   82
            Text            =   " "
            Top             =   1383
            Width           =   1095
         End
         Begin VB.TextBox txt20 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   81
            Text            =   " "
            Top             =   1744
            Width           =   1095
         End
         Begin VB.TextBox txt10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   80
            Text            =   " "
            Top             =   2105
            Width           =   1095
         End
         Begin VB.TextBox txt5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   79
            Text            =   " "
            Top             =   2466
            Width           =   1095
         End
         Begin VB.TextBox txt2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   78
            Text            =   " "
            Top             =   2827
            Width           =   1095
         End
         Begin VB.TextBox txt1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   77
            Text            =   " "
            Top             =   3188
            Width           =   1095
         End
         Begin VB.TextBox txtchange 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   76
            Text            =   " "
            Top             =   3550
            Width           =   1095
         End
         Begin VB.TextBox txt1000 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   75
            Text            =   " "
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label LBLCHG 
            AutoSize        =   -1  'True
            Caption         =   "Change"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   3600
            Width           =   555
         End
         Begin VB.Label LBL1 
            AutoSize        =   -1  'True
            Caption         =   "1s"
            Height          =   195
            Left            =   240
            TabIndex        =   73
            Top             =   3240
            Width           =   165
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "2s"
            Height          =   195
            Left            =   240
            TabIndex        =   72
            Top             =   2880
            Width           =   165
         End
         Begin VB.Label LBL5 
            AutoSize        =   -1  'True
            Caption         =   "5s"
            Height          =   195
            Left            =   240
            TabIndex        =   71
            Top             =   2520
            Width           =   165
         End
         Begin VB.Label LBL10 
            AutoSize        =   -1  'True
            Caption         =   "10s"
            Height          =   195
            Left            =   240
            TabIndex        =   70
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label LBL20 
            AutoSize        =   -1  'True
            Caption         =   "20s"
            Height          =   195
            Left            =   240
            TabIndex        =   69
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label LBL50 
            AutoSize        =   -1  'True
            Caption         =   "50s"
            Height          =   195
            Left            =   240
            TabIndex        =   68
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label LBL100 
            AutoSize        =   -1  'True
            Caption         =   "100s"
            Height          =   195
            Left            =   240
            TabIndex        =   67
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label LBL500 
            AutoSize        =   -1  'True
            Caption         =   "500s"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   720
            Width           =   345
         End
         Begin VB.Label LBL1000 
            AutoSize        =   -1  'True
            Caption         =   "1000s"
            Height          =   195
            Left            =   240
            TabIndex        =   65
            Top             =   345
            Width           =   435
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ge&nerate"
         Height          =   375
         Left            =   -67560
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton SHL_SET 
         Caption         =   "Upd&ate Shells"
         Height          =   375
         Left            =   -67560
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtempshl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74520
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox TRP_NOS 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "M_FRM_RSI1.frx":0506
         Left            =   2640
         List            =   "M_FRM_RSI1.frx":050D
         TabIndex        =   49
         Top             =   2800
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "P&rint"
         Height          =   375
         Left            =   -67440
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   -67440
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Collection Details"
         Height          =   4575
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   4335
         Begin VB.TextBox M_CRT_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   131
            Text            =   "0.00"
            Top             =   4200
            Width           =   1495
         End
         Begin VB.TextBox m_dep_rcd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   130
            Text            =   "0.00"
            Top             =   2928
            Width           =   1495
         End
         Begin VB.TextBox M_CUR_CRD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   129
            Text            =   "0.00"
            Top             =   3720
            Width           =   1095
         End
         Begin VB.TextBox M_OLD_CRD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   128
            Text            =   "0.00"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox M_GRS_SAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   121
            Text            =   "0.00"
            Top             =   240
            Width           =   1495
         End
         Begin VB.TextBox M_FRE_QTY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   120
            Text            =   "0.00"
            Top             =   576
            Width           =   1495
         End
         Begin VB.TextBox M_NET_SAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   119
            Text            =   "0.00"
            Top             =   912
            Width           =   1495
         End
         Begin VB.TextBox M_DIS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   118
            Text            =   "0.00"
            Top             =   1248
            Width           =   1495
         End
         Begin VB.TextBox m_dep_ref 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   105
            Text            =   "0.00"
            Top             =   2592
            Width           =   1495
         End
         Begin VB.TextBox M_COM_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   62
            Text            =   "0.00"
            Top             =   1584
            Width           =   1495
         End
         Begin VB.TextBox M_BRK_QTY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   54
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
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox M_TAX_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   2256
            Width           =   1495
         End
         Begin VB.TextBox M_CAS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox M_EXP_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   1920
            Width           =   1495
         End
         Begin VB.TextBox M_CHQ_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
            Height          =   375
            Left            =   4320
            TabIndex        =   134
            Top             =   120
            Width           =   15
         End
         Begin VB.Label Label11Y 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cash Short-Excess"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   330
            TabIndex        =   132
            Top             =   4245
            Width           =   1335
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4320
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Label Label1801 
            AutoSize        =   -1  'True
            Caption         =   "Credit Sale"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2160
            TabIndex        =   127
            Top             =   3765
            Width           =   765
         End
         Begin VB.Label Label1701 
            Caption         =   "Old Dues"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2280
            TabIndex        =   126
            Top             =   3405
            Width           =   660
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4320
            Y1              =   3290
            Y2              =   3290
         End
         Begin VB.Label LBLLabel4 
            AutoSize        =   -1  'True
            Caption         =   "Gross Sales"
            Height          =   195
            Left            =   825
            TabIndex        =   125
            Top             =   285
            Width           =   840
         End
         Begin VB.Label LBLLabel5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Free Quantity"
            Height          =   195
            Left            =   720
            TabIndex        =   124
            Top             =   621
            Width           =   945
         End
         Begin VB.Label LBLLabel7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Net Sales"
            Height          =   195
            Left            =   975
            TabIndex        =   123
            Top             =   957
            Width           =   690
         End
         Begin VB.Label LBLLabel8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Discount Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   750
            TabIndex        =   122
            Top             =   1293
            Width           =   915
         End
         Begin VB.Label Label101 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Dep. Refunded (-)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   390
            TabIndex        =   107
            Top             =   2637
            Width           =   1275
         End
         Begin VB.Label Label102 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Dep. Received (+)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   360
            TabIndex        =   106
            Top             =   2973
            Width           =   1305
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Commission Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   555
            TabIndex        =   63
            Top             =   1629
            Width           =   1110
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Breakage"
            Height          =   195
            Left            =   4815
            TabIndex        =   56
            Top             =   1035
            Width           =   810
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Leakage"
            Height          =   195
            Left            =   4875
            TabIndex        =   55
            Top             =   1395
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sales Tax - VAT (+)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   285
            TabIndex        =   41
            Top             =   2301
            Width           =   1380
         End
         Begin VB.Label LLabel12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Expense Rs."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   765
            TabIndex        =   40
            Top             =   1965
            Width           =   900
         End
         Begin VB.Label Label10H 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cheques"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   75
            TabIndex        =   38
            Top             =   3765
            Width           =   630
         End
         Begin VB.Label Label9L 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cash"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   345
            TabIndex        =   37
            Top             =   3405
            Width           =   360
         End
      End
      Begin MSComctlLib.ListView LIST7 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   32
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
         Height          =   1200
         Left            =   5400
         TabIndex        =   31
         Top             =   480
         Width           =   3255
         Begin VB.CheckBox Check3 
            Caption         =   "&Complete Returns"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   300
            Width           =   1575
         End
         Begin VB.CheckBox Check4 
            Caption         =   "C&hallan Close"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   705
            Width           =   1815
         End
         Begin VB.CheckBox Trolley 
            Caption         =   "P&ush Cart"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4560
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   180
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker LINDATE 
            Height          =   330
            Left            =   1860
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16646147
            CurrentDate     =   37099
         End
      End
      Begin MSComctlLib.ListView list2 
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
         TabIndex        =   26
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2250
         Width           =   1455
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
         ItemData        =   "M_FRM_RSI1.frx":0516
         Left            =   -72960
         List            =   "M_FRM_RSI1.frx":0523
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74880
         TabIndex        =   4
         Top             =   4500
         Width           =   8295
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            Caption         =   "POP"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3360
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            Caption         =   "Promotional Item"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4680
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1800
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3350
         Width           =   1455
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   2640
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16646147
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
         Height          =   4155
         Left            =   -74940
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   7329
         _Version        =   393216
         Cols            =   24
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         FocusRect       =   0
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Challan ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Route"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Reference"
            Object.Width           =   2540
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
         TabIndex        =   29
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
            Text            =   "BBD-Exp. Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSS 
         Height          =   4425
         Left            =   -74880
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   450
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7805
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   277
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   16711680
         WordWrap        =   -1  'True
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid ms1 
         Height          =   4035
         Left            =   -74910
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   960
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   7117
         _Version        =   393216
         Cols            =   19
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid ms2 
         Height          =   4035
         Left            =   -74880
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   960
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   7117
         _Version        =   393216
         Cols            =   19
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
      End
      Begin VB.Label Labelttlvariance 
         AutoSize        =   -1  'True
         Caption         =   "Cash Variance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -69960
         TabIndex        =   136
         Top             =   4860
         Width           =   1035
      End
      Begin VB.Label LBLROUTE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2025
         TabIndex        =   114
         Top             =   3945
         Width           =   435
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1710
         TabIndex        =   104
         Top             =   1775
         Width           =   750
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1560
         TabIndex        =   87
         Top             =   3425
         Width           =   900
      End
      Begin VB.Label LBLMSG 
         AutoSize        =   -1  'True
         Caption         =   "Press F6 to Enter Currency Details"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -71797
         TabIndex        =   86
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trip No"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1935
         TabIndex        =   48
         Top             =   2860
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
         TabIndex        =   28
         Top             =   4560
         Width           =   11085
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
         Left            =   4200
         TabIndex        =   25
         Top             =   2250
         Width           =   2460
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1710
         TabIndex        =   24
         Top             =   2325
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   195
         Left            =   1920
         TabIndex        =   21
         Top             =   4525
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   2100
         TabIndex        =   20
         Top             =   1225
         Width           =   345
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
         Left            =   2640
         TabIndex        =   19
         Top             =   4450
         Width           =   1455
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
         Left            =   5040
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
         Left            =   4200
         TabIndex        =   16
         Top             =   3350
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
         Top             =   3225
         Width           =   45
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
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
         Left            =   1320
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
            Picture         =   "M_FRM_RSI1.frx":0530
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":0A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":0E3C
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":1190
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":17C8
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":1B1C
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":1F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":2270
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":267C
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":2788
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":2ADC
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":2EF0
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RSI1.frx":343C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   2040
      ScaleHeight     =   3555
      ScaleWidth      =   4635
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   0
         TabIndex        =   51
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
         TabIndex        =   52
         Top             =   0
         Width           =   4695
      End
   End
End
Attribute VB_Name = "M_FRM_RSI1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, bslab, REC_SET1 As New ADODB.Recordset, rs4 As New ADODB.Recordset, REC_SET5 As New ADODB.Recordset, MSROW, INVIDY, INC, dlr_idy
Dim RET_STK As Boolean, save_rec, ACTIDY
Dim RSCUR As New ADODB.Recordset
Dim rsrdy As New ADODB.Recordset
Dim invid As String, FLD_CHK As String, M_LDS_IDY As String, IDN As Boolean
Dim BMAKEBILL As Boolean
Dim rsVan As New ADODB.Recordset
Dim AR(5, 3)
Dim RSCOM As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim RSSNV As ADODB.Recordset
Dim SHLSETCLICK As Boolean
Dim txtTemp As Integer
Dim B_SAL_RAT, B_SAL_TAX As Double
Dim LST_TOT As Double
Dim M_CPN_AMT As Double
Dim F_LST_TOT As Double
Dim F_CPN_AMT As Double
Dim R_LST_TOT As Double
Dim R_CPN_AMT As Double
Dim CTAXVAL As Double
Dim FTAXVAL As Double
Dim BAS_VAL As Double
Dim TAX_VAL As Double
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
'On Error Resume Next
M_GRS_AMT = convert(amt(Val(Total)))
M_NET_AMT = convert(amt(Round(Val(M_GRS_AMT) - Val(billdisamt), 2)))
M_NET_AMT = convert(amt(Round(Val(M_NET_AMT) - Val(addisamt), 2)))
M_NET_AMT = convert(amt(Round(Val(M_NET_AMT) + Val(taxamt), 2)))
M_NET_AMT = convert(amt(Val(M_NET_AMT) - Val(damageamt)))
M_GRS_AMT = M_NET_AMT
End Sub

Private Sub saverecord()
'On Error Resume Next
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbInformation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
Else
    If MOD_REC = True Then
        Command2.Value = True           'Take Control thru' Generate and Other Important Update Buttons
        SHL_SET.Value = True
        Command6.Value = True
        CheckShutDownStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'")
        If Trim(CheckShutDownStatus) = "IMPROPER SHUTDOWN" Then
            MsgBox ("Please reset the Improper Shutdown flag in System-Shutdown Status"), vbCritical, "DAS Version 4.0"
            CANCELRECORD
            Exit Sub
        End If
    End If
End If
If Len(Trim(cmbRoute)) <= 0 Then
    MsgBox "Select Route!", vbInformation + vbOKOnly, "DAS Version 4.0"
    SSTab1.Tab = 0
    cmbRoute.SetFocus
    Exit Sub
End If
If CHECKTB("smn.dbf,act.dbf,van.dbf,prd.dbf,brd.dbf,inl.dbf,snv.dbf,vbk.dbf,sbk.dbf,cbk.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("smn", "smn_idy", Trim(M_SMN_IDY)) = False Then
    MsgBox "Salesman Not Found, Define Salesman !", vbInformation, "DAS Version 4.0"
    M_SMN_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If
If check("act", "act_idy", Trim(M_CST_IDY)) = False Then
    MsgBox "Account Not Found, Define Account !", vbInformation, "DAS Version 4.0"
    M_CST_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If
' Day Closure Routine Added by PNR as Addon Control

Dim RS_DAYCLS As New ADODB.Recordset
If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
RS_DAYCLS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' .and. alltrim(DIR_ALS) = '" & M_INV_DAT & "'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
If RS_DAYCLS.RecordCount > 0 Then
   If Trim(RS_DAYCLS!DIR_MSG) = "Y" Then
    MsgBox "Day Closure Effective ! Cannot Alter - Create Transaction FOR " & RS_DAYCLS("DIR_ALS"), vbCritical, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
   End If
Else
   RS_DAYCLS.AddNew
   RS_DAYCLS("DIR_IDY") = "DAY"
   RS_DAYCLS("DIR_IDX") = ""
   RS_DAYCLS("DIR_ALS") = M_INV_DAT
   RS_DAYCLS("DIR_MSG") = "N"
   RS_DAYCLS.Update
End If

' End of Day Closure Routine

' Check For Day Closure Tolerance Days as per System -> Settings -> Specify Defaults Control Option Z

Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open "SELECT * FROM DIR WHERE DIR_IDY = 'DAY' AND TRIM(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveLast
    LST_DCD = DIR_RS("DIR_DAT")
End If
CHK_DATE = LST_DCD + CHK_DAYS
CHK_DAYS = GetValue("DIR_ALS", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'DCD'")

CHK_DATE = LST_DCD + CHK_DAYS
If M_INV_DAT <= CHK_DATE Then
Else
    MsgBox "Day Closure Tolerance Days " & Trim(str(CHK_DAYS)) & " Crossed! Cannot Alter - Create Transaction !", vbCritical, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
End If

' End of Day Closure Tolerance Days Check
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If REC_SET.EOF = False Then
    REC_SET.MoveFirst
    If Val(tot_lou) > Val(REC_SET("VAN_CAP")) Then
        MsgBox "Load Out is More Than Van Capacity ! Load-Out : " & Trim(str(Val(tot_lou))) & " Van Cap : " & Trim(str(Val(REC_SET("VAN_CAP")))), vbInformation, "DAS Version 4.0"
    End If
End If

If Check3.Enabled = False And Check4.Enabled = False Then
    MsgBox "Challan Closed ! Changes Not Saved!", vbInformation, "DAS Version 4.0"
    INC = Trim(M_INV_IDY)
    GoTo VAN_PRT_CHK:
    CANCELRECORD
    Exit Sub
End If

If MOD_REC = False Then
    If querycheck("van", "smn_idy", M_SMN_IDY, " and van_cls in('S','N')") = True Then MsgBox "Van Challans (Settlements) Pending for this Salesman!", vbInformation
    If querycheck("van", "ROU_idy", cmbRoute, " and van_cls in('S','N')") = True Then MsgBox "Van Challans (Settlements) Pending for this Route!", vbInformation
End If

If Check4.Value = 1 And Check4.Enabled = True Then
    If Val(MSS.TextMatrix(1, 3)) < -9 Or Val(MSS.TextMatrix(1, 3)) > 9 Then
        MsgBox "Challan Cannot Be Closed! Gross Sale Value - Variance Exists! Value " & Trim(str(Val(MSS.TextMatrix(1, 3)))), vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 4
        MSS.SetFocus
        Exit Sub
    End If
    If Val(MSS.TextMatrix(9, 3)) < -9 Or Val(MSS.TextMatrix(9, 3)) > 9 Then
        MsgBox "Challan Cannot Be Closed! Net Sale Value - Variance Exists! Value " & Trim(str(Val(MSS.TextMatrix(9, 3)))), vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 4
        MSS.SetFocus
        Exit Sub
    End If
''''    Challan Can Be Closed only on Cash Variance Position, since the closure should allow for Salesman Shortage, as per Prince Chopra FeedBack
''''    If Val(MSS.TextMatrix(10, 3)) < 0 Or Val(MSS.TextMatrix(10, 3)) > 1 Then
''''        MsgBox "Challan Cannot Be Closed! Cash Variance Exists!", vbInformation, "DAS Version 4.0"
''''        SSTab1.Tab = 4
''''        MSS.SetFocus
''''        Exit Sub
''''    End If
    C_CSH_VAR = Val(MSS.TextMatrix(10, 3))
    C_CHQ_VAR = Val(MSS.TextMatrix(11, 3))
    C_EXP_VAR = Val(MSS.TextMatrix(14, 3))
    C_CSH_SHT = Val(MSS.TextMatrix(12, 3))
    C_DEP_RFD = Val(MSS.TextMatrix(15, 3))
    
    C_DEP_RCD = Val(MSS.TextMatrix(16, 3))
    C_OLD_COL = Val(MSS.TextMatrix(17, 3))
    
    If (Val(C_CSH_VAR) + Val(C_EXP_VAR) + Val(C_CSH_SHT) + Val(C_DEP_RFD) + Val(C_CHQ_VAR)) - (Val(C_DEP_RCD) + Val(C_OLD_COL)) < -9 Or (Val(C_CSH_VAR) + Val(C_EXP_VAR) + Val(C_CSH_SHT) + Val(C_DEP_RFD) + Val(C_CHQ_VAR)) - (Val(C_DEP_RCD) + Val(C_OLD_COL)) > 9 Then
        MsgBox "Challan Cannot Be Closed! Overall CASH Variance Exists!", vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 4
        MSS.SetFocus
        Exit Sub
    End If
'    If Val(MSS.TextMatrix(11, 3)) > Val(MSS.TextMatrix(17, 3)) Then         'Cheque Variance should be less than Old Dues - Collection Amount Prince Chopra FeedBack
'        MsgBox "Challan Cannot Be Closed! Cheque Vs Old Dues Variance Exists ! ", vbInformation, "DAS Version 4.0"
'        SSTab1.Tab = 4
'        MSS.SetFocus
'        Exit Sub
'    End If
    If Val(MSS.TextMatrix(11, 3)) < 0 Then ' -9 Or Val(MSS.TextMatrix(11, 3)) > 9 Then
        MsgBox "Challan Cannot Be Closed! Cheque Variance HAS -VE Value " & Trim(str(Val(MSS.TextMatrix(11, 3)))), vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 4
        MSS.SetFocus
        Exit Sub
    End If
    If Val(MSS.TextMatrix(11, 3)) > Val(MSS.TextMatrix(17, 3)) Then ' -9 Or Val(MSS.TextMatrix(11, 3)) > 9 Then
        MsgBox "Challan Cannot Be Closed! Cheque Variance > (More than) Old Dues Variance ! Value " & Trim(str(Val(MSS.TextMatrix(11, 3)) - Val(MSS.TextMatrix(17, 3)))), vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 4
        MSS.SetFocus
        Exit Sub
    End If
    If Val(MSS.TextMatrix(18, 3)) < -9 Or Val(MSS.TextMatrix(18, 3)) > 9 Then
        MsgBox "Challan Cannot Be Closed! Credit Sale Value - Variance Exists! Value " & Trim(str(Val(MSS.TextMatrix(18, 3)))), vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 4
        MSS.SetFocus
        Exit Sub
    End If
End If
'''''    If Val(MSS.TextMatrix(1, 3)) > 1 Or Val(MSS.TextMatrix(9, 3)) > 1 Or Val(MSS.TextMatrix(18, 3)) > 1 Then
'''''        MsgBox "Challan Cannot Be Closed! Gross, Net, Credit Sale Value - Variance Exists!", vbCritical, "DAS Version 4.0"
'''''        SSTab1.Tab = 4
'''''        MSS.SetFocus
'''''        Exit Sub
'''''    End If

If Check4.Value = 1 Then 'And Check4.Enabled = True Then
    RETURNSTOCKS
    For A = 1 To MS.Rows - 1
        If Val(MS.TextMatrix(A, 11)) <> 0 Then
            MsgBox "CANNOT Close Van Challan! Stock Variance Not Reconciled!", vbInformation, "DAS Version 4.0"
            CANCELRECORD
            Exit Sub
        End If
    Next
' Code to Check for Free/Scheme Variance in Invoices made against this Van Challan ' PNR

    If REC_SET1.State = 1 Then REC_SET1.Close
    REC_SET1.Open "SELECT * FROM INV WHERE INV.ORD_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    
    If REC_SET1.RecordCount > 0 Then
        While Not REC_SET1.EOF = True
                If REC_SET.State = 1 Then REC_SET.Close
                REC_SET.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "' AND VAN.PRD_IDY = '" & Trim(REC_SET1("PRD_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If REC_SET.EOF = True Then
                    MsgBox "Batch Level MisMatch in Schemes Exist! Proceeding for Closure!" & " SKU ID " & Trim(REC_SET1("PRD_IDY")), vbCritical, "DAS Version 4.0"
                    ' Enabled Challan Closure based on Prince Chopra escalation
                    'CANCELRECORD
                    'Exit Sub
                End If
            REC_SET1.MoveNext
        Wend
'''''    Else
'''''        MsgBox "CANNOT Close Van Challan! Invoices Not Made ! ", vbInformation, "DAS Version 4.0"
'''''        CANCELRECORD
'''''        Exit Sub
    End If
    
' End of Code to Check for Free/Scheme ' PNR

End If

If Trolley.Value = 1 Then
Else
    If TRP_NOS = "All" And MOD_REC = False And Check3.Value = 0 Then
        MsgBox "Invalid Trip Number! ", vbInformation, "DAS Version 4.0"
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
    If MsgBox("Trip Wise Data will be Merged ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
        SSTab1.Tab = 0
        Exit Sub
    End If
End If
If Check3.Value = 1 And MS.TextMatrix(1, 1) = "All" And Check3.Enabled = True Then
    If MsgBox("Trip Wise Data will be Merged ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
        SSTab1.Tab = 0
        Exit Sub
    End If
End If
If Check4.Value = 1 And Check4.Enabled = True Then
    Command6_Click
    For A = 2 To MS1.Rows - 1
        Q_QTY = Val(Q_QTY) + Val(MS1.TextMatrix(A, 7))
        F_QTY = Val(F_QTY) + Val(MS1.TextMatrix(A, 8))
'''''        If Val(ms1.TextMatrix(A, 7)) <> 0 Or Val(ms1.TextMatrix(A, 8)) <> 0 Then
'''''            MsgBox "Challan Cannot Be Closed! Empties - Glass Variance Exists!", vbInformation, "DAS Version 4.0"
'''''            SSTab1.Tab = 6
'''''            ms1.SetFocus
'''''            Exit Sub
'''''        End If
    Next
End If

If Check4.Value = 1 And Check4.Enabled = True Then
    If Val(Q_QTY) <> 0 Or Val(F_QTY) <> 0 Then
        MsgBox "Challan Cannot Be Closed! Total Empties - Glass Variance Exists! " & Trim(str(Val(Q_QTY))) & " , " & Trim(str(Val(F_QTY))), vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 6
        MS1.SetFocus
        Exit Sub
    End If
End If

If CHECKSTOCK = True Then Exit Sub
If REC_SET1.State = 1 Then REC_SET1.Close
If REC_SET.State = 1 Then REC_SET.Close
If MOD_REC = True Then
    If Check3.Value = 1 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from van,PRD where VAN.PRD_IDY = PRD.PRD_IDY AND inv_idy = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset
        For A = 1 To REC_SET.RecordCount
            retqty = MTS(Val(REC_SET!RET_QTY & ""), Val(REC_SET!SUB_UNT))
            brkqty = MTS(Val(REC_SET!BRK_QTY & ""), Val(REC_SET!SUB_UNT))
            lkgqty = MTS(Val(REC_SET!LKG_QTY & ""), Val(REC_SET!SUB_UNT))
            RET_QTY = STM(Val(retqty) + Val(brkqty) + Val(lkgqty), REC_SET!SUB_UNT)
            con.Execute "update prd where prd_idy = '" & REC_SET("prd_idy") & "' set cur_lev = " & CONQTY(Val(RET_QTY), REC_SET("prd_idy"))
            con.Execute "UPDATE PRD WHERE PRD_IDY = '" & REC_SET("prd_idy") & "' set PRD_PCS = PRD_PCS - " & MTS(Val(RET_QTY), Val(REC_SET("SUB_UNT")))
            con.Execute "update Brd where BAT_idy = '" & REC_SET("BAT_idy") & "' set PRD_QTY = " & CONBRDQTY(Val(RET_QTY), REC_SET("BAT_idy"))
            REC_SET.MoveNext
        Next

    ElseIf Check3.Value = 0 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from van,PRD where VAN.PRD_IDY = PRD.PRD_IDY AND inv_idy = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset
        For A = 1 To REC_SET.RecordCount
            con.Execute "update prd where prd_idy = '" & REC_SET("prd_idy") & "' set cur_lev = " & ADDQTY(Val(REC_SET("TOT_qty")), REC_SET("prd_idy"))
            con.Execute "UPDATE PRD WHERE PRD_IDY = '" & REC_SET("prd_idy") & "' set PRD_PCS = PRD_PCS + " & MTS(Val(REC_SET("TOT_qty")), Val(REC_SET("SUB_UNT")))
            con.Execute "update Brd where BAT_idy = '" & REC_SET("BAT_idy") & "' set PRD_QTY = " & ADDBRDQTY(Val(REC_SET("TOT_qty")), REC_SET("BAT_idy"))
            REC_SET.MoveNext
        Next
    End If
    TxnMoveBackup "VAN", "VBK", M_INV_IDY, ""
    TxnMoveBackup "CUR", "CBK", M_INV_IDY, ""
    
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    con.Execute "delete from van where inv_idy = '" & Trim(M_INV_IDY) & "'", d
    con.Execute "delete from cur where inv_idy = '" & Trim(M_INV_IDY) & "'", d
    INC = Trim(M_INV_IDY)
Else
    If Trim(MS.TextMatrix(1, 0)) <> "" Then
        INC = "RSI" & T7increment("IDY_RSI")
        INVIDY = INC
        If Trim(M_INV_IDY) = "" And Trolley.Value = 1 Then
            con.Execute "UPDATE INL SET ORD_IDY = '" & Trim(INC) & "' WHERE INV_IDY = '" & Trim(invid) & "'"
        End If
    Else
         CANCELRECORD
        Exit Sub
    End If
End If
If REC_SET.State = 1 Then REC_SET.Close
If RSCUR.State = 1 Then RSCUR.Close
RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY = ''", con, adOpenKeyset, adLockPessimistic
REC_SET.Open "select * from van where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
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
        REC_SET("prd_idy") = Trim(MS.TextMatrix(A, 14))
        REC_SET("ACT_idy") = Trim(M_ACT_IDY)
        REC_SET("CST_idy") = Trim(M_CST_IDY)
        REC_SET("SMN_idy") = Trim(M_SMN_IDY)
        REC_SET("rou_idy") = Trim(cmbRoute.Text & "")
        REC_SET!LIN_DAT = LINDATE
        REC_SET!MNL_REF = Trim(MNL_REF)
        If Pieces.Value = 0 Then
            MS.TextMatrix(A, 2) = STM(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 3) = STM(Val(MS.TextMatrix(A, 3)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 4) = STM(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 6) = STM(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 5) = STM(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 19) = Val(Val(MS.TextMatrix(A, 19)) * Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 6) = STM(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 7) = STM(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 8) = STM(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 10) = STM(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
            MS.TextMatrix(A, 12) = convert(Val(MS.TextMatrix(A, 12)) * Val(MS.TextMatrix(A, 16)))
        End If
        REC_SET("OPN_QTY") = Val(MS.TextMatrix(A, 2))
        REC_SET("ISS_QTY") = Val(MS.TextMatrix(A, 3))
        REC_SET("tot_qty") = Val(MS.TextMatrix(A, 4))
        REC_SET("fre_qty") = Val(MS.TextMatrix(A, 6))
        REC_SET("brk_qty") = Val(MS.TextMatrix(A, 7))
        REC_SET("lkg_qty") = Val(MS.TextMatrix(A, 8))
        REC_SET("BIL_QTY") = Val(MS.TextMatrix(A, 10))
        REC_SET("RET_QTY") = Val(MS.TextMatrix(A, 5))
        REC_SET("BAT_idy") = Trim(MS.TextMatrix(A, 17))
        REC_SET!PRD_rat = Val(MS.TextMatrix(A, 19))
        
        REC_SET("OPN_PCS") = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 16)))
        REC_SET("ISS_PCS") = MTS(Val(MS.TextMatrix(A, 3)), Val(MS.TextMatrix(A, 16)))
        REC_SET("TOT_PCS") = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
        REC_SET("FRE_PCS") = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
        REC_SET("BRK_PCS") = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
        REC_SET("LKG_PCS") = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
        REC_SET("BIL_PCS") = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
        REC_SET("RET_PCS") = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
        REC_SET("VAN_SUT") = Val(MS.TextMatrix(A, 16))
        REC_SET("LDS_IDY") = M_LDS_IDY
        If Check3.Value = 1 Then
            REC_SET!vAN_CLS = "S"
        Else
            REC_SET!vAN_CLS = "N"
        End If
        If Check4.Value = 1 Or save_rec = "Save" Then
            save_rec = ""
            REC_SET!vAN_CLS = "C"
        End If
        REC_SET!COL_AMT = Val(MS.TextMatrix(A, 13))
        REC_SET!CAS_AMT = Val(M_CAS_AMT)
        REC_SET!CHQ_AMT = Val(M_CHQ_AMT)
        REC_SET!DIS_AMT = Val(M_DIS_AMT)
        REC_SET!EXP_AMT = Val(M_EXP_AMT)
        REC_SET!ADD_AMT = Val(M_TAX_AMT)
        REC_SET!com_amt = Val(M_COM_AMT)
        REC_SET!dep_ref = Val(m_dep_ref)
        REC_SET!dep_rcd = Val(m_dep_rcd)
        REC_SET!OLD_CRD = Val(M_OLD_CRD)
        REC_SET!CUR_CRD = Val(M_CUR_CRD)
        REC_SET!SHT_CSH = Val(MSS.TextMatrix(12, 3))
        'REC_SET!BIL_IDY = Trim(Val(M_COM_AMT)) & ""
        If Trolley.Value = 1 Then
            REC_SET!VAN_CHK = "P"
        Else
            REC_SET!VAN_CHK = "V"
        End If
        'Commission Update
        If Check3.Value = 1 And Check3.Enabled = True Then 'Calculate and Save Commission only first time
            If RSCOM.State = 1 Then RSCOM.Close
            RSCOM.Open "select * from sic WHERE act_IDY='" & Trim(M_SMN_IDY) & "' and prd_idy='" & Trim(MS.TextMatrix(A, 14)) & "' ", con, adOpenKeyset, adLockPessimistic
            If Not RSCOM.EOF Then
                com_val = 0
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select * from prd where prd_idy='" & Trim(MS.TextMatrix(A, 14)) & "' GROUP BY prd_idy", con, adOpenKeyset, adLockPessimistic
                If Pieces.Value = 0 Then
                    TOTQ = Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 8)) + Val(MS.TextMatrix(A, 7)) + Val(MS.TextMatrix(A, 6)))
                    TOTQ = STM(Val(TOTQ), Val(MS.TextMatrix(A, 16)))
                Else
                    TOTQ = Val(MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))) - (Val(MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))) + Val(MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))) + Val(MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))) + Val(MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))))
                    TOTQ = STM(Val(TOTQ), Val(MS.TextMatrix(A, 16)))
                End If
                If MTS(Val(TOTQ), MS.TextMatrix(A, 16)) >= MTS(RSCOM("min_qty"), MS.TextMatrix(A, 16)) Then
                    If Pieces.Value = 0 Then
                        If Not rsprd.EOF Then com_val = MTS(Val(TOTQ), Val(rsprd("SUB_UNT"))) * (RSCOM("min_val") / Val(rsprd("sub_unt")))
                    Else
                        'com_val = Val(TOTQ) * RSCOM("min_val")
                        com_val = (Int(TOTQ) * RSCOM("min_val")) + (((TOTQ - Int(TOTQ)) * 1000) * (RSCOM("MIN_VAL") / Val(rsprd("sub_unt"))))
                    End If
                End If
                REC_SET!BIL_IDY = Trim(amt(Round(Val(com_val), 2))) & ""
            Else
                REC_SET!BIL_IDY = ""
            End If
        End If
        'Commission Update End
        REC_SET.Update
        If Check3.Value = 1 Then
            retqty = MTS(Val(Trim(MS.TextMatrix(A, 5) & "")), Val(MS.TextMatrix(A, 16)))
            brkqty = MTS(Val(Trim(MS.TextMatrix(A, 7) & "")), Val(MS.TextMatrix(A, 16)))
            lkgqty = MTS(Val(Trim(MS.TextMatrix(A, 8) & "")), Val(MS.TextMatrix(A, 16)))
            RET_QTY = STM(Val(retqty) + Val(brkqty) + Val(lkgqty), Val(MS.TextMatrix(A, 16)))
            con.Execute "update prd where prd_idy = '" & Trim(MS.TextMatrix(A, 14)) & "' set cur_lev = " & ADDQTY(QTY(Trim(RET_QTY & "")), Trim(MS.TextMatrix(A, 14)))
            con.Execute "UPDATE PRD WHERE PRD_IDY = '" & Trim(MS.TextMatrix(A, 14)) & "' set PRD_PCS = PRD_PCS + " & MTS(Val(RET_QTY), Val(MS.TextMatrix(A, 16)))
            con.Execute "UPDATE BRD WHERE BAT_IDY = '" & Trim(MS.TextMatrix(A, 17)) & "' set prd_qty = " & ADDBRDQTY(QTY(Val(RET_QTY)), Trim(MS.TextMatrix(A, 17)))
        ElseIf Check3.Value = 0 Then
            con.Execute "update prd where prd_idy = '" & Trim(MS.TextMatrix(A, 14)) & "' set cur_lev = " & CONQTY(QTY(Trim(MS.TextMatrix(A, 4) & "")), Trim(MS.TextMatrix(A, 14)))
            con.Execute "UPDATE PRD WHERE PRD_IDY = '" & Trim(MS.TextMatrix(A, 14)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
            con.Execute "UPDATE BRD WHERE BAT_IDY = '" & Trim(MS.TextMatrix(A, 17)) & "' set prd_qty = " & CONBRDQTY(QTY(Val(MS.TextMatrix(A, 4))), Trim(MS.TextMatrix(A, 17)))
        End If
    End If
Next
'CHFID = 0
'Do While Not RSCUR.EOF
'    If RSCUR(0) = Trim(INC) Then
'        CHFID = 1
'        Exit Do
'    End If
'    RSCUR.MoveNext
'Loop
If Check4.Value = 1 And Check4.Enabled = True Then
    Dim DFL_CSH_ACT As New ADODB.Recordset
    If DFL_CSH_ACT.State = 1 Then DFL_CSH_ACT.Close
    DFL_CSH_ACT.Open ("SELECT * FROM ACT WHERE GRP_CLF = 'Cash Book'"), con, adOpenKeyset, adLockPessimistic
    If DFL_CSH_ACT.RecordCount > 0 Then
        DFL_CSH_ACT.MoveFirst
        CSH_ACT_IDY = Trim(DFL_CSH_ACT("ACT_IDY"))
        Dim JRN_ACT As New ADODB.Recordset
        If JRN_ACT.State = 1 Then JRN_ACT.Close
        JRN_ACT.Open ("SELECT * FROM VCH"), con, adOpenKeyset, adLockPessimistic
        If Abs(Val(MSS.TextMatrix(10, 3))) <> 0 Then
            If Val(MSS.TextMatrix(12, 3)) > 0 Then
                INC = "PMT" & T7increment("idy_pay")
                JRN_ACT.AddNew
                JRN_ACT("TXN_IDY") = INC
                JRN_ACT("TXN_DAT") = Trim(M_INV_DAT)
                JRN_ACT("TXN_AMT") = Abs(Val(MSS.TextMatrix(12, 3)))
                JRN_ACT("DBT_IDY") = Trim(M_CST_IDY)
                JRN_ACT("CRT_IDY") = Trim(CSH_ACT_IDY)
                JRN_ACT("TXN_RMK") = "Sys Gen Van ShortCash Challan No:- " & Trim(M_INV_IDY)
                JRN_ACT("TXN_TYP") = "PMT"
                JRN_ACT("CLG_FLG") = ""
                JRN_ACT("CLG_DAT") = Trim(M_INV_DAT)
                JRN_ACT("INV_IDY") = ""
                JRN_ACT("VAN_REF") = Trim(M_INV_IDY)
                JRN_ACT.Update
            ElseIf Val(MSS.TextMatrix(12, 3)) < 0 Then
                INC = "RCT" & T7increment("idy_rct")
                JRN_ACT.AddNew
                JRN_ACT("TXN_IDY") = INC
                JRN_ACT("TXN_DAT") = Trim(M_INV_DAT)
                JRN_ACT("TXN_AMT") = Abs(Val(MSS.TextMatrix(12, 3)))
                JRN_ACT("DBT_IDY") = Trim(CSH_ACT_IDY)
                JRN_ACT("CRT_IDY") = Trim(M_CST_IDY)
                JRN_ACT("TXN_RMK") = "Sys Gen Van ExcessCash Challan No:- " & Trim(M_INV_IDY)
                JRN_ACT("TXN_TYP") = "RCT"
                JRN_ACT("CLG_FLG") = ""
                JRN_ACT("CLG_DAT") = Trim(M_INV_DAT)
                JRN_ACT("INV_IDY") = ""
                JRN_ACT("VAN_REF") = Trim(M_INV_IDY)
                JRN_ACT.Update
            End If
            If Abs(Val(M_EXP_AMT)) <> 0 Then
                JRN_ACT.AddNew
                JRN_ACT("TXN_IDY") = INC
                JRN_ACT("TXN_DAT") = Trim(M_INV_DAT)
                JRN_ACT("TXN_AMT") = Abs(Val(M_EXP_AMT))
                JRN_ACT("DBT_IDY") = Trim(M_CST_IDY)
                JRN_ACT("CRT_IDY") = Trim(CSH_ACT_IDY)
                JRN_ACT("TXN_RMK") = "Sys Gen Van Expenses Challan No:- " & Trim(M_INV_IDY)
                JRN_ACT("TXN_TYP") = "PMT"
                JRN_ACT("CLG_FLG") = ""
                JRN_ACT("CLG_DAT") = Trim(M_INV_DAT)
                JRN_ACT("INV_IDY") = ""
                JRN_ACT("VAN_REF") = Trim(M_INV_IDY)
                JRN_ACT.Update
            End If
          End If
        MsgBox ("Financial Entries for Expenses and ShortCash Posted!"), vbInformation, "DAS Version 4.0"
    Else
        MsgBox ("Default Cash Account Not Available for Financial Entries Posting!"), vbCritical, "DAS Version 4.0"
    End If
End If
RSCUR.AddNew    ' edited by radhika
    RSCUR(0) = Trim(INC)
    RSCUR(1) = Val(Trim(txt1000) & "")
    RSCUR(2) = Val(Trim(txt500) & "")
    RSCUR(3) = Val(Trim(txt100) & "")
    RSCUR(4) = Val(Trim(txt50) & "")
    RSCUR(5) = Val(Trim(txt20) & "")
    RSCUR(6) = Val(Trim(txt10) & "")
    RSCUR(7) = Val(Trim(txt5) & "")
    RSCUR(8) = Val(Trim(txt2) & "")
    RSCUR(9) = Val(Trim(txt1) & "")
    RSCUR(10) = Val(Trim(txtchange) & "")
    For A = 1 To MS.Rows - 1
        If Val(MS.TextMatrix(A, 11)) <> 0 Or (Val(MSS.TextMatrix(9, 3)) < -9 Or Val(MSS.TextMatrix(9, 3)) > 9) Then
            RSCUR(11) = 1
        Else
            RSCUR(11) = 0
        End If
    Next

    RSCUR.Update    ' END
M_INV_IDY = INC
ShlSet_Click
SAVESHELL Trim(INC)
SHLSETCLICK = False
If REC_SET1.State = 1 Then REC_SET1.Close
If MOD_REC = False Then MsgBox "New Challan ID is " & INC, vbInformation, "DAS Version 4.0"
Label23 = labdisp("idy_rsi")
VAN_PRT_CHK:
If CHECKTB("prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_RSI1") = True Then
    con.Execute "update dir where dir_idy = 'INV' set dir_als = '" & INC & "'"
    con.Execute "update dir where dir_idy = 'REM' set dir_als = 'RSI'"
    If Trolley.Value = 1 Then
        If Pieces.Value = 0 Then
            OpenReport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Pieces)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Cases)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        End If
    Else
        If Check3.Value <> 1 Then
            OpenReport CURDIR & "\DOS_VCH.RP1", "Van MTF Gate Pass", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
            If REC_SET.State = 1 Then REC_SET.Close
            REC_SET.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(INC) & "' GROUP BY TRP_IDY", con, adOpenKeyset, adLockPessimistic
            If REC_SET.RecordCount > 0 Then
                While REC_SET.EOF = False
                    If Val(REC_SET("TRP_IDY")) > 3 Then
                        OpenReport CURDIR & "\DOS_VCH.RP1", "Van MTF Gate Pass Seq", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
                    End If
                    REC_SET.MoveNext
                Wend
                REC_SET.Close
            Else
                REC_SET.Close
            End If
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Sale Estimate", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
            SAL_QRY = "INL->ORD_IDY = '" & Trim(M_INV_IDY) & "' .AND. INL->TXN_TYP = 'SAL'"
            updatevan
            OpenReport CURDIR & "\REP_LIB.RP1", "Van Reconcile Cases", 0, SAL_QRY, 6

            'OpenReport CURDIR & "\DOS_VCH.RP1", "Van Challan", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        End If
    End If
    CANCELRECORD
End If
CANCELRECORD
End Sub

Private Sub CANCELRECORD()
clrctr M_FRM_RSI1
setval
Form_Load
MS.Rows = 1
MS1.Rows = 2
ms2.Rows = 2
MS.AddItem ""
SSTab1.Tab = 0
Toolbar1.Buttons(12).Enabled = False
Toolbar1.Buttons(19).Enabled = False
M_INV_IDY.Enabled = True
M_INV_IDY.SetFocus
Trolley.Enabled = False
For A = 1 To MSS.Rows - 1
    MSS.TextMatrix(A, 1) = "0.00"
    MSS.TextMatrix(A, 2) = "0.00"
    MSS.TextMatrix(A, 3) = "0.00"
Next
If cmbRoute.ListCount > 0 Then cmbRoute.ListIndex = 0
If Trolley.Value = 0 Then TRP_NOS.Enabled = True
Command6.Enabled = True
TRP_NOS.CLEAR
TRP_NOS.AddItem "All"
TRP_NOS = ""
invid = ""
MOD_REC = False
TRP_NOS.Text = "1"
tot_lin = "0.000"
tot_lou = "0.000"
tot_net = "0.000"
tot_var = "0.000"
Tot_fre = "0.000"
End Sub

Private Sub newrecord()
CANCELRECORD
If Trolley.Value = 1 Then
MS.ColWidth(11) = 0
MS.ColWidth(10) = 1050
End If
Check3.Enabled = False 'Changed by PNR based on Prince Chopra Update
'Check4.Enabled = False
M_INV_IDY.Enabled = False
MNL_REF.SetFocus
SSTab1.Tab = 0
MS.FixedCols = 0
If cmbRoute.ListCount > 0 Then cmbRoute.ListIndex = 0
idnew = True
IDN = True
TRP_NOS.Text = "1"
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
    REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check2.Value = 1 Then
    REC_SET.Open "select * from SMN WHERE SMN_AD4 like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
End If
list1.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(REC_SET("smn_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    list1.Visible = True
    list1.SetFocus
    
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Salesmen Not Found, Define Salesmen !", vbInformation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_SMN_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    TRP_NOS.AddItem "All"
'    TRP_NOS.Text = "All"
'    TRP_NOS.Enabled = False
    Check4.Enabled = False
    LINDATE.Enabled = True
    LINDATE.SetFocus
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

Private Sub Check5_Click()
If Check5.Value = 1 Then
    Check2.Value = 0
    If Check1.Visible = True Then Check1.Value = 0
    If Check6.Visible = True Then Check6.Value = 0
    PRODUCTLIST
Else
    PRODUCTLIST
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
    Check2.Value = 0
    If Check1.Visible = True Then Check1.Value = 0
    If Check5.Visible = True Then Check5.Value = 0
    PRODUCTLIST
Else
    PRODUCTLIST
End If
End Sub

Private Sub cmbRoute_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SSTab1.Tab = 1
    If MOD_REC = False Then
        Toolbar1.Buttons(12).Enabled = True
        Toolbar1.Buttons(19).Enabled = True
    End If
    MS.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    TRP_NOS.SetFocus
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
    MSS.TextMatrix(13, 1) = convert(Val(MSS.TextMatrix(9, 1)) - (Val(MSS.TextMatrix(10, 1)) + Val(MSS.TextMatrix(11, 1)) + Val(MSS.TextMatrix(12, 1))))
    MSS.TextMatrix(9, 3) = convert(Val(MSS.TextMatrix(9, 1)) - Val(MSS.TextMatrix(9, 2)))
    MSS.TextMatrix(10, 3) = convert(Val(MSS.TextMatrix(10, 1)) - Val(MSS.TextMatrix(10, 2)))
    MSS.TextMatrix(11, 3) = convert(Val(MSS.TextMatrix(11, 1)) - Val(MSS.TextMatrix(11, 2)))
    
'    If MSS.TextMatrix(12, 2) >= 0 Then
        MSS.TextMatrix(12, 3) = convert(MSS.TextMatrix(12, 1) - (MSS.TextMatrix(12, 2)))
'    Else
'        FST_VAL = 0
'        SEC_VAL = 0
'        FST_VAL = Abs(MSS.TextMatrix(12, 1))
'        SEC_VAL = Abs(MSS.TextMatrix(12, 2))
'        If (FST_VAL - SEC_VAL > 0) Then
'            MSS.TextMatrix(12, 3) = convert(((FST_VAL - SEC_VAL) * -1))
'        Else
'            MSS.TextMatrix(12, 3) = convert((FST_VAL - SEC_VAL))
'        End If
'    End If
    
    MSS.TextMatrix(14, 3) = convert(Val(MSS.TextMatrix(14, 1)) - Val(MSS.TextMatrix(14, 2)))
    MSS.TextMatrix(15, 3) = convert(Val(MSS.TextMatrix(15, 1)) - Val(MSS.TextMatrix(15, 2)))
    MSS.TextMatrix(16, 3) = convert(Val(MSS.TextMatrix(16, 1)) - Val(MSS.TextMatrix(16, 2)))
    MSS.TextMatrix(17, 3) = convert(Val(MSS.TextMatrix(17, 1)) - Val(MSS.TextMatrix(17, 2)))
    MSS.TextMatrix(18, 3) = convert(Val(MSS.TextMatrix(18, 1)) - Val(MSS.TextMatrix(18, 2)))

For i = 1 To 11
    MSS.Row = i
    For J = 0 To 3
        MSS.Col = J
        MSS.CellBackColor = &HFFFFC0
    Next
Next

TTL_CSH_VAR = 0#

TTL_CSH_VAR = convert((Val(MSS.TextMatrix(10, 3)) + Val(MSS.TextMatrix(11, 3)) + Val(MSS.TextMatrix(12, 3)) + (Val(MSS.TextMatrix(14, 3))) + Val(MSS.TextMatrix(15, 3)) - Val(MSS.TextMatrix(16, 3))) - (Val(MSS.TextMatrix(17, 3))))
MSS.Row = 12
MSS.Col = 1

End Sub
Private Sub Command2_Click()
If CHECKTB("inl.dbf,inv.dbf,lnv.dbf,prd.dbf,brd.dbf,vch.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
M_GRS_SAL = Round(M_GRS_SAL, 1)
M_GRS_SAL = convert(M_GRS_SAL)
M_FRE_QTY = Round(M_FRE_QTY, 1)
M_FRE_QTY = convert(M_FRE_QTY)
MSS.TextMatrix(1, 1) = M_GRS_SAL ' "Sales"
MSS.TextMatrix(2, 1) = M_FRE_QTY ' "Free Qty    (-)"
MSS.TextMatrix(7, 1) = M_DIS_AMT '"Discount    (-)"
MSS.TextMatrix(8, 1) = M_TAX_AMT '"Tax/Others  (+)"
MSS.TextMatrix(9, 1) = M_CRT_AMT 'M_NET_SAL '"Net Sales      "
MSS.TextMatrix(10, 1) = M_CAS_AMT '"Cash          "
MSS.TextMatrix(11, 1) = M_CHQ_AMT '"Cheques       "
MSS.TextMatrix(14, 1) = convert(Val(M_EXP_AMT) + Val(M_COM_AMT)) '"Exp/Comsn  (-)"
MSS.TextMatrix(15, 1) = m_dep_ref '"Deposit Refunded"
MSS.TextMatrix(16, 1) = m_dep_rcd '"Deposit Received"
MSS.TextMatrix(17, 1) = M_OLD_CRD '"Old Dues"
MSS.TextMatrix(18, 1) = M_CUR_CRD '"Credit Sale"
Dim RSINL As New ADODB.Recordset
If Len(Trim(M_INV_IDY)) = 0 Then
    MsgBox "Select Van Challan For Settlement Variance Generation ", vbInformation, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
End If
RSINL.Open "SELECT SUM(SPD_AMT),SUM(CSD_AMT),SUM(TAX_AMT),SUM(EXD_AMT),SUM(PKG_CST), SUM(FRT_AMT) FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "' AND TXN_TYP = 'SAL'", con, adOpenKeyset, adLockPessimistic
If RSINL.RecordCount > 0 Then
    inldiscount = Val(RSINL(0) & "") + Val(RSINL(1) & "") + Val(RSINL(5) & "")
    InlTax = Val(RSINL(2) & "") + Val(RSINL(3) & "") + Val(RSINL(4) & "")
End If

' Code functionality partly disabled by PNR since INV taxes are no more shown separately in the final settlement part
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "SELECT SUM(DIS_AMT),SUM(ADD_AMT),SUM(TAX_AMT) FROM INV WHERE INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') AND STK_TYP IN('Salable','Free') AND TXN_TYP = 'SAL'", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        invdiscount = Val(RSINL(0) & "") + Val(RSINL(1) & "")
        'invtax = Val(RSINL(2) & "")
    End If

If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(FRT_AMT) FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "' and rtrim(ltrim(TXN_RM2))<>'Scheme Amount' AND TXN_TYP = 'SAL'", con, adOpenKeyset, adLockPessimistic
If RSINL.RecordCount > 0 Then inldiscount = Val(inldiscount) ' + Val(RSINL(0) & "")
MSS.TextMatrix(7, 2) = convert(Val(inldiscount) + Val(invdiscount))
'MSS.TextMatrix(7, 2) = Val(invdiscount) 'Modified by Ramanesh & Radhika
MSS.TextMatrix(8, 2) = convert(Val(InlTax) + Val(InvTax))
'grossamt
'If RSINL.State = 1 Then RSINL.Close
'RSINL.Open "SELECT * FROM INV,BRD,PRD WHERE INV.BAT_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "')", CON, adOpenKeyset, adlockpessimistic
'If RSINL.RecordCount > 0 Then
'    If Trim(RSINL!stk_typ) = "Salable" Then
'        RSamt = Val(RSINL!COL_AMT) + Val(RSINL!TAX_AMT) + RSamt
'    End If
'    If Trim(RSINL!stk_typ) = "Free" Then
'        RSamt = RSINL! + RSamt
'    End If
'End If
'
'
Dim FS_INL As New ADODB.Recordset
Dim FS_INV As New ADODB.Recordset
Dim FS_PRD As New ADODB.Recordset
Dim M_SKU_VAL, M_SKU_RAT, M_CRT_RAT, M_SKU_TAX, M_CRT_TAX, M_INC_LES, P_SAL_TAX As Double
Dim T_GRS_AMT, F_GRS_AMT, B_GRS_AMT As Double

T_GRS_AMT = 0#
F_GRS_AMT = 0#
B_GRS_AMT = 0#

If FS_PRD.State = 1 Then FS_PRD.Close
FS_PRD.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
If FS_INL.State = 1 Then FS_INL.Close
FS_INL.Open "SELECT * FROM INV, BRD WHERE INV.BAT_IDY = BRD.BAT_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') AND INV.STK_TYP IN('Salable','TakeBack','Damage','Free','Breakage','Leakage') AND TXN_TYP = 'SAL'", con, adOpenKeyset, adLockPessimistic
If FS_INL.EOF = False Then
    FS_INL.MoveFirst
End If

    If FS_INL.RecordCount > 0 Then
        While Not FS_INL.EOF
            
            M_SKU_RAT = 0
            M_CRT_RAT = 0
            M_SKU_TAX = 0
            M_CRT_TAX = 0
            M_SKU_VAL = 0
            M_GRS_AMT = 0
            P_SAL_TAX = 0
            P_LST_RAT = 0
            P_LST_DIS = 0
            x_SUB_UNT = 0
            M_INC_LES = 0
            
            If FS_PRD.RecordCount > 0 Then FS_PRD.MoveFirst
                FS_PRD.Find "PRD_IDY='" & Trim(FS_INL("PRD_IDY")) & "'"
                If FS_PRD.EOF = False Then
                    P_SAL_TAX = Val(FS_PRD("SAL_TAX"))
                    x_SUB_UNT = Val(FS_PRD("SUB_UNT"))
                    If IsNull(FS_PRD("LST_RAT")) Or IsNull(FS_PRD("LST_DIS")) Then
                        P_LST_RAT = 0#
                        P_LST_DIS = 0#
                    Else
                        P_LST_RAT = Val(FS_PRD("LST_RAT"))
                        P_LST_DIS = Val(FS_PRD("LST_DIS"))
                    End If
                End If
        
            If Trim(FS_INL("STK_TYP")) = "Salable" Or Trim(FS_INL("STK_TYP")) = "Free" Then
                If Val(FS_INL("SUB_UNT")) > 0 Then
                    M_SKU_RAT = Val(FS_INL("PRD_SDR")) / Val(FS_INL("SUB_UNT"))
                Else
                    M_SKU_RAT = Val(FS_INL("PRD_SDR")) / 1
                End If
                'M_CRT_RAT = Val(FS_INL("LST_RAT")) / Val(FS_INL("SUB_UNT"))
                M_CRT_RAT = (P_LST_RAT / x_SUB_UNT)
                If IsNull(Val(FS_INL("S_INC_DIS"))) = False Then
                    If Val(FS_INL("S_INC_DIS")) <> 0 Then
                        M_INC_LES = Val(FS_INL("S_INC_DIS")) / Val(FS_INL("SUB_UNT"))
                    End If
                Else
                    M_INC_LES = 0#
                End If
                M_SKU_TAX = (((M_SKU_RAT - Val(M_INC_LES)) * P_SAL_TAX) / 100)
                'M_CRT_TAX = ((M_CRT_RAT * Val(FS_INL("LST_DIS"))) / 100)
                M_CRT_TAX = ((M_CRT_RAT * P_LST_DIS) / 100)
                M_SKU_VAL = (M_SKU_RAT + M_CRT_RAT + M_SKU_TAX + M_CRT_TAX)
                M_GRS_AMT = (M_SKU_VAL * Val(FS_INL("TOT_PCS")))
                T_GRS_AMT = (M_GRS_AMT + T_GRS_AMT)
            End If
            If Trim(FS_INL("STK_TYP")) = "Breakage" Or Trim(FS_INL("STK_TYP")) = "Leakage" Or Trim(FS_INL("STK_TYP")) = "TakeBack" Or Trim(FS_INL("STK_TYP")) = "Damage" Then
                M_SKU_RAT = Val(FS_INL("PRD_SDR")) / Val(FS_INL("SUB_UNT"))
                'M_CRT_RAT = Val(FS_INL("LST_RAT")) / Val(FS_INL("SUB_UNT"))
                M_CRT_RAT = (P_LST_RAT / x_SUB_UNT)
                M_INC_LES = Val(FS_INL("S_INC_DIS")) / Val(FS_INL("SUB_UNT"))
                M_SKU_TAX = (((M_SKU_RAT - Val(M_INC_LES)) * P_SAL_TAX) / 100)
                'M_CRT_TAX = ((M_CRT_RAT * Val(FS_INL("LST_DIS"))) / 100)
                M_CRT_TAX = ((M_CRT_RAT * P_LST_DIS) / 100)
                M_SKU_VAL = (M_SKU_RAT + M_CRT_RAT + M_SKU_TAX + M_CRT_TAX)
                M_GRS_AMT = (M_SKU_VAL * Val(FS_INL("TOT_PCS")))
                B_GRS_AMT = (M_GRS_AMT + B_GRS_AMT)
            End If
            If Trim(FS_INL("STK_TYP")) = "Free" Then
                M_SKU_RAT = Val(FS_INL("PRD_SDR")) / Val(FS_INL("SUB_UNT"))
                'M_CRT_RAT = Val(FS_INL("LST_RAT")) / Val(FS_INL("SUB_UNT"))
                M_CRT_RAT = (P_LST_RAT / x_SUB_UNT)
                M_INC_LES = Val(FS_INL("S_INC_DIS")) / Val(FS_INL("SUB_UNT"))
                M_SKU_TAX = (((M_SKU_RAT - Val(M_INC_LES)) * P_SAL_TAX) / 100)
                'M_CRT_TAX = ((M_CRT_RAT * Val(FS_INL("LST_DIS"))) / 100)
                M_CRT_TAX = ((M_CRT_RAT * P_LST_DIS) / 100)
                M_SKU_VAL = (M_SKU_RAT + M_CRT_RAT + M_SKU_TAX + M_CRT_TAX)
                M_GRS_AMT = (M_SKU_VAL * Val(FS_INL("TOT_PCS")))
                F_GRS_AMT = (M_GRS_AMT + F_GRS_AMT)
            End If
            
            
            FS_INL.MoveNext
            
        Wend
    End If
    MSS.TextMatrix(1, 2) = convert(Round(Val(T_GRS_AMT - B_GRS_AMT), 1))
    MSS.TextMatrix(2, 2) = convert(Round(Val(F_GRS_AMT), 1))
    
''''''''''If RSINL.State = 1 Then RSINL.Close
''''''''''RSINL.Open "SELECT sum(((prd_SDR + inv.lst_rat)/ sub_unt) * tot_pcs),stk_typ FROM INV,BRD WHERE BRD.BAT_IDY = INV.BAT_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') AND INV.STK_TYP IN('Salable','TakeBack','Damage','Free','Breakage','Leakage')group by stk_typ", CON, adOpenKeyset, adlockpessimistic
''''''''''RSamt = 0
''''''''''If RSINL.RecordCount > 0 Then
''''''''''    For A = 1 To RSINL.RecordCount
''''''''''        If Trim(RSINL!stk_typ) = "Salable" Then
''''''''''            RSamt = Val(RSINL(0) & "") + RSamt
''''''''''        ElseIf Trim(RSINL!stk_typ) = "Free" Then
''''''''''            RSamt = Val(RSINL(0) & "") + RSamt
''''''''''        ElseIf Trim(RSINL!stk_typ) = "Damage" Then
''''''''''            RSamt = RSamt - Val(RSINL(0) & "")
''''''''''        ElseIf Trim(RSINL!stk_typ) = "TakeBack" Or Trim(RSINL!stk_typ) = "Breakage" Or Trim(RSINL!stk_typ) = "Leakage" Then
''''''''''            RSamt = RSamt - Val(RSINL(0) & "")
''''''''''        End If
''''''''''        RSINL.MoveNext
''''''''''    Next
''''''''''    MSS.TextMatrix(1, 2) = convert(Val(RSamt))
''''''''''End If
''''''''''
'''''''''''freeqty
''''''''''If RSINL.State = 1 Then RSINL.Close
''''''''''RSINL.Open "SELECT sum(((BRD.PRD_SDR + INV.LST_RAT)/INV.SUB_UNT) * INV.TOT_PCS) FROM INV,BRD WHERE INV.BAT_IDY = BRD.BAT_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') AND STK_TYP IN('Free') and sch_idy IS NULL", CON, adOpenKeyset, adlockpessimistic
''''''''''If RSINL.RecordCount > 0 Then MSS.TextMatrix(2, 2) = convert(Val(RSINL(0)))
'''''''''''schqty
''''''''''If RSINL.State = 1 Then RSINL.Close
''''''''''RSINL.Open "SELECT * FROM lnv,brd,prd WHERE lnv.INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') and lev_des = brd.bat_idy and brd.prd_idy = prd.prd_idy and lnv.txn_typ = 'FRE'", CON, adOpenKeyset, adlockpessimistic
''''''''''If RSINL.RecordCount > 0 Then
''''''''''    For A = 1 To RSINL.RecordCount
''''''''''        b = Val(b) + (((Val(RSINL!PRD_SDR) + Val(RSINL!LST_RAT)) / Val(RSINL!SUB_UNT)) * Val(RSINL!INV_GRS))
''''''''''        RSINL.MoveNext
''''''''''
''''''''''    Next
''''''''''    MSS.TextMatrix(5, 2) = convert(Val(b))
''''''''''    If RSINL.State = 1 Then RSINL.Close
''''''''''    'RSINL.Open "SELECT sum(lnv_amt) FROM lnv WHERE lnv.INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') and (len(alltrim(lev_des))=0 or inv_grs = 0) ", CON, adOpenKeyset, adlockpessimistic
''''''''''    RSINL.Open "SELECT sum(frt_amt) FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adlockpessimistic 'Modified by Ramanesh & Radhika
''''''''''    If RSINL.RecordCount > 0 Then C = Val(RSINL(0) & "")
''''''''''    MSS.TextMatrix(6, 2) = convert(Val(C))
''''''''''End If


'CASH
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH,ACT WHERE VCH.DBT_IDY = ACT.ACT_IDY AND ACT.GRP_CLF LIKE 'Cash Book' and inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "')", con, adOpenKeyset, adLockPessimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(10, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH,ACT WHERE VCH.DBT_IDY = ACT.ACT_IDY AND ACT.GRP_CLF LIKE 'Bank Book' and inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "')", con, adOpenKeyset, adLockPessimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(11, 2) = convert(Val(RSINL(0) & ""))

If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(NET_AMT-RCD_AMT) FROM IRT WHERE INV_IDY IN (SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') AND (NET_AMT - RCD_AMT) > 0 AND TXN_TYP = 'SAL'", con, adOpenKeyset, adLockPessimistic
' The Above Query brings the Total Net Sale for that particular Van Challan with all positive amounts

If RSINL.RecordCount > 0 Then MSS.TextMatrix(18, 2) = convert(Val(RSINL(0) & "")) ' Final Reco Credit Amount

If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(NET_AMT-RCD_AMT) FROM IRT WHERE INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') AND (NET_AMT-RCD_AMT) < 0 AND TXN_TYP = 'SAL'", con, adOpenKeyset, adLockPessimistic
' The Above Query brings Total Receipt Amount for that particular Van Challan

If RSINL.RecordCount > 0 Then MSS.TextMatrix(12, 2) = convert((Val(RSINL(0)) & ""))   ' Final Reco Excess Amount
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT sum(dep_amt) FROM SNV WHERE TXN_TYP = 'SAL' AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "')", con, adOpenKeyset, adLockPessimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(16, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT sum(dep_amt) FROM SNV WHERE TXN_TYP = 'DPI' AND INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(16, 2) = Val(MSS.TextMatrix(16, 2)) + convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT sum(dep_amt) FROM SNV WHERE TXN_TYP = 'DPO' AND INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If RSINL.RecordCount > 0 Then MSS.TextMatrix(15, 2) = convert(Val(RSINL(0) & ""))
SETCALC
End Sub

Private Sub Command3_Click()
If Len(Trim(M_INV_IDY)) > 0 Then
    Dim rs As New ADODB.Recordset
    con.Execute "delete from stl"
    rs.Open "SELECT * FROM STL", con, adOpenKeyset, adLockPessimistic
    For A = 1 To MSS.Rows - 1
    If MSS.RowHeight(A) > 0 Then
        rs.AddNew
        rs!TXN_RMK = Trim(MSS.TextMatrix(A, 0))
        rs!rct_Qty = Val(MSS.TextMatrix(A, 1))
        rs!ISS_QTY = Val(MSS.TextMatrix(A, 2))
        rs!INV_IDY = Trim(M_INV_IDY)
        rs!prd_idy = ""
        rs.Update
    End If
    Next
    rs.Close
    OpenReport CURDIR & "\DOS_VCH.RP1", "Van Settlement", 0, "", 6
Else
    MsgBox "Select Van Challan for Settlement Report Print ", vbInformation, "DAS Version 4.0"
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
If Trim(M_INV_IDY) <> "" Then
    DIR_FLD_UPD
    OpenReport CURDIR & "\DOS_VCH.RP1", "Van MTF Gate Pass", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "' GROUP BY TRP_IDY", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        While REC_SET.EOF = False
            If Val(REC_SET("TRP_IDY")) > 3 Then
                OpenReport CURDIR & "\DOS_VCH.RP1", "Van MTF Gate Pass Seq", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
            End If
            REC_SET.MoveNext
        Wend
        REC_SET.Close
    Else
        REC_SET.Close
    End If
Else
    MsgBox "Select Van Challan for Gate Pass Print !", vbInformation, "DAS Version 4.0"
    CANCELRECORD
End If
End Sub
Private Sub Command6_Click()
Verify_Glass_Variance   'Module separately made by PNR for autoclick feature as per Prince Chopra input
End Sub
Private Sub Verify_Glass_Variance()
If Len(Trim(M_INV_IDY)) = 0 Then
    MsgBox "Select Van Challan ", vbInformation, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("snv.dbf,shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
With MS1
.Rows = 2
.Rows = ms2.Rows
For A = 2 To ms2.Rows - 1
    .TextMatrix(A, 0) = ms2.TextMatrix(A, 0)
    .TextMatrix(A, 1) = ms2.TextMatrix(A, 1)
    .TextMatrix(A, 2) = QTY(ms2.TextMatrix(A, 2))
    .TextMatrix(A, 10) = ms2.TextMatrix(A, 11)
    .TextMatrix(A, 11) = ms2.TextMatrix(A, 10)
    .TextMatrix(A, 3) = Val(ms2.TextMatrix(A, 5)) + Val(ms2.TextMatrix(A, 12))
    MS5 = MTS(Val(ms2.TextMatrix(A, 6)), Val(QTY(ms2.TextMatrix(A, 11))))
    MS12 = MTS(Val(ms2.TextMatrix(A, 13)), Val(QTY(ms2.TextMatrix(A, 11))))
    .TextMatrix(A, 4) = QTY(STM(Val(MS5) + Val(MS12), Val(ms2.TextMatrix(A, 11))))
    .TextMatrix(A, 9) = ms2.TextMatrix(A, 9)
Next
.SetFocus
'.Row = 1
.Col = 1
Dim rs As New ADODB.Recordset
rs.Open "SELECT SUM(R_SHL_NOS),SUM(R_PCK_NOS),SNV.SHL_IDY,PRD_PCK,PRD_UPC FROM SNV WHERE TXN_TYP = 'SAL' AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "') GROUP BY SNV.SHL_IDY,PRD_PCK,PRD_UPC", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 2 To .Rows - 1
        If Trim(rs!PRD_PCK & "") = Trim(.TextMatrix(b, 9)) And Val(rs!prd_upc) = Val(.TextMatrix(b, 10)) And Trim(rs!shl_idy) = Trim(.TextMatrix(b, 11)) Then
'            .TextMatrix(b, 5) = Val(rs(0) & "")    'COMMENTED BY PNR PROCESS CHANGE AS PER PRINCE CHOPRA
'            .TextMatrix(b, 6) = Val(rs(1) & "")    'COMMENTED BY PNR PROCESS CHANGE AS PER PRINCE CHOPRA
            GoTo n
        End If
    Next
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = GetValue("SHL_NME", "SHL WHERE SHL_IDY = '" & Trim(rs!shl_idy) & "'") & "-" & Trim(rs!PRD_PCK) & Val(rs!prd_upc) 'To be Checked in QUERY PNR
'    .TextMatrix(.Rows - 1, 5) = (Val(rs(0) & ""))
'    .TextMatrix(.Rows - 1, 6) = Val(rs(1) & "")
    .TextMatrix(.Rows - 1, 5) = 0
    .TextMatrix(.Rows - 1, 6) = 0
    .TextMatrix(.Rows - 1, 9) = Trim(rs(3) & "")
    .TextMatrix(.Rows - 1, 10) = Val(rs(4) & "")
    .TextMatrix(.Rows - 1, 11) = Trim(rs!shl_idy)
n:
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(d_SHL_NOS),SUM(d_PCK_NOS),SNV.SHL_IDY,PRD_PCK,PRD_UPC,TXN_TYP FROM SNV WHERE TXN_TYP in('AIN','AOU','DPI','DPO') AND INV_IDY = '" & Trim(M_INV_IDY) & "' GROUP BY SNV.SHL_IDY,PRD_PCK,PRD_UPC,TXN_TYP", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 2 To .Rows - 1
        If Trim(rs!PRD_PCK & "") = Trim(.TextMatrix(b, 9)) And Val(rs!prd_upc) = Val(.TextMatrix(b, 10)) Then
            If rs!TXN_TYP = "AIN" Or rs!TXN_TYP = "DPI" Then
                .TextMatrix(b, 5) = Val(.TextMatrix(b, 5)) - Val(rs(0) & "") ' + MADE TO - BY PNR AS PER PRINCE CHOPRA
                .TextMatrix(b, 6) = Val(.TextMatrix(b, 6)) - Val(rs(1) & "") ' + MADE TO - BY PNR AS PER PRINCE CHOPRA
            ElseIf rs!TXN_TYP = "AOU" Or rs!TXN_TYP = "DPO" Then
                .TextMatrix(b, 5) = Val(.TextMatrix(b, 5)) + Val(rs(0) & "") ' - MADE TO + BY PNR AS PER PRINCE CHOPRA
                .TextMatrix(b, 6) = Val(.TextMatrix(b, 6)) + Val(rs(1) & "") ' - MADE TO + BY PNR AS PER PRINCE CHOPRA
            End If
            GoTo M
        End If
    Next
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = GetValue("SHL_NME", "SHL WHERE SHL_IDY = '" & Trim(rs!shl_idy) & "'") & "-" & Trim(rs!PRD_PCK) & Val(rs!prd_upc) 'To be Checked in QUERY PNR
    .TextMatrix(.Rows - 1, 5) = (Val(rs(0) & ""))
    .TextMatrix(.Rows - 1, 6) = Val(rs(1) & "")
    .TextMatrix(.Rows - 1, 9) = Trim(rs(3) & "")
    .TextMatrix(.Rows - 1, 10) = Val(rs(4) & "")
M:
    rs.MoveNext
Next

'On Error Resume Next
con.Execute "DELETE FROM RP1"
Dim SNV_SMRY As New ADODB.Recordset
SNV_SMRY.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic
SNV_SMRY.AddNew
SNV_SMRY("CHR_ONE") = Trim(M_INV_IDY)
For b = 2 To .Rows - 1
    .TextMatrix(b, 6) = QTY(STM(Val(.TextMatrix(b, 6)), Val(.TextMatrix(b, 10))))
    .TextMatrix(b, 7) = Val(.TextMatrix(b, 1)) - (Val(.TextMatrix(b, 5)) + Val(.TextMatrix(b, 3))) 'Final Variance Formula = Net Load - (RGB Load In + COL Entered) As per Prince and modified by PNR
    .TextMatrix(b, 8) = QTY(STM(Val(MTS(Val(.TextMatrix(b, 2)), Val(.TextMatrix(b, 10))) - (MTS(Val(.TextMatrix(b, 6)), Val(.TextMatrix(b, 10))) + MTS(Val(.TextMatrix(b, 4)), Val(.TextMatrix(b, 10))))), Val(.TextMatrix(b, 10))))
'    SNV_SMRY.AddNew
'    SNV_SMRY("CHR_ONE") = Trim(M_INV_IDY)
'    SNV_SMRY("CHR_TWO") = Left(Trim(.TextMatrix(b, 0)), 15)
'    If (Len(Trim(.TextMatrix(b, 0))) - 15) > 0 Then
'        SNV_SMRY("CHR_THR") = Right(Trim(.TextMatrix(b, 0)), (Len(Trim(.TextMatrix(b, 0))) - 15))
'    End If
    SNV_SMRY("NUM_ONE") = Val(SNV_SMRY("NUM_ONE")) + Val(.TextMatrix(b, 1))
    SNV_SMRY("NUM_TWO") = Val(SNV_SMRY("NUM_TWO")) + Val(.TextMatrix(b, 2))
    SNV_SMRY("NUM_THR") = Val(SNV_SMRY("NUM_THR")) + Val(.TextMatrix(b, 3))
    SNV_SMRY("NUM_FOR") = Val(SNV_SMRY("NUM_FOR")) + Val(.TextMatrix(b, 4))
    SNV_SMRY("NUM_FIV") = Val(SNV_SMRY("NUM_FIV")) + Val(.TextMatrix(b, 5))
    SNV_SMRY("NUM_SIX") = Val(SNV_SMRY("NUM_SIX")) + Val(.TextMatrix(b, 6))
    SNV_SMRY("NUM_SEV") = Val(SNV_SMRY("NUM_SEV")) + Val(.TextMatrix(b, 7))
    SNV_SMRY("NUM_EIG") = Val(SNV_SMRY("NUM_EIG")) + Val(.TextMatrix(b, 8))
    SNV_SMRY.Update
Next
''SNV_SMRY.Close
Set SNV_SMRY = Nothing
If .Rows >= 3 Then .Row = 2
.SetFocus
End With
End Sub

Private Sub Command7_Click()
If Trim(M_INV_IDY) <> "" Then
    SAL_QRY = "INL->ORD_IDY = '" & Trim(M_INV_IDY) & "' .AND. INL->TXN_TYP = 'SAL'"
    updatevan
    OpenReport CURDIR & "\REP_LIB.RP1", "Van Reconcile Cases", 0, SAL_QRY, 6
Else
    MsgBox "Select Van Challan for Reconcile Report Print !", vbInformation, "DAS Version 4.0"
    CANCELRECORD
End If
End Sub

Private Sub Command8_Click()
If Trim(M_INV_IDY) <> "" Then
    OpenReport CURDIR & "\DOS_VCH.RP1", "Sale Estimate", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
Else
    MsgBox "Select Van Challan for Sale Estimate Print !", vbInformation, "DAS Version 4.0"
    CANCELRECORD
End If
End Sub

Private Sub Command9_Click()
Shell "C:\PROGRAM FILES\INTERNET EXPLORER\iexplore.exe " & CURDIR & "\SCH_SKP.TXT", vbMaximizedFocus
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command4_Click
End If
End Sub

Private Sub LINDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If M_INV_DAT > LINDATE Then
        MsgBox "Invalid Return Date!", vbCritical
        LINDATE.SetFocus
    Else
        M_INV_DAT.SetFocus
    End If
End If
End Sub

Private Sub LINDATE_Validate(cancel As Boolean)
    If M_INV_DAT > LINDATE Then
        MsgBox "Invalid Return Date!", vbCritical
        cancel = True
    End If
End Sub
Private Sub list3_DblClick()
List3_KeyPress 13
End Sub
Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list3.ListItems.count > 0 Then
        M_CST_IDY = list3.SelectedItem.ListSubItems(1)
        M_CST_IDY_keydown 13, 0
        list3.Visible = False
    End If
End If
End Sub
Private Sub exitrecord()
Unload Me
Set Form = Nothing
End Sub

Private Sub list2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If List2.SortOrder = lvwAscending Then
List2.SortOrder = lvwDescending
ElseIf List2.SortOrder = lvwDescending Then
List2.SortOrder = lvwAscending
End If
List2.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub list7_DblClick()
List7_keypress 13
End Sub
Private Sub list7_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If List7.SortOrder = lvwAscending Then
    List7.SortOrder = lvwDescending
ElseIf List7.SortOrder = lvwDescending Then
    List7.SortOrder = lvwAscending
End If
    List7.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub list2_DblClick()
list2_keypress 13
End Sub

Private Sub list3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list3.SortOrder = lvwAscending Then
list3.SortOrder = lvwDescending
ElseIf list3.SortOrder = lvwDescending Then
list3.SortOrder = lvwAscending
End If
list3.SortKey = ColumnHeader.Index - 1
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
    If list4.ListItems.count > 0 Then
        list4.Visible = False
        If list4.ListItems.count = 0 Then
            Option3.Visible = False
            Exit Sub
        End If
        Set LS1 = list4.SelectedItem.ListSubItems.Item(6)
        dispprd LS1.Text
        Check5.Visible = False
        Check6.Visible = False
    End If
End If
End Sub

Private Sub list2_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If List2.ListItems.count > 0 Then
        M_ACT_IDY = List2.SelectedItem.ListSubItems.Item(1)
        M_ACT_NME = List2.SelectedItem
        SSTab1.Tab = 0
        List2.Visible = False
        TRP_NOS.SetFocus
    End If
End If
End Sub
Private Sub List7_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If List7.ListItems.count > 0 Then
    List7.Visible = False
    cmbRoute.Enabled = True
    cmbRoute.SetFocus
'    SSTab1.Tab = 1
'    MS.SetFocus
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
    REC_SET.Open "select * from prd,BRD where BRD.PRD_IDY = PRD.PRD_IDY AND BRD.prd_idy like '" & MS.TextMatrix(MS.Row, 10) & "'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then MS.TextMatrix(MS.Row, 4) = amt(REC_SET("PRD_MRP") & "")
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
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' and ACT_NME LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' and ACT_AD4 LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
End If
List2.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(REC_SET("ACT_NME")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_IDY"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Account"
    List2.Visible = True
    List2.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Stock Issue Account Not Found, Create Stock Issue Account !", vbInformation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_ACT_IDY.SetFocus
End If
End If
REC_SET.Close
Set REC_SET = Nothing
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
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' AND ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from act where grp_clf like '" & TYP & "' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
End If
List7.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = List7.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Account"
    List7.Visible = True
    List7.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Stock Issue Account Not Found, Create Stock Issues Account !", vbInformation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_ACT_IDY.SetFocus
End If
End If
REC_SET.Close
Set REC_SET = Nothing
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
    REC_SET.Open "select ACT.*,MAR.MAR_NME,rdy.ROU_IDY,smn.smn_nme from act,MAR,RDY,smn where (ACT_TYP = 'BOT' OR ACT_TYP = 'SAL') and  act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.ACT_IDY = RDY.ACT_IDY AND RDY.SMN_IDY = SMN.SMN_IDY AND DBT_TYP='Y' and act.cst_typ = 'Van' GROUP BY ACT.ACT_IDY", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select ACT.*,MAR.MAR_NME,rdy.ROU_IDY,smn.smn_nme from act,MAR,RDY,smn where (act.ACT_TYP = 'BOT' OR ACT_TYP = 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY = MAR.MAR_IDY AND ACT.ACT_IDY = RDY.ACT_IDY AND RDY.SMN_IDY = SMN.SMN_IDY AND DBT_TYP='Y' and act.cst_typ = 'Van' GROUP BY ACT.ACT_IDY", con, adOpenKeyset, adLockPessimistic
End If
list3.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list3.ListItems.Add(, , Trim(REC_SET("ACT_NME")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_IDY"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("SMN_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("MAR_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("rou_idy") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Check5.Visible = False
    Check6.Visible = False
    Label49 = "Select Vehicle"
    list3.Visible = True
    list3.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Van - Vehicle Not Found, Create Van - Vehicle !", vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
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
    If Check3.Value = 1 And Trolley.Value = 1 Then
        MakeBill
    Else
        saverecord
    End If
ElseIf KeyCode = 38 Then
    M_CUR_CRD.SetFocus
End If
End Sub
Private Sub M_CUR_CRD_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_CUR_CRD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CRT_AMT.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_OLD_CRD.SetFocus
    calc
End If
End Sub

Private Sub M_CUR_CRD_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_GRS_SAL_GotFocus()
'SendKeys "{HOME}+{END}"
'Shl_Set_Click   'Added by PNR as per Prince Chopra input to update Shells
'Verify_Glass_Variance
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
'SendKeys "{HOME}+{END}"
End Sub

Private Sub M_CHQ_AMT_GotFocus()
'SendKeys "{HOME}+{END}"
End Sub
Private Sub M_CRT_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_CHQ_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_OLD_CRD.SetFocus
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
Shl_Set_Click   'Added by PNR as per Prince Chopra Input to update Shells
Verify_Glass_Variance
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
    SendKeys "{HOME}+{END}"
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    m_dep_rcd.SetFocus
    calc
End If
End Sub
Private Sub M_OLD_CRD_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub M_OLD_CRD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CUR_CRD.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CHQ_AMT.SetFocus
    calc
End If
End Sub

Private Sub M_OLD_CRD_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_TAX_AMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TAX_AMT_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    m_dep_ref.SetFocus
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
Private Sub M_dep_ref_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_dep_ref_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    m_dep_rcd.SetFocus
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_EXP_AMT.SetFocus
    calc
End If
End Sub

Private Sub M_dep_ref_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub m_dep_rcd_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub m_dep_rcd_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    M_CAS_AMT.SetFocus
    SendKeys "{HOME}+{END}"
    calc
ElseIf KeyCode = 38 Then
    Me.ActiveControl.Text = convert(Me.ActiveControl.Text)
    m_dep_ref.SetFocus
    calc
End If
End Sub

Private Sub m_dep_rcd_keypress(K As Integer)
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

Private Sub M_CST_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If CHECKTB("act.dbf,smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_CST_IDY)) & "' AND (ACT_TYP = 'BOT' OR ACT_TYP = 'SAL') AND DBT_TYP='Y' and cst_typ = 'Van'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        M_CST_IDY = Sincrement(M_CST_IDY)
        If Len(Trim(M_CST_IDY)) = 0 Then Exit Sub
        M_CST_NME = Trim(REC_SET("act_nme") & "")
        If IDN = True Then
            SMN = Split(GetSalesman(M_CST_IDY, M_INV_DAT), "|")
            If UBound(SMN) > 0 Then
                M_SMN_IDY = Trim(SMN(0))
                M_SMN_NME = Trim(SMN(1))
            End If
        End If
        If Len(Trim(M_INV_IDY)) = 0 Then If GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT) <> "" Then cmbRoute = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
        If REC_SET1.State = 1 Then REC_SET1.Close
        M_DBT_RAP = Trim(REC_SET("dbt_rap"))
        SSTab1.Tab = 0
        If K = 38 Then
            MNL_REF.SetFocus
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
If KeyCode = 13 Then MNL_REF.SetFocus
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
    If SSTab1.Tab = 0 Then M_INV_DAT.SetFocus
End If

Label23 = labdisp("idy_rsi")
SSTab1.TabEnabled(3) = False
TRP_NOS.Text = "1"
'setval
SETYEAR M_FRM_RSI1
txt1000 = "0"
txt500 = "0"
txt100 = "0"
txt50 = "0"
txt20 = "0"
txt10 = "0"
txt5 = "0"
txt2 = "0"
txt1 = "0"
txtchange = "0.00"
LINDATE.Enabled = Check3.Enabled
If rsrdy.State = 1 Then rsrdy.Close
rsrdy.Open "select rou_idy from rdy group by rou_idy", con, adOpenKeyset, adLockPessimistic
 cmbRoute.AddItem " "
While Not rsrdy.EOF
    cmbRoute.AddItem Trim(rsrdy("rou_Idy") & "")
    rsrdy.MoveNext
Wend
If cmbRoute.ListCount > 0 Then cmbRoute.ListIndex = 0
'calc
'discalc
'GRS_CAL
M_GRS_AMT = M_NET_SAL
For i = 1 To 11
MSS.Row = i
For J = 0 To 3
MSS.Col = J
MSS.CellBackColor = &HFFFFC0
Next
Next
MSS.Row = 12
MSS.Col = 1
End Sub
Private Sub INVOICELIST()
If CHECKTB("smn.dbf,van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RS_DAYCLS As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Trolley.Value = 1 Then
    qry = " and van_chk= 'P'"
Else
    qry = " and van_chk= 'V'"
End If
DateDayCls = Date
If Check1.Value = 1 Then
    If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY = 'SYS' AND DIR_IDX = 'MDC'") = True Then
        If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
        RS_DAYCLS.Open "SELECT MAX(DIR_DAT) FROM DIR WHERE DIR_IDY = 'DAY' .AND. TRIM(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
        If RS_DAYCLS.EOF = False Then
            RS_DAYCLS.MoveFirst
            DateDayCls = RS_DAYCLS(0)
        End If
    End If
    If DateDayCls = Date Then
        REC_SET.Open "select INV_IDY,INV_DAT,SMN_NME,SMN_AD4,VAN_CLS,ROU_IDY,MNL_REF from SMN,VAN where SMN_nme like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy " & qry & " GROUP BY INV_IDY ", con, adOpenKeyset, adLockPessimistic
    Else
        REC_SET.Open "select INV_IDY,INV_DAT,SMN_NME,SMN_AD4,VAN_CLS,ROU_IDY,MNL_REF from SMN,VAN where SMN_nme like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy " & qry & " and VAN.INV_DAT > CTOD('" & DateDayCls & "') GROUP BY INV_IDY ", con, adOpenKeyset, adLockPessimistic
    End If
Else
    If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY = 'SYS' AND DIR_IDX = 'MDC'") = True Then
        If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
        RS_DAYCLS.Open "SELECT MAX(DIR_DAT) FROM DIR WHERE DIR_IDY = 'DAY' .AND. TRIM(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
        If RS_DAYCLS.EOF = False Then
            RS_DAYCLS.MoveFirst
            DateDayCls = Format(RS_DAYCLS(0), "DD/MM/YYYY")
        End If
    End If
    If DateDayCls = Date Then
        REC_SET.Open "select INV_IDY,INV_DAT,SMN_NME,SMN_AD4,VAN_CLS,ROU_IDY,MNL_REF from SMN,VAN where SMN_ad4 like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy " & qry & " GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    Else
        REC_SET.Open "select INV_IDY,INV_DAT,SMN_NME,SMN_AD4,VAN_CLS,ROU_IDY,MNL_REF from SMN,VAN where SMN_ad4 like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy " & qry & " and VAN.INV_DAT > CTOD('" & DateDayCls & "') GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    End If
End If
list5.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("SMN_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ROU_IDY") & "")
        If REC_SET!vAN_CLS = "N" Then
            ls.ListSubItems.Add , , "Open"
        ElseIf REC_SET!vAN_CLS = "S" Then
            ls.ListSubItems.Add , , "Returns Posted"
        Else
            ls.ListSubItems.Add , , "Closed"
        End If
        ls.ListSubItems.Add , , Trim(REC_SET("MNL_REF") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Van Challan Document"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Van Challan Number Not Found!", vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If SSTab1.Tab = 2 And KeyCode = 117 Then
    txt1000.SetFocus
    txt1000.Alignment = 1
End If
If list1.Visible = False And list4.Visible = False And list5.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 82 And Shift = 2 Then RETURNSTOCKS
    If KeyCode = 65 And Shift = 2 Then
        If Toolbar1.Buttons(19).Enabled = True Then
            M_FRM_RSI_PRD.Show 1
            MS.Rows = 1
            MS.SetFocus
            For A = 1 To PrdCollection.count
                MS.AddItem ""
                MS.Row = A
                dispprd PrdCollection.Item(A)
            Next
            Do While PrdCollection.count > 0
                PrdCollection.Remove 1
            Loop
            If MS.Rows = 1 Then
                MS.AddItem ""
            End If
            SSTab1.Tab = 1
            MS.SetFocus
            Exit Sub
        End If
    End If
    If KeyCode = 83 And Shift = 2 Then
    If Check3.Value = 1 And Trolley.Value = 1 Then
        MakeBill
    Else
        saverecord
    End If
    End If
    If KeyCode = 88 And Shift = 2 Then exitrecord
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 67 And Shift = 2 Then ChallanCopy
    If KeyCode = 82 And Shift = 2 Then ChallanReturnCopy
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_RSI1.ActiveControl.name = "M_INV_IDY" Then
            INVOICELIST
        ElseIf M_FRM_RSI1.ActiveControl.name = "M_SMN_IDY" Then
            SALESMANLIST
        ElseIf M_FRM_RSI1.ActiveControl.name = "Text3" Then
            PRODUCTLIST
        ElseIf M_FRM_RSI1.ActiveControl.name = "M_CST_IDY" Then
            V_DLR_LST
        ElseIf M_FRM_RSI1.ActiveControl.name = "M_ACT_IDY" Then
            V_ACT_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_RSI1.ActiveControl Is TextBox Then M_FRM_RSI1.ActiveControl.Text = FUNKEY(KeyCode)
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
If list5.ListItems.count > 0 Then
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
    If list1.Visible = True Then
        list1.Visible = False
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
    ElseIf List2.Visible = True Then
        List2.Visible = False
        M_ACT_IDY.SetFocus
        SSTab1.Tab = 0
    ElseIf list3.Visible = True Then
        list3.Visible = False
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        Text3.Visible = False
        amount.Visible = False
        GRS_CAL
        discalc
        SSTab1.Tab = 2
    ElseIf list3.Visible = True Then
        list3.Visible = False
        M_CST_IDY.SetFocus
        SSTab1.Tab = 0
    Else
        If M_FRM_RSI1.ActiveControl.name = "M_INV_IDY" Then
            Unload Me
            Set Form = Nothing
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
SHLSETCLICK = False
MS.ColWidth(0) = 2200
MS.ColWidth(1) = 500
MS.ColWidth(2) = 1050
MS.ColWidth(3) = 1050
MS.ColWidth(4) = 1050
MS.ColWidth(5) = 1050
MS.ColWidth(6) = 1050
MS.ColWidth(7) = 1050
MS.ColWidth(8) = 1050
MS.ColWidth(9) = 1050
MS.ColWidth(10) = 1050
MS.ColWidth(11) = 1050
MS.ColWidth(12) = 1000
MS.ColWidth(13) = 1200
MS.ColWidth(14) = 1100
MS.ColWidth(15) = 1500
MS.ColWidth(16) = 0
MS.ColWidth(17) = 0
MS.ColWidth(18) = 1600
MS.ColWidth(19) = 1000
ms2.RowHeight(0) = 450
MS.TextMatrix(0, 0) = "Product Name"
SCl MS, 0
MS.ColAlignment(1) = 7
MS.TextMatrix(0, 1) = "Trip"
MS.TextMatrix(0, 2) = "Opg Stk"
MS.ColWidth(2) = 0 'Opg Stk is made to zero, column will not be displayed - Prince Chopra, PNR
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
MS1.RowHeight(0) = 450
MS.TextMatrix(0, 9) = "Net Load"
MS.ColAlignment(9) = 7
MS.TextMatrix(0, 10) = "Bill Stk"
MS.ColAlignment(10) = 7
MS.TextMatrix(0, 11) = "Variance"
MS.ColAlignment(11) = 7
MS.TextMatrix(0, 12) = "MRP"
MS.ColAlignment(12) = 7
MS.TextMatrix(0, 13) = "Amount"
MS.ColAlignment(13) = 7
MS.TextMatrix(0, 14) = "Product ID"
MS.TextMatrix(0, 15) = "Short Name"
MS.TextMatrix(0, 16) = "SUBUNTS"
MS.TextMatrix(0, 17) = "batchidy"
MS.TextMatrix(0, 18) = "Batch"
MS.TextMatrix(0, 19) = "Sale Rate"
MS.TextMatrix(0, 20) = "List Rate"
MS.TextMatrix(0, 21) = "List %+"
MS.TextMatrix(0, 22) = "Pur-Inc'tve"
MS.TextMatrix(0, 23) = "Sal-Inc'tve"
MS.TextMatrix(0, 6) = "Free. Stk"
MS.ColAlignment(19) = 7
MS.ColAlignment(6) = 7
MS.ColAlignment(20) = 7
MS.ColAlignment(21) = 7
MS.ColAlignment(22) = 7
MS.ColAlignment(23) = 7
SCl MS, 6
MS.TextMatrix(0, 7) = "Transit Breakage"
MS.ColAlignment(7) = 7
SCl MS, 7
MS.TextMatrix(0, 8) = "Transit Leakage"
MS.ColAlignment(8) = 7
SCl MS, 8
MS.Col = 0
If CHECKTB("dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If querycheck("DIR", "DIR_MSG", "C", " AND DIR_IDY = 'SYS' AND DIR_IDX = 'PCS'") = True Then
    Pieces.Value = 1
Else
    Pieces.Value = 0
End If
MSS.ColAlignment(1) = 7
MSS.ColAlignment(2) = 7
MSS.ColAlignment(3) = 7
MSS.TextMatrix(0, 1) = "Quick Setl-Van Stk"
MSS.TextMatrix(0, 2) = "Final Setl-Bill Entry"
MSS.TextMatrix(0, 3) = "  Variance"
MSS.ColWidth(0) = 2350
MSS.ColWidth(1) = 1850 '1700
MSS.ColWidth(2) = 1850 '1700
MSS.ColWidth(3) = 1200 '1500
MSS.Row = 0
MSS.Col = 1
MSS.CellFontBold = True
MSS.Col = 2
MSS.CellFontBold = True
MSS.Col = 3
MSS.CellFontBold = True
MSS.Rows = 1
MSS.Rows = 19
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
MSS.TextMatrix(12, 0) = "Cash Short-Excess"
MSS.TextMatrix(13, 0) = "Variance      "
MSS.TextMatrix(14, 0) = "Exp/Comsn  (-)"
MSS.TextMatrix(15, 0) = "Deposit Refunded"
MSS.TextMatrix(16, 0) = "Deposit Received"
MSS.TextMatrix(17, 0) = "Old Dues - Collection"
MSS.TextMatrix(18, 0) = "Credit Sales   "
MSS.RowHeight(13) = 0
MSS.Col = 0
For A = 1 To 18
MSS.Row = A
MSS.CellFontBold = True
Next
For i = 17 To 18
    MSS.Row = i
    For J = 0 To 3
        MSS.Col = J
        MSS.CellBackColor = &HFFFFC0
    Next
Next
MSS.Row = 12
MSS.Col = 1

With MS1
.Cols = 12
.Rows = 3
.FixedRows = 2
.FixedCols = 1
.MergeCells = flexMergeFree
.MergeRow(0) = True
.ColWidth(0) = 2500 'EARLIER VALUE 2200
.ColWidth(1) = 850
.ColWidth(2) = 850
.ColWidth(3) = 850
.ColWidth(4) = 850
.ColWidth(5) = 850
.ColWidth(6) = 850
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(9) = 0
.ColWidth(10) = 0
.ColWidth(11) = 0
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
.TextMatrix(0, 1) = "RGB NetLoad"
.Col = 2
.CellAlignment = 4
.TextMatrix(0, 2) = "RGB NetLoad"
.Row = 0
.Col = 3
.CellAlignment = 4
.TextMatrix(0, 1 + 2) = "Glass Collected"
.Col = 4
.CellAlignment = 4
.TextMatrix(0, 2 + 2) = "Glass Collected"
.Col = 5
.CellAlignment = 4
.TextMatrix(0, 3 + 2) = "RGB-COL Entry"
.Col = 6
.CellAlignment = 4
.TextMatrix(0, 4 + 2) = "RGB-COL Entry"
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

.TextMatrix(0, 0) = "Shell - Pack - UPC"
.TextMatrix(1, 1) = "Shells"
.TextMatrix(1, 2) = "Cases"
.TextMatrix(1, 3) = "Shells"
.TextMatrix(1, 4) = "Cases"
.TextMatrix(1, 5) = "Shells"
.TextMatrix(1, 6) = "Cases"
.TextMatrix(1, 7) = "Shells"
.TextMatrix(1, 8) = "Cases"
End With

With ms2
'.Cols = 13 Edited By Radhika
.Cols = 15
.Rows = 3
.FixedRows = 2
.Rows = 2
.FixedCols = 1
.MergeCells = flexMergeFree
.MergeRow(0) = True
.MergeCol(0) = True
.MergeRow(1) = True
For A = 0 To .Cols - 1
.Row = 0
.Col = A
.CellFontBold = True
.Row = 1
.Col = A
.CellFontBold = True
Next
.ColWidth(0) = 2500
.ColWidth(1) = 900
.ColWidth(2) = 900
.ColWidth(3) = 900
.ColWidth(4) = 900
.ColWidth(5) = 900
.ColWidth(6) = 900
.ColWidth(7) = 900
.ColWidth(8) = 900
.ColWidth(9) = 0
.ColWidth(10) = 0
.ColWidth(11) = 0
.ColWidth(12) = 900
.ColWidth(13) = 900
.ColWidth(14) = 0
.ColAlignment(1) = 7
.ColAlignment(2) = 7
.ColAlignment(3) = 7
.ColAlignment(4) = 7
.ColAlignment(5) = 7
.ColAlignment(6) = 7
.ColAlignment(7) = 7
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
.Col = 7
.CellAlignment = 4
.Col = 11
.CellAlignment = 4
.Col = 12
.CellAlignment = 4
.Col = 13
.CellAlignment = 4
SCl ms2, 12
SCl ms2, 13
SCl ms2, 3
SCl ms2, 4
.TextMatrix(0, 3) = "RGB Load-Out"
.TextMatrix(0, 4) = "RGB Load-Out"
.TextMatrix(0, 5) = "Glass Collected"
SCl ms2, 4
.TextMatrix(0, 6) = "Glass Collected"
SCl ms2, 5
.TextMatrix(0, 7) = "  Variance"
.TextMatrix(0, 8) = "  Variance"
.TextMatrix(0, 12) = "Breakage"
.TextMatrix(0, 13) = "Breakage"
.TextMatrix(1, 0) = "Shell - Pack - UPC"
.TextMatrix(0, 0) = "Shell - Pack - UPC"
.TextMatrix(0, 1) = "RGB NetLoad"
.TextMatrix(0, 2) = "RGB NetLoad"
.TextMatrix(1, 1) = "Shells"
.TextMatrix(1, 2) = "Cases"
.TextMatrix(1, 3) = "Shells"
.TextMatrix(1, 4) = "Cases"
.TextMatrix(1, 5) = "Shells"
.TextMatrix(1, 6) = "Cases"
.TextMatrix(1, 7) = "Shells"
.TextMatrix(1, 8) = "Cases"
.TextMatrix(1, 12) = "Shells"
.TextMatrix(1, 13) = "Cases"
End With
MSS.Row = 13
MSS.Col = 1
MSS.CellFontBold = True
MSS.Col = 2
MSS.CellFontBold = True
MSS.Col = 3
MSS.CellFontBold = True
For A = 1 To 18
MSS.TextMatrix(A, 1) = "0.00"
MSS.TextMatrix(A, 2) = "0.00"
MSS.TextMatrix(A, 3) = "0.00"
Next
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
    con.Execute "INSERT INTO PRT VALUES('PUSH CART')"
Else
    con.Execute "DELETE FROM PRT WHERE PRT_NME LIKE 'PUSH CART'"
End If
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Challan Number"
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
If M_INV_IDY = "" And K = 27 Then Unload Me
Set Form = Nothing
End Sub

Private Sub LIST5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If list5.ListItems.count > 0 Then
    If Label49 = "Select Van Challan" Then
        AddOpeningStock
        DIR_FLD_UPD
    Else
        M_INV_IDY = list5.SelectedItem
        DIR_FLD_UPD
        DISPDETAILS
    End If
End If
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
    If Check5.Visible = True Then Check5.Value = 0
    If Check6.Visible = True Then Check6.Value = 0
    'PRODUCTLIST
Else
    'PRODUCTLIST
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
    If Check5.Visible = True Then Check5.Value = 0
    If Check6.Visible = True Then Check6.Value = 0
    'PRODUCTLIST
Else
    'PRODUCTLIST
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
    con.Execute "insert into prt values('prdnum')"
    Else
    con.Execute "delete from prt where prt_nme like 'prdnum'"
    End If
    PRODUCTLIST
End If
End Sub

Private Sub MNL_REF_KEYDOWN(K As Integer, s As Integer)
If Trim(MNL_REF) = "NULLIFY" Then
    If MsgBox("Nullify this VanChallan ? All data will be reset to zeros ?", vbQuestion + vbOKCancel, "DAS Version 4.0 - Day Closure Confirmation") = vbOK Then
        ChallanNullify
    End If
End If
If K = 13 Then
    If Trim(MNL_REF) = "NULLIFY" Then
        CANCELRECORD
    Else
        M_CST_IDY.SetFocus
    End If
ElseIf K = 38 Then
    M_INV_DAT.SetFocus
End If
End Sub

Private Sub MNL_REF_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
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
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Bill Stock"
ElseIf MS.Col = 11 Then
    M_MDI.STS_BAR.Panels(1).Text = "Variance"
ElseIf MS.Col = 12 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf MS.Col = 13 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Amount"
ElseIf MS.Col = 14 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product ID"
ElseIf MS.Col = 15 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
ElseIf MS.Col = 16 Then
    M_MDI.STS_BAR.Panels(1).Text = "Batch"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Free Stock"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Breakage Stock"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Leakage Stock"
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
'    If Check3.Value = 1 Then
        If ms2.Col = 3 Or ms2.Col = 4 Or ms2.Col = 5 Or ms2.Col = 6 Or ms2.Col = 12 Or ms2.Col = 13 Then
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
 '   Else
'        If ms2.Col = 1 Then
'            txttemp = Val(ms2.Text)
'            If (K > 47 And K < 58) Or K = 45 Then
'                txtempshl = ""
'                txtempshl.MaxLength = 8
'                txtempshl.Visible = True
'                txtempshl.Top = ms2.Top + ms2.CellTop
'                txtempshl.Left = ms2.CellLeft + ms2.Left
'                txtempshl.Width = ms2.CellWidth
'                txtempshl.Height = ms2.CellHeight
'                txtempshl.Visible = True
'                txtempshl.SetFocus
'                txtempshl = Chr(K)
'                txtempshl.SelStart = Len(txtempshl)
'            Else
'                K = 0
'            End If
'        End If
'    End If
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
Dim BB(50) As Double
If Len(Trim(M_INV_IDY)) = 0 And MOD_REC = True Then
    MsgBox "Select Van Challan ", vbInformation, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
End If
SHLCAS = 0
Dim rs As New ADODB.Recordset
ms2.Rows = 2
If rs.State = 1 Then rs.Close
rs.Open "select * from snv,SHL where inv_idy = '" & Trim(M_INV_IDY) & "' and SNV.SHL_IDY = SHL.SHL_IDY AND  txn_typ = 'VAN'", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 2 To ms2.Rows - 1
        If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(rs!shl_idy) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!prd_upc) = Val(ms2.TextMatrix(b, 11)) Then
            ms2.TextMatrix(b, 3) = Val(rs!I_shl_nos & "")
            ms2.TextMatrix(b, 4) = QTY(STM(Val(rs!I_pck_nos & ""), ms2.TextMatrix(b, 11)))
            ms2.TextMatrix(b, 5) = Val(rs!R_shl_nos & "")
            ms2.TextMatrix(b, 6) = QTY(STM(Val(rs!R_pck_nos & ""), ms2.TextMatrix(b, 11)))
            GoTo l
        End If
    Next
    ms2.AddItem ""
    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(rs!shl_nme) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!prd_upc)
    ms2.TextMatrix(ms2.Rows - 1, 1) = 0
    ms2.TextMatrix(ms2.Rows - 1, 2) = 0
    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!prd_upc)
    ms2.TextMatrix(b, 3) = Val(rs!I_shl_nos & "")
    ms2.TextMatrix(b, 4) = QTY(STM(Val(rs!I_pck_nos & ""), ms2.TextMatrix(b, 11)))
    ms2.TextMatrix(b, 5) = Val(rs!R_shl_nos & "")
    ms2.TextMatrix(b, 6) = QTY(STM(Val(rs!R_pck_nos & ""), ms2.TextMatrix(b, 11)))
    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(rs!shl_idy & "")
    ms2.TextMatrix(ms2.Rows - 1, 12) = Trim(rs!BRK_SHL & "")
    ms2.TextMatrix(ms2.Rows - 1, 13) = Trim(Int(Val(rs!BRK_PCK) / Val(rs!prd_upc)) & "")
   
l:
    rs.MoveNext
Next
If rs.State = 1 Then rs.Close
rs.Open "select * from prd", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1

    If MS.TextMatrix(A, 0) <> "" Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            rs.Find "prd_idy='" & Trim(MS.TextMatrix(A, 14)) & "'"
            SHLID = GetValue("SHL_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            SHLNME = GetValue("SHL_NME", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            shlchk = GetValue("PRD_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            If rs.EOF = False Then
                For b = 2 To ms2.Rows - 1
                    If shlchk <> "" Then 'Added by PNR
                    If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(SHLID) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!SUB_UNT) = Val(ms2.TextMatrix(b, 11)) Then
                        If Pieces.Value = 0 Then
                            SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 7))), Val(MS.TextMatrix(A, 16)))
                            ms2.TextMatrix(b, 12) = Val(ms2.TextMatrix(b, 12)) + Int(STM(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16))))
                            ms2.TextMatrix(b, 13) = QTY(STM(Val(MS.TextMatrix(A, 7)) + MTS(ms2.TextMatrix(b, 13), Val(ms2.TextMatrix(b, 11))), Val(MS.TextMatrix(A, 14))))
                            CC = Val(ms2.TextMatrix(b, 3))
                            BB(b) = Val(MS.TextMatrix(A, 5))
                            e = (Val(MS.TextMatrix(A, 7)))
                            f = (Val(MS.TextMatrix(A, 8)))

                        Else
                            AA = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                            BB(b) = Val(BB(b)) + MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                            C = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
                            d = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
                            e = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
                            f = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
                            CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                            SHLCAS = QTY(STM(Val(AA) - (Val(BB(b)) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
                        End If
                            ms2.TextMatrix(b, 1) = Int(STM(Val(CC) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                            ms2.TextMatrix(b, 2) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                        GoTo n
                    End If
                    End If 'Added by PNR
                Next
                shlchk = GetValue("PRD_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'") 'Added by PNR
                If shlchk <> "" Then 'Added by PNR
                ms2.AddItem ""
                ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(SHLNME) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!SUB_UNT)
                ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
                ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(SHLID & "")
                ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!SUB_UNT)
                If Pieces.Value = 0 Then
                    SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 7))), Val(MS.TextMatrix(A, 16)))
                    ms2.TextMatrix(b, 12) = Int(Val(MS.TextMatrix(A, 7)))
                    ms2.TextMatrix(b, 13) = QTY(Val(MS.TextMatrix(A, 7)))
                    CC = Val(ms2.TextMatrix(b, 3))
                    BB(b) = Val(MS.TextMatrix(A, 5))
                    e = (Val(MS.TextMatrix(A, 7)))
                    f = (Val(MS.TextMatrix(A, 8)))
                Else
                    AA = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                    BB(b) = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                    C = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
                    d = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
                    e = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
                    f = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
                    CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                    SHLCAS = QTY(STM(Val(AA) - (Val(BB(b)) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
                End If
                    ms2.TextMatrix(b, 1) = Int(STM(Val(CC) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                    ms2.TextMatrix(b, 2) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) - (Val(BB(b)) + Val(d) + Val(f)), ms2.TextMatrix(b, 11)))
                End If
            End If
n:
        End If
    End If
Next

If rs.State = 1 Then rs.Close
rs.Open "select * from snv,SHL where inv_idy = '" & Trim(M_INV_IDY) & "' and SNV.SHL_IDY = SHL.SHL_IDY AND  txn_typ = 'VAN'", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 2 To ms2.Rows - 1
        If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(rs!shl_idy) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!prd_upc) = Val(ms2.TextMatrix(b, 11)) Then
            ms2.TextMatrix(b, 3) = Val(rs!I_shl_nos & "")
            ms2.TextMatrix(b, 4) = QTY(STM(Val(rs!I_pck_nos & ""), ms2.TextMatrix(b, 11)))
            ms2.TextMatrix(b, 5) = Val(rs!R_shl_nos & "")
            ms2.TextMatrix(b, 6) = QTY(STM(Val(rs!R_pck_nos & ""), ms2.TextMatrix(b, 11)))
            GoTo KR
        End If
    Next
    ms2.AddItem ""
    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(rs!shl_nme) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!prd_upc)
    ms2.TextMatrix(ms2.Rows - 1, 1) = 0
    ms2.TextMatrix(ms2.Rows - 1, 2) = 0
    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!prd_upc)
    ms2.TextMatrix(b, 3) = Val(rs!I_shl_nos & "")
    ms2.TextMatrix(b, 4) = QTY(STM(Val(rs!I_pck_nos & ""), ms2.TextMatrix(b, 11)))
    ms2.TextMatrix(b, 5) = Val(rs!R_shl_nos & "")
    ms2.TextMatrix(b, 6) = QTY(STM(Val(rs!R_pck_nos & ""), ms2.TextMatrix(b, 11)))
    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(rs!shl_idy & "")
KR:
    rs.MoveNext
Next

For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 7) = Val(ms2.TextMatrix(b, 1)) - Val(ms2.TextMatrix(b, 5))
    ms2.TextMatrix(b, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(b, 2)), Val(ms2.TextMatrix(b, 11))) - MTS(Val(ms2.TextMatrix(b, 6)), Val((ms2.TextMatrix(b, 11))))), Val((ms2.TextMatrix(b, 11)))))
Next

For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 7) = Val(ms2.TextMatrix(b, 1)) - (Val(ms2.TextMatrix(b, 5)) + Val(ms2.TextMatrix(b, 12)))
    ms2.TextMatrix(b, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(b, 2)), Val(ms2.TextMatrix(b, 11))) - (MTS(Val(ms2.TextMatrix(b, 13)), Val(ms2.TextMatrix(ms2.Row, 11))) + MTS(Val(ms2.TextMatrix(b, 6)), Val((ms2.TextMatrix(b, 11)))))), Val((ms2.TextMatrix(b, 11)))))
Next
If rs.State = 1 Then rs.Close
rs.Open "SELECT SUB_UNT,SHL_IDY,PRD_PCK,SHL_NME FROM PRD,SHL WHER PRD.PRD_IDY = SHL.PRD_IDY GROUP BY SUB_UNT,SHL_IDY,PRD_PCK", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 2 To ms2.Rows - 1
        If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(rs!shl_idy) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!SUB_UNT) = Val(ms2.TextMatrix(b, 11)) Then
            GoTo o:
        End If
    Next
    ms2.AddItem ""
    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(rs!shl_nme) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!SUB_UNT)
    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(rs!shl_idy & "")
    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!SUB_UNT)
o:
    rs.MoveNext
Next
If SSTab1.Tab = 5 Then
    ms2.SetFocus
    ms2.Col = 2
End If
End Sub
Private Sub Shl_Set_Click()
Dim CHK_MOR As Double
Dim CAS_UTS As Double
Dim CAS_STS As Double
Dim BB(50) As Double
CHK_MOR = 0
CAS_UTS = 0
CAS_STS = 0
SHLCAS = 0
If Len(Trim(M_INV_IDY)) = 0 And MOD_REC = True Then
    MsgBox "Select Van Challan ", vbInformation, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
End If
' *
Dim rs As New ADODB.Recordset
For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 1) = 0
    ms2.TextMatrix(b, 2) = 0
'    ms2.TextMatrix(b, 12) = 0
'    ms2.TextMatrix(b, 13) = 0
    If Check3.Value = 0 Then
'        If Val(ms2.TextMatrix(b, 3)) = 0 Then
            ms2.TextMatrix(b, 3) = 0
'        End If
'        If Val(ms2.TextMatrix(b, 4)) = 0 Then
            ms2.TextMatrix(b, 4) = 0
'        End If
    Else
'        If Val(ms2.TextMatrix(b, 3)) = 0 Then
            ms2.TextMatrix(b, 3) = 0
'        End If
'        If Val(ms2.TextMatrix(b, 4)) = 0 Then
            ms2.TextMatrix(b, 4) = 0
'        End If
    End If
Next

rs.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
    If MS.TextMatrix(A, 0) <> "" Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            rs.Find "prd_idy='" & Trim(MS.TextMatrix(A, 14)) & "'"
            SHLID = GetValue("SHL_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            SHLNME = GetValue("SHL_NME", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            shlchk = GetValue("PRD_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            If rs.EOF = False Then
                For b = 2 To ms2.Rows - 1
                    If shlchk <> "" Then 'Added by PNR
                        If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(SHLID) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!SUB_UNT) = Val(ms2.TextMatrix(b, 11)) Then
                            If Pieces.Value = 0 Then
                                SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 7))), Val(MS.TextMatrix(A, 16)))
                                ms2.TextMatrix(b, 12) = Val(ms2.TextMatrix(b, 12)) + Int(STM(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16))))
                                ms2.TextMatrix(b, 13) = QTY(STM(Val(MS.TextMatrix(A, 7)) + MTS(ms2.TextMatrix(b, 13), Val(ms2.TextMatrix(b, 11))), Val(MS.TextMatrix(A, 14))))
                                CC = Val(ms2.TextMatrix(b, 3))
                                BB(b) = Val(MS.TextMatrix(A, 5))
                                e = (Val(MS.TextMatrix(A, 7)))
                                f = (Val(MS.TextMatrix(A, 8)))
                            Else
                                AA = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                                BB(b) = Val(BB(b)) + MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                                C = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
                                d = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
                                e = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
                                f = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
                                'ms2.TextMatrix(b, 12) = Val(ms2.TextMatrix(b, 12)) + Int(Val(MS.TextMatrix(A, 7)))
                                'ms2.TextMatrix(b, 13) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 13)), Val(ms2.TextMatrix(b, 11))) + MTS(Val(MS.TextMatrix(A, 7)), Val(ms2.TextMatrix(b, 11))), Val(ms2.TextMatrix(b, 11))))
                                SHLCAS = QTY(STM(Val(AA) - (Val(BB(b)) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
                                CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                                ms2.TextMatrix(b, 1) = Int(STM(Val(CC) - Val(BB(b)), ms2.TextMatrix(b, 11)))
                                ms2.TextMatrix(b, 2) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) - (Val(BB(b))), ms2.TextMatrix(b, 11)))
                            End If
                            ms2.TextMatrix(b, 1) = Int(SHLCAS) + Val(ms2.TextMatrix(b, 1))
                            ms2.TextMatrix(b, 2) = QTY(STM(MTS(ms2.TextMatrix(b, 2), ms2.TextMatrix(b, 11)) + (Val(Val(AA) - (Val(BB(b)) + Val(e) + Val(f)))), ms2.TextMatrix(b, 11)))   ' MODIFIED

                            If Check3.Value = 0 Or Check3.Value = 1 Then  'MODIFIED BY PARVATHI
'                                If Val(ms2.TextMatrix(b, 3)) = 0 Then
                                    ms2.TextMatrix(b, 3) = Int(MS.TextMatrix(A, 4)) + Val(ms2.TextMatrix(b, 3))
'                                End If
'                                If Val(ms2.TextMatrix(b, 4)) = 0 Then
                                    ms2.TextMatrix(b, 4) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) + (Val(AA)), ms2.TextMatrix(b, 11)))
'                                End If
                                CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                            Else 'Part Added by PNR for checking the sequence as per Prince Chopra input
                                'If Val(ms2.TextMatrix(b, 3)) = 0 Then
                                    ms2.TextMatrix(b, 3) = Int(MS.TextMatrix(A, 4)) + Val(ms2.TextMatrix(b, 3))
                                'End If
                                'If Val(ms2.TextMatrix(b, 4)) = 0 Then
                                    ms2.TextMatrix(b, 4) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) + (Val(AA)), ms2.TextMatrix(b, 11)))
                                'End If
                                CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
 
                            End If
                            
                            'COMMENTED BY PKB
                            ms2.TextMatrix(b, 1) = Int(STM(Val(CC) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                            ms2.TextMatrix(b, 2) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                            
                            'ADDED BY PKB - 05/Jun/2005
'                            ms2.TextMatrix(b, 1) = Int(ms2.TextMatrix(b, 1) + Int((MS.TextMatrix(A, 11))))
'                            ms2.TextMatrix(b, 2) = QTY(ms2.TextMatrix(b, 2) + Int((MS.TextMatrix(A, 11))))
                            GoTo n
                        End If
                    End If 'Added by PNR
                Next
                
                shlchk = GetValue("PRD_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'") 'Added by PNR
                If shlchk <> "" Then 'Added by PNR
                    ms2.AddItem ""
                    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(SHLNME) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!SUB_UNT)
                    'ms2.TextMatrix(ms2.Rows - 1, 1) = Trim(RS!PRD_PCK) & " - " & Val(RS!sub_unt)
                    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
                    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(SHLID & "")
                    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!SUB_UNT)
                    If Pieces.Value = 0 Then
                        SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 7))), Val(MS.TextMatrix(A, 16)))
                        ms2.TextMatrix(b, 12) = Int(Val(MS.TextMatrix(A, 7)))
                        ms2.TextMatrix(b, 13) = QTY(Val(MS.TextMatrix(A, 7)))
                        CC = Val(ms2.TextMatrix(b, 3))
                        BB(b) = Val(MS.TextMatrix(A, 5))
                        e = (Val(MS.TextMatrix(A, 7)))
                        f = (Val(MS.TextMatrix(A, 8)))
                    Else
                        AA = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                        BB(b) = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                        C = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
                        d = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
                        e = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
                        f = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
                        CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                        SHLCAS = QTY(STM(Val(AA) - (Val(BB(b)) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
                    End If
                    If Check3.Value = 0 Then
'                        If Val(ms2.TextMatrix(b, 3)) = 0 Then
                            ms2.TextMatrix(b, 3) = Int(Val(MS.TextMatrix(A, 4))) + Val(ms2.TextMatrix(b, 3))
'                        End If
'                        If Val(ms2.TextMatrix(b, 4)) = 0 Then
                            ms2.TextMatrix(b, 4) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) + (Val(AA)), ms2.TextMatrix(b, 11)))
'                        End If
                        CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                    Else 'Part Added by PNR as per Prince Chopras inputs to check sequence
                        'If Val(ms2.TextMatrix(b, 3)) = 0 Then
                            ms2.TextMatrix(b, 3) = Int(Val(MS.TextMatrix(A, 4))) + Val(ms2.TextMatrix(b, 3))
                        'End If
                        'If Val(ms2.TextMatrix(b, 4)) = 0 Then
                            ms2.TextMatrix(b, 4) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) + (Val(AA)), ms2.TextMatrix(b, 11)))
                        'End If
                        CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                    End If
                    ms2.TextMatrix(b, 1) = Int(STM(Val(CC) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                    ms2.TextMatrix(b, 2) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                    
                End If
            End If
n:
        End If
    End If
Next

'Check to equal shells values to shells values of glass - PNR
For b = 2 To ms2.Rows - 1
    If Int(Val(ms2.TextMatrix(b, 3))) <> Int(Val(ms2.TextMatrix(b, 4))) Then
        ms2.TextMatrix(b, 3) = Int(Val(ms2.TextMatrix(b, 4)))               'Load Out
    End If
    If Int(Val(ms2.TextMatrix(b, 1))) <> Int(Val(ms2.TextMatrix(b, 2))) Then
        ms2.TextMatrix(b, 1) = Int(Val(ms2.TextMatrix(b, 2)))               'Net Load
    End If
Next


' Code added by PNR to auto increment Shells by 1 when Glass Count is more than half crate compared to Sub Units
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY = 'SYS' AND DIR_IDX = 'AIS'") = True Then
    For b = 2 To ms2.Rows - 1
        CHK_MOR = (Val(ms2.TextMatrix(b, 11)) / 2)
        CAS_UTS = Val(ms2.TextMatrix(b, 2))
        CAS_STS = Int(Val(ms2.TextMatrix(b, 2)))
        If ((Val(CAS_UTS) - Val(CAS_STS)) * 1000) > CHK_MOR Then
            ms2.TextMatrix(b, 1) = (Val(ms2.TextMatrix(b, 1)) + 1)
            ms2.TextMatrix(b, 3) = (Val(ms2.TextMatrix(b, 3)) + 1)
        End If
    Next
End If
' End of Code Added by PNR

For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 7) = Val(ms2.TextMatrix(b, 1)) - Val(ms2.TextMatrix(b, 5))
    ms2.TextMatrix(b, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(b, 2)), Val(ms2.TextMatrix(b, 11))) - MTS(Val(ms2.TextMatrix(b, 6)), Val((ms2.TextMatrix(b, 11))))), Val((ms2.TextMatrix(b, 11)))))
Next

For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 7) = Val(ms2.TextMatrix(b, 1)) - (Val(ms2.TextMatrix(b, 5)) + Val(ms2.TextMatrix(b, 12)))
    ms2.TextMatrix(b, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(b, 2)), Val(ms2.TextMatrix(b, 11))) - (MTS(Val(ms2.TextMatrix(b, 13)), Val(ms2.TextMatrix(ms2.Row, 11))) + MTS(Val(ms2.TextMatrix(b, 6)), Val((ms2.TextMatrix(b, 11)))))), Val((ms2.TextMatrix(b, 11)))))
Next
If rs.State = 1 Then rs.Close
rs.Open "SELECT SUB_UNT,SHL_IDY,PRD_PCK,SHL_NME FROM PRD,SHL WHER PRD.PRD_IDY = SHL.PRD_IDY GROUP BY SUB_UNT,SHL_IDY,PRD_PCK", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 2 To ms2.Rows - 1
        If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(rs!shl_idy) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!SUB_UNT) = Val(ms2.TextMatrix(b, 11)) Then
            GoTo o:
        End If
    Next
    ms2.AddItem ""
    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(rs!shl_nme) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!SUB_UNT)
    'ms2.TextMatrix(ms2.Rows - 1, 1) = Trim(RS!PRD_PCK) & " - " & Val(RS!sub_unt)
    'ms2.TextMatrix(ms2.Rows - 1, 2) = Trim(RS!PRD_PCK) & " - " & Val(RS!sub_unt)
    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(rs!shl_idy & "")
    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!SUB_UNT)
o:
    rs.MoveNext
Next

If SSTab1.Tab = 5 Then
    ms2.SetFocus
    ms2.Col = 2
End If
End Sub
Private Sub ShlSet_Click()
Dim BB(50) As Double
If Len(Trim(M_INV_IDY)) = 0 And MOD_REC = True Then
    MsgBox "Select Van Challan ", vbInformation, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
End If
SHLCAS = 0
If ms2.Rows > 2 Then
    SHLSETCLICK = True
Exit Sub
End If
Dim rs As New ADODB.Recordset
For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 1) = 0
    ms2.TextMatrix(b, 2) = 0
    ms2.TextMatrix(b, 13) = 0
    ms2.TextMatrix(b, 12) = 0
Next
rs.Open "select * from prd", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
    If MS.TextMatrix(A, 0) <> "" Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            rs.Find "prd_idy='" & Trim(MS.TextMatrix(A, 14)) & "'"
            SHLID = GetValue("SHL_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            SHLNME = GetValue("SHL_NME", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            shlchk = GetValue("PRD_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'")
            If rs.EOF = False Then
                For b = 2 To ms2.Rows - 1
                    If shlchk <> "" Then 'Added by PNR
                        If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(SHLID) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!SUB_UNT) = Val(ms2.TextMatrix(b, 11)) Then
                            If Pieces.Value = 0 Then
                                SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 7))), Val(MS.TextMatrix(A, 16)))
                                ms2.TextMatrix(b, 12) = Val(ms2.TextMatrix(b, 12)) + Int(STM(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16))))
                                ms2.TextMatrix(b, 13) = QTY(STM(Val(MS.TextMatrix(A, 7)) + MTS(ms2.TextMatrix(b, 13), Val(ms2.TextMatrix(b, 11))), Val(MS.TextMatrix(A, 14))))
                                CC = Val(ms2.TextMatrix(b, 3))
                                BB(b) = Val(MS.TextMatrix(A, 5))
                                e = (Val(MS.TextMatrix(A, 7)))
                                f = (Val(MS.TextMatrix(A, 8)))
                            Else
                                AA = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                                BB(b) = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                                C = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
                                d = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
                                e = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
                                f = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
                                CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                                'ms2.TextMatrix(b, 12) = Val(ms2.TextMatrix(b, 12)) + Int(Val(MS.TextMatrix(A, 7)))
                                'ms2.TextMatrix(b, 13) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 13)), Val(ms2.TextMatrix(b, 11))) + MTS(Val(MS.TextMatrix(A, 7)), Val(ms2.TextMatrix(b, 11))), Val(ms2.TextMatrix(b, 11))))
                                SHLCAS = QTY(STM(Val(AA) - (Val(BB(b)) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
                            End If
                            If Check3.Value = 0 Then
                                ms2.TextMatrix(b, 3) = Int(Val(MS.TextMatrix(A, 4))) + Val(ms2.TextMatrix(b, 3))
                                ms2.TextMatrix(b, 4) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) + (Val(AA)), ms2.TextMatrix(b, 11)))
                                CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                            End If
                            ms2.TextMatrix(b, 1) = Int(STM(Val(CC) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                            ms2.TextMatrix(b, 2) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                            GoTo n
                        End If
                    End If 'Added by PNR
                Next
                shlchk = GetValue("PRD_IDY", "SHL WHERE PRD_IDY = '" & Trim(rs!prd_idy) & "'") 'Added by PNR
                If shlchk <> "" Then 'Added by PNR
                    ms2.AddItem ""
                    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(SHLNME) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!SUB_UNT)
                    'ms2.TextMatrix(ms2.Rows - 1, 1) = Trim(RS!PRD_PCK) & " - " & Val(RS!sub_unt)
                    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
                    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(SHLID & "")
                    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!SUB_UNT)
                    If Pieces.Value = 0 Then
                        SHLCAS = STM(MS.TextMatrix(A, 4) - (MS.TextMatrix(A, 5) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 7))), Val(MS.TextMatrix(A, 16)))
                        ms2.TextMatrix(b, 12) = Int(Val(MS.TextMatrix(A, 7)))
                        ms2.TextMatrix(b, 13) = QTY(Val(MS.TextMatrix(A, 7)))
                        CC = Val(ms2.TextMatrix(b, 3))
                        BB(b) = Val(MS.TextMatrix(A, 5))
                        e = (Val(MS.TextMatrix(A, 7)))
                        f = (Val(MS.TextMatrix(A, 8)))
                    Else
                        AA = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                        BB(b) = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                        C = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
                        d = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
                        e = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
                        f = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
                        CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                        'ms2.TextMatrix(b, 12) = Int(Val(MS.TextMatrix(A, 7)))
                        'ms2.TextMatrix(b, 13) = QTY(Val(MS.TextMatrix(A, 7)))
                        SHLCAS = QTY(STM(Val(AA) - (Val(BB(b)) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
                    End If
                    If Check3.Value = 0 Then
                        ms2.TextMatrix(b, 3) = Int(Val(MS.TextMatrix(A, 4))) + Val(ms2.TextMatrix(b, 3))
                        ms2.TextMatrix(b, 4) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) + (Val(AA)), ms2.TextMatrix(b, 11)))
                        CC = MTS(Val(ms2.TextMatrix(b, 3)), ms2.TextMatrix(b, 11))
                    End If
                    ms2.TextMatrix(b, 1) = Int(STM(Val(CC) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                    ms2.TextMatrix(b, 2) = QTY(STM(MTS(Val(ms2.TextMatrix(b, 4)), ms2.TextMatrix(b, 11)) - (Val(BB(b)) + Val(e) + Val(f)), ms2.TextMatrix(b, 11)))
                End If
            End If
n:
        End If
    End If
Next

'If RS.State = 1 Then RS.Close
'RS.Open "select * from snv,SHL where inv_idy = '" & Trim(M_INV_IDY) & "' and SNV.SHL_IDY = SHL.SHL_IDY AND  txn_typ = 'VAN'", CON, adOpenKeyset, adlockpessimistic
'For A = 1 To RS.RecordCount
'    For b = 2 To ms2.Rows - 1
'        If Trim(RS!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(RS!shl_idy) = Trim(ms2.TextMatrix(b, 10)) And Val(RS!prd_upc) = Val(ms2.TextMatrix(b, 11)) Then
'            ms2.TextMatrix(b, 5) = Val(RS!R_shl_nos & "")
'            ms2.TextMatrix(b, 6) = QTY(STM(Val(RS!R_pck_nos & ""), ms2.TextMatrix(b, 11)))
'            GoTo l
'        End If
'    Next
'    ms2.AddItem ""
'    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(RS!shl_nme) & "-" & Trim(RS!PRD_PCK) & " - " & Val(RS!prd_upc)
''    ms2.TextMatrix(ms2.Rows - 1, 1)
''    ms2.TextMatrix(ms2.Rows - 1, 1)
'    ms2.TextMatrix(b, 5) = Val(RS!R_shl_nos & "")
'    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(RS!prd_upc)
'    ms2.TextMatrix(b, 6) = QTY(STM(Val(RS!R_pck_nos & ""), ms2.TextMatrix(b, 11)))
'    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(RS!PRD_PCK & "")
'    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(RS!shl_idy & "")
'
'l:
'    RS.MoveNext
'Next

For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 7) = Val(ms2.TextMatrix(b, 3)) - Val(ms2.TextMatrix(b, 5))
    ms2.TextMatrix(b, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(b, 4)), Val(ms2.TextMatrix(b, 11))) - MTS(Val(ms2.TextMatrix(b, 6)), Val((ms2.TextMatrix(b, 11))))), Val((ms2.TextMatrix(b, 11)))))
Next

For b = 2 To ms2.Rows - 1
    ms2.TextMatrix(b, 7) = Val(ms2.TextMatrix(b, 3)) - (Val(ms2.TextMatrix(b, 5)) + Val(ms2.TextMatrix(b, 12)))
    ms2.TextMatrix(b, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(b, 4)), Val(ms2.TextMatrix(b, 11))) - (MTS(Val(ms2.TextMatrix(b, 13)), Val(ms2.TextMatrix(ms2.Row, 11))) + MTS(Val(ms2.TextMatrix(b, 6)), Val((ms2.TextMatrix(b, 11)))))), Val((ms2.TextMatrix(b, 11)))))
Next
If rs.State = 1 Then rs.Close
rs.Open "SELECT SUB_UNT,SHL_IDY,PRD_PCK,SHL_NME FROM PRD,SHL WHER PRD.PRD_IDY = SHL.PRD_IDY GROUP BY SUB_UNT,SHL_IDY,PRD_PCK", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 2 To ms2.Rows - 1
        If Trim(rs!PRD_PCK) = Trim(ms2.TextMatrix(b, 9)) And Trim(rs!shl_idy) = Trim(ms2.TextMatrix(b, 10)) And Val(rs!SUB_UNT) = Val(ms2.TextMatrix(b, 11)) Then
            GoTo o:
        End If
    Next
    ms2.AddItem ""
    ms2.TextMatrix(ms2.Rows - 1, 0) = Trim(rs!shl_nme) & "-" & Trim(rs!PRD_PCK) & " - " & Val(rs!SUB_UNT)
    'ms2.TextMatrix(ms2.Rows - 1, 1) = Trim(RS!PRD_PCK) & " - " & Val(RS!sub_unt)
    'ms2.TextMatrix(ms2.Rows - 1, 2) = Trim(RS!PRD_PCK) & " - " & Val(RS!sub_unt)
    ms2.TextMatrix(ms2.Rows - 1, 9) = Trim(rs!PRD_PCK & "")
    ms2.TextMatrix(ms2.Rows - 1, 10) = Trim(rs!shl_idy & "")
    ms2.TextMatrix(ms2.Rows - 1, 11) = Val(rs!SUB_UNT)
o:
    rs.MoveNext
Next

If SSTab1.Tab = 5 Then
    ms2.SetFocus
    ms2.Col = 2
End If
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
ElseIf Button.KEY = "CAS_CRT" Then
ChallanCopy
ElseIf Button.KEY = "crt" Then
    Picture2.Visible = True
    DTPicker1.SetFocus
    SSTab1.Enabled = False
    Toolbar1.Enabled = False
ElseIf Button.KEY = "Bill" Then
    M_FRM_RSI_PRD.Show 1
    MS.Rows = 1
    For A = 1 To PrdCollection.count
        MS.AddItem ""
        MS.Row = A
        dispprd PrdCollection.Item(A)
    Next
    Do While PrdCollection.count > 0
        PrdCollection.Remove 1
    Loop
    If MS.Rows = 1 Then
        MS.AddItem ""
    End If
    MS.SetFocus
    SSTab1.Tab = 1
    MS.SetFocus
ElseIf Button.KEY = "ed" Then
    If M_FRM_RSI1.ActiveControl.name = "M_INV_IDY" Then
        Label49 = "Select Invoice"
        INVOICELIST
    ElseIf M_FRM_RSI1.ActiveControl.name = "M_SMN_IDY" Then
        Label49 = "Select Salesman"
        SALESMANLIST
    ElseIf M_FRM_RSI1.ActiveControl.name = "M_CST_IDY" Then
        Label49 = "Select Vehicle"
        V_DLR_LST
    ElseIf M_FRM_RSI1.ActiveControl.name = "M_ACT_IDY" Then
        Label49 = "Select Account"
        V_ACT_LST
    ElseIf M_FRM_RSI1.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        PRODUCTLIST
    End If
ElseIf Button.KEY = "DAY" Then
''''''    If Button.Value = tbrPressed Then
''''''        Text1.Visible = True
''''''        Text1 = Format(Date, "dd/mm/yyyy") & " " & Format(Date, "dddd")
''''''        Text1.BackColor = vbBlue
''''''        Text1.ForeColor = vbWhite
''''''    Else
''''''        Text1.Visible = False
''''''    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
If check("prt", "prt_nme", "M_FRM_RSI1") = False Then
Check3.Value = 0
Else
Check3.Value = 1
End If
Check3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Private Sub M_INV_IDY_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
If Len(Trim(M_INV_IDY)) = 0 Then
    INVOICELIST
Else
    inV = M_INV_IDY
    clrctr M_FRM_RSI1
    txt1000 = "0"
    txt500 = "0"
    txt100 = "0"
    txt50 = "0"
    txt20 = "0"
    txt10 = "0"
    txt5 = "0"
    txt2 = "0"
    txt1 = "0"
    txtchange = "0"
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
M_GRS_AMT = M_NET_SAL
End Sub
Public Sub DISPDETAILS()
'On Error GoTo LBL
If CHECKTB("van.dbf,prd.dbf,brd.dbf,act.dbf,smn.dbf,shl.dbf,snv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rsVan As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
If RSCUR.State = 1 Then RSCUR.Close
If Trolley.Value = 1 Then
    If Len(Trim(M_INV_IDY)) = 10 Then
        rsVan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY like BRD.PRD_IDY AND VAN.bat_IDY like brd.bat_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "' AND VAN_CHK='P'", con, adOpenKeyset, adLockPessimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY like '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    Else
        rsVan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY like BRD.PRD_IDY AND VAN.bat_IDY like brd.bat_IDY AND inv_idy like 'RSI" & Tincrement(M_INV_IDY) & "' AND VAN_CHK='P'", con, adOpenKeyset, adLockPessimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY LIKE 'RSI" & Tincrement(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    End If
Else
    If Len(Trim(M_INV_IDY)) = 10 Then
        rsVan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY like BRD.PRD_IDY AND VAN.bat_IDY like brd.bat_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "' AND VAN_CHK='V'", con, adOpenKeyset, adLockPessimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY like '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    Else
        rsVan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY like BRD.PRD_IDY AND VAN.bat_IDY like brd.bat_IDY AND inv_idy like 'RSI" & Tincrement(M_INV_IDY) & "' AND VAN_CHK='V' ", con, adOpenKeyset, adLockPessimistic
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY LIKE 'RSI" & Tincrement(Trim(M_INV_IDY)) & "'", con, adOpenKeyset, adLockPessimistic
    End If
End If
If rsVan.RecordCount > 0 Then
    If Len(Trim(M_INV_IDY)) = 10 Then
        M_INV_IDY = Trim(M_INV_IDY)
    Else
        M_INV_IDY = "RSI" & Tincrement(M_INV_IDY)
    End If
    TRP_NOS.CLEAR
    TRP_NOS.AddItem "All"
    If rsact.State = 1 Then rsact.Close
    rsact.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "' GROUP BY TRP_IDY", con, adOpenKeyset, adLockPessimistic
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
    M_INV_DAT = rsVan("INV_DAT")
    Trolley.Enabled = False
    M_CST_IDY = Trim(rsVan!CST_IDY)
    M_CST_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY = '" & Trim(rsVan!CST_IDY) & "'")
    M_ACT_IDY = Trim(rsVan!ACT_IDY)
    M_ACT_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY = '" & Trim(rsVan!ACT_IDY) & "'")
    M_SMN_IDY = Trim(rsVan!SMN_IDY)
    MNL_REF = Trim(rsVan!MNL_REF & "")
    If IsNull(Trim(rsVan!LIN_DAT & "")) = False Then
        LINDATE = Trim(rsVan!INV_DAT & "")
    Else
        LINDATE = Trim(rsVan!LIN_DAT & "")
    End If
    If IsNull(Trim(rsVan!LDS_IDY)) = False Then
        M_LDS_IDY = Trim(rsVan!LDS_IDY)
    End If
    For K = 0 To cmbRoute.ListCount - 1
        If Trim(rsVan!rou_idy & "") = Trim(cmbRoute.List(K)) Then
            cmbRoute.ListIndex = K
            Exit For
        End If
    Next
    M_SMN_NME = GetValue("SMN_NME", "SMN WHERE SMN_IDY = '" & Trim(rsVan!SMN_IDY) & "'")
    IDN = False
    If rsVan!VAN_CHK = "V" Then
        Trolley.Value = 0
        MS.ColWidth(11) = 1050
    Else
        Trolley.Value = 1
        MS.ColWidth(11) = 0
    End If
    
    If rsVan!vAN_CLS = "S" Then
        Check3.Value = 1
        Check3.Enabled = False
        Check4.Enabled = True
    ElseIf rsVan!vAN_CLS = "C" Then
        Check3.Value = 1
        Check3.Enabled = False
        Check4.Value = 1
        Check4.Enabled = False
    ElseIf rsVan!vAN_CLS = "N" Then
        Check3.Value = 0
        Check3.Enabled = True
        Check4.Value = 0
        Check4.Enabled = False
    Else
        Check3.Enabled = False
        Check3.Value = 0
    End If
    
    MS.Rows = rsVan.RecordCount + 1
    'MS.FixedCols = 1
    For A = 1 To rsVan.RecordCount
        MS.TextMatrix(A, 0) = Trim(rsVan!PRD_NME)
        If Val(rsVan!trp_idy & "") = 0 And Trolley.Value = 0 Then
            MS.TextMatrix(A, 1) = "All"
        ElseIf Val(rsVan!trp_idy & "") = 0 And Trolley.Value = 1 Then
            MS.TextMatrix(A, 1) = "0"
        Else
            MS.TextMatrix(A, 1) = Trim(Val(rsVan!trp_idy & ""))
        End If
        If Pieces.Value = 0 Then
            MS.TextMatrix(A, 2) = MTS(Val(rsVan!OPN_QTY), rsVan!SUB_UNT)
            MS.TextMatrix(A, 3) = MTS(Val(rsVan!ISS_QTY), rsVan!SUB_UNT)
            MS.TextMatrix(A, 4) = MTS(Val(rsVan!tot_qty), rsVan!SUB_UNT)
            MS.TextMatrix(A, 11) = MTS(Val(rsVan!BIL_QTY), rsVan!SUB_UNT)
            MS.TextMatrix(A, 5) = MTS(Val(rsVan!RET_QTY), rsVan!SUB_UNT)
            MS.TextMatrix(A, 6) = MTS(Val(rsVan!FRE_QTY & ""), rsVan!SUB_UNT)
            MS.TextMatrix(A, 7) = MTS(Val(rsVan!BRK_QTY & ""), rsVan!SUB_UNT)
            MS.TextMatrix(A, 8) = MTS(Val(rsVan!LKG_QTY & ""), rsVan!SUB_UNT)
            MS.TextMatrix(A, 12) = convert(Val(rsVan!PRD_MRP) / Val(rsVan!SUB_UNT))
            MS.TextMatrix(A, 19) = convert(Val(rsVan!PRD_rat) / Val(rsVan!SUB_UNT))
            If IsNull(rsVan!LST_RAT) = True Then
                MS.TextMatrix(A, 20) = 0#
                MS.TextMatrix(A, 21) = 0#
            Else
                MS.TextMatrix(A, 20) = convert(Val(rsVan!LST_RAT))
                MS.TextMatrix(A, 21) = convert(Val(rsVan!LST_DIS))
            End If
        Else
            MS.TextMatrix(A, 2) = QTY(rsVan!OPN_QTY)
            MS.TextMatrix(A, 3) = QTY(rsVan!ISS_QTY)
            MS.TextMatrix(A, 4) = QTY(rsVan!tot_qty)
            MS.TextMatrix(A, 10) = QTY(rsVan!BIL_QTY)
            MS.TextMatrix(A, 5) = QTY(rsVan!RET_QTY)
            MS.TextMatrix(A, 6) = QTY(rsVan!FRE_QTY)
            MS.TextMatrix(A, 7) = QTY(rsVan!BRK_QTY)
            MS.TextMatrix(A, 8) = QTY(rsVan!LKG_QTY)
            MS.TextMatrix(A, 12) = convert(rsVan!PRD_MRP & "")
            MS.TextMatrix(A, 19) = convert(rsVan!PRD_rat)
            If IsNull(rsVan!P_INC_DIS) = False Then
                MS.TextMatrix(A, 22) = convert(Val(rsVan!P_INC_DIS))
            Else
                MS.TextMatrix(A, 22) = convert(0#)
            End If
            If IsNull(rsVan!S_INC_DIS) = False Then
                MS.TextMatrix(A, 23) = convert(Val(rsVan!S_INC_DIS))
            Else
                MS.TextMatrix(A, 23) = convert(0#)
            End If
            If IsNull(rsVan!LST_RAT) = True Then
                MS.TextMatrix(A, 20) = 0#
                MS.TextMatrix(A, 21) = 0#
            Else
                MS.TextMatrix(A, 20) = convert(Val(rsVan!LST_RAT))
                MS.TextMatrix(A, 21) = convert(Val(rsVan!LST_DIS))
            End If
        End If
        M_CAS_AMT = convert(rsVan!CAS_AMT & "")
        M_CHQ_AMT = convert(rsVan!CHQ_AMT & "")
        M_DIS_AMT = convert(rsVan!DIS_AMT & "")
        M_EXP_AMT = convert(rsVan!EXP_AMT & "")
        M_TAX_AMT = convert(rsVan!ADD_AMT & "")
        M_COM_AMT = convert(rsVan!com_amt & "")
        m_dep_ref = convert(rsVan!dep_ref & "")
        m_dep_rcd = convert(rsVan!dep_rcd & "")
        M_OLD_CRD = convert(rsVan!OLD_CRD & "")
        M_CUR_CRD = convert(rsVan!CUR_CRD & "")
        MS.TextMatrix(A, 13) = convert(rsVan!COL_AMT & "")
        'CHANGED
        MS.TextMatrix(A, 14) = Trim(rsVan!prd_idy)
        MS.TextMatrix(A, 15) = Trim(rsVan!prd_prn)
        MS.TextMatrix(A, 16) = Val(rsVan!SUB_UNT)
        MS.TextMatrix(A, 17) = Trim(rsVan!BAT_IDY)
        MS.TextMatrix(A, 18) = Trim(rsVan!Bat_nme & "")
        rsVan.MoveNext
    Next
    If Not RSCUR.EOF Then
        txt1000 = Val(Trim(RSCUR(1)) & "")
        txt500 = Val(Trim(RSCUR(2)) & "")
        txt100 = Val(Trim(RSCUR(3)) & "")
        txt50 = Val(Trim(RSCUR(4)) & "")
        txt20 = Val(Trim(RSCUR(5)) & "")
        txt10 = Val(Trim(RSCUR(6)) & "")
        txt5 = Val(Trim(RSCUR(7)) & "")
        txt2 = Val(Trim(RSCUR(8)) & "")
        txt1 = Val(Trim(RSCUR(9)) & "")
        txtchange = Val(Trim(RSCUR(10)) & "")
    End If
    LINDATE.Enabled = Check3.Enabled
    ShlSet
'    SETDISPDETAILS
    discalc
    calc
    list5.Visible = False
    MOD_REC = True
    
    Command2.Value = True           'Take Control thru' Generate and Update Buttons
    SHL_SET.Value = True
    Command6.Value = True
    RETURNSTOCKS
    SSTab1.Tab = 0
    MNL_REF.SetFocus
Else
    INVOICELIST
End If
Exit Sub
LBL:
If err.Number = 35773 Then
    MsgBox "Invalid Voucher Selected for the Current Financial Year !", vbInformation, "DAS Version 4.0"
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
Dim rsVan As New ADODB.Recordset
If Len(Trim(M_INV_IDY)) = 10 Then
    rsVan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY = BRD.PRD_IDY AND VAN.bat_IDY = brd.bat_IDY AND inv_idy = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rsVan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY = BRD.PRD_IDY AND VAN.bat_IDY = brd.bat_IDY AND inv_idy like 'RSI" & Tincrement(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
End If
TOTVAL = 0#
CTAXVAL = 0#
RetVal = 0#
freval = 0#
brkval = 0#
lkgval = 0#
If rsVan.RecordCount > 0 Then
    For A = 1 To rsVan.RecordCount
        CTAXVAL = 0#
        TOTVAL = Val(TOTVAL) + Val(RATCAL(QTY(Val(rsVan!tot_qty & "")), Val(rsVan!PRD_rat) + Val(rsVan!LST_RAT), Val(rsVan!SUB_UNT)))
        'TAXVAL = ((TOTVAL * Val(rsVan!LST_DIS) / 100) + (TOTVAL * Val(rsVan!SAL_TAX) / 100)) 'PNR
        CTAXVAL = Val(rsVan!LST_RAT) * Val(rsVan!LST_DIS)
        CTAXVAL = (CTAXVAL / 100)
        
        CTAXVAL = CTAXVAL * (MTS(Val(rsVan!tot_qty), Val(rsVan!SUB_UNT)))
'        TOTVAL = TOTVAL + TAXVAL 'PNR
        RetVal = Val(RetVal) + Val(RATCAL(QTY(Val(rsVan!RET_QTY & "")), Val(rsVan!PRD_rat), Val(rsVan!SUB_UNT)))
        freval = Val(freval) + Val(RATCAL(QTY(Val(rsVan!FRE_QTY & "")), (Val(rsVan!PRD_rat) + Val(rsVan!LST_RAT)), Val(rsVan!SUB_UNT)))
'        FTAXVAL = ((freval * Val(rsVan!LST_DIS) / 100) + (freval * Val(rsVan!SAL_TAX) / 100)) 'PNR
        freval = freval + FTAXVAL 'PNR
        brkval = Val(brkval) + Val(RATCAL(QTY(Val(rsVan!BRK_QTY & "")), Val(rsVan!PRD_rat), Val(rsVan!SUB_UNT)))
        lkgval = Val(lkgval) + Val(RATCAL(QTY(Val(rsVan!LKG_QTY & "")), Val(rsVan!PRD_rat), Val(rsVan!SUB_UNT)))
        rsVan.MoveNext
    Next
    rsVan.MoveFirst
    MSS.TextMatrix(1, 1) = convert((Val(TOTVAL) + CTAXVAL) - (Val(RetVal) + Val(brkval) + Val(lkgval)))

    MSS.TextMatrix(2, 1) = convert(Val(freval) + FTAXVAL)
    MSS.TextMatrix(3, 1) = convert(Val(brkval))
    MSS.TextMatrix(4, 1) = convert(Val(lkgval))
    MSS.TextMatrix(7, 1) = convert(Val(rsVan!DIS_AMT & ""))
    
    MSS.TextMatrix(14, 1) = convert(Val(rsVan!EXP_AMT & "") + Val(rsVan!com_amt & ""))
    MSS.TextMatrix(8, 1) = convert(Val(rsVan!ADD_AMT & ""))
    MSS.TextMatrix(10, 1) = convert(Val(rsVan!CAS_AMT & ""))
    MSS.TextMatrix(11, 1) = convert(Val(rsVan!CHQ_AMT & ""))
    MSS.TextMatrix(15, 1) = convert(Val(rsVan!dep_ref & ""))
    MSS.TextMatrix(16, 1) = convert(Val(rsVan!dep_rcd & ""))
    MSS.TextMatrix(17, 1) = convert(Val(rsVan!OLD_CRD & ""))
    MSS.TextMatrix(18, 1) = convert(Val(rsVan!CUR_CRD & ""))
    
'''''''''    MSS.TextMatrix(1, 1) = convert(0#)
'''''''''    MSS.TextMatrix(2, 1) = convert(0#)
'''''''''    MSS.TextMatrix(3, 1) = convert(0#)
'''''''''    MSS.TextMatrix(4, 1) = convert(0#)
'''''''''    MSS.TextMatrix(7, 1) = convert(0#)
'''''''''    MSS.TextMatrix(8, 1) = convert(0#)
'''''''''    MSS.TextMatrix(10, 1) = convert(0#)
'''''''''    MSS.TextMatrix(11, 1) = convert(0#)
'''''''''    MSS.TextMatrix(14, 1) = convert(0#)
'''''''''    MSS.TextMatrix(15, 1) = convert(0#)
'''''''''    MSS.TextMatrix(16, 1) = convert(0#)
End If
SETCALC
End Sub
Private Sub M_SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Salesmen"
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub List1_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list1.ListItems.count > 0 Then
        M_SMN_IDY = list1.SelectedItem.ListSubItems.Item(1)
        M_SMN_IDY_KeyDown 13, 0
        list1.Visible = False
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
    If rsVan.State = 1 Then rsVan.Close
    If Len(Trim(M_INV_IDY)) = 10 Then
        rsVan.Open "SELECT * FROM VAN WHERE INV_IDY='" & Trim(M_INV_IDY) & "'", con
    Else
        rsVan.Open "SELECT * FROM VAN WHERE INV_IDY='RSI" & Tincrement(M_INV_IDY) & "'", con
    End If
    Shl_Set_Click   'Added by PNR for update shells routine to get updated as per Prince Chopra input
    'If rsvan.EOF Then SHL_SET.Enabled = False
    SHL_SET.SetFocus
'    If ms2.Rows > 1 Then ms2.Row = 1
'    ms2.SetFocus
ElseIf SSTab1.Tab = 6 Then
    MS1.SetFocus
    If rsVan.State = 1 Then rsVan.Close
    If Len(Trim(M_INV_IDY)) = 10 Then
        rsVan.Open "SELECT * FROM VAN WHERE INV_IDY='" & Trim(M_INV_IDY) & "'", con
    Else
        rsVan.Open "SELECT * FROM VAN WHERE INV_IDY='RSI" & Tincrement(M_INV_IDY) & "'", con
    End If
    If rsVan.EOF Then Command6.Enabled = False
    If MS1.Row > 1 Then MS1.Row = 2
    MS.Col = 1
    Verify_Glass_Variance
    If Command6.Enabled = True Then Command6.SetFocus
ElseIf SSTab1.Tab = 2 Then
    M_GRS_SAL.SetFocus
    ADDTOTCUR
    If Check3.Value = 1 And Check3.Enabled = True Then COMMISSION
End If
If list1.Visible = True Then list1.Visible = False
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

Private Sub M_SMN_IDY_KeyDown(K As Integer, s As Integer)
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
    REC_SET.Open "select * from smn where smn_idy = '" & Sincrement(Trim(M_SMN_IDY)) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        If REC_SET1.State = 1 Then REC_SET1.Close
        REC_SET1.Open "SELECT * FROM INL WHERE SMN_IDY LIKE '" & Sincrement(Trim(M_SMN_IDY)) & "' AND TXN_TYP = 'RSI'", con, adOpenKeyset, adLockPessimistic
        If REC_SET1.RecordCount > 0 Then
            For A = 1 To REC_SET1.RecordCount
                If REC_SET2.State = 1 Then REC_SET2.Close
                REC_SET2.Open "SELECT * FROM INL WHERE ORD_IDY = '" & Trim(REC_SET1("INV_IDY")) & "' AND TXN_TYP = 'RSR'", con, adOpenKeyset, adLockPessimistic
                If REC_SET2.RecordCount = 0 Then
                    If MOD_REC = False Then
                        MsgBox "Return Stock Should be posted for Salesman Mr." & Trim(REC_SET("smn_nme"))
                        M_SMN_IDY.SetFocus
                        M_SMN_IDY = ""
                        SSTab1.Tab = 0
                    Else
                        SSTab1.Tab = 0
                        cmbRoute.SetFocus
'                        SSTab1.Tab = 1
'                        If MOD_REC = False Then
'                        Toolbar1.Buttons(12).Enabled = True
'                        Toolbar1.Buttons(19).Enabled = True
'                        End If
'                        MS.SetFocus
                    End If
                    Exit Sub
                Else
                    M_SMN_IDY = Sincrement(M_SMN_IDY)
                    M_SMN_NME = Trim(REC_SET("smn_nme"))
                    If K <> 38 Then
                    SSTab1.Tab = 0
                    cmbRoute.SetFocus
'                        SSTab1.Tab = 1
'                        If MOD_REC = False Then
'                            Toolbar1.Buttons(12).Enabled = True
'                            Toolbar1.Buttons(19).Enabled = True
'                        End If
'                        MS.SetFocus
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
'                If MOD_REC = False Then
'                Toolbar1.Buttons(12).Enabled = True
'                Toolbar1.Buttons(19).Enabled = True
'                End If
'                SSTab1.Tab = 1
'                MS.SetFocus
                 SSTab1.Tab = 0
                 cmbRoute.SetFocus
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
    ElseIf MS.Col = 3 And MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
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
    ElseIf MS.Col = 5 Or MS.Col = 6 Or MS.Col = 7 Or MS.Col = 8 Then
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
Private Sub Text3_KeyDown(K As Integer, s As Integer)
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
        REC_SET.Open "SELECT * FROM prd,brd WHERE brd.prd_idy = prd.prd_idy and brd.prd_IDY = '" & Sincrement(Text3) & "' AND PRD_QTY >0 AND PRD.STK_TYP IN('Salable','Free')", con, adOpenKeyset
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
M_GRS_AMT = convert(amt(M_GRS_AMT & ""))
Next
M_NET_AMT = convert(Val(M_GRS_AMT))
M_NET_SAL = convert(Val(M_GRS_SAL) - Val(M_FRE_QTY))
'M_CRT_AMT.Text = convert((Val(M_NET_SAL) + Val(m_dep_rcd) + Val(M_TAX_AMT)) - (Val(M_CAS_AMT) + Val(m_dep_ref) + Val(M_CHQ_AMT) + Val(M_DIS_AMT) + Val(M_COM_AMT) + Val(M_EXP_AMT)))
M_CRT_AMT.Text = convert((Val(M_NET_SAL) + Val(m_dep_rcd) + Val(M_TAX_AMT) + Val(M_OLD_CRD)) - (Val(M_CAS_AMT) + Val(m_dep_ref) + Val(M_CHQ_AMT) + Val(M_DIS_AMT) + Val(M_EXP_AMT) + Val(M_CUR_CRD)))
'TTL_CSH_VAR = convert((Val(M_CRT_AMT) + Val(M_EXP_AMT) + Val(m_dep_ref)) - (Val(m_dep_rcd) + Val(M_OLD_CRD)))
'If Val(M_CRT_AMT) < 0 Then M_CRT_AMT = convert(0#)
MSS.TextMatrix(12, 1) = convert(M_CRT_AMT.Text)

M_GRS_AMT = M_NET_SAL
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
            REC_SET2.Open "select TOT_QTY,SUB_UNT from VAN,PRD where VAN.PRD_IDY = PRD.PRD_IDY AND BAT_IDY = '" & Trim(MS.TextMatrix(MS.Row, 17)) & "' and inv_idy = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
            If REC_SET2.RecordCount > 0 Then
                For A = 1 To REC_SET2.RecordCount
                    INVTOT = Val(INVTOT) + MTS(Val(REC_SET2(0)), Val(REC_SET2!SUB_UNT))
                    REC_SET2.MoveNext
                Next
                REC_SET2.MoveFirst
                INVTOT = QTY(STM(Val(INVTOT), Val(REC_SET2!SUB_UNT)))
            End If
        End If
        If Pieces.Value = 1 Then
            STK = Val(MTS(Val(MS.TextMatrix(MS.Row, 2)), Val(MS.TextMatrix(MS.Row, 16))))
            STK = STK + Val(MTS(Val(amount), Val(MS.TextMatrix(MS.Row, 16))))
        ElseIf Pieces.Value = 0 Then
            amount = Int(amt(Val(amount)))
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
        If b > 0 And Val(MS.TextMatrix(MS.Row, 16)) > 0 Then
            If InStr(b + 1, amount, ".") > 0 Then
                MS.TextMatrix(MS.Row, 4) = "0.000"
                Exit Sub
            End If
            M = Mid(amount, b + 1, 3)
            If M >= Val(MS.TextMatrix(MS.Row, 16)) Then
                MsgBox "Product Sub Units Are " & Val(MS.TextMatrix(MS.Row, 16)), vbInformation, "DAS Version 4.0"
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
            REMIND = MTS(Val(ADDBRDQTY(QTY(Val(INVTOT)), Trim(MS.TextMatrix(MS.Row, 17)))), Val(MS.TextMatrix(MS.Row, 16)))
            For A = 1 To MS.Rows - 1
                If MS.Row <> A Then
                    If Trim(MS.TextMatrix(MS.Row, 17)) = Trim(MS.TextMatrix(A, 17)) Then
                        If Pieces = 0 Then
                            ENTQTY = Val(ENTQTY) + Val(MS.TextMatrix(A, 4))
                        Else
                            ENTQTY = Val(ENTQTY) + MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                        End If
                    End If
                End If
            Next
            CURQTY = Val(M_QTY) + Val(ENTQTY)
            If querycheck("DIR", "DIR_MSG", "U", " and dir_idy = 'SYS' AND DIR_IDX = 'STK'") = False Then
                If Val(REMIND) - Val(CURQTY) < 0 Then
                    MsgBox "Stock Out Of Range, Check Stock !", vbInformation, "DAS Version 4.0"
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
                MS.TextMatrix(MS.Row, 4) = QTY(STM(Val(STK), MS.TextMatrix(MS.Row, 16)))
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
            REC_SET2.Open "select TOT_QTY,PRD.SUB_UNT from VAN,PRD where VAN.BAT_IDY = '" & Trim(MS.TextMatrix(MS.Row, 13)) & "' and VAN.INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
            'Code modified by PNR since there is no SUB_UNT field in VAN and KRB is trying to bring this field for SUB UNIT calculation
            For A = 1 To REC_SET2.RecordCount
                INVTOT = Val(INVTOT) + MTS(Val(REC_SET2(0)), Val(REC_SET2!SUB_UNT))
                REC_SET2.MoveNext
            Next
            If Not REC_SET2.EOF Then  'Added by PNR, Since EOF is encountered for the defined RECORDSET
               REC_SET2.MoveFirst
               INVTOT = QTY(STM(Val(INVTOT), Val(REC_SET2!SUB_UNT)))
            End If
        End If
        If Pieces.Value = 1 Then
        STK = Val(MTS(Val(MS.TextMatrix(MS.Row, 3)), Val(MS.TextMatrix(MS.Row, 16))))
        STK = STK + Val(MTS(Val(amount), Val(MS.TextMatrix(MS.Row, 16))))
        ElseIf Pieces.Value = 0 Then
            amount = Int(amt(Val(amount)))
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
        If b > 0 And Val(MS.TextMatrix(MS.Row, 16)) > 0 Then
            If InStr(b + 1, amount, ".") > 0 Then
                MS.TextMatrix(MS.Row, 4) = "0.000"
                Exit Sub
            End If
            M = Mid(amount, b + 1, 3)
            If M >= Val(MS.TextMatrix(MS.Row, 16)) Then
                MsgBox "Product Sub Units Are " & Val(MS.TextMatrix(MS.Row, 16)), vbInformation, "DAS Version 4.0"
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
            REMIND = MTS(Val(ADDBRDQTY(QTY(Val(INVTOT)), Trim(MS.TextMatrix(MS.Row, 17)))), Val(MS.TextMatrix(MS.Row, 16)))
            For A = 1 To MS.Rows - 1
                If MS.Row <> A Then
                    If Trim(MS.TextMatrix(MS.Row, 17)) = Trim(MS.TextMatrix(A, 17)) Then
                        If Pieces = 0 Then
                            ENTQTY = Val(ENTQTY) + Val(MS.TextMatrix(A, 4))
                        Else
                            ENTQTY = Val(ENTQTY) + MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                        End If
                    End If
                End If
            Next
            CURQTY = Val(M_QTY)
            If querycheck("DIR", "DIR_MSG", "U", " and dir_idy = 'SYS' AND DIR_IDX = 'STK'") = False Then
                If Val(REMIND) - Val(CURQTY) < 0 Then
                    MsgBox "Stock Out Of Range, Check Stock !", vbInformation, "DAS Version 4.0"
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
                MS.TextMatrix(MS.Row, 4) = QTY(STM(Val(STK), Val(MS.TextMatrix(MS.Row, 16))))
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
ElseIf MS.Col = 5 Or MS.Col = 6 Or MS.Col = 7 Or MS.Col = 8 Then
        If Pieces.Value = 1 Then
            amount = QTY(amount)
            b = InStr(amount, ".")
            If b > 0 And Val(MS.TextMatrix(MS.Row, 16)) > 0 Then
                If InStr(b + 1, amount, ".") > 0 Then
                    MS.TextMatrix(MS.Row, 4) = "0.000"
                    Exit Sub
                End If
                M = Mid(amount, b + 1, 4)
                If M >= Val(MS.TextMatrix(MS.Row, 16)) Then
                    MsgBox "Product Sub Units Are " & Val(MS.TextMatrix(MS.Row, 16)), vbInformation, "DAS Version 4.0"
                    If amount.Visible = True Then
                        amount.SetFocus
                        SendKeys "{home}+{end}"
                        Exit Sub
                    Else
                        amount.Visible = False
                        MS.SetFocus
                    End If
                Else
                    If Val(amount) > Val(MS.TextMatrix(MS.Row, 4)) Then
                        MsgBox "Return, Free, Breakage, Leakage Stock is More Than Load Out !", vbInformation, "DAS Version 4.0"
                        If amount.Visible = True Then
                            amount.SetFocus
                            SendKeys "{home}+{end}"
                            Exit Sub
                        End If
                    Else
                        amount.Visible = False
                        MS.Text = QTY(amount)
                        P_BIL_QTY
                        discalc
                        If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                        MS.SetFocus
                    End If
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
            MS.Text = QTY(amount)   'ADDED BY PNR
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
Public Sub dispprd(PRDID As String)
If CHECKTB("prd.dbf,brd.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
MSROW = MS.Row
Dim REC_SET As New ADODB.Recordset
For A = 1 To MS.Rows - 1
    If Trim(PRDID) = MS.TextMatrix(A, 17) And Val(TRP_NOS) = Val(MS.TextMatrix(A, 1)) Then
        SSTab1.Tab = 1
        list4.Visible = False
        MsgBox "Product - SKU Already Exists In Invoice !", vbInformation, "DAS Version 4.0"
        Text3.Visible = False
        MS.SetFocus
        MS.Col = 0
        MS.Row = A
        Exit Sub
    End If
Next
REC_SET.Open "select * from PRD,BRD,ACT where BRD.PRD_IDY = PRD.PRD_IDY AND BRD.BAT_idy = '" & Trim(PRDID) & "' AND ACT.ACT_IDY LIKE '" & Trim(M_CST_IDY) & "%'", con, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    list4.Visible = False
    MS.TextMatrix(MS.Row, 0) = REC_SET("PRD_NME")
    
    ' Formatting of cells into 0.000 or 0 ie either cases or pieces
            A = "0.000"
            If Pieces.Value = 0 Then A = 0
            MS.TextMatrix(MS.Row, 1) = Val(TRP_NOS)
            MS.TextMatrix(MS.Row, 2) = A
            MS.TextMatrix(MS.Row, 2) = A
            MS.TextMatrix(MS.Row, 3) = A
            MS.TextMatrix(MS.Row, 4) = A
            MS.TextMatrix(MS.Row, 5) = A
            MS.TextMatrix(MS.Row, 10) = A
            MS.TextMatrix(MS.Row, 11) = A
    ' End of cells formatting
    
    ' Default rates for cases or pieces
            If Pieces.Value = 1 Then
                MS.TextMatrix(MS.Row, 12) = convert(REC_SET!PRD_MRP)
                If REC_SET("DBT_RAP") = "A" Then
                    If IsNull(REC_SET!BAS_RAT) = True Then
                        MS.TextMatrix(MS.Row, 19) = 0 'convert(REC_SET!BAS_RAT)
                    Else
                        MS.TextMatrix(MS.Row, 19) = convert(REC_SET!BAS_RAT)
                    End If
                ElseIf REC_SET("DBT_RAP") = "B" Then
                    B_SAL_RAT = (Val(REC_SET("PRD_SDR")) - Val(REC_SET("S_INC_DIS")))
                    'B_SAL_TAX = (REC_SET("PRD_SDR") * REC_SET("SAL_TAX") / 100)
                    B_SAL_TAX = (Val(B_SAL_RAT) * REC_SET("SAL_TAX") / 100)
                    MS.TextMatrix(MS.Row, 19) = convert(REC_SET!PRD_SDR + B_SAL_TAX)
                ElseIf REC_SET("DBT_RAP") = "C" Then
                    MS.TextMatrix(MS.Row, 19) = convert(REC_SET!PRD_MRP)
                End If
            Else
                MS.TextMatrix(MS.Row, 12) = convert(Val(REC_SET!PRD_MRP) / Val(REC_SET!SUB_UNT))
                If REC_SET("DBT_RAP") = "A" Then
                    MS.TextMatrix(MS.Row, 19) = convert(Val(REC_SET!BAS_RAT) / Val(REC_SET!SUB_UNT))
                ElseIf REC_SET("DBT_RAP") = "B" Then
                    B_SAL_RAT = (Val(REC_SET("PRD_SDR")) - Val(REC_SET("S_INC_DIS")))
                    'B_SAL_TAX = (REC_SET("PRD_SDR") * REC_SET("SAL_TAX") / 100)
                    B_SAL_TAX = (Val(B_SAL_RAT) * REC_SET("SAL_TAX") / 100)
                    MS.TextMatrix(MS.Row, 19) = convert(Val(REC_SET!PRD_SDR + B_SAL_TAX) / Val(REC_SET!SUB_UNT))
                ElseIf REC_SET("DBT_RAP") = "C" Then
                    MS.TextMatrix(MS.Row, 19) = convert(Val(REC_SET!PRD_MRP) / Val(REC_SET!SUB_UNT))
                End If
            End If
    ' End of default rates for cases or pieces
    
    MS.TextMatrix(MS.Row, 13) = ""
    MS.TextMatrix(MS.Row, 14) = Trim(REC_SET("PRD_IDY"))
    MS.TextMatrix(MS.Row, 15) = Trim(REC_SET("PRD_PRN"))
    MS.TextMatrix(MS.Row, 16) = Val(REC_SET!SUB_UNT)
    MS.TextMatrix(MS.Row, 17) = Trim(REC_SET!BAT_IDY)
    MS.TextMatrix(MS.Row, 18) = Trim(REC_SET!Bat_nme)
    If IsNull(REC_SET!P_INC_DIS) = False Then
        MS.TextMatrix(MS.Row, 22) = convert(Val(REC_SET!P_INC_DIS))
    Else
        MS.TextMatrix(MS.Row, 22) = convert(0#)
    End If
    If IsNull(REC_SET!S_INC_DIS) = False Then
        MS.TextMatrix(MS.Row, 23) = convert(Val(REC_SET!S_INC_DIS))
    Else
        MS.TextMatrix(MS.Row, 23) = convert(0#)
    End If
    ' Crate rental and tax by default
        If IsNull(REC_SET!LST_RAT) = True Then
            MS.TextMatrix(MS.Row, 20) = 0#
            MS.TextMatrix(MS.Row, 21) = 0#
        Else
            MS.TextMatrix(MS.Row, 20) = convert(REC_SET!LST_RAT)
            MS.TextMatrix(MS.Row, 21) = convert(REC_SET!LST_DIS)
        End If
    ' End of crate rental and tax by default
    Text3.Visible = False
    calc
    SSTab1.Tab = 1
    If REC_SET("BRD_DOE") <= Date Then 'BBD Check
        MsgBox "SKU Batch BBD - Expiry Date is " & Format(REC_SET("BRD_DOE"), "dd/mmm/yyyy"), vbCritical, "DAS Version 4.0"
    End If
    MS.Row = MSROW
    MS.Col = 3
    MS.SetFocus
End If
End Sub
Public Sub discalc()
units = 0
tax = 0
Total = 0
tot_lou = 0
tot_lin = 0
Tot_fre = 0
tot_var = 0
tot_net = 0
'Dim L_SAL_VAL, P_GRS_SAL, P_FRE_QTY, M_GRS_SAL, M_FRE_QTY, M_BRK_QTY, M_LKG_QTY As Double
L_SAL_VAL = 0
P_GRS_SAL = 0
P_FRE_QTY = 0
M_GRS_SAL = 0
M_FRE_QTY = 0
M_BRK_QTY = 0
M_LKG_QTY = 0
LST_TOT = 0
F_LST_TOT = 0
F_CPN_AMT = 0
R_LST_TOT = 0
R_CPN_AMT = 0
M_CPN_AMT = 0
BRK_LST_TOT = 0
LKG_LST_TOT = 0
BRK_CPN_AMT = 0
LKG_CPN_AMT = 0

'Customer Discount calculating on units and rate
    
For A = 1 To MS.Rows - 1

    If MS.TextMatrix(A, 0) <> "" Then
        ' Quantity in pieces modules calculation
        
        If Pieces.Value = 0 Then
            'MS.TextMatrix(A, 13) = convert(Val(MS.TextMatrix(A, 19)) * Val(MS.TextMatrix(A, 10)))
                        
            ' Crate rental and tax calculation
                LST_TOT = ((Val(MS.TextMatrix(A, 20)) * Val(MS.TextMatrix(A, 21)) / 100)) / Val(MS.TextMatrix(A, 16))
                M_CPN_AMT = (LST_TOT * Val(MS.TextMatrix(A, 4))) + (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)) * Val(MS.TextMatrix(A, 4)))
            ' End of crate rental and tax calculation
           
           ' Leakage quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 8)) > 0 Then
                    LKG_LST_TOT = convert((Val(MS.TextMatrix(A, 20)) * Val(MS.TextMatrix(A, 21)) / 100)) / Val(MS.TextMatrix(A, 16))
                    LKG_CPN_AMT = convert(F_LST_TOT * Val(MS.TextMatrix(A, 8))) + (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)) * Val(MS.TextMatrix(A, 8)))
                Else
                    LKG_LST_TOT = 0#
                    LKG_CPN_AMT = 0#
                End If
            ' End of leakage quantity crate rental and tax calculation
            
           ' Breakage quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 7)) > 0 Then
                    BRK_LST_TOT = convert((Val(MS.TextMatrix(A, 20)) * Val(MS.TextMatrix(A, 21)) / 100)) / Val(MS.TextMatrix(A, 16))
                    BRK_CPN_AMT = convert(F_LST_TOT * Val(MS.TextMatrix(A, 7))) + (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)) * Val(MS.TextMatrix(A, 7)))
                Else
                    BRK_LST_TOT = 0#
                    BRK_CPN_AMT = 0#
                End If
            ' End of breakage quantity crate rental and tax calculation
            
            ' Free quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 6)) > 0 Then
                    F_LST_TOT = convert((Val(MS.TextMatrix(A, 20)) * Val(MS.TextMatrix(A, 21)) / 100)) / Val(MS.TextMatrix(A, 16))
                    F_CPN_AMT = convert(F_LST_TOT * Val(MS.TextMatrix(A, 6))) + (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)) * Val(MS.TextMatrix(A, 6)))
                Else
                    F_LST_TOT = 0#
                    F_CPN_AMT = 0#
                End If
            ' End of free quantity crate rental and tax calculation
            
            
            ' Return quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 5)) > 0 Then
                    R_LST_TOT = convert((Val(MS.TextMatrix(A, 20)) * Val(MS.TextMatrix(A, 21)) / 100)) / Val(MS.TextMatrix(A, 16))
                    R_CPN_AMT = convert(R_LST_TOT * Val(MS.TextMatrix(A, 5))) + (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)) * Val(MS.TextMatrix(A, 5)))
                Else
                    R_LST_TOT = 0#
                    R_CPN_AMT = 0#
                End If
            ' End of return quantity crate rental and tax calculation
            
            
            
            M_GRS_SAL = convert(M_GRS_SAL + LST_TOT + M_CPN_AMT) + (Val(MS.TextMatrix(A, 4)) - Val(MS.TextMatrix(A, 5))) * Val(MS.TextMatrix(A, 19))
            M_GRS_SAL = convert(M_GRS_SAL - (R_LST_TOT + R_CPN_AMT + BRK_LST_TOT + BRK_CPN_AMT + LKG_LST_TOT + LKG_CPN_AMT))
            M_FRE_QTY = convert((M_FRE_QTY + F_LST_TOT + F_CPN_AMT) + (Val(MS.TextMatrix(A, 19)) * Val(MS.TextMatrix(A, 6))))
            
            MS.TextMatrix(A, 9) = convert(Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 7)) + Val(MS.TextMatrix(A, 8))))
            MS.TextMatrix(A, 13) = convert(Val(MS.TextMatrix(A, 19)) * Val(MS.TextMatrix(A, 9)))
            MS.TextMatrix(A, 13) = convert(MS.TextMatrix(A, 13) + LST_TOT + M_CPN_AMT) 'PNR
            If Trolley.Value = 1 Then MS.TextMatrix(A, 10) = Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 7)) + Val(MS.TextMatrix(A, 8)))
            MS.TextMatrix(A, 11) = convert(Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 10)) + Val(MS.TextMatrix(A, 7)) + Val(MS.TextMatrix(A, 8))))

        
        End If
            
            'M_BRK_QTY = convert(Val(M_BRK_QTY) + Val(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 17))))
            'M_LKG_QTY = convert(Val(M_LKG_QTY) + Val(Val(MS.TextMatrix(A, 15)) * Val(MS.TextMatrix(A, 18))))
        
        If Pieces.Value = 1 Then
            'MS.TextMatrix(A, 13) = convert(RATCAL(QTY(Val(MS.TextMatrix(A, 16) & "")), Val(MS.TextMatrix(A, 19) & ""), Val(MS.TextMatrix(A, 16))))
            
            AA = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
            b = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
            C = MTS(Val(MS.TextMatrix(A, 10)), Val(MS.TextMatrix(A, 16)))
            d = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
            e = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
            f = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
            ' + VAL(C)) AFTER VAL(B)
            
            ' Crate rental and tax calculation
                LST_TOT = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                LST_TOT = LST_TOT * (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)))
                M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 21))) / 100)
            ' End of crate rental and tax calculation
            
            ' Leakage quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 8)) > 0 Then
                    LKG_LST_TOT = MTS(Val(MS.TextMatrix(A, 8)), Val(MS.TextMatrix(A, 16)))
                    LKG_LST_TOT = LKG_LST_TOT * (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)))
                    LKG_CPN_AMT = ((LKG_LST_TOT * Val(MS.TextMatrix(A, 21))) / 100)
                Else
                    LKG_LST_TOT = 0#
                    LKG_CPN_AMT = 0#
                End If
            ' End of Leakage quantity crate rental and tax calculation
            
            ' Breakage quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 7)) > 0 Then
                    BRK_LST_TOT = MTS(Val(MS.TextMatrix(A, 7)), Val(MS.TextMatrix(A, 16)))
                    BRK_LST_TOT = BRK_LST_TOT * (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)))
                    BRK_CPN_AMT = ((BRK_LST_TOT * Val(MS.TextMatrix(A, 21))) / 100)
                Else
                    BRK_LST_TOT = 0#
                    BRK_CPN_AMT = 0#
                End If
            ' End of Breakage quantity crate rental and tax calculation
            
            ' Free quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 6)) > 0 Then
                    F_LST_TOT = MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 16)))
                    F_LST_TOT = F_LST_TOT * (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16)))
                    F_CPN_AMT = ((F_LST_TOT * Val(MS.TextMatrix(A, 21))) / 100)
                Else
                    F_LST_TOT = 0#
                    F_CPN_AMT = 0#
                End If
            ' End of free quantity crate rental and tax calculation
            
            ' Return quantity crate rental and tax calculation
                If Val(MS.TextMatrix(A, 5)) > 0 Then
                    R_LST_TOT = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                    R_LST_TOT = (R_LST_TOT * (Val(MS.TextMatrix(A, 20)) / Val(MS.TextMatrix(A, 16))))
                    R_CPN_AMT = ((R_LST_TOT * Val(MS.TextMatrix(A, 21))) / 100)
                Else
                    R_LST_TOT = 0#
                    R_CPN_AMT = 0#
                End If
            ' End of return quantity crate rental and tax calculation
            
            
            M_GRS_SAL = convert(Val(M_GRS_SAL) + (RATCAL(STM(Val(AA) - (Val(b) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))), Val(MS.TextMatrix(A, 19)), Val(MS.TextMatrix(A, 16)))))
            M_GRS_SAL = convert(M_GRS_SAL - (R_LST_TOT + R_CPN_AMT + BRK_LST_TOT + BRK_CPN_AMT + LKG_LST_TOT + LKG_CPN_AMT))
            M_GRS_SAL = convert(Round(M_GRS_SAL + LST_TOT + M_CPN_AMT, 1))
            M_FRE_QTY = convert(Val(M_FRE_QTY) + Val(RATCAL(STM(Val(d), Val(MS.TextMatrix(A, 16))), Val(MS.TextMatrix(A, 19)), Val(MS.TextMatrix(A, 16)))))
            M_FRE_QTY = convert(Round((M_FRE_QTY + F_LST_TOT + F_CPN_AMT), 1))
            
            'M_BRK_QTY = convert(Val(M_BRK_QTY) + Val(RATCAL(STM(Val(e), Val(MS.TextMatrix(A, 12))), Val(MS.TextMatrix(A, 15)), Val(MS.TextMatrix(A, 12)))))
            'M_LKG_QTY = convert(Val(M_LKG_QTY) + Val(RATCAL(STM(Val(f), Val(MS.TextMatrix(A, 12))), Val(MS.TextMatrix(A, 15)), Val(MS.TextMatrix(A, 12)))))
            
            MS.TextMatrix(A, 9) = QTY(STM(Val(AA) - (Val(b) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
            MS.TextMatrix(A, 13) = convert(RATCAL(QTY(Val(MS.TextMatrix(A, 9) & "")), Val(MS.TextMatrix(A, 19) & ""), Val(MS.TextMatrix(A, 16)))) ' Radhika
            MS.TextMatrix(A, 13) = convert(MS.TextMatrix(A, 13) + LST_TOT + M_CPN_AMT) 'PNR
            MS.TextMatrix(A, 11) = QTY(STM(Val(AA) - (Val(b) + Val(C) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
            If Trolley.Value = 1 Then MS.TextMatrix(A, 10) = QTY(STM(Val(AA) - (Val(b) + Val(d) + Val(e) + Val(f)), Val(MS.TextMatrix(A, 16))))
        End If
        Total = convert(amt(Val(Total) + Val(MS.TextMatrix(A, 13))))
    End If
    tot_lou = Val(tot_lou) + Val(MS.TextMatrix(A, 4))
    tot_lin = Val(tot_lin) + Val(MS.TextMatrix(A, 5))
    tot_net = Val(tot_net) + Val(MS.TextMatrix(A, 9))
    tot_var = Val(tot_var) + Val(MS.TextMatrix(A, 11))
    Tot_fre = Val(Tot_fre) + Val(MS.TextMatrix(A, 6))
Next
If Pieces.Value = 1 Then
    tot_lou = QTY(Val(tot_lou))
    tot_lin = QTY(Val(tot_lin))
    tot_net = QTY(Val(tot_net))
    tot_var = QTY(Val(tot_var))
    Tot_fre = QTY(Val(Tot_fre))
End If
'If Check3.Value = 1 And Check3.Enabled = True Then COMMISSION  ' mODIFIED RADHIKA
If Check3.Value = 1 And Check3.Enabled = True Then COMMISSION
calc
M_GRS_AMT = convert(amt(Total))
Total = convert(amt(Total))
M_NET_AMT = convert(amt(Val(M_GRS_AMT)))
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
    REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY = PRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item','POP') AND PRD_QTY > 0 ORDER BY BRW_SEQ", con, adOpenKeyset
ElseIf Check5.Value = 1 Then
    REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY = PRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item','POP') AND PRD_QTY > 0 and prd.stk_typ = 'POP' ORDER BY BRW_SEQ", con, adOpenKeyset
ElseIf Check6.Value = 1 Then
    REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY = PRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item','POP') AND PRD_QTY > 0 and prd.stk_typ = 'Promotional Item' ORDER BY BRW_SEQ", con, adOpenKeyset
ElseIf Check2.Value = 1 Then
    REC_SET.Open "SELECT * FROM brd,prd WHERE BRD.PRD_IDY = PRD.PRD_IDY AND PRD_PRN LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item','POP') AND PRD_QTY > 0 ORDER BY BRW_SEQ", con, adOpenKeyset
Else
    REC_SET.Open "SELECT * FROM brd,prd WHERE BRD.PRD_IDY = PRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item','POP') AND PRD_QTY > 0 ORDER BY BRW_SEQ", con, adOpenKeyset
End If
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(REC_SET!PRD_MRP)
        ls.ListSubItems.Add , , convert(REC_SET!PRD_SDR)
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
    Label49 = "Select Product - SKU"
    SSTab1.Tab = 3
    Check2.Caption = "Short Name"
    Check5.Visible = True
    Check6.Visible = True
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product - SKU Not Found, Define Product,SKU !", vbInformation, "DAS Version 4.0"
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
REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY = PRD.PRD_IDY AND prd.PRD_idy LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') AND PRD_QTY >0 ORDER BY BRW_SEQ", con, adOpenKeyset
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(REC_SET!PRD_MRP)
        ls.ListSubItems.Add , , convert(REC_SET!PRD_SDR)
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
    Label49 = "Select Product - SKU"
    SSTab1.Tab = 3
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product - SKU Not Found, Define Product,SKU !", vbInformation, "DAS Version 4.0"
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
If check("PRT", "PRT_NME", "M_FRM_RSI1") = True Then
M_FRM_PRV.Check3.Value = 1
End If
M_FRM_PRV.Tag = "RSI"
M_FRM_PRV.Show 1
End Sub

Public Sub setval()
If list5.Visible = False Then clrctr M_FRM_RSI1
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
M_OLD_CRD = "0.00"
M_CUR_CRD = "0.00"
M_COM_AMT = "0.00"
m_dep_ref = "0.00"
m_dep_rcd = "0.00"
M_CRT_AMT = "0.00"
TTL_CSH_VAR = "0.00"
M_NET_AMT = "0.00"
M_TAX_AMT = "0.00"
txt1000 = "0"
txt500 = "0"
txt100 = "0"
txt50 = "0"
txt20 = "0"
txt10 = "0"
txt5 = "0"
txt2 = "0"
txt1 = "0"
txtchange = "0"
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
LINDATE = Date
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
REC_SET.Open "select SUM((INT(RET_QTY)*SUB_UNT)+((RET_QTY-INT(RET_QTY))*1000)),* from VAN,PRD,BRD where inv_idy like '" & Trim(list5.SelectedItem) & "' and VAN.BAT_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY GROUP BY VAN.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
REC_SET1.Open "SELECT * FROM ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'", con, adOpenKeyset, adLockPessimistic
MS.Rows = REC_SET.RecordCount + 1
For A = 1 To REC_SET.RecordCount
    MS.TextMatrix(A, 0) = REC_SET("prd_nme")
    MS.TextMatrix(A, 1) = Val(TRP_NOS)
    If Pieces.Value = 1 Then
        MS.TextMatrix(A, 2) = QTY(STM(Val(REC_SET(0) & ""), Val(REC_SET!SUB_UNT)))
        MS.TextMatrix(A, 4) = QTY(STM(Val(REC_SET(0) & ""), Val(REC_SET!SUB_UNT)))
        MS.TextMatrix(A, 12) = convert(REC_SET("PRD_MRP"))
        If REC_SET1.RecordCount > 0 Then
            If REC_SET1("DBT_RAP") = "A" Then
                MS.TextMatrix(A, 19) = convert(Val(REC_SET!BAS_RAT) - Val(REC_SET!S_INC_DIS))
            ElseIf REC_SET1("DBT_RAP") = "B" Then
'                MS.TextMatrix(A, 19) = ((REC_SET!PRD_SDR * REC_SET!SAL_TAX) / 100) 'added by PNR
'                MS.TextMatrix(A, 19) = (MS.TextMatrix(A, 19) + (REC_SET!PRD_SDR))
                MS.TextMatrix(A, 19) = convert(Val(REC_SET!PRD_SDR) - Val(REC_SET!S_INC_DIS))
            ElseIf REC_SET1("DBT_RAP") = "C" Then
                MS.TextMatrix(A, 19) = convert(REC_SET!PRD_MRP)
            End If
        Else
'            MS.TextMatrix(A, 19) = ((REC_SET!PRD_SDR * REC_SET!SAL_TAX) / 100) 'added by PNR
'            MS.TextMatrix(A, 19) = (MS.TextMatrix(A, 19) + (REC_SET!PRD_SDR))
            MS.TextMatrix(A, 19) = convert(Val(REC_SET!PRD_SDR) - Val(REC_SET!S_INC_DIS))
        End If
    Else
        MS.TextMatrix(A, 2) = Val(REC_SET(0))
        MS.TextMatrix(A, 4) = MS.TextMatrix(A, 2)
        MS.TextMatrix(A, 12) = convert(REC_SET("prd_mrp") / Val(REC_SET!SUB_UNT))
        If REC_SET1.RecordCount > 0 Then
            If REC_SET1("DBT_RAP") = "A" Then
                MS.TextMatrix(A, 19) = (convert(Val(REC_SET!BAS_RAT) - Val(REC_SET!S_INC_DIS)) / Val(REC_SET!SUB_UNT))
            ElseIf REC_SET1("DBT_RAP") = "B" Then
'                MS.TextMatrix(A, 19) = ((REC_SET!PRD_SDR * REC_SET!SAL_TAX) / 100) 'added by PNR
'                MS.TextMatrix(A, 19) = (MS.TextMatrix(A, 19) + (REC_SET!PRD_SDR))
                MS.TextMatrix(A, 19) = (convert(Val(REC_SET!PRD_SDR) - Val(REC_SET!S_INC_DIS)) / Val(REC_SET!SUB_UNT))
            ElseIf REC_SET1("DBT_RAP") = "C" Then
                MS.TextMatrix(A, 19) = convert(Val(REC_SET!PRD_MRP) / Val(REC_SET!SUB_UNT))
            End If
        Else
'            MS.TextMatrix(A, 19) = ((REC_SET!PRD_SDR * REC_SET!SAL_TAX) / 100) 'added by PNR
'            MS.TextMatrix(A, 19) = (MS.TextMatrix(A, 19) + (REC_SET!PRD_SDR))
            MS.TextMatrix(A, 19) = (convert(Val(REC_SET!PRD_SDR) - Val(REC_SET!S_INC_DIS)) / Val(REC_SET!SUB_UNT))
        End If
    End If
    MS.TextMatrix(A, 14) = Trim(REC_SET("PRD_IDY"))
    MS.TextMatrix(A, 15) = Trim(REC_SET("PRD_PRN"))
    MS.TextMatrix(A, 16) = Trim(REC_SET("SUB_UNT"))
    MS.TextMatrix(A, 17) = Trim(REC_SET("BAT_IDY"))
    MS.TextMatrix(A, 19) = Trim(REC_SET("BAT_NME"))
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
    REC_SET.Open "select * from SMN,VAN where VAN.smn_idy = smn.smn_idy AND INV_DAT = CTOD('" & Trim(DTPicker1) & "') AND VAN_CHK= 'P' AND VAN_CLS='C' GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from SMN,VAN where VAN.smn_idy = smn.smn_idy AND INV_DAT = CTOD('" & Trim(DTPicker1) & "') AND VAN_CHK= 'V' AND VAN_CLS='C' GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
End If
list5.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET!INV_IDY))
        ls.ListSubItems.Add , , Format(REC_SET!INV_DAT, "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("SMN_NME") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Van Challan"
    list5.Visible = True
    list5.SetFocus
Else
    MsgBox "Van Challan Number Not Found!", vbInformation, "DAS Version 4.0"
    MS.SetFocus
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub
Public Function CHECKSTOCK() As Boolean
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Function
End If
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
For A = 1 To MS.Rows - 1
    If MS.TextMatrix(A, 0) <> "" Then
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM PRD,brd WHERE prd.PRD_IDY = brd.prd_idy and brd.BAT_idy like '" & Trim(MS.TextMatrix(A, 16)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            If MOD_REC = True Then
            Else
            If Pieces.Value = 1 Then
                For d = 1 To MS.Rows - 1
                    If A <> d Then
                        If Trim(MS.TextMatrix(A, 16)) = Trim(MS.TextMatrix(d, 16)) Then
                            b = Val(b) + MTS(Val(MS.TextMatrix(d, 4)), Val(MS.TextMatrix(d, 16)))
                        End If
                    End If
                Next
                b = Val(b) + MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                C = MTS(Val(rs!prd_qty), Val(MS.TextMatrix(A, 16)))
            Else
                For d = 1 To MS.Rows - 1
                    If A <> d Then
                        If Trim(MS.TextMatrix(A, 16)) = Trim(MS.TextMatrix(d, 16)) Then
                            b = Val(b) + Val(MS.TextMatrix(d, 4))
                        End If
                    End If
                Next
                b = Val(b) + Val(MS.TextMatrix(A, 4))
                C = MTS(Val(rs!prd_qty), Val(MS.TextMatrix(A, 16)))
            End If
                If Val(b) > Val(C) Then
                    MsgBox "Stock out of Range, Check Stock", vbInformation, "DAS Version 4.0"
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
Dim rs As New ADODB.Recordset
With MDIForm1
If ButtonMenu.KEY = "stkStmt" Or ButtonMenu.KEY = "VAN_CNQ" Then
    If ButtonMenu.KEY = "stkStmt" Then
       M_FRM_DSN.Tag = "stkStmt"
    Else
       M_FRM_DSN.Tag = "VAN_CNQ"
    End If
    M_FRM_DSN.Show 1
    ElseIf ButtonMenu.KEY = "VAN_SCR" Then
            M_REP_SMN.Tag = "VAN_SCR"
            M_REP_SMN.Show 1
    ElseIf ButtonMenu.KEY = "SHL_SOH_NLD" Then
            OpenReport CURDIR & "\SNN.RP1", "Shells SOH Net Load", 0, "", 6
    ElseIf ButtonMenu.KEY = "SCB_CUM" Then
            M_REP_SMN.Tag = "CUM"
            M_REP_SMN.Show 1
    ElseIf ButtonMenu.KEY = "SHT_CSH" Then
            M_REP_SMN.Tag = "SHT_CSH"
            M_REP_SMN.Show 1
    ElseIf ButtonMenu.KEY = "EXP_RPT" Then
            M_REP_SMN.Tag = "EXP_RPT"
            M_REP_SMN.Show 1
    ElseIf ButtonMenu.KEY = "SCB_SMY" Then
            M_FRM_ASD.Tag = "SCB_SMY"
            M_FRM_ASD.Show 1
Else
    If Trim(M_INV_IDY) <> "" Then
        SAL_QRY = "INL->ORD_IDY = '" & Trim(M_INV_IDY) & "' .AND. INL->TXN_TYP = 'SAL'"
        If ButtonMenu.KEY = "SAL_REP" Then
            OpenReport CURDIR & "\REPORT.RP1", "Sales Report (Van)", 2, SAL_QRY, 6
        ElseIf ButtonMenu.KEY = "SAL_REP_SKU" Then
            OpenReport CURDIR & "\REPORT.RP1", "Sales Product Van Cases", 2, SAL_QRY, 6
        ElseIf ButtonMenu.KEY = "CAS_REC" Then
            updatevan
            OpenReport CURDIR & "\REP_LIB.RP1", "Van Reconcile Cases", 0, SAL_QRY, 6
        ElseIf ButtonMenu.KEY = "UNT_REC" Then
            updatevan
            OpenReport CURDIR & "\REP_LIB.RP1", "Van Reconcile", 0, SAL_QRY, 6
        ElseIf ButtonMenu.KEY = "VAN_STL" Then
            OpenReport CURDIR & "\DOS_VCH.RP1", "Sale Estimate", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
        ElseIf ButtonMenu.KEY = "VAN_MTF" Then
            DIR_FLD_UPD
            OpenReport CURDIR & "\DOS_VCH.RP1", "Van MTF Gate Pass", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
            Dim REC_SET As New ADODB.Recordset
            If REC_SET.State = 1 Then REC_SET.Close
            REC_SET.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "' GROUP BY TRP_IDY", con, adOpenKeyset, adLockPessimistic
            If REC_SET.RecordCount > 0 Then
                While REC_SET.EOF = False
                    If Val(REC_SET("TRP_IDY")) > 3 Then
                        OpenReport CURDIR & "\DOS_VCH.RP1", "Van MTF Gate Pass Seq", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
                    End If
                    REC_SET.MoveNext
                Wend
                REC_SET.Close
            Else
                REC_SET.Close
            End If
        ElseIf ButtonMenu.KEY = "VAN_VRE" Then
            DIR_FLD_UPD
            OpenReport CURDIR & "\REP_LIB.RP1", "Van Reconcile Shells", 2, "SHL->SHL_TYP = 'Shell' .and. SNV->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
        ElseIf ButtonMenu.KEY = "VAN_FRE" Then
            DIR_FLD_UPD
            OpenReport CURDIR & "\REP_LIB.RP1", "Van Settlement Shells", 2, "SNV->TXN_IDY='" & Trim(M_INV_IDY) & "'", 6
        ElseIf ButtonMenu.KEY = "VAN_CCS" Then
            DIR_FLD_UPD
            OpenReport CURDIR & "\DOS_VCH.RP1", "Van Currency", 2, "CUR->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
        ElseIf ButtonMenu.KEY = "VAN_CHL_ALT" Then
            DIR_FLD_UPD
            OpenReport CURDIR & "\DOS_VCH.RP1", "Van Challan", 2, "VAN->INV_IDY='" & Trim(M_INV_IDY) & "'", 6
        End If
    Else
        MsgBox "Select Van Challan for Reports Print !", vbInformation, "DAS Version 4.0"
        CANCELRECORD
    End If
End If
'CANCELRECORD
End With
End Sub
Public Sub updatevan()
If CHECKTB("cld.dbf,dir.dbf,van.dbf,prd.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim RSCLD As New ADODB.Recordset
Dim rsVan As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
con.Execute "DELETE FROM CLD"
con.Execute "update dir set dir_msg = '" & Trim(cmbRoute.Text) & "' where dir_idy = 'ROU' and dir_idx = 'IDY'"
con.Execute "update dir set dir_msg = '" & M_SMN_NME & "' where dir_idy = 'SMN' and dir_idx = 'IDY'"
con.Execute "update dir set dir_msg = '" & M_CST_NME & "' where dir_idy = 'DLR' and dir_idx = 'IDY'"
con.Execute "update dir set dir_msg = '" & M_INV_IDY & "' where dir_idy = 'INV' and dir_idx = 'NUM'"
con.Execute "update dir set dir_msg = '" & Format(M_INV_DAT, "dd/MM/yyyy") & "' where dir_idy = 'HDR' and dir_idx = 'DAT'"
If rsVan.State = 1 Then rsVan.Close
rsVan.Open "SELECT SUM(((tot_qty- int(tot_qty))*1000)+ (int(tot_qty)*sub_unt)) p_tot_qty, SUM(((fre_qty- int(fre_qty))*1000)+ (int(fre_qty)*sub_unt)) p_fre_qty, SUM(((ret_qty- int(ret_qty))*1000)+ (int(ret_qty)*sub_unt)+((brk_qty- int(brk_qty))*1000)+ (int(brk_qty)*sub_unt)+((lkg_qty- int(lkg_qty))*1000)+ (int(lkg_qty)*sub_unt)) p_ret_qty,*  FROM VAN,PRD WHERE PRD.PRD_IDY = VAN.PRD_IDY AND INV_IDY LIKE '" & Trim(M_INV_IDY) & "' group by VAN.BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rsVan.RecordCount
    con.Execute "INSERT INTO CLD VALUES('" & Trim(rsVan!BAT_IDY) & "'," & Val(rsVan!p_tot_qty) & "," & Val(rsVan!p_ret_qty) & ",0,0,0," & Val(rsVan!P_FRE_QTY) & ",0,0)"
    'changed rsVan!BAT_IDY to rsVan!PRD_IDY by PNR
    rsVan.MoveNext
Next
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT SUM(TOT_PCS),BAT_IDY,STK_TYP FROM INV WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "' AND TXN_TYP = 'SAL' AND STK_TYP IN('Salable','Free','Damage','TakeBack','Breakage','Leakage') GROUP BY BAT_IDY,STK_TYP", con, adOpenKeyset, adLockPessimistic
'RSINV.Open "SELECT SUM(TOT_PCS),PRD_IDY,STK_TYP FROM INV WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "' AND STK_TYP IN('Salable','Free','Damage','TakeBack','Breakage','Leakage') GROUP BY PRD_IDY,STK_TYP", CON, adOpenKeyset, adlockpessimistic
For A = 1 To RSINV.RecordCount
    If Trim(RSINV(2)) = "Salable" Then
        con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!BAT_IDY) & "'", d
        If d = 0 Then con.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!BAT_IDY) & "',0,0," & Val(RSINV(0)) & ",0,0,0,0,0)"
'        CON.Execute "UPDATE CLD SET PRT_STK= PRT_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!prd_idy) & "'", d
'        If d = 0 Then CON.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!prd_idy) & "',0,0," & Val(RSINV(0)) & ",0,0,0,0,0)"
    ElseIf Trim(RSINV(2)) = "Free" Then
        con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!BAT_IDY) & "'", d
        If d = 0 Then con.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!BAT_IDY) & "',0,0," & Val(RSINV(0)) & ",0,0,0,0,0)"
    Else
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!BAT_IDY) & "'", d
        If d = 0 Then con.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!BAT_IDY) & "',0,0,0," & Val(RSINV(0)) & ",0,0,0,0)"
'        CON.Execute "UPDATE CLD SET SAL_STK= SAL_STK + " & Val(RSINV(0)) & "WHERE PRD_IDY LIKE '" & Trim(RSINV!prd_idy) & "'", d
'        If d = 0 Then CON.Execute "INSERT INTO CLD VALUES('" & Trim(RSINV!prd_idy) & "',0,0,0," & Val(RSINV(0)) & ",0,0,0,0)"
    End If
    RSINV.MoveNext
Next
End Sub
Public Sub DIR_FLD_UPD()
con.Execute "update dir set dir_msg = '" & M_SMN_NME & "' where dir_idy = 'SMN' and dir_idx = 'IDY'"
con.Execute "update dir set dir_msg = '" & M_CST_NME & "' where dir_idy = 'DLR' and dir_idx = 'IDY'"
con.Execute "update dir set dir_msg = '" & M_INV_IDY & "' where dir_idy = 'INV' and dir_idx = 'NUM'"
con.Execute "update dir set dir_msg = '" & Format(M_INV_DAT, "dd/MM/yyyy") & "' where dir_idy = 'HDR' and dir_idx = 'DAT'"
End Sub

Private Sub Trolley_Click()
If Trolley.Value = 1 Then
   Check4.Enabled = False
   TRP_NOS.Enabled = False
   MS.ColWidth(11) = 0
   MS.ColWidth(10) = 1050
Else
    TRP_NOS.Enabled = True
    MS.ColWidth(11) = 1050
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
'On Error Resume Next
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbInformation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If

If CHECKTB("smn.dbf,act.dbf,prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("smn", "smn_idy", Trim(M_SMN_IDY)) = False Then
    MsgBox "Salesman Not Found, Define Salesman !", vbInformation, "DAS Version 4.0"
    M_SMN_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If check("act", "act_idy", Trim(M_CST_IDY)) = False Then
    MsgBox "Account Not Found, Define Account !", vbInformation, "DAS Version 4.0"
    M_CST_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If Check3.Enabled = False And Check4.Enabled = False Then
    MsgBox "Challan Closed ! Changes Not Saved!", vbInformation, "DAS Version 4.0"
    INC = Trim(M_INV_IDY)
    If check("PRT", "PRT_NME", "M_FRM_RSI1") = True Then
        con.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
        con.Execute "update dir where dir_idy like 'REM' set dir_als = 'RSI'"
        If Pieces.Value = 0 Then
            OpenReport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Pieces)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Cases)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        End If
    End If
    CANCELRECORD
    Exit Sub
End If

Dim RSINL As New ADODB.Recordset
Dim REC_PNL As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
Dim REC_IRT As New ADODB.Recordset
If MsgBox("Convert to Bill", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
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
REC_SET.Open "select * from act where grp_clf like 'Sales' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
ListView1.ListItems.CLEAR
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
        RSINL.Open "select * from inl where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
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
        RSINL("ROU_IDY") = GetValue("ROU_IDY", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'")
        RATEVAL = GetValue("DBT_RAP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'")
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
            REC_IRT.Open "select * from IRT", con, adOpenKeyset, adLockPessimistic
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
    RSINV.Open "select * from inv where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
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
            RSINV!SUB_UNT = Val(MS.TextMatrix(A, 12))
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
                con.Execute "update PRD where prd_idy = '" & Trim(MS.TextMatrix(A, 10)) & "' set CUR_LEV = " & CONQTY(Trim(MS.TextMatrix(A, 6)), Trim(MS.TextMatrix(A, 10))), d
                con.Execute "UPDATE PRD WHERE PRD_IDY = '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 12))), d
                con.Execute "UPDATE BRD WHERE BAT_IDY = '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(Trim(MS.TextMatrix(A, 6)), Trim(MS.TextMatrix(A, 13))), d
            Else
                con.Execute "update PRD where prd_idy = '" & Trim(MS.TextMatrix(A, 10)) & "' set CUR_LEV = " & CONQTY(QTY(STM(MS.TextMatrix(A, 6), Val(MS.TextMatrix(A, 12)))), Trim(MS.TextMatrix(A, 10)))
                con.Execute "UPDATE PRD WHERE PRD_IDY = '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS - " & Val(MS.TextMatrix(A, 6))
                con.Execute "UPDATE BRD WHERE BAT_IDY = '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(STM(MS.TextMatrix(A, 6), Val(MS.TextMatrix(A, 12))), Trim(MS.TextMatrix(A, 13)))
            End If
        End If
    Next
    'PRINT DATABASE LOGIC BEGIN
    If MOD_REC = True Then
        TxnMoveBackup "PNV", "XNV", Trim(INC), "SAL"
        TxnMoveBackup "PNL", "XNL", Trim(INC), "SAL"
        
        PNL.Execute "delete from pnv where inv_idy = '" & Trim(INC) & "'"
        PNL.Execute "delete from pnl where inv_idy = '" & Trim(INC) & "'"
    End If
    REC_PNL.Open "SELECT * FROM PNL", PNL, adOpenKeyset, adLockPessimistic
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
    REC_PNL!SMN_NME = GetValue("smn_nme", "smn where smn_idy = '" & Trim(M_SMN_IDY) & "'")
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
    REC_PNL!ins_nme = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'TT1'")
    REC_PNL!ins_adr = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'TT2'")
    REC_PNL!sys_rm1 = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'RM1'")
    REC_PNL!sys_rm2 = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'RM2'")
    REC_PNL!sys_rm3 = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'RM3'")
    REC_PNL!sys_rm4 = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'RM4'")
    REC_PNL!ins_gst = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'GST'")
    REC_PNL!ins_cst = GetValue("dir_msg ", "dir where dir_idy = 'SYS' AND DIR_IDX = 'CST'")
    REC_PNL!FRM_DAT = Date
    REC_PNL!TO_DAT = Date
    REC_PNL!rou_act = ""
    REC_PNL!dch_dat = Date
    REC_PNL!DIS_CHK = ""
    REC_PNL.Update
    REC_PNL.Close
    If REC_PNL.State = 1 Then REC_PNL.Close
    REC_PNL.Open "select * from pnv where inv_idy = ''", PNL, adOpenKeyset, adLockPessimistic
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
            REC_PNL!SUB_UNT = Val(MS.TextMatrix(A, 12))
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
    MsgBox "New Invoice ID " & INC, vbInformation
    BMAKEBILL = True
Else
End If
If REC_SET1.State = 1 Then REC_SET1.Close
If REC_SET.State = 1 Then REC_SET.Close
If MOD_REC = True Then
    If Check3.Value = 1 And Check3.Enabled = False Then
    Else
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from van,PRD where VAN.PRD_IDY = PRD.PRD_IDY AND inv_idy = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset
    For A = 1 To REC_SET.RecordCount
    con.Execute "update prd where prd_idy = '" & REC_SET("prd_idy") & "' set cur_lev = " & ADDQTY(Val(REC_SET("TOT_qty")), Trim(REC_SET("prd_idy")))
    con.Execute "UPDATE PRD WHERE PRD_IDY = '" & REC_SET("prd_idy") & "' set PRD_PCS = PRD_PCS + " & MTS(Val(REC_SET("TOT_qty")), Val(REC_SET("SUB_UNT")))
    con.Execute "update Brd where BAT_idy = '" & REC_SET("BAT_idy") & "' set PRD_QTY = " & ADDBRDQTY(Val(REC_SET("TOT_qty")), Trim(REC_SET("BAT_idy")))
    REC_SET.MoveNext
    Next
    End If
    
    TxnMoveBackup "VAN", "VBK", Trim(M_INV_IDY), ""
    TxnMoveBackup "CUR", "CBK", Trim(M_INV_IDY), ""
    
    con.Execute "delete from van where inv_idy = '" & Trim(M_INV_IDY) & "'", d
    con.Execute "delete from cur where inv_idy = '" & Trim(M_INV_IDY) & "'", d
    INC = Trim(M_INV_IDY)
Else
    If Trim(MS.TextMatrix(1, 0)) <> "" Then
        INC = "RSI" & T7increment("IDY_RSI")
        INVIDY = INC
        If Trim(M_INV_IDY) = "" And Trolley.Value = 1 Then
            con.Execute "UPDATE INL SET ORD_IDY = '" & Trim(INC) & "' WHERE INV_IDY = '" & Trim(invid) & "'"
        End If
    Else
         CANCELRECORD
        Exit Sub
    End If
End If
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from van where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
If RSCUR.State = 1 Then RSCUR.Close
RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY=''", con, adOpenKeyset, adLockPessimistic
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
        REC_SET("rou_idy") = Trim(cmbRoute.Text & "")
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
            REC_SET!vAN_CLS = "S"
        Else
            REC_SET!vAN_CLS = "N"
        End If
        If Check4.Value = 1 Or save_rec = "Save" Then
        save_rec = ""
        REC_SET!vAN_CLS = "C"
        End If
        REC_SET!COL_AMT = Val(MS.TextMatrix(A, 9))
        REC_SET!CAS_AMT = Val(M_CAS_AMT)
        REC_SET!CHQ_AMT = Val(M_CHQ_AMT)
        REC_SET!OLD_CRD = Val(M_OLD_CRD)
        REC_SET!CUR_CRD = Val(M_CUR_CRD)
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
            con.Execute "update prd where prd_idy = '" & Trim(MS.TextMatrix(A, 10)) & "' set cur_lev = " & CONQTY(QTY(Trim(MS.TextMatrix(A, 4) & "")), Trim(MS.TextMatrix(A, 10)))
            con.Execute "UPDATE PRD WHERE PRD_IDY = '" & Trim(MS.TextMatrix(A, 10)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 12)))
            con.Execute "UPDATE BRD WHERE BAT_IDY = '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(QTY(Val(MS.TextMatrix(A, 4))), Trim(MS.TextMatrix(A, 13)))
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
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 11)) <> 0 Or (Val(MSS.TextMatrix(9, 3)) < -9 Or Val(MSS.TextMatrix(9, 3)) > 9) Then
        RSCUR(11) = 1
    Else
        RSCUR(11) = 0
    End If
Next

RSCUR.Update    'END
If REC_SET1.State = 1 Then REC_SET1.Close
If MOD_REC = False Then MsgBox "New Challan ID is " & INC, vbInformation, "DAS Version 4.0"
Label23 = labdisp("idy_rsi")
VAN_PRT_CHK:
If CHECKTB("prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_RSI1") = True Then
    con.Execute "update dir where dir_idy = 'INV' set dir_als = '" & INC & "'"
    con.Execute "update dir where dir_idy = 'REM' set dir_als = 'RSI'"
    If Trolley.Value = 1 Then
        If Pieces.Value = 0 Then
            OpenReport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Pieces)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        Else
            OpenReport CURDIR & "\DOS_VCH.RP1", "Push Cart Challan (Cases)", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
        End If
    Else
        OpenReport CURDIR & "\DOS_VCH.RP1", "Van Challan", 2, "VAN->INV_IDY='" & Trim(INC) & "'", 6
    End If
    CANCELRECORD
End If
CANCELRECORD
End Sub
Public Sub ChallanCopy()
If Trim(M_INV_IDY) <> "" Then
    Dim CAS_VAN As New ADODB.Recordset
    Dim SRC_VAN As New ADODB.Recordset
    Dim SHL_CAS_VAN As New ADODB.Recordset
    Dim SHL_SRC_VAN As New ADODB.Recordset
    Dim BAT_STK As New ADODB.Recordset
    Dim PRD_STK As New ADODB.Recordset
    
    If BAT_STK.State = 1 Then BAT_STK.Close
    BAT_STK.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    
    If PRD_STK.State = 1 Then PRD_STK.Close
    PRD_STK.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    
    If SHL_CAS_VAN.State = 1 Then SHL_CAS_VAN.Close
    SHL_CAS_VAN.Open "SELECT * FROM SNV WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    
    If SHL_SRC_VAN.State = 1 Then SHL_SRC_VAN.Close
    SHL_SRC_VAN.Open "SELECT * FROM SNV WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    
    If CAS_VAN.State = 1 Then CAS_VAN.Close
    CAS_VAN.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    
    If SRC_VAN.State = 1 Then SRC_VAN.Close
    SRC_VAN.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If SRC_VAN.EOF = False Then SRC_VAN.MoveFirst
    
    If MsgBox("Confirm Challan Copy and Save ? ", vbYesNo + vbInformation, "DAS Version 4.0") = vbYes Then
        INC = "RSI" & T7increment("IDY_RSI")
        INVIDY = INC
        While SRC_VAN.EOF = False
            PRD_STK.MoveFirst
            PRD_STK.Find "PRD_IDY='" & Trim(SRC_VAN("PRD_IDY")) & "'"
            If PRD_STK.EOF = False Then
                BAT_STK.MoveFirst
                BAT_STK.Find "BAT_IDY='" & Trim(SRC_VAN("BAT_IDY")) & "'"
                If BAT_STK.EOF = False Then
                    PRD_STK.Find "PRD_IDY='" & Trim(SRC_VAN("PRD_IDY")) & "'"
                    CAS_QTY = MTS(SRC_VAN("ISS_QTY"), PRD_STK("SUB_UNT"))
                    CUR_QTY = MTS(BAT_STK("PRD_QTY"), PRD_STK("SUB_UNT"))
                    If CAS_QTY <= CUR_QTY Then
                        UPD_QTY = STM(Val(CUR_QTY) - Val(CAS_QTY), PRD_STK("SUB_UNT"))
                        BAT_STK("PRD_QTY") = UPD_QTY
                        BAT_STK.Update
                        PRD_STK.MoveFirst
                        PRD_STK.Find "PRD_IDY='" & Trim(BAT_STK("PRD_IDY")) & "'"
                        If PRD_STK.EOF = False Then
                            P_CUR_QTY = MTS(PRD_STK("CUR_LEV"), PRD_STK("SUB_UNT"))
                            P_UPD_QTY = STM(Val(P_CUR_QTY) - Val(CAS_QTY), PRD_STK("SUB_UNT"))
                            PRD_STK("CUR_LEV") = P_UPD_QTY
                            PRD_STK.Update
                        End If
                        CAS_VAN.AddNew
                        CAS_VAN("INV_IDY") = Trim(INC)
                        CAS_VAN("INV_DAT") = Date
                        CAS_VAN("PRD_IDY") = SRC_VAN("PRD_IDY")
                        CAS_VAN("OPN_QTY") = 0#
                        CAS_VAN("ISS_QTY") = SRC_VAN("ISS_QTY")
                        CAS_VAN("TOT_QTY") = SRC_VAN("TOT_QTY")
                        CAS_VAN("RET_QTY") = 0#
                        CAS_VAN("BIL_QTY") = 0#
                        CAS_VAN("FRE_QTY") = 0#
                        CAS_VAN("LKG_QTY") = 0#
                        CAS_VAN("BRK_QTY") = 0#
                        CAS_VAN("CAS_AMT") = 0#
                        CAS_VAN("CHQ_AMT") = 0#
                        CAS_VAN("DIS_AMT") = 0#
                        CAS_VAN("EXP_AMT") = 0#
                        CAS_VAN("ADD_AMT") = 0#
                        CAS_VAN("COM_AMT") = 0#
                        CAS_VAN("DEP_RCD") = 0#
                        CAS_VAN("DEP_REF") = 0#
                        CAS_VAN("ACT_IDY") = ""
                        CAS_VAN("CST_IDY") = SRC_VAN("CST_IDY")
                        CAS_VAN("BAT_IDY") = SRC_VAN("BAT_IDY")
                        CAS_VAN("SMN_IDY") = SRC_VAN("SMN_IDY")
                        CAS_VAN("VAN_CLS") = "N"
                        CAS_VAN("VAN_CHK") = SRC_VAN("VAN_CHK")
                        CAS_VAN("PRD_RAT") = SRC_VAN("PRD_RAT")
                        CAS_VAN("COL_AMT") = SRC_VAN("COL_AMT")
                        CAS_VAN("TRP_IDY") = 1
                        CAS_VAN("ROU_IDY") = SRC_VAN("ROU_IDY")
                        CAS_VAN("LIN_DAT") = SRC_VAN("LIN_DAT")
                        CAS_VAN("MNL_REF") = SRC_VAN("MNL_REF")
                        CAS_VAN.Update
                    Else
                        CAS_VAN.AddNew
                        CAS_VAN("INV_IDY") = Trim(INC)
                        CAS_VAN("INV_DAT") = Date
                        CAS_VAN("PRD_IDY") = SRC_VAN("PRD_IDY")
                        CAS_VAN("OPN_QTY") = 0#
                        CAS_VAN("ISS_QTY") = 0#
                        CAS_VAN("TOT_QTY") = 0#
                        CAS_VAN("RET_QTY") = 0#
                        CAS_VAN("BIL_QTY") = 0#
                        CAS_VAN("FRE_QTY") = 0#
                        CAS_VAN("LKG_QTY") = 0#
                        CAS_VAN("BRK_QTY") = 0#
                        CAS_VAN("CAS_AMT") = 0#
                        CAS_VAN("CHQ_AMT") = 0#
                        CAS_VAN("DIS_AMT") = 0#
                        CAS_VAN("EXP_AMT") = 0#
                        CAS_VAN("ADD_AMT") = 0#
                        CAS_VAN("COM_AMT") = 0#
                        CAS_VAN("DEP_RCD") = 0#
                        CAS_VAN("DEP_REF") = 0#
                        CAS_VAN("ACT_IDY") = ""
                        CAS_VAN("CST_IDY") = SRC_VAN("CST_IDY")
                        CAS_VAN("BAT_IDY") = SRC_VAN("BAT_IDY")
                        CAS_VAN("SMN_IDY") = SRC_VAN("SMN_IDY")
                        CAS_VAN("VAN_CLS") = "N"
                        CAS_VAN("VAN_CHK") = SRC_VAN("VAN_CHK")
                        CAS_VAN("PRD_RAT") = SRC_VAN("PRD_RAT")
                        CAS_VAN("COL_AMT") = SRC_VAN("COL_AMT")
                        CAS_VAN("TRP_IDY") = 1
                        CAS_VAN("ROU_IDY") = SRC_VAN("ROU_IDY")
                        CAS_VAN("LIN_DAT") = SRC_VAN("LIN_DAT")
                        CAS_VAN("MNL_REF") = SRC_VAN("MNL_REF")
                        CAS_VAN.Update
                    End If
                End If
            End If
            SRC_VAN.MoveNext
        Wend
        
        While SHL_SRC_VAN.EOF = False
            SHL_CAS_VAN.AddNew
            SHL_CAS_VAN("INV_IDY") = Trim(INC)
            SHL_CAS_VAN("INV_DAT") = Date
            SHL_CAS_VAN("TXN_TYP") = "VAN"
            SHL_CAS_VAN("SHL_IDY") = SHL_SRC_VAN("SHL_IDY")
            SHL_CAS_VAN("PRD_IDY") = ""
            SHL_CAS_VAN("I_SHL_NOS") = SHL_SRC_VAN("I_SHL_NOS")
            SHL_CAS_VAN("I_PCK_NOS") = SHL_SRC_VAN("I_PCK_NOS")
            SHL_CAS_VAN("R_SHL_NOS") = 0#
            SHL_CAS_VAN("R_PCK_NOS") = 0#
            SHL_CAS_VAN("D_SHL_NOS") = 0#
            SHL_CAS_VAN("D_PCK_NOS") = 0#
            SHL_CAS_VAN("BRK_SHL") = 0#
            SHL_CAS_VAN("BRK_PCK") = 0#
            SHL_CAS_VAN("TXN_DAT") = SHL_SRC_VAN("TXN_DAT")
            SHL_CAS_VAN("TXN_IDY") = SHL_SRC_VAN("TXN_IDY")
            SHL_CAS_VAN("CST_IDY") = SHL_SRC_VAN("CST_IDY")
            SHL_CAS_VAN("STK_UPD") = SHL_SRC_VAN("STK_UPD")
            SHL_CAS_VAN("PRD_UPC") = SHL_SRC_VAN("PRD_UPC")
            SHL_CAS_VAN("PRD_PCK") = SHL_SRC_VAN("PRD_PCK")
            SHL_CAS_VAN("SMN_IDY") = SHL_SRC_VAN("SMN_IDY")
            SHL_CAS_VAN("ROU_IDY") = SHL_SRC_VAN("ROU_IDY")
            SHL_CAS_VAN("S_SHL_NOS") = 0#
            SHL_CAS_VAN("S_PCK_NOS") = 0#
            SHL_CAS_VAN.Update
            SHL_SRC_VAN.MoveNext
        Wend
        
        If RSCUR.State = 1 Then RSCUR.Close
            RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY=''", con, adOpenKeyset, adLockPessimistic
            RSCUR.AddNew    ' edited by radhika
            RSCUR(0) = Trim(INC)
            RSCUR(1) = Val(Trim(txt1000) & "")
            RSCUR(2) = Val(Trim(txt500) & "")
            RSCUR(3) = Val(Trim(txt100) & "")
            RSCUR(4) = Val(Trim(txt50) & "")
            RSCUR(5) = Val(Trim(txt20) & "")
            RSCUR(6) = Val(Trim(txt10) & "")
            RSCUR(7) = Val(Trim(txt5) & "")
            RSCUR(8) = Val(Trim(txt2) & "")
            RSCUR(9) = Val(Trim(txt1) & "")
            RSCUR(10) = Val(Trim(txtchange) & "")
            For A = 1 To MS.Rows - 1
                If Val(MS.TextMatrix(A, 11)) <> 0 Or (Val(MSS.TextMatrix(9, 3)) < -9 Or Val(MSS.TextMatrix(9, 3)) > 9) Then
                    RSCUR(11) = 1
                Else
                    RSCUR(11) = 0
                End If
            Next
            RSCUR.Update    ' END
            MsgBox "Newly Created Challan ID is " & INC, vbInformation, "DAS Version 4.0"
            CANCELRECORD
    Else
        CANCELRECORD
    End If
Else
    MsgBox "Select Existing Van Challan for Challan Cascade - Copy !", vbInformation, "DAS Version 4.0"
    CANCELRECORD
End If
End Sub
Public Sub ChallanReturnCopy()
If Trim(M_INV_IDY) <> "" Then
    Dim CAS_VAN As New ADODB.Recordset
    Dim SRC_VAN As New ADODB.Recordset
    Dim SHL_CAS_VAN As New ADODB.Recordset
    Dim SHL_SRC_VAN As New ADODB.Recordset
    Dim BAT_STK As New ADODB.Recordset
    Dim PRD_STK As New ADODB.Recordset
    
    If BAT_STK.State = 1 Then BAT_STK.Close
    BAT_STK.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    
    If PRD_STK.State = 1 Then PRD_STK.Close
    PRD_STK.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    
    If SHL_CAS_VAN.State = 1 Then SHL_CAS_VAN.Close
    SHL_CAS_VAN.Open "SELECT * FROM SNV WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    
    If SHL_SRC_VAN.State = 1 Then SHL_SRC_VAN.Close
    SHL_SRC_VAN.Open "SELECT * FROM SNV WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    
    If CAS_VAN.State = 1 Then CAS_VAN.Close
    CAS_VAN.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    
    If SRC_VAN.State = 1 Then SRC_VAN.Close
    SRC_VAN.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If SRC_VAN.EOF = False Then SRC_VAN.MoveFirst
    
    If MsgBox("Confirm Challan Copy and Save ? ", vbYesNo + vbInformation, "DAS Version 4.0") = vbYes Then
        INC = "RSI" & T7increment("IDY_RSI")
        INVIDY = INC
        While SRC_VAN.EOF = False
            PRD_STK.MoveFirst
            PRD_STK.Find "PRD_IDY='" & Trim(SRC_VAN("PRD_IDY")) & "'"
            If PRD_STK.EOF = False Then
                BAT_STK.MoveFirst
                BAT_STK.Find "BAT_IDY='" & Trim(SRC_VAN("BAT_IDY")) & "'"
                If BAT_STK.EOF = False Then
                    PRD_STK.Find "PRD_IDY='" & Trim(SRC_VAN("PRD_IDY")) & "'"
                    CAS_QTY = MTS(SRC_VAN("ISS_QTY"), PRD_STK("SUB_UNT"))
                    CUR_QTY = MTS(BAT_STK("PRD_QTY"), PRD_STK("SUB_UNT"))
                    If CAS_QTY <= CUR_QTY Then
                        UPD_QTY = STM(Val(CUR_QTY) - Val(CAS_QTY), PRD_STK("SUB_UNT"))
                        BAT_STK("PRD_QTY") = UPD_QTY
                        BAT_STK.Update
                        PRD_STK.MoveFirst
                        PRD_STK.Find "PRD_IDY='" & Trim(BAT_STK("PRD_IDY")) & "'"
                        If PRD_STK.EOF = False Then
                            P_CUR_QTY = MTS(PRD_STK("CUR_LEV"), PRD_STK("SUB_UNT"))
                            P_UPD_QTY = STM(Val(P_CUR_QTY) - Val(CAS_QTY), PRD_STK("SUB_UNT"))
                            PRD_STK("CUR_LEV") = P_UPD_QTY
                            PRD_STK.Update
                        End If
'                        If Val(SRC_VAN("RET_QTY")) > 0 Then
                            CAS_VAN.AddNew
                            CAS_VAN("INV_IDY") = Trim(INC)
                            CAS_VAN("INV_DAT") = Date
                            CAS_VAN("PRD_IDY") = SRC_VAN("PRD_IDY")
                            CAS_VAN("OPN_QTY") = 0#
                            CAS_VAN("ISS_QTY") = SRC_VAN("RET_QTY")
                            CAS_VAN("TOT_QTY") = SRC_VAN("RET_QTY")
                            CAS_VAN("RET_QTY") = 0#
                            CAS_VAN("BIL_QTY") = 0#
                            CAS_VAN("FRE_QTY") = 0#
                            CAS_VAN("LKG_QTY") = 0#
                            CAS_VAN("BRK_QTY") = 0#
                            CAS_VAN("CAS_AMT") = 0#
                            CAS_VAN("CHQ_AMT") = 0#
                            CAS_VAN("DIS_AMT") = 0#
                            CAS_VAN("EXP_AMT") = 0#
                            CAS_VAN("ADD_AMT") = 0#
                            CAS_VAN("COM_AMT") = 0#
                            CAS_VAN("DEP_RCD") = 0#
                            CAS_VAN("DEP_REF") = 0#
                            CAS_VAN("ACT_IDY") = ""
                            CAS_VAN("CST_IDY") = SRC_VAN("CST_IDY")
                            CAS_VAN("BAT_IDY") = SRC_VAN("BAT_IDY")
                            CAS_VAN("SMN_IDY") = SRC_VAN("SMN_IDY")
                            CAS_VAN("VAN_CLS") = "N"
                            CAS_VAN("VAN_CHK") = SRC_VAN("VAN_CHK")
                            CAS_VAN("PRD_RAT") = SRC_VAN("PRD_RAT")
                            CAS_VAN("COL_AMT") = SRC_VAN("COL_AMT")
                            CAS_VAN("TRP_IDY") = 1
                            CAS_VAN("ROU_IDY") = SRC_VAN("ROU_IDY")
                            CAS_VAN("LIN_DAT") = SRC_VAN("LIN_DAT")
                            CAS_VAN("MNL_REF") = SRC_VAN("MNL_REF")
                            CAS_VAN.Update
 '                       End If
                    Else
 '                       If Val(SRC_VAN("RET_QTY")) > 0 Then
                            CAS_VAN.AddNew
                            CAS_VAN("INV_IDY") = Trim(INC)
                            CAS_VAN("INV_DAT") = Date
                            CAS_VAN("PRD_IDY") = SRC_VAN("PRD_IDY")
                            CAS_VAN("OPN_QTY") = 0#
                            CAS_VAN("ISS_QTY") = 0#
                            CAS_VAN("TOT_QTY") = 0#
                            CAS_VAN("RET_QTY") = 0#
                            CAS_VAN("BIL_QTY") = 0#
                            CAS_VAN("FRE_QTY") = 0#
                            CAS_VAN("LKG_QTY") = 0#
                            CAS_VAN("BRK_QTY") = 0#
                            CAS_VAN("CAS_AMT") = 0#
                            CAS_VAN("CHQ_AMT") = 0#
                            CAS_VAN("DIS_AMT") = 0#
                            CAS_VAN("EXP_AMT") = 0#
                            CAS_VAN("ADD_AMT") = 0#
                            CAS_VAN("COM_AMT") = 0#
                            CAS_VAN("DEP_RCD") = 0#
                            CAS_VAN("DEP_REF") = 0#
                            CAS_VAN("ACT_IDY") = ""
                            CAS_VAN("CST_IDY") = SRC_VAN("CST_IDY")
                            CAS_VAN("BAT_IDY") = SRC_VAN("BAT_IDY")
                            CAS_VAN("SMN_IDY") = SRC_VAN("SMN_IDY")
                            CAS_VAN("VAN_CLS") = "N"
                            CAS_VAN("VAN_CHK") = SRC_VAN("VAN_CHK")
                            CAS_VAN("PRD_RAT") = SRC_VAN("PRD_RAT")
                            CAS_VAN("COL_AMT") = SRC_VAN("COL_AMT")
                            CAS_VAN("TRP_IDY") = 1
                            CAS_VAN("ROU_IDY") = SRC_VAN("ROU_IDY")
                            CAS_VAN("LIN_DAT") = SRC_VAN("LIN_DAT")
                            CAS_VAN("MNL_REF") = SRC_VAN("MNL_REF")
                            CAS_VAN.Update
'                        End If
                    End If
                End If
            End If
            SRC_VAN.MoveNext
        Wend
        
        While SHL_SRC_VAN.EOF = False
            SHL_CAS_VAN.AddNew
            SHL_CAS_VAN("INV_IDY") = Trim(INC)
            SHL_CAS_VAN("INV_DAT") = Date
            SHL_CAS_VAN("TXN_TYP") = "VAN"
            SHL_CAS_VAN("SHL_IDY") = SHL_SRC_VAN("SHL_IDY")
            SHL_CAS_VAN("PRD_IDY") = ""
            SHL_CAS_VAN("I_SHL_NOS") = SHL_SRC_VAN("R_SHL_NOS") 'changed to return value
            SHL_CAS_VAN("I_PCK_NOS") = SHL_SRC_VAN("R_PCK_NOS") 'changed to return value
            SHL_CAS_VAN("R_SHL_NOS") = 0#
            SHL_CAS_VAN("R_PCK_NOS") = 0#
            SHL_CAS_VAN("D_SHL_NOS") = 0#
            SHL_CAS_VAN("D_PCK_NOS") = 0#
            SHL_CAS_VAN("BRK_SHL") = 0#
            SHL_CAS_VAN("BRK_PCK") = 0#
            SHL_CAS_VAN("TXN_DAT") = SHL_SRC_VAN("TXN_DAT")
            SHL_CAS_VAN("TXN_IDY") = SHL_SRC_VAN("TXN_IDY")
            SHL_CAS_VAN("CST_IDY") = SHL_SRC_VAN("CST_IDY")
            SHL_CAS_VAN("STK_UPD") = SHL_SRC_VAN("STK_UPD")
            SHL_CAS_VAN("PRD_UPC") = SHL_SRC_VAN("PRD_UPC")
            SHL_CAS_VAN("PRD_PCK") = SHL_SRC_VAN("PRD_PCK")
            SHL_CAS_VAN("SMN_IDY") = SHL_SRC_VAN("SMN_IDY")
            SHL_CAS_VAN("ROU_IDY") = SHL_SRC_VAN("ROU_IDY")
            SHL_CAS_VAN("S_SHL_NOS") = 0#
            SHL_CAS_VAN("S_PCK_NOS") = 0#
            SHL_CAS_VAN.Update
            SHL_SRC_VAN.MoveNext
        Wend
        
        If RSCUR.State = 1 Then RSCUR.Close
            RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY=''", con, adOpenKeyset, adLockPessimistic
            RSCUR.AddNew    ' edited by radhika
            RSCUR(0) = Trim(INC)
            RSCUR(1) = Val(Trim(txt1000) & "")
            RSCUR(2) = Val(Trim(txt500) & "")
            RSCUR(3) = Val(Trim(txt100) & "")
            RSCUR(4) = Val(Trim(txt50) & "")
            RSCUR(5) = Val(Trim(txt20) & "")
            RSCUR(6) = Val(Trim(txt10) & "")
            RSCUR(7) = Val(Trim(txt5) & "")
            RSCUR(8) = Val(Trim(txt2) & "")
            RSCUR(9) = Val(Trim(txt1) & "")
            RSCUR(10) = Val(Trim(txtchange) & "")
            For A = 1 To MS.Rows - 1
                If Val(MS.TextMatrix(A, 11)) <> 0 Or (Val(MSS.TextMatrix(9, 3)) < -9 Or Val(MSS.TextMatrix(9, 3)) > 9) Then
                    RSCUR(11) = 1
                Else
                    RSCUR(11) = 0
                End If
            Next
            RSCUR.Update    ' END
            MsgBox "Newly Created Challan ID is " & INC, vbInformation, "DAS Version 4.0"
            CANCELRECORD
    Else
        CANCELRECORD
    End If
Else
    MsgBox "Select Existing Van Challan for Challan Cascade - Copy !", vbInformation, "DAS Version 4.0"
    CANCELRECORD
End If
End Sub
Public Sub RETURNSTOCKS()
Dim rsprd As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
If MOD_REC = True Then
    If CHECKTB("inv.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    For b = 1 To MS.Rows - 1
        If Pieces.Value = 0 Then
            MS.TextMatrix(b, 10) = 0
        Else
            MS.TextMatrix(b, 10) = "0.000"
        End If
    Next
    RSINV.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT,PRD_IDY,STK_TYP FROM INV WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "' AND INV.TXN_TYP = 'SAL' AND STK_TYP IN('Salable','Free') GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
    If RSINV.RecordCount > 0 Then
        For A = 1 To RSINV.RecordCount
            For b = 1 To MS.Rows - 1
                If MS.TextMatrix(b, 17) = Trim(RSINV(1)) Then
                    If Pieces.Value = 0 Then
                        MS.TextMatrix(b, 10) = Val(RSINV(0))
                    Else
                        MS.TextMatrix(b, 10) = QTY(STM(Val(RSINV(0)), Val(RSINV(2))))
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
    RSINV.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT,PRD_IDY,STK_TYP FROM INV WHERE ORD_IDY = '" & Trim(M_INV_IDY) & "' AND INV.TXN_TYP = 'SAL' AND STK_TYP IN('Damage','TakeBack','Breakage','Leakage') GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
    If RSINV.RecordCount > 0 Then
        For A = 1 To RSINV.RecordCount
            For b = 1 To MS.Rows - 1
                If MS.TextMatrix(b, 17) = Trim(RSINV(1)) Then
                    If Pieces.Value = 0 Then
                        MS.TextMatrix(b, 10) = Val(MS.TextMatrix(b, 10)) - Val(RSINV(0))
                    Else
                        MS.TextMatrix(b, 10) = QTY(STM(MTS(Val(MS.TextMatrix(b, 10)), Val(RSINV(2))) - Val(RSINV(0)), Val(RSINV(2))))
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
                MS3 = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 16)))
                MS4 = MTS(Val(MS.TextMatrix(A, 5)), Val(MS.TextMatrix(A, 16)))
                MS5 = Val(MS3) - Val(MS4)
                MS.TextMatrix(A, 10) = QTY(STM(Val(MS5), Val(MS.TextMatrix(A, 16))))
            Else
                MS.TextMatrix(A, 10) = Val(MS.TextMatrix(A, 4)) - Val(MS.TextMatrix(A, 5))
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
            'MS.TextMatrix(A, 13) = convert(Val(MS.TextMatrix(A, 10)) * Val(MS.TextMatrix(A, 12)))
            MS.TextMatrix(A, 13) = convert(Val(MS.TextMatrix(A, 9)) * Val(MS.TextMatrix(A, 19)))
        Else
            'MS.TextMatrix(A, 13) = convert(RATCAL(QTY(Val(MS.TextMatrix(A, 10) & "")), Val(MS.TextMatrix(A, 12) & ""), Val(MS.TextMatrix(A, 16))))
            MS.TextMatrix(A, 13) = convert(RATCAL(QTY(Val(MS.TextMatrix(A, 9) & "")), Val(MS.TextMatrix(A, 19) & ""), Val(MS.TextMatrix(A, 16))))
        End If
        Total = convert(amt(Val(Total) + Val(MS.TextMatrix(A, 13))))
    End If
Next
M_GRS_AMT = convert(amt(Total))
Total = convert(amt(Total))
End Sub

Private Sub TRP_NOS_GotFocus()
Dim rsact As New ADODB.Recordset
If rsact.State = 1 Then rsact.Close
rsact.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "' GROUP BY TRP_IDY", con, adOpenKeyset, adLockPessimistic
' GROUP BY TRP_IDY", CON, adOpenKeyset, adLockPessimistic
If rsact.RecordCount > 1 And Check3.Value = 1 Then
       TRP_NOS.Text = "All"
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select ACT.* from ACT where ACT_IDY = '" & Trim(M_CST_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then
    If REC_SET("DBT_RAP") = "A" Then
        MsgBox ("Applicable SKU Rates for Settlement is A Rate for this VAN !"), vbCritical, "DAS Version 4.0"
    End If
End If
End Sub

Private Sub TRP_NOS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Dim rs As New ADODB.Recordset
    If Check3.Value = 1 And TRP_NOS = "All" And MOD_REC = True Then
        If CHECKTB("van.dbf,brd.dbf,prd.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
        If RSCUR.State = 1 Then RSCUR.Close
        RSCUR.Open "SELECT * FROM CUR WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        rs.Open "SELECT VAN.*,BRD.*,SUM(INT(FRE_QTY) * SUB_UNT)+(FRE_QTY-INT(FRE_QTY))*1000 P_FRE_QTY,SUM(INT(TOT_QTY) * SUB_UNT)+(TOT_QTY-INT(TOT_QTY))*1000 P_TOT_QTY,SUM(INT(OPN_QTY) * SUB_UNT)+(OPN_QTY-INT(OPN_QTY))*1000 P_OPN_QTY,SUM(INT(ISS_QTY) * SUB_UNT)+(ISS_QTY-INT(ISS_QTY))*1000 P_ISS_QTY,SUM(INT(RET_QTY) * SUB_UNT)+(RET_QTY-INT(RET_QTY))*1000 P_RET_QTY,SUM(INT(BIL_QTY) * SUB_UNT)+(BIL_QTY-INT(BIL_QTY))*1000 P_BIL_QTY,SUB_UNT,PRD_NME,PRD_PRN FROM VAN,BRD,PRD WHERE VAN.BAT_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY AND INV_IDY = '" & Trim(M_INV_IDY) & "' GROUP BY BRD.BAT_IDY", con, adOpenKeyset, adLockPessimistic
        'rs.Open "SELECT VAN.*,BRD.*,SUM((INT(FRE_QTY) * SUB_UNT)+INT(FRE_QTY-INT(FRE_QTY))*1000) P_FRE_QTY,SUM((INT(TOT_QTY) * SUB_UNT)+INT((TOT_QTY-INT(TOT_QTY))*1000)) P_TOT_QTY,SUM((INT(OPN_QTY) * SUB_UNT)+INT(OPN_QTY-INT(OPN_QTY))*1000) P_OPN_QTY,SUM((INT(ISS_QTY) * SUB_UNT)+(ISS_QTY-INT(ISS_QTY))*1000) P_ISS_QTY,SUM((INT(RET_QTY) * SUB_UNT)+(RET_QTY-INT(RET_QTY))*1000) P_RET_QTY,SUM((INT(BIL_QTY) * SUB_UNT)+INT(BIL_QTY-INT(BIL_QTY))*1000) P_BIL_QTY,SUB_UNT,PRD_NME,PRD_PRN FROM VAN,BRD,PRD WHERE VAN.BAT_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND INV_IDY LIKE '" & Trim(M_INV_IDY) & "' GROUP BY BRD.BAT_IDY", CON, adOpenKeyset, adLockPessimistic 'MODIFIED BY PARVATHI
        MS.Rows = 1
        For A = 1 To rs.RecordCount
            MS.AddItem ""
            MS.TextMatrix(A, 0) = Trim(rs!PRD_NME)
            MS.TextMatrix(A, 1) = "All"
            If Pieces.Value = 0 Then
                MS.TextMatrix(A, 2) = Val(rs!P_OPN_QTY & "")
                MS.TextMatrix(A, 3) = Val(rs!P_ISS_QTY & "")
                MS.TextMatrix(A, 4) = Val(rs!p_tot_qty & "")
                MS.TextMatrix(A, 10) = Val(rs!P_BIL_QTY & "")
                MS.TextMatrix(A, 5) = Val(rs!p_ret_qty & "")
                MS.TextMatrix(A, 6) = Val(rs!P_FRE_QTY & "")
                MS.TextMatrix(A, 12) = convert(Val(rs!PRD_MRP) / Val(rs!SUB_UNT))
                MS.TextMatrix(A, 19) = (convert(Val(rs!PRD_rat) - Val(rs!S_INC_DIS)) / Val(rs!SUB_UNT))
'                MS.TextMatrix(A, 19) = convert((Val(rs!prd_rat) + ((Val(rs!prd_rat) * Val(rs!SAL_TAX)) / 100))) / Val(rs!SUB_UNT)
            Else
                MS.TextMatrix(A, 2) = QTY(STM(Val(rs!P_OPN_QTY), Val(rs!SUB_UNT)))
                MS.TextMatrix(A, 3) = QTY(STM(Val(rs!P_ISS_QTY), Val(rs!SUB_UNT)))
                MS.TextMatrix(A, 4) = QTY(STM(Val(rs!p_tot_qty), Val(rs!SUB_UNT)))
                MS.TextMatrix(A, 10) = QTY(STM(Val(rs!P_BIL_QTY), Val(rs!SUB_UNT)))
                MS.TextMatrix(A, 5) = QTY(STM(Val(rs!p_ret_qty), Val(rs!SUB_UNT)))
                MS.TextMatrix(A, 6) = QTY(STM(Val(rs!P_FRE_QTY), Val(rs!SUB_UNT)))
                MS.TextMatrix(A, 12) = convert(rs!PRD_MRP & "")
                MS.TextMatrix(A, 19) = convert(Val(rs!PRD_rat) - Val(rs!S_INC_DIS))
'                MS.TextMatrix(A, 19) = convert((Val(rs!prd_rat) + (Val(rs!prd_rat) * Val(rs!SAL_TAX)) / 100))
            End If
            M_CAS_AMT = convert(rs!CAS_AMT & "")
            M_CHQ_AMT = convert(rs!CHQ_AMT & "")
            M_OLD_CRD = convert(rs!OLD_CRD & "")
            M_CUR_CRD = convert(rs!CUR_CRD & "")
            M_DIS_AMT = convert(rs!DIS_AMT & "")
            M_EXP_AMT = convert(rs!EXP_AMT & "")
            M_TAX_AMT = convert(rs!ADD_AMT & "")
            MS.TextMatrix(A, 13) = convert(rs!COL_AMT & "")
            MS.TextMatrix(A, 14) = Trim(rs!prd_idy)
            MS.TextMatrix(A, 15) = Trim(rs!prd_prn)
            MS.TextMatrix(A, 16) = Val(rs!SUB_UNT)
            MS.TextMatrix(A, 17) = Trim(rs!BAT_IDY)
            MS.TextMatrix(A, 18) = Trim(rs!Bat_nme)
            rs.MoveNext
        Next
        If RSCUR.RecordCount > 0 Then
            txt1000 = Val(RSCUR(1) & "")
            txt500 = Val(RSCUR(2) & "")
            txt100 = Val(RSCUR(3) & "")
            txt50 = Val(RSCUR(4) & "")
            txt20 = Val(RSCUR(5) & "")
            txt10 = Val(RSCUR(6) & "")
            txt5 = Val(RSCUR(7) & "")
            txt2 = Val(RSCUR(8) & "")
            txt1 = Val(RSCUR(9) & "")
            txtchange = Val(RSCUR(10) & "")
        End If
        TRP_NOS.Enabled = False
        P_BIL_QTY
        discalc
        M_SMN_IDY.SetFocus
    Else
        If MOD_REC = False Then
            If TRP_NOS = "All" Then
                MsgBox "Enter Trip number!", vbInformation, "DAS Version 4.0"
                SendKeys "{home}+{end}"
                TRP_NOS.SetFocus
            Else
                If Len(Trim(TRP_NOS)) = 0 Or Val(TRP_NOS) = 0 Then
                    MsgBox "Enter Trip number!", vbInformation, "DAS Version 4.0"
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
            If rs.State = 1 Then rs.Close
            rs.Open "select * from van where inv_idy = '" & Trim(M_INV_IDY) & "' and trp_idy = " & Val(TRP_NOS), con, adOpenKeyset, adLockPessimistic
            If rs.RecordCount > 0 Then
                TRP_NOS.Enabled = False
                M_SMN_IDY.SetFocus
            Else
                If TRP_NOS = "All" Or Val(TRP_NOS) = 0 Then
                    MsgBox "Enter Trip Number!", vbInformation, "DAS Version 4.0"
                    TRP_NOS.SetFocus
                    SendKeys "{home}+{end}"
                ElseIf Check3.Value = 0 Then
                    If MsgBox("Create New Trip ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                        TRP_NOS.Enabled = False
                        M_SMN_IDY.SetFocus
                    Else
                        SendKeys "{home}+{end}"
                        TRP_NOS.SetFocus
                    End If
                End If
            End If
            rs.Close
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
    ADDTOTCUR
    If Val(M_CAS_AMT) >= Val(TXTTOTAL) Then
        txtchange = Val(M_CAS_AMT) - Val(TXTTOTAL)
    End If
    txtchange.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt2.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
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
    ADDTOTCUR
    txt500.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    M_CRT_AMT.SetFocus
End If
txt1000.Alignment = 1
End Sub

Private Sub txt2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
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
    ADDTOTCUR
    txt2.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 And Shift = 1 Then
    txt10.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txt50_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
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
    If Val(M_CAS_AMT) = Val(TXTTOTAL.Text) Then
        If Check3.Value = 1 And Trolley.Value = 1 Then
            MakeBill
        Else
            saverecord
        End If
    Else
        If MsgBox("Cash Currency Details DO NOT TALLY with Cash Collected! Proceed ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
            If Check3.Value = 1 And Trolley.Value = 1 Then
                MakeBill
            Else
                saverecord
            End If
        Else
            txtchange.SetFocus
            SendKeys "{home}+{end}"
        End If
    End If
ElseIf KeyCode = 38 And Shift = 1 Then
    txt1.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtempshl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If ms2.Col = 2 Then
    If Val(txtempshl) >= Val(txtTemp) Then
        ms2.Text = Val(txtempshl)
        txtempshl = ""
        txtempshl.Visible = False
        ms2.Col = ms2.Col + 1
        ms2.SetFocus
    Else
        MsgBox "Invalid Value!", vbInformation + vbOKOnly, "DAS Version 4.0"
        txtempshl.SetFocus
        SendKeys "{home}+{end}"
    End If
    Exit Sub
End If
With ms2
    If .Col = 3 Then
        .Text = Int(Val(txtempshl))
        .TextMatrix(.Row, 7) = Val(.TextMatrix(.Row, 3)) - (Val(.TextMatrix(.Row, 5)) + Val(.TextMatrix(.Row, 12)))
        .TextMatrix(.Row, 8) = QTY(STM(Val(MTS(Val(.TextMatrix(.Row, 4)), Val(.TextMatrix(.Row, 11))) - (MTS(Val(.TextMatrix(.Row, 13)), Val(.TextMatrix(.Row, 11))) + MTS(Val(.TextMatrix(.Row, 6)), Val((.TextMatrix(.Row, 11)))))), Val((.TextMatrix(.Row, 11)))))
        txtempshl = ""
        txtempshl.Visible = False
        .Col = 3
        .SetFocus
    ElseIf .Col = 4 Then
        If CSUBUNT(Val(txtempshl), Val(ms2.TextMatrix(.Row, 11))) = True Then
            .Text = QTY(Val(txtempshl))
            txtempshl.Visible = False
            MS3 = MTS(Val(.TextMatrix(.Row, 4)), Val(.TextMatrix(.Row, 11)))
            MS5 = MTS(Val(.TextMatrix(.Row, 6)), Val(.TextMatrix(.Row, 11)))
            MS5 = Val(MS5) + MTS(Val(.TextMatrix(.Row, 13)), Val(.TextMatrix(.Row, 11)))
            .TextMatrix(.Row, 8) = QTY(STM(Val(MS3) - Val(MS5), Val(.TextMatrix(.Row, 11))))
            .Col = 5
            .SetFocus
        End If
    ElseIf .Col = 5 Then
        .Text = Int(Val(txtempshl))
        .TextMatrix(.Row, 6) = QTY(.Text)
        txtempshl.Visible = False
        .TextMatrix(.Row, 7) = Val(.TextMatrix(.Row, 1)) - (Val(.TextMatrix(.Row, 5)) + Val(.TextMatrix(.Row, 12)))
        .TextMatrix(ms2.Row, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(ms2.Row, 2)), Val(ms2.TextMatrix(ms2.Row, 11))) - (MTS(Val(ms2.TextMatrix(ms2.Row, 13)), Val(ms2.TextMatrix(ms2.Row, 11))) + MTS(Val(ms2.TextMatrix(ms2.Row, 6)), Val((ms2.TextMatrix(ms2.Row, 11)))))), Val((ms2.TextMatrix(ms2.Row, 11)))))
        .Col = 6
        .SetFocus
    '.TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 2)) - Val(.TextMatrix(.Row, 4))
    ElseIf .Col = 6 Then
        If CSUBUNT(Val(txtempshl), Val(ms2.TextMatrix(.Row, 11))) = True Then
            .Text = QTY(Val(txtempshl))
            txtempshl.Visible = False
            '.TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 1)) - Val(.TextMatrix(.Row, 3))
            MS3 = MTS(Val(.TextMatrix(.Row, 2)), Val(.TextMatrix(.Row, 11)))
            MS5 = MTS(Val(.TextMatrix(.Row, 6)), Val(.TextMatrix(.Row, 11)))
            MS5 = Val(MS5) + MTS(Val(.TextMatrix(.Row, 13)), Val(.TextMatrix(.Row, 11)))
            .TextMatrix(.Row, 8) = QTY(STM(Val(MS3) - Val(MS5), Val(.TextMatrix(.Row, 11))))
            If .Rows - 1 <> .Row Then .Row = .Row + 1
            .Col = 5
            .SetFocus
        End If
    ElseIf .Col = 12 Then
        .Text = Int(Val(txtempshl))
        ms2.TextMatrix(ms2.Row, 7) = Val(ms2.TextMatrix(ms2.Row, 1)) - (Val(ms2.TextMatrix(ms2.Row, 5)) + Val(ms2.TextMatrix(ms2.Row, 12)))
        .TextMatrix(.Row, 13) = QTY(.Text)
        ms2.TextMatrix(ms2.Row, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(ms2.Row, 2)), Val(ms2.TextMatrix(ms2.Row, 11))) - (MTS(Val(ms2.TextMatrix(ms2.Row, 13)), Val(ms2.TextMatrix(ms2.Row, 11))) + MTS(Val(ms2.TextMatrix(ms2.Row, 6)), Val((ms2.TextMatrix(ms2.Row, 11)))))), Val((ms2.TextMatrix(ms2.Row, 11)))))
        txtempshl.Visible = False
        .SetFocus
    ElseIf .Col = 13 Then
        If CSUBUNT(Val(txtempshl), Val(ms2.TextMatrix(.Row, 11))) = True Then
            .Text = QTY(Val(txtempshl))
            ms2.TextMatrix(ms2.Row, 8) = QTY(STM(Val(MTS(Val(ms2.TextMatrix(ms2.Row, 2)), Val(ms2.TextMatrix(ms2.Row, 11))) - (MTS(Val(ms2.TextMatrix(ms2.Row, 13)), Val(ms2.TextMatrix(ms2.Row, 11))) + MTS(Val(ms2.TextMatrix(ms2.Row, 6)), Val((ms2.TextMatrix(ms2.Row, 11)))))), Val((ms2.TextMatrix(ms2.Row, 11)))))
            txtempshl.Visible = False
            .SetFocus
        End If
    End If
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
Dim rs As New ADODB.Recordset
TxnMoveBackup "SNV", "SBK", Trim(IDY), "VAN"
con.Execute "DELETE FROM SNV WHERE INV_IDY = '" & Trim(IDY) & "' AND TXN_TYP='VAN'"
rs.Open "SELECT * FROM SNV WHERE INV_IDY = ''", con, adOpenKeyset, adLockPessimistic
For A = 2 To ms2.Rows - 1
    If ms2.TextMatrix(A, 10) <> "" Then
        rs.AddNew
        rs!INV_IDY = Trim(IDY)
        rs!TXN_IDY = Trim(IDY) 'Added by PNR
        rs!INV_DAT = M_INV_DAT
        rs!shl_idy = ms2.TextMatrix(A, 10)
        rs!prd_idy = Trim(ms2.TextMatrix(A, 14))    ' Modified
        rs!PRD_PCK = Trim(ms2.TextMatrix(A, 9))
        rs!prd_upc = Val(ms2.TextMatrix(A, 11))
        If SHLSETCLICK = False Then
            rs!I_shl_nos = Val(ms2.TextMatrix(A, 1))
            rs!I_pck_nos = MTS(Val(ms2.TextMatrix(A, 2)), Val(ms2.TextMatrix(A, 11)))
        Else
            rs!I_shl_nos = Val(ms2.TextMatrix(A, 3))
            rs!I_pck_nos = MTS(Val(ms2.TextMatrix(A, 4)), Val(ms2.TextMatrix(A, 11)))
        End If
        rs!R_shl_nos = Val(ms2.TextMatrix(A, 5))
        rs!R_pck_nos = MTS(Val(ms2.TextMatrix(A, 6)), Val(ms2.TextMatrix(A, 11)))
        rs!d_shl_nos = Val(ms2.TextMatrix(A, 7))
        rs!d_pck_nos = MTS(Val(ms2.TextMatrix(A, 8)), Val(ms2.TextMatrix(A, 11)))
        rs!BRK_SHL = Val(ms2.TextMatrix(A, 12))
        rs!BRK_PCK = MTS(Val(ms2.TextMatrix(A, 13)), Val(ms2.TextMatrix(A, 11)))
        rs!s_shl_nos = Val(ms2.TextMatrix(A, 1))
        rs!s_pck_nos = MTS(Val(ms2.TextMatrix(A, 2)), Val(ms2.TextMatrix(A, 11)))
        rs!CST_IDY = M_CST_IDY
        rs!TXN_TYP = "VAN"
        rs!txn_dat = M_INV_DAT
        rs!stk_upd = "Y"
        rs!SMN_IDY = M_SMN_IDY
        rs!rou_idy = Trim(cmbRoute.Text & "")
        rs.Update
    End If
Next
End Sub
Public Sub COMMISSION()
Dim TOTCOM As Double, TOTQ As Double, TOTVAL As Double, TTAMT As Double, com_val As Double
Dim rsVan As ADODB.Recordset
If RSCOM.State = 1 Then RSCOM.Close
RSCOM.Open "select * from sic WHERE act_IDY='" & Trim(M_SMN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
Set rsVan = New ADODB.Recordset
If rsVan.State = 1 Then rsVan.Close
rsVan.Open "SELECT INV_IDY,PRD_IDY,BIL_IDY FROM VAN WHERE INV_IDY='" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If RSCOM.EOF Then
    M_COM_AMT = "0.00"
Else
    com_val = 0
    TOTCOM = 0
    While Not RSCOM.EOF
        For r = 1 To MS.Rows - 1
            If Trim(RSCOM("prd_idy")) = Trim(MS.TextMatrix(r, 14)) Then 'And Trim(RSCOM("BAT_IDY")) = Trim(MS.TextMatrix(R, 17)) Then
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select * from prd where prd_idy='" & Trim(RSCOM("prd_idy")) & "' GROUP BY prd_idy", con, adOpenKeyset, adLockPessimistic
                If Pieces.Value = 0 Then
                    TOTQ = Val(MS.TextMatrix(r, 4)) - (Val(MS.TextMatrix(r, 5)) + Val(MS.TextMatrix(r, 8)) + Val(MS.TextMatrix(r, 7)) + Val(MS.TextMatrix(r, 6)))
                    TOTQ = STM(Val(TOTQ), Val(MS.TextMatrix(r, 16)))
                Else
                    TOTQ = Val(MTS(Val(MS.TextMatrix(r, 4)), Val(MS.TextMatrix(r, 16)))) - (Val(MTS(Val(MS.TextMatrix(r, 5)), Val(MS.TextMatrix(r, 16)))) + Val(MTS(Val(MS.TextMatrix(r, 8)), Val(MS.TextMatrix(r, 16)))) + Val(MTS(Val(MS.TextMatrix(r, 7)), Val(MS.TextMatrix(r, 16)))) + Val(MTS(Val(MS.TextMatrix(r, 6)), Val(MS.TextMatrix(r, 16)))))
                    TOTQ = STM(Val(TOTQ), Val(MS.TextMatrix(r, 16)))
                End If
                If MTS(Val(TOTQ), MS.TextMatrix(r, 16)) >= MTS(RSCOM("min_qty"), MS.TextMatrix(r, 16)) Then
                    If Pieces.Value = 0 Then
                        If Not rsprd.EOF Then com_val = MTS(Val(TOTQ), Val(rsprd("SUB_UNT"))) * (RSCOM("min_val") / Val(rsprd("sub_unt")))
                    Else
                        com_val = (Int(TOTQ) * RSCOM("min_val")) + (((TOTQ - Int(TOTQ)) * 1000) * (RSCOM("MIN_VAL") / Val(rsprd("sub_unt"))))
                    End If
                End If
'                If rsVan.RecordCount > 0 Then rsVan.MoveFirst
'                rsVan.Find "PRD_IDY='" & Trim(RSCOM("PRD_IDY")) & "'"
'                If Not rsVan.EOF Then
'                    rsVan("BIL_IDY") = Trim(Val(com_val)) & ""
'                    rsVan.Update
'                End If
                TOTCOM = TOTCOM + com_val
                Exit For
            End If
        Next
        RSCOM.MoveNext
    Wend
    'TOTCOM = TOTCOM + com_val 'VALUE GETTING DOUBLED IN COMMISSION CALCULATION, COMMENTED BY PNR
    'M_COM_AMT = convert(AMT(Trim(TOTCOM))) 'Commission automatic calculation and effect on settlement disabled - as per Prince Chopra DAS Request
    M_COM_AMT = convert(0#)
End If
End Sub
Private Sub ADDTOTCUR()
TXTTOTAL = convert(Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txt2) * 2 + Val(txt1) * 1 + Val(txtchange))
End Sub
Public Function CSUBUNT(STK As Double, SUBUNT As Integer) As Boolean
S_UNT = (Val(STK) - Int(STK)) * 1000
If S_UNT < SUBUNT Then
    CSUBUNT = True
Else
    MsgBox "Product Subunits are " & SUBUNT, vbInformation, "DAS Version 4.0"
    txtempshl.SetFocus
    SendKeys "{HOME}+{END}"
    CSUBUNT = False
End If
End Function
Private Sub ChallanNullify()
Dim CNlfy As New ADODB.Recordset
Dim SNlfy As New ADODB.Recordset
Dim RNlfy As New ADODB.Recordset
If CNlfy.State = 1 Then CNlfy.Close
CNlfy.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If CNlfy.EOF = False Then
    If Trim(CNlfy("VAN_CLS")) = "C" Then
        MsgBox ("This Challan is Closed! Cannot Process Nullify Request!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    Else
        While Not CNlfy.EOF
            If Trim(CNlfy("VAN_CLS")) = "N" Then
                CNlfy("OPN_QTY") = 0#
                CNlfy("ISS_QTY") = 0#
                CNlfy("TOT_QTY") = 0#
                CNlfy("RET_QTY") = 0#
                CNlfy("BIL_QTY") = 0#
                CNlfy("PRD_RAT") = 0#
                CNlfy("COL_AMT") = 0#
                CNlfy("FRE_QTY") = 0#
                CNlfy("LKG_QTY") = 0#
                CNlfy("BRK_QTY") = 0#
                CNlfy("CAS_AMT") = 0#
                CNlfy("CHQ_AMT") = 0#
                CNlfy("DIS_AMT") = 0#
                CNlfy("EXP_AMT") = 0#
                CNlfy("ADD_AMT") = 0#
                CNlfy("COM_AMT") = 0#
                CNlfy("DEP_RCD") = 0#
                CNlfy("DEP_REF") = 0#
                CNlfy("OLD_CRD") = 0#
                CNlfy("CUR_CRD") = 0#
                CNlfy("SHT_CSH") = 0#
                CNlfy.Update
            End If
            If Trim(CNlfy("VAN_CLS")) = "S" Then
                CNlfyRtn = 0#
                CNlfyRtn = CNlfy("TOT_QTY")
                CNlfy("RET_QTY") = CNlfyRtn
                CNlfy("BIL_QTY") = 0#
                CNlfy("PRD_RAT") = 0#
                CNlfy("COL_AMT") = 0#
                CNlfy("FRE_QTY") = 0#
                CNlfy("LKG_QTY") = 0#
                CNlfy("BRK_QTY") = 0#
                CNlfy("CAS_AMT") = 0#
                CNlfy("CHQ_AMT") = 0#
                CNlfy("DIS_AMT") = 0#
                CNlfy("EXP_AMT") = 0#
                CNlfy("ADD_AMT") = 0#
                CNlfy("COM_AMT") = 0#
                CNlfy("DEP_RCD") = 0#
                CNlfy("DEP_REF") = 0#
                CNlfy("OLD_CRD") = 0#
                CNlfy("CUR_CRD") = 0#
                CNlfy("SHT_CSH") = 0#
                CNlfy.Update
            End If
            CNlfy.MoveNext
        Wend
        If SNlfy.State = 1 Then SNlfy.Close
        SNlfy.Open ("SELECT * FROM SNV WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'"), con, adOpenKeyset, adLockPessimistic
        While Not SNlfy.EOF
            SNlfy("I_SHL_NOS") = 0#
            SNlfy("I_PCK_NOS") = 0#
            SNlfy("R_SHL_NOS") = 0#
            SNlfy("R_PCK_NOS") = 0#
            SNlfy("D_SHL_NOS") = 0#
            SNlfy("D_PCK_NOS") = 0#
            SNlfy("DP_SHL_NOS") = 0#
            SNlfy("DP_PCK_NOS") = 0#
            SNlfy("DEP_AMT") = 0#
            SNlfy("BRK_SHL") = 0#
            SNlfy("BRK_PCK") = 0#
            SNlfy("S_SHL_NOS") = 0#
            SNlfy("S_PCK_NOS") = 0#
            SNlfy("N_SHL_NOS") = 0#
            SNlfy("N_PCK_NOS") = 0#
            SNlfy.Update
            SNlfy.MoveNext
        Wend
        If RNlfy.State = 1 Then RNlfy.Close
        RNlfy.Open ("SELECT * FROM CUR WHERE INV_IDY = '" & Trim(M_INV_IDY) & "'"), con, adOpenKeyset, adLockPessimistic
        While Not RNlfy.EOF
            RNlfy("ONE_IDY") = 0#
            RNlfy("TWO_IDY") = 0#
            RNlfy("THR_IDY") = 0#
            RNlfy("FOR_IDY") = 0#
            RNlfy("FIV_IDY") = 0#
            RNlfy("SIX_IDY") = 0#
            RNlfy("SEV_IDY") = 0#
            RNlfy("EIG_IDY") = 0#
            RNlfy("NIN_IDY") = 0#
            RNlfy("TEN_IDY") = 0#
            RNlfy("CHG_IDY") = 0#
            RNlfy.Update
            RNlfy.MoveNext
        Wend
    End If
    MsgBox ("Challan Nullified!"), vbInformation, "DAS Version 4.0"
End If
CNlfy.Close
SNlfy.Close
RNlfy.Close
Set CNlfy = Nothing
Set SNlfy = Nothing
Set RNlfy = Nothing
End Sub


