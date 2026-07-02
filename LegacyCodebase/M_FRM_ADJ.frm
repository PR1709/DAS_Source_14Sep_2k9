VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_ADJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Adjustment"
   ClientHeight    =   5625
   ClientLeft      =   1455
   ClientTop       =   330
   ClientWidth     =   8820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "M_FRM_ADJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&Voucher Details"
      TabPicture(0)   =   "M_FRM_ADJ.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label23"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "M_FFR_NME"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "M_ACT_NME"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "M_SUP_NME"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "M_INV_DAT"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "M_INV_IDY"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "M_CRT_IDY"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "M_DBT_IDY"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&Products"
      TabPicture(1)   =   "M_FRM_ADJ.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "M_FLX_AMT"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MS"
      Tab(1).Control(3)=   "TOTUNT"
      Tab(1).Control(4)=   "TOTAMT"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&Taxes and  Levies"
      TabPicture(2)   =   "M_FRM_ADJ.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "M_DUE_DAT"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(3)=   "M_CRT_DAY"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label9"
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(6)=   "Label7"
      Tab(2).Control(7)=   "Label5"
      Tab(2).Control(8)=   "Label42"
      Tab(2).Control(9)=   "Label44"
      Tab(2).Control(10)=   "Label24"
      Tab(2).Control(11)=   "Label25"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "&Shipment Details"
      TabPicture(3)   =   "M_FRM_ADJ.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "M_TXN_RMK"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label54"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "View"
      TabPicture(4)   =   "M_FRM_ADJ.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1"
      Tab(4).Control(1)=   "list4"
      Tab(4).Control(2)=   "list5"
      Tab(4).Control(3)=   "list2"
      Tab(4).Control(4)=   "list1"
      Tab(4).Control(5)=   "list3"
      Tab(4).Control(6)=   "Label49"
      Tab(4).ControlCount=   7
      Begin VB.TextBox M_FLX_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   -70800
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame4 
         Height          =   2400
         Left            =   -69600
         TabIndex        =   87
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox M_CSD_AMT 
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
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   200
            Width           =   855
         End
         Begin VB.TextBox M_SPD_AMT 
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
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   560
            Width           =   855
         End
         Begin VB.TextBox M_TAX_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   910
            Width           =   855
         End
         Begin VB.TextBox M_TAX_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   910
            Width           =   615
         End
         Begin VB.TextBox M_SPD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   560
            Width           =   615
         End
         Begin VB.TextBox M_CSD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   200
            Width           =   615
         End
         Begin VB.TextBox M_PKG_CST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   1640
            Width           =   855
         End
         Begin VB.TextBox M_FRT_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   2000
            Width           =   855
         End
         Begin VB.TextBox M_EXD_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1280
            Width           =   855
         End
         Begin VB.TextBox M_EXD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   1280
            Width           =   615
         End
         Begin VB.TextBox M_AMT_ADD 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   240
            MaxLength       =   70
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   1640
            Width           =   1695
         End
         Begin VB.TextBox M_AMT_LES 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   240
            MaxLength       =   70
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   2000
            Width           =   1695
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Additonal Discount % "
            Height          =   195
            Left            =   510
            TabIndex        =   105
            Top             =   630
            Width           =   1545
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bill Discount % "
            Height          =   195
            Left            =   975
            TabIndex        =   104
            Top             =   270
            Width           =   1080
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tax % "
            Height          =   195
            Left            =   1575
            TabIndex        =   103
            Top             =   975
            Width           =   480
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sur - Charge on Tax % "
            Height          =   195
            Left            =   435
            TabIndex        =   102
            Top             =   1350
            Width           =   1635
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount (-)"
            Height          =   195
            Left            =   2115
            TabIndex        =   101
            Top             =   2070
            Width           =   720
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount (+)"
            Height          =   195
            Left            =   2070
            TabIndex        =   100
            Top             =   1680
            Width           =   765
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Load Products"
         Height          =   375
         Left            =   -67560
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Transport Receipt Details"
         Height          =   2175
         Left            =   -73800
         TabIndex        =   70
         Top             =   960
         Width           =   5655
         Begin VB.TextBox M_WBL_IDY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox M_TXN_TPT 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   1620
            Width           =   4335
         End
         Begin VB.TextBox M_TTL_CAS 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4305
            MaxLength       =   5
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   765
            Width           =   1215
         End
         Begin VB.TextBox M_LRN_IDY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox M_LRN_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   765
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker M_LRN_DAT 
            Height          =   330
            Left            =   4305
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74186755
            CurrentDate     =   37097
         End
         Begin MSComCtl2.DTPicker M_WBL_DAT 
            Height          =   330
            Left            =   4305
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74186755
            CurrentDate     =   37097
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WB Number"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   180
            TabIndex        =   84
            Top             =   1275
            Width           =   870
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WB Date"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3540
            TabIndex        =   83
            Top             =   1275
            Width           =   660
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transporter"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   240
            TabIndex        =   80
            Top             =   1688
            Width           =   810
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cases"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   3600
            TabIndex        =   78
            Top             =   833
            Width           =   435
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LR Number"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   240
            TabIndex        =   76
            Top             =   435
            Width           =   810
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LR Date"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3540
            TabIndex        =   75
            Top             =   435
            Width           =   600
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LR Amount"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   255
            TabIndex        =   74
            Top             =   833
            Width           =   795
         End
      End
      Begin VB.TextBox M_DUE_DAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Height          =   2490
         Left            =   -72000
         TabIndex        =   30
         Top             =   1395
         Width           =   3375
         Begin VB.TextBox G_PRD_DIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox T_PRD_VAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox G_FRT_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   3480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox G_SCT_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox G_PKG_CST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   3120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox G_PRD_VAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox G_PRD_TAX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox G_GRS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox M_ROU_OFF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            MaxLength       =   4
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox G_CSD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox G_SPD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox G_TAX_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox T_FRT_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   3480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox T_SCT_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox T_PKG_CST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   3120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox T_PRD_DIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox T_PRD_TAX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox T_NET_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox T_GRS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox T_ROU_OFF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox T_CSD_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox T_SPD_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox T_TAX_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "+/-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   63
            Top             =   1695
            Width           =   225
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1605
            TabIndex        =   62
            Top             =   3135
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1600
            TabIndex        =   61
            Top             =   2055
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1600
            TabIndex        =   60
            Top             =   1695
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1605
            TabIndex        =   59
            Top             =   3525
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   1600
            TabIndex        =   58
            Top             =   2805
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   1600
            TabIndex        =   57
            Top             =   2445
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1600
            TabIndex        =   56
            Top             =   1005
            Width           =   105
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1600
            TabIndex        =   55
            Top             =   615
            Width           =   75
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "="
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
            Left            =   1600
            TabIndex        =   54
            Top             =   1388
            Width           =   120
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "="
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
            Index           =   0
            Left            =   1600
            TabIndex        =   53
            Top             =   315
            Width           =   120
         End
      End
      Begin VB.TextBox M_TXN_RMK 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72580
         MaxLength       =   70
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74760
         TabIndex        =   18
         Top             =   4560
         Width           =   8295
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.TextBox M_CRT_DAY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73320
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox M_DBT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox M_CRT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2482
         Width           =   1215
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   5640
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74186755
         CurrentDate     =   37099
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   3915
         Left            =   -74935
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   795
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   6906
         _Version        =   393216
         Cols            =   21
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   0
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
         Height          =   3615
         Left            =   -74640
         TabIndex        =   21
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Pur. Rate"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Short Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2540
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
            Text            =   "Sale Rate"
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
      Begin MSComctlLib.ListView list5 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   23
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
            Text            =   "Voucher ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Voucher Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   24
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   25
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Salesman Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list3 
         Height          =   3615
         Left            =   -74640
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label TOTUNT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -70920
         TabIndex        =   107
         Top             =   4800
         Width           =   45
      End
      Begin VB.Label TOTAMT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -66480
         TabIndex        =   106
         Top             =   4800
         Width           =   45
      End
      Begin VB.Label M_SUP_NME 
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
         Left            =   5040
         TabIndex        =   85
         Top             =   1920
         Width           =   3405
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Amount"
         Height          =   195
         Left            =   -73230
         TabIndex        =   68
         Top             =   2700
         Width           =   990
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Tax"
         Height          =   195
         Left            =   -73110
         TabIndex        =   67
         Top             =   2400
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Discount"
         Height          =   195
         Left            =   -73470
         TabIndex        =   66
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Value"
         Height          =   195
         Left            =   -73245
         TabIndex        =   65
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73390
         TabIndex        =   28
         Top             =   3435
         Width           =   630
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
         Left            =   -74640
         TabIndex        =   22
         Top             =   480
         Width           =   8295
      End
      Begin VB.Label M_ACT_NME 
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
         Left            =   5040
         TabIndex        =   16
         Top             =   2482
         Width           =   3405
      End
      Begin VB.Label M_FFR_NME 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   5040
         TabIndex        =   15
         Top             =   3960
         Width           =   45
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R/Off"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -72645
         TabIndex        =   13
         Top             =   3045
         Width           =   405
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -73080
         TabIndex        =   12
         Top             =   3375
         Width           =   840
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Days"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -74250
         TabIndex        =   11
         Top             =   675
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payable Date"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -74400
         TabIndex        =   10
         Top             =   1035
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2370
         TabIndex        =   9
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Account"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1410
         TabIndex        =   8
         Top             =   2550
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   675
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   5040
         TabIndex        =   5
         Top             =   675
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Account"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   1988
         Width           =   1020
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "M_FRM_ADJ.frx":04CE
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":0DDA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":112E
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":1766
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":1ABA
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":220E
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":261A
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":2726
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":2A7A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":2E8E
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ADJ.frx":33DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   29
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
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "Adjustment Report"
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
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_ADJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, bslab, MSROW, INVIDY, INC, dlr_idy

Private Sub CREATE()
MOD_REC = False
M_INV_IDY = ""
saverecord
End Sub

Private Sub saverecord()
On Error GoTo lab
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
Dim REC_IRT As New ADODB.Recordset
Dim REC_PNL As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If

If MOD_REC = True Then
    MsgBox "Changes Cannot be Saved", vbExclamation
    If check("PRT", "PRT_NME", "M_FRM_ADJ") = True Then
    OpenDocument CURDIR & "\Vouch.RP1", "Adjustment Voucher", 2, Trim(M_INV_IDY), Trim(M_INV_IDY), 6, CURDIR & "\inv.cdx,c,inv_idy"
    End If
    CANCELRECORD
    Exit Sub
End If
If check("ACT", "ACT_TYP LIKE '' AND ACT_idy", Trim(M_DBT_IDY)) = False Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_DBT_IDY.SetFocus
    Exit Sub
End If

If check("act", " ACT_TYP LIKE '' AND act_idy", Trim(M_CRT_IDY)) = False Then
    MsgBox " Account Not Found, Define Account...!", vbExclamation, "DAS Version 4.0"
    M_CRT_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inl where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
REC_SET.AddNew
REC_SET("inv_idy") = "ADJ" & T7increment("IDY_ADJ")
INC = REC_SET("inv_idy")
INVIDY = INC
REC_SET("inv_dat") = M_INV_DAT
REC_SET("cst_idy") = Trim(M_DBT_IDY & "")
REC_SET("act_idy") = Trim(M_CRT_IDY & "")
REC_SET!ORD_IDY = ""
REC_SET!USR_IDY = Trim(USERNAME)
REC_SET!sys_dat = Date
REC_SET!sys_tim = Format(Time, "HH:MM:SS")
REC_SET!TXN_TYP = "ADJ"
REC_SET("SCH_CHK") = "Y"
REC_SET!SRN_AMT = 0
REC_SET!vch_idy = ""
REC_SET!Sch_Idy = ""
REC_SET!ORD_IDY = ""
REC_SET!DCH_IDY = ""
For A = 1 To MS.Rows - 1
TAMT = Val(TAMT) + Abs(Val(MS.TextMatrix(A, 5)))
Next
REC_SET!GRS_AMT = TAMT
REC_SET!NET_AMT = TAMT
REC_SET.Update

If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inv where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" And Abs(Val(MS.TextMatrix(A, 4))) > 0 Then
        REC_SET.AddNew
        REC_SET("inv_idy") = Trim(INC)
        REC_SET("inv_dat") = Trim(M_INV_DAT)
        REC_SET("prd_idy") = MS.TextMatrix(A, 12)
        REC_SET("opn_qty") = Abs(Val(MS.TextMatrix(A, 2)))
        REC_SET("iss_qty") = Abs(Val(MS.TextMatrix(A, 3)))
        REC_SET("prd_qty") = Val(MS.TextMatrix(A, 4))
        REC_SET("prd_rat") = convert(amt(Val(MS.TextMatrix(A, 6))))
        REC_SET("trd_dis") = 0
        REC_SET("trd_tax") = 0
        REC_SET("ADD_AMT") = 0
        REC_SET("DIS_AMT") = 0
        REC_SET("TAX_AMT") = 0
        REC_SET("txn_typ") = "ADJ"
        REC_SET!BAT_IDY = Trim(MS.TextMatrix(A, 13))
        C = GetValue("CON_FAC", "PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 12)) & "'")
        REC_SET!PRD_UOM = Val(C) * MTS(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 14)))
        REC_SET!SUB_UNT = Val(MS.TextMatrix(A, 14))
        REC_SET!PRD_UNT = Abs(Int(Val(MS.TextMatrix(A, 4))))
        REC_SET!PRD_PCS = (Abs(Int(Val(MS.TextMatrix(A, 4)))) - Abs(Int(Val(MS.TextMatrix(A, 4))))) * 1000
        REC_SET!TOT_PCS = MTS(Abs(Val(MS.TextMatrix(A, 4))), Val(MS.TextMatrix(A, 14)))
        REC_SET("COL_AMT") = Val(MS.TextMatrix(A, 5))
        REC_SET!stk_typ = MS.TextMatrix(A, 1)
        REC_SET!stk_upd = "Y"
        REC_SET!ORD_IDY = ""
        REC_SET.Update
        con.Execute "update PRD where prd_idy like '" & Trim(MS.TextMatrix(A, 12)) & "' set CUR_LEV = " & CONQTY(Abs(Trim(MS.TextMatrix(A, 4))), Trim(MS.TextMatrix(A, 12)))
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 12)) & "' set PRD_PCS = PRD_PCS - " & MTS(Abs(Val(MS.TextMatrix(A, 4))), Val(MS.TextMatrix(A, 14)))
        If Val(MS.TextMatrix(A, 2)) < Val(MS.TextMatrix(A, 3)) Then
            con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & ADDBRDQTY(Abs(Val(Trim(MS.TextMatrix(A, 4)))), Trim(MS.TextMatrix(A, 13)))
        Else
            con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & CONBRDQTY(Abs(Val(Trim(MS.TextMatrix(A, 4)))), Trim(MS.TextMatrix(A, 13)))
        End If
    End If
Next
'pnv update close
If MOD_REC = False Then MsgBox "New Adjustment is " & INC, vbExclamation, "DAS Version 4.0"
Label23 = labdisp("IDY_ADJ")
If check("PRT", "PRT_NME", "M_FRM_ADJ") = True Then
    OpenDocument CURDIR & "\Vouch.RP1", "Adjustment Voucher", 2, Trim(INC), Trim(INC), 6, CURDIR & "\inv.cdx,c,inv_idy"
End If
CANCELRECORD
Exit Sub
lab:
MsgBox err.Description
End Sub

Private Sub CANCELRECORD()
clrctr M_FRM_ADJ
setval
Form_Load
MS.Rows = 1
TOTAMT = ""
TOTUNT = ""
MS.AddItem ""
SSTab1.Tab = 0
MOD_REC = False
M_INV_IDY.Enabled = True
Command1.Enabled = True
M_INV_IDY.SetFocus
End Sub

Private Sub newrecord()
CANCELRECORD
M_INV_IDY.Enabled = False
M_INV_DAT.SetFocus
SSTab1.Tab = 0
End Sub

Private Sub v_dbt_lst()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where act_nme like '" & Trim(M_DBT_IDY) & "%' and act_idy not like '" & Sincrement(M_CRT_IDY) & "%' and act_typ like ''", con, adOpenKeyset, adLockPessimistic
ElseIf Check2.Value = 1 Then
    REC_SET.Open "select * from act where act_ad4 like '" & Trim(M_DBT_IDY) & "%' and act_idy not like '" & Sincrement(M_CRT_IDY) & "%' and act_typ like ''", con, adOpenKeyset, adLockPessimistic
End If
list3.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list3.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    Label49 = "Select Debit Account"
    list3.Visible = True
    list3.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Debit Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_DBT_IDY.SetFocus
    End If
End If
End Sub

Private Sub v_crt_lst()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where ACT_NME LIKE '" & Trim(M_CRT_IDY) & "%' and act_idy not like '" & Sincrement(M_DBT_IDY) & "%' and ACT_TYP LIKE ''", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from act where ACT_AD4 LIKE '" & Trim(M_CRT_IDY) & "%' and act_idy not like '" & Sincrement(M_DBT_IDY) & "%'and ACT_TYP LIKE ''", con, adOpenKeyset, adLockPessimistic
End If
List2.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    Label49 = "Select Credit Account"
    List2.Visible = True
    List2.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Account Not Found, Define Account...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_CRT_IDY.SetFocus
End If
End If
End Sub

Private Sub EXT_REC()
Unload Me
End Sub

Private Sub Command1_Click()
Dim RSBRD As New ADODB.Recordset
MS.Cols = 16
MS.ColWidth(15) = 0
If RSBRD.State = 1 Then RSBRD.Close
RSBRD.Open "SELECT * FROM BRD,PRD WHERE PRD.STK_TYP='Salable' and prd.prd_idy LIKE brd.prd_idy AND PRD_QTY>0", con, adOpenKeyset, adLockPessimistic
MS.Rows = RSBRD.RecordCount + 1
For r = 1 To MS.Rows
    If RSBRD.EOF = False Then
        MS.TextMatrix(r, 0) = Trim(RSBRD!PRD_NME)
        MS.TextMatrix(r, 1) = "Salable"
        MS.TextMatrix(r, 2) = QTY(RSBRD!prd_qty)
        MS.TextMatrix(r, 3) = 0
        MS.TextMatrix(r, 4) = "0.000"
        MS.TextMatrix(r, 5) = "0.00"
        MS.TextMatrix(r, 6) = convert(RSBRD!PRD_PDR)
        MS.TextMatrix(r, 7) = convert(RSBRD!PRD_MRP)
        MS.TextMatrix(r, 8) = Trim(RSBRD!Bat_nme)
        MS.TextMatrix(r, 9) = convert(RSBRD!PRD_SDR)
        MS.TextMatrix(r, 10) = Format(RSBRD!bRD_dom, "DD/MMM/YYYY")
        MS.TextMatrix(r, 11) = Format(RSBRD!brd_doe, "DD/MMM/YYYY")
        MS.TextMatrix(r, 12) = Trim(RSBRD!prd_idy)
        MS.TextMatrix(r, 13) = Trim(RSBRD!BAT_IDY)
        MS.TextMatrix(r, 14) = Val(RSBRD!SUB_UNT)
        RSBRD.MoveNext
    End If
Next
LESSDAMAGE
TOTALS
MS.SetFocus
MS.Col = 2
End Sub
Private Sub LESSDAMAGE()
Me.MousePointer = vbHourglass
Dim rs As New ADODB.Recordset
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('PUR','SAL','CNV') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 13) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('STO','ADJ') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 13) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) - Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Salable' and txn_typ in('CNV') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 13) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) - Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 15)) > 0 Then
        FLXQTY = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 14)))
        MS.TextMatrix(A, 2) = QTY(STM((Val(FLXQTY) - Val(MS.TextMatrix(A, 15))), Val(MS.TextMatrix(A, 14))))
    End If
Next

For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 15)) > 0 Then
        MS.AddItem ""
        For b = 0 To 14
            MS.TextMatrix(MS.Rows - 1, b) = Trim(MS.TextMatrix(A, b))
        Next
        MS.TextMatrix(MS.Rows - 1, 1) = "Damage"
        MS.TextMatrix(MS.Rows - 1, 2) = QTY(STM(Val(MS.TextMatrix(A, 15)), Val(MS.TextMatrix(A, 14))))
    End If
Next

For A = 1 To MS.Rows - 1
    MS.TextMatrix(A, 3) = MS.TextMatrix(A, 2)
Next
MS.Col = 0
MS.Sort = 1
Me.MousePointer = vbDefault
End Sub

Private Sub M_CRT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Account ID or Press <Enter> to View Existing Accounts"
End Sub

Private Sub M_CRT_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub M_FLX_AMT_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If CHKSUBUNT(Val(M_FLX_AMT), Val(MS.TextMatrix(MS.Row, 14))) = False Then
        M_FLX_AMT = QTY(M_FLX_AMT)
        SendKeys "{HOME}+{END}"
        M_FLX_AMT.Visible = True
        M_FLX_AMT.SetFocus
        Exit Sub
    End If
    MS.Text = QTY(M_FLX_AMT)
    FLXQTY = MTS(Val(M_FLX_AMT), Val(MS.TextMatrix(MS.Row, 14)))
    CURQTY = MTS(Val(MS.TextMatrix(MS.Row, 2)), Val(MS.TextMatrix(MS.Row, 14)))
'    If Val(FLXQTY) - Val(CURQTY) > 0 Then
'        MsgBox "Invalid Quantity", vbExclamation
'        m_flx_amt.Visible = True
'        m_flx_amt.SetFocus
'        SendKeys "{home}+{end}"
'        Exit Sub
'    Else
        MS.TextMatrix(MS.Row, 4) = QTY(STM(Val(FLXQTY) - Val(CURQTY), Val(MS.TextMatrix(MS.Row, 14))))
        MS.TextMatrix(MS.Row, 5) = convert(RATCAL(QTY(Abs(MS.TextMatrix(MS.Row, 4))), Val(MS.TextMatrix(MS.Row, 6)), Val(MS.TextMatrix(MS.Row, 14))))
    'End If
    M_FLX_AMT.Visible = False
    If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
    MS.SetFocus
    TOTALS
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub M_INV_DAT_CloseUp()
M_INV_DAT.SetFocus
End Sub

Private Sub M_INV_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date Using Arrow Keys Or Press F4"
End Sub

Private Sub M_INV_DAT_KeyDown(KeyCode As Integer, s As Integer)
If KeyCode = 13 Then
M_DBT_IDY.SetFocus
End If
End Sub

Private Sub M_DBT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Account ID or Press <Enter> to View Account List"
End Sub

Private Sub M_DBT_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_DBT_IDY)) & "' AND ACT_TYP LIKE ''", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        M_DBT_IDY = Sincrement(M_DBT_IDY)
        If Len(Trim(M_DBT_IDY)) = 0 Then Exit Sub
        M_SUP_NME = Trim(REC_SET("act_nme") & "")
        SSTab1.Tab = 0
        If K = 38 Then
            M_INV_DAT.SetFocus
        Else
            M_CRT_IDY.SetFocus
        End If
    Else
        v_dbt_lst
    End If
End If
End Sub

Private Sub M_DBT_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Form_Activate()
checkdongle
If Me.Tag <> "Batch" Then
SSTab1.Tab = 0
List7 = "Y"
If M_INV_IDY.Enabled = True Then
    M_INV_IDY.SetFocus
Else
    M_INV_DAT.SetFocus
End If
Label23 = labdisp("IDY_ADJ")
SSTab1.TabEnabled(4) = False
SSTab1.TabVisible(2) = False
SSTab1.TabVisible(3) = False
setval
SETYEAR M_FRM_ADJ
End If
End Sub

Private Sub V_INV_LST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inl WHERE txn_typ ='ADJ'", con, adOpenKeyset, adLockPessimistic
list5.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , convert(REC_SET("NET_AMT") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    Label49 = "Select Voucher"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Voucher Number Not Found...!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If list1.Visible = False And List2.Visible = False And list3.Visible = False And list4.Visible = False And list5.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then EXT_REC
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        
        If M_FRM_ADJ.ActiveControl.name = "M_INV_IDY" Then
            V_INV_LST
        ElseIf M_FRM_ADJ.ActiveControl.name = "M_DBT_IDY" Then
            v_dbt_lst
        ElseIf M_FRM_ADJ.ActiveControl.name = "M_SMN_IDY" Then
            
        ElseIf M_FRM_ADJ.ActiveControl.name = "M_CRT_IDY" Then
            v_crt_lst
        ElseIf M_FRM_ADJ.ActiveControl.name = "Text3" Then
            
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_ADJ.ActiveControl Is TextBox Then M_FRM_ADJ.ActiveControl.Text = FUNKEY(KeyCode)
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
    
    M_INV_IDY = list5.SelectedItem
    DISPDETAILS
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_SMN_IDY.SetFocus
    ElseIf List2.Visible = True Then
        List2.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_CRT_IDY.SetFocus
    ElseIf list3.Visible = True Then
        list3.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_DBT_IDY.SetFocus
    ElseIf list5.Visible = True Then
        list5.Visible = False
        SSTab1.Tab = 0
        If M_INV_IDY.Enabled = True Then
            M_INV_IDY.SetFocus
        Else
            M_INV_DAT.SetFocus
        End If
    ElseIf list4.Visible = True Then
        list4.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
        If Text3.Visible = True Then
            Text3.SetFocus
        Else
            Check2.Caption = "Place"
            SSTab1.Tab = 1
            MS.Row = MS.Rows - 1
            MS.SetFocus
        End If
    Else
        If M_FRM_ADJ.ActiveControl.name = "M_INV_IDY" Then
            Unload Me
        Else
            saverecord
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
MOD_REC = False
With MS
.Cols = 16
MS.ColWidth(0) = 2500
MS.ColWidth(1) = 1400
MS.ColWidth(2) = 1050
MS.ColWidth(3) = 1050
MS.ColWidth(4) = 1050
MS.ColWidth(5) = 1050
MS.ColWidth(6) = 1050
MS.ColWidth(7) = 1050
MS.ColWidth(8) = 1500
MS.ColWidth(9) = 1050
MS.ColWidth(10) = 1200
MS.ColWidth(11) = 1200
MS.ColWidth(12) = 0
MS.ColWidth(13) = 0
MS.ColWidth(14) = 0
MS.ColWidth(15) = 0
.TextMatrix(0, 0) = "Product Name"
.TextMatrix(0, 1) = "Stock Type"
.TextMatrix(0, 2) = "Quantity"
.TextMatrix(0, 3) = "Physical Qty"
SCl MS, 3
.TextMatrix(0, 4) = "Variance"
.TextMatrix(0, 5) = "Amount"
.TextMatrix(0, 6) = "Pur.Rate"
.TextMatrix(0, 7) = "     MRP"
.TextMatrix(0, 8) = "Batch"
.TextMatrix(0, 9) = "Sale Rate(B)"
.TextMatrix(0, 10) = "Mfg.Date"
.TextMatrix(0, 11) = "BBD-Exp.Date"
.TextMatrix(0, 12) = "Product"
.TextMatrix(0, 13) = "Batch ID"
.TextMatrix(0, 14) = "SUBUNT"
For A = 2 To 7
    .ColAlignment(A) = 7
Next
list5.ColumnHeaders(1).Width = List2.Width / 4
list5.ColumnHeaders(2).Width = List2.Width / 4
list5.ColumnHeaders(3).Width = List2.Width / 4

list1.ColumnHeaders(1).Width = 3800
List2.ColumnHeaders(1).Width = 3800
list3.ColumnHeaders(1).Width = 3800
Label23 = labdisp("IDY_ADJ")
M_LRN_AMT = "0.00"
deliverycharges = "0.00"
M_LRN_IDY = 0
waybill = 0
M_NET_AMT = "0.00"
M_GRS_AMT = "0.00"
M_CRT_DAY = 0
waybill = 0
M_SPD_DED = "0.00"
M_PRD_TOT = "0.00"
tax = "0.00"
End With
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press Ctrl+N to Generate New Opening Voucher ID by System, Press <Enter> to View Existing Vouchers"
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
If M_INV_IDY = "" And K = 27 Then Unload Me
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
        M_FFR_NME = list1.SelectedItem
        SSTab1.Tab = 0
        M_CRT_IDY.SetFocus
        list1.Visible = False
    End If
End If
End Sub

Private Sub List1_DblClick()
List1_keypress 13
End Sub

Private Sub list2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If List2.SortOrder = lvwAscending Then
List2.SortOrder = lvwDescending
ElseIf List2.SortOrder = lvwDescending Then
List2.SortOrder = lvwAscending
End If
List2.SortKey = ColumnHeader.Index - 1
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

Private Sub list2_DblClick()
list2_keypress 13
End Sub

Private Sub list3_DblClick()
List3_KeyPress 13
End Sub

Private Sub LIST5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If list5.ListItems.count > 0 Then
    M_INV_IDY = list5.SelectedItem
    list5.Visible = False
    M_INV_IDY_KeyDown 13, 0
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
    
ElseIf Label49.Caption = "Select Debit Account" Then
    v_dbt_lst
ElseIf Label49.Caption = "Select Voucher" Then
    V_INV_LST
ElseIf Label49.Caption = "Select Credit Account" Then
    v_crt_lst
ElseIf Label49.Caption = "Select Product" Then
    
End If
End Sub

Private Sub M_WBL_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Way Bill Date, Press F4 to get View"
End Sub


Private Sub MS_KeyPress(K As Integer)
If MS.Col = 3 Then
    If Len(Trim(MS.TextMatrix(MS.Row, 0))) > 0 Then
        If K > 47 And K < 58 Or K = 46 Then
            With M_FLX_AMT
                .Text = ""
                .Visible = True
                .Left = MS.CellLeft + MS.Left
                .Top = MS.Top + MS.CellTop
                .Height = MS.CellHeight
                .Width = MS.CellWidth
                .Text = Chr(K)
                .SelStart = Len(Trim(.Text))
                .SetFocus
            End With
        End If
    End If
End If
End Sub

Private Sub MS_Scroll()
If M_FLX_AMT.Visible = True Then
    M_FLX_AMT = ""
    M_FLX_AMT.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
newrecord
ElseIf Button.KEY = "s" Then
saverecord
ElseIf Button.KEY = "e" Then
EXT_REC
ElseIf Button.KEY = "p" Then
Printrecord
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "pre" Then
M_FRM_PRE.Caption = "Prefix Purchase"
M_FRM_PRE.Show 1
ElseIf Button.KEY = "crt" Then
CREATE
ElseIf Button.KEY = "ed" Then
    If M_FRM_ADJ.ActiveControl.name = "M_INV_IDY" Then
        Label49 = "Select Voucher"
        V_INV_LST
    ElseIf M_FRM_ADJ.ActiveControl.name = "M_DBT_IDY" Then
        Label49 = "Select Debit Account"
        
        v_dbt_lst
    ElseIf M_FRM_ADJ.ActiveControl.name = "M_SMN_IDY" Then
        Label49 = "Select Salesman"
        
    ElseIf M_FRM_ADJ.ActiveControl.name = "M_CRT_IDY" Then
        Label49 = "Select Credit Account"
        v_crt_lst
    ElseIf M_FRM_ADJ.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        
    End If
ElseIf Button.KEY = "DAY" Then
'''''    If Button.Value = tbrPressed Then
'''''        Text1.Visible = True
'''''        Text1 = Format(Date, "dd/mm/yyyy") & " " & Format(Date, "dddd")
'''''        Text1.BackColor = vbBlue
'''''        Text1.ForeColor = vbWhite
'''''    Else
'''''        Text1.Visible = False
'''''    End If
ElseIf Button.KEY = "h" Then

M_HELP
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
If check("prt", "prt_nme", "M_FRM_ADJ") = False Then
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
    V_INV_LST
Else
    inV = M_INV_IDY
    clrctr M_FRM_ADJ
    M_SUP_NME = ""
    M_SUP_AD1 = ""
    M_SUP_AD2 = ""
    M_SUP_AD3 = ""
    M_SUP_AD4 = ""
    Label17 = "0.00"
    Label18 = "0"
    Label19 = "0.00"
    Label20 = "0.00"
    M_GRS_AMT = "0.00"
    M_NET_AMT = "0.00"
    M_FFR_NME = ""
    M_ACT_NME = ""
    M_INV_IDY = inV
    DISPDETAILS
End If
End If
End Sub

Public Sub DISPDETAILS()
Dim rs As New ADODB.Recordset
If Len(Trim(M_INV_IDY)) <> 10 Then
    rs.Open "SELECT * FROM INL WHERE INV_IDY LIKE 'ADJ" & Tincrement(Trim(M_INV_IDY)) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "SELECT * FROM INL WHERE INV_IDY LIKE '" & (Trim(M_INV_IDY)) & "'", con, adOpenKeyset, adLockPessimistic
End If
If rs.EOF = False Then
Command1.Enabled = False
    If Len(Trim(M_INV_IDY)) <> 10 Then
        M_INV_IDY = "ADJ" & Tincrement(Trim(M_INV_IDY))
    Else
        M_INV_IDY = Trim(M_INV_IDY)
    End If
    M_INV_DAT = rs!INV_DAT
    M_CRT_IDY = rs!ACT_IDY
    M_DBT_IDY = rs!CST_IDY
    M_SUP_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(M_DBT_IDY) & "'")
    M_ACT_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(M_CRT_IDY) & "'")
End If
If rs.State = 1 Then rs.Close
rs.Open "SELECT INV.*,PRD_NME,BAT_NME,BRD_DOM,BRD_DOE,PRD_SDR,PRD_MRP FROM INV,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND INV.BAT_IDY LIKE BRD.BAT_IDY AND INV_IDY LIKE '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
MS.Rows = rs.RecordCount + 1
For A = 1 To rs.RecordCount
    MS.TextMatrix(A, 0) = Trim(rs!PRD_NME)
    MS.TextMatrix(A, 1) = Trim(rs!stk_typ)
    MS.TextMatrix(A, 2) = QTY(rs!OPN_QTY & "")
    MS.TextMatrix(A, 3) = QTY(rs!ISS_QTY & "")
    MS.TextMatrix(A, 4) = QTY(rs!prd_qty & "")
    MS.TextMatrix(A, 5) = 0
    MS.TextMatrix(A, 6) = convert(rs!PRD_rat & "")
    MS.TextMatrix(A, 7) = convert(rs!PRD_MRP)
    MS.TextMatrix(A, 8) = Trim(rs!Bat_nme)
    MS.TextMatrix(A, 9) = convert(rs!PRD_SDR)
    MS.TextMatrix(A, 10) = Format(rs!bRD_dom, "DD/MMM/YYYY")
    MS.TextMatrix(A, 11) = Format(rs!brd_doe, "DD/MMM/YYYY")
    MS.TextMatrix(A, 12) = Trim(rs!prd_idy)
    MS.TextMatrix(A, 13) = Trim(rs!BAT_IDY)
    MS.TextMatrix(A, 14) = Val(rs!SUB_UNT)
    MS.TextMatrix(A, 5) = convert(RATCAL(QTY(Abs(MS.TextMatrix(A, 4))), Val(MS.TextMatrix(A, 6)), Val(MS.TextMatrix(A, 14))))
    rs.MoveNext
Next
TOTALS
SSTab1.Tab = 0
CURRENTSTOCK
M_INV_IDY.Enabled = False
M_INV_DAT.SetFocus
MOD_REC = True
End Sub

Private Sub list2_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If List2.ListItems.count > 0 Then
        M_CRT_IDY = List2.SelectedItem.ListSubItems.Item(1)
        M_ACT_NME = List2.SelectedItem
        SSTab1.Tab = 1
            If MOD_REC = True Then
                MS.SetFocus
            Else
            If Command1.Enabled = True Then Command1.SetFocus
            End If

        
        List2.Visible = False
    End If
End If
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list3.ListItems.count > 0 Then
        M_DBT_IDY = list3.SelectedItem.ListSubItems(1)
        M_DBT_IDY_keydown 13, 0
        SSTab1.Tab = 0
        list3.Visible = False
        M_CRT_IDY.SetFocus
    End If
End If
End Sub

Private Sub list4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list4.ListItems.count > 0 Then
        list4.Visible = False
        If list4.ListItems.count = 0 Then
            Option3.Visible = False
            Exit Sub
        End If
        Set ls = list4.SelectedItem
        
    End If
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
'    M_PRD_TOT.Visible = True
    If M_INV_IDY.Enabled = True Then
        M_INV_IDY.SetFocus
    Else
        M_INV_DAT.SetFocus
    End If
ElseIf SSTab1.Tab = 1 Then
    MS.Col = 0
    MS.Row = 1
 '   M_PRD_TOT.Visible = True
    MS.SetFocus
    If Command1.Enabled = True Then Command1.SetFocus
ElseIf SSTab1.Tab = 2 Then
  '  M_PRD_TOT.Visible = True
    M_ROU_OFF.SetFocus
End If
If list1.Visible = True Then list1.Visible = False
If List2.Visible = True Then List2.Visible = False
If list3.Visible = True Then list3.Visible = False
If list4.Visible = True Then list4.Visible = False
If list5.Visible = True Then list5.Visible = False

M_FLX_AMT.Visible = False
If SSTab1.Tab = 3 Then
    M_LRN_IDY.SetFocus
End If
If SSTab1.Tab = 4 Then
Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTab1.Tab = 1 Then MS.SetFocus
    If SSTab1.Tab = 0 Then M_INV_DAT.SetFocus
End If
End Sub


Private Sub M_CRT_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_CRT_IDY)) & "' AND act_idy not like '" & Sincrement(M_DBT_IDY) & "%' and ACT_TYP like '' ", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_CRT_IDY = Sincrement(M_CRT_IDY)
        If K = 38 Then
        M_DBT_IDY.SetFocus
        Else
        M_ACT_NME = Trim(REC_SET("act_nme"))
        SSTab1.Tab = 1
            If MOD_REC = True Then
                MS.SetFocus
            Else
            If Command1.Enabled = True Then Command1.SetFocus
            End If
        End If
    Else
        v_crt_lst
    End If
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_LeaveCell()
If M_FLX_AMT.Visible = True Then M_FLX_AMT.Visible = False
End Sub

Public Sub Printrecord()
If check("PRT", "PRT_NME", "M_FRM_ADJ") = True Then
M_FRM_PRV.Check3.Value = 1
Else
M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "ADJ"
M_FRM_PRV.Show 1
End Sub

Public Sub setval()
clrctr M_FRM_ADJ
'M_ROU_IDY.Selected(0) = True
M_SUP_NME = ""
M_SUP_AD1 = ""
M_SUP_AD2 = ""
M_SUP_AD3 = ""
M_SUP_AD4 = ""
Label17 = "0.00"
Label18 = "0"
Label19 = "0.00"
Label20 = "0.00"
M_GRS_AMT = "0.00"
M_NET_AMT = "0.00"
M_ROU_OFF = "0.00"
PRD_NME = ""
M_FFR_NME = ""
M_ACT_NME = ""
M_SPD_DED = "0.00"
cd = "0.00"
units = "0.00"
tax = "0.00"
M_PRD_TOT = "0.00"
M_INV_DAT = Date
M_DCH_DAT = Date
M_CSD_PCG = "0.00"
M_CSD_AMT = "0.00"
M_SPD_PCG = "0.00"
M_SPD_AMT = "0.00"
M_TAX_PCG = "0.00"
M_TAX_AMT = "0.00"
M_EXD_PCG = "0.00"
M_EXD_AMT = "0.00"
M_FRT_AMT = "0.00"
M_PKG_CST = "0.00"
M_CRT_DAY = "0"
End Sub

Public Sub TOTALS()
For A = 1 To MS.Rows - 1
tunt = Val(tunt) + Abs(Val(MS.TextMatrix(A, 4)))
TAMT = Val(TAMT) + Abs(Val(MS.TextMatrix(A, 5)))
Next
TOTUNT = "Units:  " & QTY(Val(tunt))
TOTAMT = "Amount:  " & convert(Val(TAMT))
End Sub

Public Sub CURRENTSTOCK()
Me.MousePointer = vbHourglass
Me.Refresh
If rs.State = 1 Then rs.Close
con.Execute "DELETE FROM CLD"
rs.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
If RS1.State = 1 Then RS1.Close
RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat < ctod('" & M_INV_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic
For A = 1 To RS1.RecordCount
    If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0)) * -1
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "OPS" Then
        con.Execute "UPDATE CLD SET OPN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!opn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PUR" Then
    con.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PUR_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PRT" Then
    con.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PRT_STK = Val(RS1(0))
            rs!PUR_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STI" Then
    con.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!STI_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!PUR_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SRN" Then
    con.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
    con.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs(7) = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs!PUR_STK = 0
            rs!cls_stk = 0
            rs.Update
        End If
    End If
    RS1.MoveNext
Next
rs.Close
con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
con.Execute "UPDATE CLD SET OPN_STK=CLS_STK"
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM CLD,BRD,PRD WHERE CLD.PRD_IDY LIKE BRD.BAT_IDY AND PRD.PRD_IDY LIKE BRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If rs!BAT_IDY = Trim(MS.TextMatrix(b, 14)) And Trim(MS.TextMatrix(b, 1)) = "Salable" Then MS.TextMatrix(b, 2) = STM(Val(rs!cls_stk & ""), Val(rs!SUB_UNT))
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE inv_dat < ctod('" & M_INV_DAT & "') and STK_TYP LIKE 'Damage' and txn_typ in('PUR','SAL','CNV') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 13) = Trim(rs!BAT_IDY) And Trim(MS.TextMatrix(b, 1)) = "Salable" Then
            MS.TextMatrix(b, 15) = Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE inv_dat < ctod('" & M_INV_DAT & "') and STK_TYP LIKE 'Damage' and txn_typ in('STO','ADJ') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 13) = Trim(rs!BAT_IDY) And Trim(MS.TextMatrix(b, 1)) = "Salable" Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) - Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE inv_dat < ctod('" & M_INV_DAT & "') and STK_TYP LIKE 'Salable' and txn_typ in('CNV') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 13) = Trim(rs!BAT_IDY) And Trim(MS.TextMatrix(b, 1)) = "Salable" Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) - Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 15)) > 0 Then
        FLXQTY = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 14)))
        MS.TextMatrix(A, 2) = QTY(STM((Val(FLXQTY) - Val(MS.TextMatrix(A, 15))), Val(MS.TextMatrix(A, 14))))
    End If
Next

'DAMAGE STOCK
If rs.State = 1 Then rs.Close
rs.Open "SELECT sum(tot_pcs),* FROM inv,BRD,PRD WHERE INV.STK_TYP='Damage' and BRD.PRD_QTY >0 AND prd.prd_idy=brd.prd_idy and inv.prd_idy like brd.prd_idy and txn_typ in('SAL','PUR','CNV') group by inv.bat_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 13) = Trim(rs!BAT_IDY) And Trim(MS.TextMatrix(b, 1)) = "Damage" Then
            MS.TextMatrix(b, 2) = QTY(STM(Val(rs(0) & ""), Val(MS.TextMatrix(A, 14))))
        End If
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('STO','ADJ') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 6) = Trim(rs!BAT_IDY) And Trim(MS.TextMatrix(b, 1)) = "Damage" Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) + Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 15)) <> 0 Then
        FLXQTY = MTS(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 14)))
        MS.TextMatrix(A, 2) = QTY(STM((Val(FLXQTY) - Val(MS.TextMatrix(A, 15))), Val(MS.TextMatrix(A, 14))))
    End If
Next
Me.MousePointer = vbDefault
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
M_FRM_DSN.Tag = "Adjustment"
M_FRM_DSN.Show 1
End Sub
