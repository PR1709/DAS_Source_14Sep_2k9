VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_OPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening Stock Entry"
   ClientHeight    =   5625
   ClientLeft      =   1455
   ClientTop       =   330
   ClientWidth     =   8820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "M_FRM_OPS.frx":0000
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
      TabPicture(0)   =   "M_FRM_OPS.frx":0442
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
      TabPicture(1)   =   "M_FRM_OPS.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MS"
      Tab(1).Control(1)=   "M_FLX_AMT"
      Tab(1).Control(2)=   "Text3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Picture1"
      Tab(1).Control(4)=   "DTPicker1"
      Tab(1).Control(5)=   "List6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Taxes and  Levies"
      TabPicture(2)   =   "M_FRM_OPS.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label25"
      Tab(2).Control(1)=   "Label24"
      Tab(2).Control(2)=   "Label44"
      Tab(2).Control(3)=   "Label42"
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(6)=   "Label8"
      Tab(2).Control(7)=   "Label9"
      Tab(2).Control(8)=   "Frame4"
      Tab(2).Control(9)=   "M_CRT_DAY"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Frame3"
      Tab(2).Control(11)=   "M_DUE_DAT"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "&Shipment Details"
      TabPicture(3)   =   "M_FRM_OPS.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label54"
      Tab(3).Control(1)=   "M_TXN_RMK"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame5"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "View"
      TabPicture(4)   =   "M_FRM_OPS.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label49"
      Tab(4).Control(1)=   "list3"
      Tab(4).Control(2)=   "list1"
      Tab(4).Control(3)=   "list2"
      Tab(4).Control(4)=   "list5"
      Tab(4).Control(5)=   "list4"
      Tab(4).Control(6)=   "Frame1"
      Tab(4).ControlCount=   7
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
         ItemData        =   "M_FRM_OPS.frx":04CE
         Left            =   -70320
         List            =   "M_FRM_OPS.frx":04E4
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -73440
         TabIndex        =   115
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   86114307
         CurrentDate     =   37634
      End
      Begin VB.Frame Frame5 
         Caption         =   "Transport Receipt Details"
         Height          =   2175
         Left            =   -73800
         TabIndex        =   98
         Top             =   960
         Width           =   5655
         Begin VB.TextBox M_WBL_IDY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox M_TXN_TPT 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   1620
            Width           =   4335
         End
         Begin VB.TextBox M_TTL_CAS 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4305
            MaxLength       =   5
            TabIndex        =   105
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
            TabIndex        =   100
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
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   765
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker M_LRN_DAT 
            Height          =   330
            Left            =   4305
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   86114307
            CurrentDate     =   37097
         End
         Begin MSComCtl2.DTPicker M_WBL_DAT 
            Height          =   330
            Left            =   4305
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   86114307
            CurrentDate     =   37097
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WB Number"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   180
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   108
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
            TabIndex        =   106
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
            TabIndex        =   104
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
            TabIndex        =   103
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
            TabIndex        =   102
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
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Height          =   2490
         Left            =   -72000
         TabIndex        =   41
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
            TabIndex        =   92
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
            TabIndex        =   73
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox M_ROU_OFF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            MaxLength       =   4
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   84
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
            Top             =   315
            Width           =   120
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   -74940
         ScaleHeight     =   195
         ScaleWidth      =   8640
         TabIndex        =   31
         Top             =   4800
         Width           =   8700
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty."
            Height          =   195
            Left            =   2640
            TabIndex        =   39
            Top             =   0
            Width           =   285
         End
         Begin VB.Label PRD_NME 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   1080
            TabIndex        =   38
            Top             =   0
            Width           =   45
         End
         Begin VB.Label AMT_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   45
         End
         Begin VB.Label TAX_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   45
         End
         Begin VB.Label ADD_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   45
         End
         Begin VB.Label DIS_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4080
            TabIndex        =   34
            Top             =   0
            Width           =   45
         End
         Begin VB.Label UOM_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3240
            TabIndex        =   33
            Top             =   0
            Width           =   45
         End
         Begin VB.Label UNT_TOT 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   1200
            TabIndex        =   32
            Top             =   0
            Width           =   45
         End
      End
      Begin VB.TextBox M_TXN_RMK 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72580
         MaxLength       =   70
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74640
         TabIndex        =   20
         Top             =   4560
         Width           =   8295
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   100
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
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   100
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73680
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
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
         Format          =   86114307
         CurrentDate     =   37099
      End
      Begin MSMask.MaskEdBox M_FLX_AMT 
         Height          =   375
         Left            =   -69240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
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
         Height          =   4395
         Left            =   -74940
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   26
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
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
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74640
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
      Begin VB.Frame Frame4 
         Height          =   2400
         Left            =   -69615
         TabIndex        =   63
         Top             =   3120
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox M_AMT_LES 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   240
            MaxLength       =   70
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   2000
            Width           =   1695
         End
         Begin VB.TextBox M_AMT_ADD 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   240
            MaxLength       =   70
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   1640
            Width           =   1695
         End
         Begin VB.TextBox M_EXD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1280
            Width           =   615
         End
         Begin VB.TextBox M_EXD_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   1280
            Width           =   855
         End
         Begin VB.TextBox M_FRT_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   2000
            Width           =   855
         End
         Begin VB.TextBox M_PKG_CST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   1640
            Width           =   855
         End
         Begin VB.TextBox M_CSD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   200
            Width           =   615
         End
         Begin VB.TextBox M_SPD_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   560
            Width           =   615
         End
         Begin VB.TextBox M_TAX_PCG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2220
            MaxLength       =   5
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   910
            Width           =   615
         End
         Begin VB.TextBox M_TAX_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3075
            MaxLength       =   9
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   910
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
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   560
            Width           =   855
         End
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
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   200
            Width           =   855
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount (+)"
            Height          =   195
            Left            =   2070
            TabIndex        =   89
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount (-)"
            Height          =   195
            Left            =   2115
            TabIndex        =   88
            Top             =   2070
            Width           =   720
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sur - Charge on Tax % "
            Height          =   195
            Left            =   435
            TabIndex        =   85
            Top             =   1350
            Width           =   1635
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tax % "
            Height          =   195
            Left            =   1575
            TabIndex        =   72
            Top             =   975
            Width           =   480
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bill Discount % "
            Height          =   195
            Left            =   975
            TabIndex        =   71
            Top             =   270
            Width           =   1080
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Additonal Discount % "
            Height          =   195
            Left            =   510
            TabIndex        =   70
            Top             =   630
            Width           =   1545
         End
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
         TabIndex        =   116
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   30
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
         TabIndex        =   24
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
         TabIndex        =   18
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
         TabIndex        =   17
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
            Picture         =   "M_FRM_OPS.frx":051C
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":0A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":0E28
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":117C
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":17B4
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":1B08
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":1F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":225C
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":2668
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":2774
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":2AC8
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":2EDC
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPS.frx":3428
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
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
Attribute VB_Name = "M_FRM_OPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, bslab, MSROW, INVIDY, INC, dlr_idy
Dim M_CPN_AMT As Double
Dim M_INC_AMT As Double
Dim LST_TOT As Double
Private Sub CREATE()
MOD_REC = False
M_INV_IDY = ""
saverecord
End Sub

Public Sub GRS_CAL()
On Error Resume Next
G_PRD_VAL = convert(Val(AMT_TOT))
T_PRD_VAL = convert(Val(AMT_TOT))

G_PRD_DIS = convert(Val(DIS_TOT) + Val(ADD_TOT))
T_PRD_DIS = convert(AMT_TOT - convert(Val(DIS_TOT) + Val(ADD_TOT)))

G_PRD_TAX = convert(Val(TAX_TOT))
T_PRD_TAX = convert(Val(TAX_TOT) + Val(T_PRD_DIS))

G_GRS_AMT = T_PRD_TAX
T_GRS_AMT = T_PRD_TAX

G_CSD_PCG = convert(Val(M_CSD_AMT))
T_CSD_AMT = convert(Val(T_GRS_AMT) - Val(G_CSD_PCG))

G_SPD_PCG = convert(Val(M_SPD_AMT))
T_SPD_AMT = convert(Val(T_CSD_AMT) - Val(G_SPD_PCG))

G_TAX_PCG = convert(Val(M_TAX_AMT))
T_TAX_AMT = convert(Val(M_TAX_AMT) + Val(T_SPD_AMT))

G_SCT_PCG = convert(Val(M_EXD_AMT))
T_SCT_PCG = convert(Val(M_EXD_AMT) + Val(T_TAX_AMT))

G_PKG_CST = convert(Val(M_PKG_CST))
T_PKG_CST = convert(Val(G_PKG_CST) + Val(T_SCT_PCG))

G_FRT_AMT = convert(Val(M_FRT_AMT))
T_FRT_AMT = convert(Val(T_PKG_CST) - Val(G_FRT_AMT))
ROU_OFF
T_ROU_OFF = convert(Val(T_FRT_AMT) + Val(M_ROU_OFF))
T_NET_AMT = convert(T_ROU_OFF)
End Sub

Private Sub saverecord()
On Error Resume Next
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
Dim REC_IRT As New ADODB.Recordset
Dim REC_PNL As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
Else
    CheckShutDownStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'")
    If Trim(CheckShutDownStatus) = "IMPROPER SHUTDOWN" Then
        MsgBox ("Please reset the Improper Shutdown flag in System-Shutdown Status"), vbCritical, "DAS Version 4.0"
        CANCELRECORD
        Exit Sub
    End If
End If
If CHECKTB("act.dbf,inv.dbf,inl.dbf,prd.dbf,brd.dbf,pnl.dbf,pnv.dbf,dir.dbf,prt.dbf,rri.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
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

If Val(T_NET_AMT) < 0 Then
    MsgBox "Can't Save Opening Voucher, NetAmount is Invalid...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 1
    MS.SetFocus
    Exit Sub
End If
' Day Closure Routine Added by PNR and PKB as Addon Control

Dim RS_DAYCLS As New ADODB.Recordset
If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
RS_DAYCLS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) like 'DAY' .and. alltrim(DIR_ALS) like '" & M_INV_DAT & "'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
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

GRS_CAL
If MOD_REC = True Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'OPS'", con, adOpenKeyset
    For A = 1 To REC_SET.RecordCount
        con.Execute "update prd where prd_idy like '" & REC_SET("prd_idy") & "' set cur_lev = " & CONQTY(Val(REC_SET("prd_qty")), REC_SET("prd_idy"))
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & REC_SET("prd_idy") & "' set PRD_PCS = PRD_PCS - " & MTS(Val(REC_SET("prd_qty")), Val(REC_SET("SUB_UNT")))
        con.Execute "update Brd where BAT_idy like '" & REC_SET("BAT_idy") & "' set PRD_QTY = " & CONBRDQTY(Val(REC_SET("PRD_qty")), REC_SET("BAT_idy"))
        REC_SET.MoveNext
    Next
    con.Execute "delete from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'OPS'"
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'OPS'", con, adOpenKeyset, adLockPessimistic
    INC = Trim(M_INV_IDY)
Else
    REC_SET.Open "select * from inl where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
    REC_SET.AddNew
    REC_SET("inv_idy") = "OPS" & T7increment("idy_ops")
    INC = REC_SET("inv_idy")
    INVIDY = INC
End If
REC_SET("inv_dat") = M_INV_DAT
REC_SET("ord_idy") = Trim(M_PUR_IDY & "")
REC_SET("cst_idy") = Trim(M_DBT_IDY & "")
REC_SET("act_idy") = Trim(M_CRT_IDY & "")
REC_SET("dch_idy") = M_DCH_IDY & ""
REC_SET("dch_dat") = M_DCH_DAT
REC_SET("csd_pcg") = discount(M_CSD_PCG)
REC_SET("csd_amt") = convert(amt(M_CSD_AMT))
REC_SET("spd_pcg") = discount(M_SPD_PCG)
REC_SET("spd_amt") = convert(amt(M_SPD_AMT))
REC_SET("tax_pcg") = discount(M_TAX_PCG)
REC_SET("tax_amt") = convert(amt(M_TAX_AMT))
REC_SET("exd_pcg") = Val(M_EXD_PCG)             'SURCHARGE PCG
REC_SET("EXD_AMT") = Val(M_EXD_AMT)             'SURCHARGE AMT
REC_SET("pkg_cst") = Val(M_PKG_CST)             'AMOUNT (+)
REC_SET("FRT_AMT") = Val(M_FRT_AMT & "")        'AMOUNT (-)
REC_SET("LRN_IDY") = Trim(M_LRN_IDY)
REC_SET("LRN_AMT") = Val(M_LRN_AMT & "")
REC_SET("LRN_DAT") = M_LRN_DAT
REC_SET("WBL_IDY") = Trim(M_WBL_IDY & "")
REC_SET("WBL_DAT") = M_WBL_DAT
REC_SET("TXN_RMK") = Trim(M_TXN_RMK)
REC_SET("TTL_CAS") = Val(M_TTL_CAS)
REC_SET("txn_typ") = "OPS"
REC_SET("grs_amt") = convert(amt(Val(T_GRS_AMT)))
REC_SET("net_amt") = Val(T_NET_AMT)
REC_SET("crt_day") = Val(M_CRT_DAY)
REC_SET("due_dat") = Date
REC_SET("rou_off") = Val(M_ROU_OFF)
REC_SET("ROU_IDY") = Trim(M_ROU_IDY & "")
REC_SET("TXN_TPT") = Trim(M_TXN_TPT)
REC_SET("TXN_RM1") = Trim(M_AMT_ADD)
REC_SET("TXN_RM2") = Trim(M_AMT_LES)
REC_SET!USR_IDY = Trim(USERNAME)
REC_SET!sys_dat = Date
REC_SET!sys_tim = Format(Time, "HH:MM:SS")
'If SCH_CHK.Value = 1 Then
    REC_SET("SCH_CHK") = "Y"
'Else
'    REC_SET("SCH_CHK") = "N"
'End If
REC_SET!SRN_AMT = 0
REC_SET!vch_idy = ""
REC_SET!Sch_Idy = ""
REC_SET.Update
'PRINT DATABASE LOGIC BEGIN
If MOD_REC = True Then
    PNL.Execute "delete from pnv where inv_idy like '" & Trim(INC) & "'"
    PNL.Execute "delete from pnl where inv_idy like '" & Trim(INC) & "'"
End If
REC_PNL.Open "SELECT * FROM PNL", PNL, adOpenKeyset, adLockPessimistic
REC_PNL.AddNew
REC_PNL!INV_IDY = INC
REC_PNL("inv_dat") = M_INV_DAT
REC_PNL("ord_idy") = Trim(M_PUR_IDY & "")
REC_PNL("cst_idy") = Trim(M_DBT_IDY & "")
REC_PNL("cst_nme") = Trim(M_SUP_NME)
REC_PNL("cst_ad1") = Trim(M_SUP_AD1)
REC_PNL("cst_ad2") = Trim(M_SUP_AD2)
REC_PNL("cst_ad3") = Trim(M_SUP_AD3)
REC_PNL("cst_ad4") = Trim(M_SUP_AD4)
REC_PNL("cst_zip") = ""
REC_PNL("Act_idy") = Trim(M_CRT_IDY & "")
REC_PNL("dch_idy") = M_DCH_IDY & ""
REC_PNL("dch_dat") = Date
REC_PNL("csd_pcg") = discount(M_CSD_PCG)
REC_PNL("csd_amt") = convert(amt(M_CSD_AMT))
REC_PNL("spd_pcg") = discount(M_SPD_PCG)
REC_PNL("spd_amt") = convert(amt(M_SPD_AMT))
REC_PNL("tax_pcg") = discount(M_TAX_PCG)
REC_PNL("tax_amt") = convert(amt(M_TAX_AMT))
REC_PNL("exd_pcg") = Val(M_EXD_PCG)             'SURCHARGE PCG
REC_PNL("EXD_AMT") = Val(M_EXD_AMT)             'SURCHARGE AMT
REC_PNL("pkg_cst") = Val(M_PKG_CST)             'AMOUNT (+)
REC_PNL("FRT_AMT") = Val(M_FRT_AMT & "")        'AMOUNT (-)
REC_PNL("LRN_IDY") = Trim(M_LRN_IDY)
REC_PNL("LRN_AMT") = Val(M_LRN_AMT & "")
REC_PNL("LRN_DAT") = M_LRN_DAT
REC_PNL("WBL_IDY") = Trim(M_WBL_IDY & "")
REC_PNL("WBL_DAT") = M_WBL_DAT
REC_PNL("TXN_RMK") = Trim(M_TXN_RMK)
REC_PNL("TTL_CAS") = Val(M_TTL_CAS)
REC_PNL("txn_typ") = "OPS"
REC_PNL("grs_amt") = convert(amt(Val(T_GRS_AMT)))
REC_PNL("net_amt") = Val(T_NET_AMT)
REC_PNL("crt_day") = Val(M_CRT_DAY)
REC_PNL("due_dat") = Date
REC_PNL("rou_off") = Val(M_ROU_OFF)
REC_PNL("ROU_IDY") = Trim(M_ROU_IDY & "")
REC_PNL("TXN_TPT") = Trim(M_TXN_TPT)
REC_PNL("TXN_RM1") = Trim(M_AMT_ADD)
REC_PNL("TXN_RM2") = Trim(M_AMT_LES)
REC_PNL!usr_nme = Trim(USERNAME)
REC_PNL!sys_dat = Date
REC_PNL!sys_tim = Format(Time, "HH:MM:SS")
'If SCH_CHK.Value = 1 Then
'    REC_PNL("SCH_CHK") = "Y"
'Else
    REC_PNL("SCH_CHK") = "N"
'End If
REC_PNL!SRN_AMT = 0
REC_PNL!vch_idy = ""
REC_PNL!Sch_Idy = ""
REC_PNL!cst_gst = ""
REC_PNL!cst_cst = ""
REC_PNL!mar_nme = ""
REC_PNL!SMN_NME = ""
REC_PNL!loc_nme = ""
REC_PNL!cat_nme = ""
REC_PNL!exd_cal = ""
REC_PNL!SMN_IDY = ""
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
REC_PNL!loc_IDY = ""
REC_PNL!MAR_IDY = ""
REC_PNL!cat_IDY = ""
REC_PNL!FRM_DAT = Date
REC_PNL!TO_DAT = Date
REC_PNL!rou_act = ""
REC_PNL!DIS_CHK = ""
REC_PNL.Update
REC_PNL.Close
'PRINT DATABASE LOGIC END

If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inv where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" And Val(MS.TextMatrix(A, 1)) > 0 Then
        REC_SET.AddNew
        REC_SET("inv_idy") = Trim(INC)
        REC_SET("inv_dat") = Trim(M_INV_DAT)
        REC_SET("prd_idy") = MS.TextMatrix(A, 0)
        REC_SET("prd_qty") = QTY(Val(MS.TextMatrix(A, 1)))
        REC_SET("prd_rat") = convert(amt(Val(MS.TextMatrix(A, 2))))
        REC_SET("trd_dis") = discount(MS.TextMatrix(A, 5))
        REC_SET("trd_tax") = discount(MS.TextMatrix(A, 6))
        REC_SET("LST_RAT") = Val(MS.TextMatrix(A, 22))
        REC_SET("LST_DIS") = Val(MS.TextMatrix(A, 23))
        REC_SET("CPN_AMT") = Val(MS.TextMatrix(A, 1)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        INCAMT = RATCAL(Val(MS.TextMatrix(A, 1) & ""), Val(MS.TextMatrix(A, 24) & ""), Val(MS.TextMatrix(A, 12)))
        PRDNET = RATCAL(Val(MS.TextMatrix(A, 1) & ""), Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 12)))
        LST_TOT = MTS(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12)))
        LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 12))))
        M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
        'M_CPN_AMT = Val(MS.TextMatrix(A, 1)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        'LST_TOT = Val(MS.TextMatrix(A, 1)) * Val(MS.TextMatrix(A, 22))
        REC_SET("ADD_AMT") = convert(amt(Val(MS.TextMatrix(A, 4))))
        REC_SET("DIS_AMT") = ((PRDNET - REC_SET("ADD_AMT")) * Val(MS.TextMatrix(A, 5)) / 100)
        TRDTAX = (PRDNET - (Val(REC_SET("ADD_AMT")) + Val(REC_SET("DIS_AMT")) + Val(INCAMT))) * Val(MS.TextMatrix(A, 6)) / 100
        TRDTAX = TRDTAX + M_CPN_AMT
        REC_SET("TAX_AMT") = Round(Val(TRDTAX & ""), 2)
        REC_SET("txn_typ") = "OPS"
        REC_SET!BAT_IDY = Trim(MS.TextMatrix(A, 13))
        REC_SET!PRD_UOM = convert(amt(MS.TextMatrix(A, 3)))
        REC_SET!SUB_UNT = Val(MS.TextMatrix(A, 12))
        REC_SET!PRD_UNT = Int(Val(MS.TextMatrix(A, 1)))
        REC_SET!PRD_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12))) Mod Val(MS.TextMatrix(A, 12))
        REC_SET!TOT_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12)))
        REC_SET("COL_AMT") = convert(amt(MS.TextMatrix(A, 7)) + LST_TOT)
        REC_SET!stk_typ = Trim(MS.TextMatrix(A, 21))
        REC_SET!stk_upd = "Y"
        REC_SET!ORD_IDY = ""
        REC_SET!INC_TTL = Val(INCAMT)
        REC_SET.Update
        con.Execute "update PRD where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "' set CUR_LEV = " & ADDQTY(Trim(MS.TextMatrix(A, 1)), Trim(MS.TextMatrix(A, 0)))
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS + " & MTS(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12)))
        con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 13)) & "' set prd_qty = " & ADDBRDQTY(Trim(MS.TextMatrix(A, 1)), Trim(MS.TextMatrix(A, 13)))
    End If
Next
'pnv update begin
If REC_PNL.State = 1 Then REC_PNL.Close
REC_PNL.Open "select * from pnv where inv_idy = ''", PNL, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" And Val(MS.TextMatrix(A, 1)) > 0 Then
        REC_PNL.AddNew
        REC_PNL("inv_idy") = Trim(INC)
        REC_PNL("inv_dat") = Trim(M_INV_DAT)
        REC_PNL("prd_idy") = MS.TextMatrix(A, 0)
        REC_PNL("GRP_IDY") = Trim(MS.TextMatrix(A, 18))
        REC_PNL("GRP_CAT") = Trim(MS.TextMatrix(A, 19))
        REC_PNL("GRP_CLS") = Trim(MS.TextMatrix(A, 20))
        REC_PNL("prd_qty") = QTY(Val(MS.TextMatrix(A, 1)))
        REC_PNL("prd_rat") = convert(amt(Val(MS.TextMatrix(A, 2))))
        REC_PNL("MRP_rat") = convert(amt(Val(MS.TextMatrix(A, 8))))
        REC_PNL("trd_dis") = discount(MS.TextMatrix(A, 5))
        REC_PNL("trd_tax") = discount(MS.TextMatrix(A, 6))
        REC_PNL("LST_RAT") = Val(MS.TextMatrix(A, 22))
        REC_PNL("LST_DIS") = Val(MS.TextMatrix(A, 23))
        REC_PNL("CPN_AMT") = Val(MS.TextMatrix(A, 1)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        INCAMT = RATCAL(Val(MS.TextMatrix(A, 1) & ""), Val(MS.TextMatrix(A, 24) & ""), Val(MS.TextMatrix(A, 12)))
        PRDNET = RATCAL(Val(MS.TextMatrix(A, 1) & ""), Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 12)))
        M_CPN_AMT = Val(MS.TextMatrix(A, 1)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        REC_PNL("ADD_AMT") = convert(amt(Val(MS.TextMatrix(A, 4))))
        REC_PNL("DIS_AMT") = ((PRDNET - REC_PNL("ADD_AMT")) * Val(MS.TextMatrix(A, 5)) / 100)
        TRDTAX = (PRDNET - (Val(REC_PNL("ADD_AMT")) + Val(REC_PNL("DIS_AMT")) + Val(INCAMT))) * Val(MS.TextMatrix(A, 6)) / 100
        TRDTAX = TRDTAX + M_CPN_AMT
        REC_PNL("TAX_AMT") = Val(TRDTAX & "")
        REC_PNL("txn_typ") = "OPS"
        REC_PNL!BAT_IDY = Trim(MS.TextMatrix(A, 13))
        REC_PNL!PRD_UOM = convert(amt(MS.TextMatrix(A, 3)))
        REC_PNL!SUB_UNT = Val(MS.TextMatrix(A, 12))
        REC_PNL!PRD_UNT = Int(Val(MS.TextMatrix(A, 1)))
        REC_PNL!PRD_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12))) Mod Val(MS.TextMatrix(A, 12))
        REC_PNL!TOT_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12)))
        REC_PNL("COL_AMT") = convert(amt(MS.TextMatrix(A, 7)) + LST_TOT)
        REC_PNL!stk_upd = "Y"
        REC_PNL("prd_nme") = Trim(MS.TextMatrix(A, 10))
        'REC_PNL("lst_rat") = 0
        'REC_PNL("lst_dis") = 0
        'REC_PNL("lst_rat") = 0
        REC_PNL("ret_qty") = 0
        REC_PNL("iss_qty") = 0
        REC_PNL("tot_qty") = 0
        REC_PNL("rat_app") = ""
        REC_PNL("cpn_rat") = 0
'        REC_PNL("cpn_amt") = 0
        REC_PNL("add_dis") = 0
        REC_PNL("fst_flg") = ""
        REC_PNL("opn_qty") = 0
        REC_PNL("sch_amt") = 0
        REC_PNL("ord_idy") = 0
        REC_PNL("bat_nme") = GetValue("bat_nme", "brd where bat_idy like '" & Trim(MS.TextMatrix(A, 13)) & "'")
        REC_PNL!stk_typ = GetValue("stk_typ", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")
        REC_PNL("sch_idy") = ""
        REC_PNL("INC_TTL") = Val(INCAMT)
        REC_PNL.Update
    End If
Next
'pnv update close

If MOD_REC = False Then MsgBox "New Ops ID is " & INC, vbExclamation, "DAS Version 4.0"
Label23 = labdisp("idy_ops")
If check("PRT", "PRT_NME", "M_FRM_OPS") = True Then
        If querycheck("DIR", "DIR_MSG", "D", "AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
        If REC_PNL.State = 1 Then REC_PNL.Close
        REC_PNL.Open "SELECT * FROM RRI WHERE RI_REPORT LIKE 'Opening Stock Note'", PNL, adOpenKeyset, adLockPessimistic
        If REC_PNL.RecordCount > 0 Then
            REC_PNL!ri_loscope = Trim(INC)
            REC_PNL!ri_hiscope = Trim(INC)
            REC_PNL!ri_scope = "O"
            REC_PNL!ri_filter = ""
            REC_PNL!ri_copies = 1
            REC_PNL!ri_begpage = 1
            REC_PNL!ri_endpage = 9999999
            REC_PNL!ri_printer = ""
            REC_PNL!ri_mindex = CURDIR & "\PNV.CDX"
            REC_PNL.Update
        End If
        REC_PNL.Close
        Shell CURDIR & "\RRUNTIME RRI 2", vbHide
    Else
        OpenDocument CURDIR & "\DOS_VCH.RP1", "Opening Stock Note", 2, Trim(INC), Trim(INC), 6, CURDIR & "\pnv.cdx,c,inv_idy"
    End If
    CANCELRECORD
End If
CANCELRECORD
End Sub

Private Sub CANCELRECORD()
clrctr M_FRM_OPS
setval
Form_Load
MS.Rows = 1
MS.AddItem ""
SSTab1.Tab = 0
MOD_REC = False
M_INV_IDY.Enabled = True
M_INV_IDY.SetFocus
End Sub

Private Sub newrecord()
CANCELRECORD
M_INV_IDY.Enabled = False
M_INV_DAT.SetFocus
SSTab1.Tab = 0
End Sub

Private Sub v_dbt_lst()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where act_nme like '" & Trim(M_DBT_IDY) & "%' and act_idy not like '" & Sincrement(M_CRT_IDY) & "%' and act_typ like '' and (GRP_CLF in('OpeningStockPL','ClsStockPL'))", con, adOpenKeyset, adLockPessimistic
ElseIf Check2.Value = 1 Then
    REC_SET.Open "select * from act where act_ad4 like '" & Trim(M_DBT_IDY) & "%' and act_idy not like '" & Sincrement(M_CRT_IDY) & "%' and act_typ like '' and (GRP_CLF in('OpeningStockPL','ClsStockPL'))", con, adOpenKeyset, adLockPessimistic
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
Private Sub V_SMN_LST()
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
End If
list1.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(REC_SET("smn_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    list1.Visible = True
    list1.SetFocus
    
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_SMN_IDY.SetFocus
    End If
End If
End Sub

Private Sub v_crt_lst()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where ACT_NME LIKE '" & Trim(M_CRT_IDY) & "%' and act_idy not like '" & Sincrement(M_DBT_IDY) & "%' and ACT_TYP LIKE '' and (GRP_CLF in('OpeningStockPL','ClsStockPL'))", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from act where ACT_AD4 LIKE '" & Trim(M_CRT_IDY) & "%' and act_idy not like '" & Sincrement(M_DBT_IDY) & "%'and ACT_TYP LIKE '' and (GRP_CLF in('OpeningStockPL','ClsStockPL'))", con, adOpenKeyset, adLockPessimistic
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

Private Sub List7_Keydown(K As Integer, s As Integer)
If K = 13 Then
    M_CRT_IDY.SetFocus
ElseIf K = 38 And s = 1 Then
    M_DBT_IDY.SetFocus
End If
End Sub

Private Sub M_FRT_TYP_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Freight Type(T->To Pay,P->Paid"
End Sub

Private Sub CHECK3_KeyPress(K As Integer)
If K = 13 Then
M_INV_DAT.SetFocus
End If
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If DateDiff("d", Format(DTPicker1, "dd/MM/yyyy"), Format(MS.Text, "dd/MM/YYYY")) <> 0 Then
        DTPicker1.Visible = False
        createbatch
    Else
        DTPicker1.Visible = False
        MS.SetFocus
    End If
End If
End Sub

Private Sub M_AMT_ADD_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Describe Add Amount Particulars"
End Sub

Private Sub M_AMT_ADD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    M_PKG_CST.SetFocus
ElseIf KeyCode = 38 Then
    M_EXD_AMT.SetFocus
End If
End Sub

Private Sub M_AMT_ADD_KeyPress(KeyAscii As Integer)
KeyAscii = CHGCASE(KeyAscii)
End Sub

Private Sub M_AMT_LES_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Describe Less Amount Particulars"
End Sub

Private Sub M_AMT_LES_KEYDOWN(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    M_FRT_AMT.SetFocus
ElseIf K = 38 Then
    M_PKG_CST.SetFocus
End If
End Sub

Private Sub M_AMT_LES_KeyPress(KeyAscii As Integer)
KeyAscii = CHGCASE(KeyAscii)
End Sub

Private Sub M_EXD_PCG_Change()
If Len(Trim(M_EXD_PCG)) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_EXD_PCG_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Surcharge On Tax Percent"
SendKeys "{home}+{end}"
End Sub

Private Sub M_EXD_PCG_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 38 Or K = 40 Then
M_EXD_PCG = discount(M_EXD_PCG)
M_EXD_AMT = convert(PERCENT(Val(G_PRD_TAX) + Val(M_TAX_AMT), Val(M_EXD_PCG)))
    If K = 38 Then
        M_TAX_AMT.SetFocus
    Else
        M_EXD_AMT.SetFocus
    End If
End If
End Sub

Private Sub M_EXD_PCG_KEYPRESS(K As Integer)
K = CHKNUM(K)
End Sub

Private Sub M_LRN_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Receipt Date, Press F4 to view"
End Sub

Private Sub M_PUR_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Debit Account Reference Number"
End Sub

Private Sub M_ROU_OFF_Change()
If Val(M_ROU_OFF) >= 1 Then M_ROU_OFF = 0
End Sub

Private Sub MS_RowColChange()
PRD_NME = MS.TextMatrix(MS.Row, 10)
End Sub
Private Sub M_PUR_IDY_KeYDOWN(K As Integer, s As Integer)
If K = 13 Then
    M_DCH_DAT.SetFocus
ElseIf K = 38 Then
    M_INV_DAT.SetFocus
End If
End Sub

Private Sub M_ROU_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Route ID or Select using Arrow Keys"
End Sub

Private Sub M_SPD_AMT_Change()
If Len(Trim(M_SPD_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_SPD_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Additional Discount Amount"
End Sub

Private Sub M_SPD_PCG_Change()
If Len(Trim(M_SPD_PCG)) = 5 Then SendKeys "{home}+{end}"
M_SPD_AMT = convert(PERCENT(Val(T_CSD_AMT), Val(M_SPD_PCG)))
End Sub

Private Sub M_SPD_PCG_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Additional Discount Percentage"
End Sub

Private Sub M_CSD_AMT_Change()
If Len(Trim(M_CSD_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_CSD_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Bill Discount Amount"
End Sub

Private Sub M_CSD_PCG_Change()
If Len(Trim(M_CSD_PCG)) = 5 Then SendKeys "{home}+{end}"
M_CSD_AMT = convert(PERCENT(Val(G_GRS_AMT), Val(M_CSD_PCG)))
End Sub

Private Sub M_CSD_PCG_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Bill Discount Percentage"
End Sub

Private Sub M_CSD_PCG_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_CSD_PCG_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    M_CSD_PCG = discount(M_CSD_PCG)
    If M_CSD_PCG > 0 Then M_CSD_AMT = convert(PERCENT(Val(G_GRS_AMT), M_CSD_PCG))
    GRS_CAL
    If K = 38 Then
        M_CRT_DAY.SetFocus
    Else
        M_CSD_AMT.SetFocus
    End If
End If
End Sub

Private Sub M_CSD_AMT_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_CSD_AMT_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    M_CSD_AMT = convert(amt(M_CSD_AMT))
    GRS_CAL
    M_SPD_PCG.SetFocus
ElseIf K = 38 Then
    M_CSD_AMT = convert(amt(M_CSD_AMT))
    GRS_CAL
    M_CSD_PCG.SetFocus
End If
End Sub
Private Sub M_SPD_PCG_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_SPD_PCG_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    M_SPD_PCG = discount(M_SPD_PCG)
    If M_SPD_PCG > 0 Then M_SPD_AMT = convert(PERCENT(Val(T_CSD_AMT), Val(M_SPD_PCG)))
    GRS_CAL
    If K = 13 Or K = 40 Then
        M_SPD_AMT.SetFocus
    Else
        M_CSD_AMT.SetFocus
    End If
End If
End Sub
Private Sub M_SPD_AMT_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_SPD_AMT_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    M_SPD_AMT = convert(amt(M_SPD_AMT))
    GRS_CAL
    M_TAX_PCG.SetFocus
ElseIf K = 38 Then
    M_SPD_AMT = convert(amt(M_SPD_AMT))
    GRS_CAL
    M_SPD_PCG.SetFocus
End If
End Sub

Private Sub M_TTL_CAS_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Total Cases"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TTL_CAS_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    M_WBL_IDY.SetFocus
ElseIf K = 38 Then
    M_LRN_AMT.SetFocus
End If
End Sub

Private Sub M_TTL_CAS_KeYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_EXD_AMT_Change()
If Len(Trim(M_EXD_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_EXD_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Surcharge on Tax Amount"
End Sub

Private Sub M_EXD_AMT_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_FRT_AMT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Less Amount"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_FRT_AMT_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    M_FRT_AMT = convert(amt(M_FRT_AMT))
    GRS_CAL
    M_ROU_OFF.SetFocus
ElseIf K = 38 Then
    M_FRT_AMT = convert(amt(M_FRT_AMT))
    GRS_CAL
    M_AMT_LES.SetFocus
End If
End Sub

Private Sub M_FRT_AMT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_FRT_TYP_keydown(K As Integer, s As Integer)
If K = 13 Then
    M_FRT_AMT.SetFocus
ElseIf K = 38 And s = 1 Then
    M_ROU_IDY.SetFocus
End If
End Sub

Private Sub M_ROU_OFF_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Zero to retain paise values"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_ROU_OFF_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
If Val(M_ROU_OFF) = 0 Then
    T_NET_AMT = T_FRT_AMT
    M_ROU_OFF = "0.00"
Else
    GRS_CAL
End If
    saverecord
End If
End Sub

Private Sub M_ROU_OFF_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub List6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    For A = 1 To MS.Rows - 1
        If MS.TextMatrix(MS.Row, 13) = MS.TextMatrix(A, 13) And Trim(List6) = MS.TextMatrix(A, 21) Then
            SSTab1.Tab = 1
            list4.Visible = False
            MsgBox "Product Already Exists In Invoice By this Type !", vbExclamation, "DAS Version 4.0"
            Text3.Visible = False
            MS.SetFocus
            MS.Col = 0
            MS.Row = A
            Exit Sub
        End If
    Next
    List6.Visible = False
    If List6 <> "" Then MS.Text = List6
    List6.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub M_ROU_IDY_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    GRS_CAL
    M_FRT_TYP.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    M_EXD_AMT.SetFocus
End If
End Sub

Private Sub M_LRN_AMT_GotFocus()
SendKeys "{HOME}+{END}"
M_MDI.STS_BAR.Panels(1).Text = "Transporter Freight Amount (No Effect on Bill Value)"
End Sub

Private Sub M_LRN_AMT_keypress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub M_LRN_AMT_Keydown(K As Integer, s As Integer)
If K = 38 Then
M_LRN_DAT.SetFocus
ElseIf K = 13 Or K = 40 Then
M_TTL_CAS.SetFocus
M_LRN_AMT = convert(M_LRN_AMT)
End If
End Sub

Private Sub M_LRN_DAT_Keydown(K As Integer, s As Integer)
If K = 38 And s = 1 Then
M_LRN_IDY.SetFocus
ElseIf K = 13 Then
M_LRN_AMT.SetFocus
End If
End Sub

Private Sub M_LRN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Transporter Receipt Number"
SendKeys "{home}+{End}"
End Sub

Private Sub M_LRN_IDY_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
M_LRN_DAT.SetFocus
ElseIf KeyCode = 38 Then
SSTab1.Tab = 2
M_ROU_OFF.SetFocus
End If
End Sub

Private Sub M_LRN_IDY_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub M_PKG_CST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Add Amount"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_PKG_CST_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    M_PKG_CST = convert(amt(M_PKG_CST))
    GRS_CAL
    M_AMT_LES.SetFocus
ElseIf K = 38 Then
    M_PKG_CST = convert(amt(M_PKG_CST))
    GRS_CAL
    M_AMT_ADD.SetFocus
End If
End Sub

Private Sub M_PKG_CST_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_TXN_RMK_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Remarks (if any)"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TAX_AMT_Change()
If Len(Trim(M_TAX_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_TAX_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Tax Amount"
End Sub

Private Sub M_TAX_PCG_Change()
If Len(Trim(M_TAX_PCG)) = 5 Then SendKeys "{home}+{end}"
M_TAX_AMT = convert(PERCENT(Val(T_CSD_AMT), Val(M_TAX_PCG)))
End Sub

Private Sub M_TAX_PCG_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Tax Percentage"
End Sub

Private Sub M_TAX_PCG_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_TAX_AMT_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    M_TAX_AMT = convert(amt(M_TAX_AMT))
    GRS_CAL
    M_EXD_PCG.SetFocus
ElseIf K = 38 Then
    M_TAX_AMT = convert(amt(M_TAX_AMT))
    GRS_CAL
    M_TAX_PCG.SetFocus
End If
End Sub
Private Sub M_EXD_AMT_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    M_EXD_AMT = convert(amt(M_EXD_AMT))
    GRS_CAL
    If K = 38 Then
        M_EXD_PCG.SetFocus
    Else
        M_AMT_ADD.SetFocus
    End If
End If
End Sub

Private Sub M_TAX_AMT_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_TAX_PCG_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    M_TAX_PCG = discount(M_TAX_PCG)
    If M_TAX_PCG > 0 Then M_TAX_AMT = convert(PERCENT(Val(T_CSD_AMT), Val(M_TAX_PCG)))
    GRS_CAL
    If K = 38 Then
        M_SPD_AMT.SetFocus
    Else
        M_TAX_AMT.SetFocus
    End If
End If
End Sub

Private Sub M_CRT_DAY_Change()
If Len(M_CRT_DAY) = 3 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_CRT_DAY_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Credit Days"
End Sub

Private Sub M_CRT_DAY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Val(M_CRT_DAY) > 365 Then
        MsgBox "Credit Days Exceeds 365 Days!", vbExclamation, "DAS Version 4.0"
        M_CRT_DAY.SetFocus
    Else
        M_CSD_PCG.SetFocus
        M_DUE_DAT = Format(DateAdd("d", Val(Trim(M_CRT_DAY)), M_INV_DAT.Value), "DD/MM/YYYY")
    End If
ElseIf KeyCode = 38 Then
    If Val(M_CRT_DAY) >= 365 Then
        MsgBox "Credit Days Exceeds 365 Days!", vbExclamation, "DAS Version 4.0"
        M_CRT_DAY.SetFocus
    Else
        SSTab1.Tab = 0
        M_DUE_DAT = Format(DateAdd("d", Val(Trim(M_CRT_DAY)), M_INV_DAT.Value), "DD/MM/YYYY")
        M_CRT_IDY.SetFocus
    End If
End If
End Sub

Private Sub M_CRT_DAY_Keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub M_CRT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Account ID or Press <Enter> to View Existing Accounts"
End Sub

Private Sub M_CRT_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
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
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_DBT_IDY)) & "' AND  ACT_TYP LIKE '' and (GRP_CLF in ('OpeningStockPL','ClsStockPL'))", con, adOpenKeyset, adLockPessimistic
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
If K = 39 Or K = 41 Or K = 37 Then K = 0
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
Label23 = labdisp("idy_ops")
SSTab1.TabEnabled(4) = False
SSTab1.TabVisible(2) = True
SSTab1.TabVisible(3) = False
setval
SETYEAR M_FRM_OPS
DTPicker1.MinDate = "1 / 1 / 1999"
DTPicker1.MaxDate = "1 / 1 / 2015"
End If
End Sub

Private Sub V_INV_LST()
If CHECKTB("inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inl WHERE txn_typ ='OPS'", con, adOpenKeyset, adLockPessimistic
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
    
    
        If M_FRM_OPS.ActiveControl.name = "M_INV_IDY" Then
            V_INV_LST
        ElseIf M_FRM_OPS.ActiveControl.name = "M_DBT_IDY" Then
            v_dbt_lst
        ElseIf M_FRM_OPS.ActiveControl.name = "M_SMN_IDY" Then
            V_SMN_LST
        ElseIf M_FRM_OPS.ActiveControl.name = "M_CRT_IDY" Then
            v_crt_lst
        ElseIf M_FRM_OPS.ActiveControl.name = "Text3" Then
            V_PRD_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_OPS.ActiveControl Is TextBox Then M_FRM_OPS.ActiveControl.Text = FUNKEY(KeyCode)
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
    ElseIf SSTab1.Tab = 1 Then
        Text3.Visible = False
        M_FLX_AMT.Visible = False
        SSTab1.Tab = 2
        Frame4.Enabled = False
    Else
        If M_FRM_OPS.ActiveControl.name = "M_INV_IDY" Then
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
MS.ColWidth(0) = 980
MS.ColWidth(21) = 1400
MS.ColWidth(1) = 1050
MS.ColWidth(2) = 1000
MS.ColWidth(3) = 900
MS.ColWidth(4) = 1200
MS.ColWidth(5) = 1000
MS.ColWidth(6) = 1000
MS.ColWidth(7) = 1200
MS.ColWidth(8) = 1200
MS.ColWidth(9) = 1200
MS.ColWidth(10) = 2500
MS.ColWidth(11) = 0
MS.ColWidth(12) = 0
MS.ColWidth(13) = 0
MS.ColWidth(14) = 1500
MS.ColWidth(15) = 1500
MS.ColWidth(16) = 1500
MS.ColWidth(17) = 1500
MS.ColWidth(18) = 0
MS.ColWidth(19) = 0
MS.ColWidth(20) = 0

MS.TextMatrix(0, 0) = "Product ID"
SCl MS, 0
MS.TextMatrix(0, 1) = "            Units"
SCl MS, 1
MS.TextMatrix(0, 2) = "            Rate"
SCl MS, 2
MS.TextMatrix(0, 3) = "         UOM"
SCl MS, 3
MS.TextMatrix(0, 4) = "         Discount"
SCl MS, 4
MS.TextMatrix(0, 5) = "           PD %"
SCl MS, 5
MS.TextMatrix(0, 6) = "    Pur.Tax %"
SCl MS, 6
MS.TextMatrix(0, 7) = "           Amount"
MS.TextMatrix(0, 8) = "             MRP"
MS.ColAlignment(8) = 7
SCl MS, 8
MS.TextMatrix(0, 9) = "      Sale Price"
MS.ColAlignment(9) = 7
SCl MS, 9
MS.TextMatrix(0, 10) = "Product Name"
MS.TextMatrix(0, 11) = "uom"
MS.TextMatrix(0, 12) = "UPC"
MS.TextMatrix(0, 13) = "batch"
MS.TextMatrix(0, 14) = "Margin % w/o-VAT"
MS.ColAlignment(14) = 7
MS.TextMatrix(0, 15) = "Margin % inc.tax"
MS.ColAlignment(15) = 7
MS.ColWidth(15) = 0
MS.TextMatrix(0, 16) = "Mfg. Date"
MS.ColAlignment(16) = 1
SCl MS, 16
MS.TextMatrix(0, 17) = "BBD-Exp.Date"
MS.ColAlignment(17) = 1
MS.TextMatrix(0, 21) = "Stock Type"
MS.ColAlignment(22) = 7
MS.TextMatrix(0, 22) = "List Rate"
MS.ColAlignment(23) = 7
MS.TextMatrix(0, 23) = "List %+"
MS.ColAlignment(24) = 7
MS.TextMatrix(0, 24) = "Inc'tve"
MS.ColAlignment(25) = 7
MS.TextMatrix(0, 25) = "Inc'TTL"

SCl MS, 21
SCl MS, 17
MS.Col = 0
MS.Col = 1
UNT_TOT.Left = Label2.Left + Label2.Width + 300
MS.Col = 6
TAX_TOT.Left = ((MS.CellWidth + MS.CellLeft + MS.Left) - TAX_TOT.Width) - Picture1.Left
MS.Col = 5
ADD_TOT.Left = ((MS.CellWidth + MS.CellLeft + MS.Left) - ADD_TOT.Width) - Picture1.Left
MS.Col = 7
AMT_TOT.Left = ((MS.CellWidth + MS.CellLeft + MS.Left) - AMT_TOT.Width) - Picture1.Left
MS.Col = 4
DIS_TOT.Left = ((MS.CellWidth + MS.CellLeft + MS.Left) - DIS_TOT.Width) - Picture1.Left
PRD_NME.Left = 10
MS.Col = 0
M_INV_DAT = Date
M_LRN_DAT = Date
M_DCH_DAT = Date
waydate = Date

list5.ColumnHeaders(1).Width = List2.Width / 4
list5.ColumnHeaders(2).Width = List2.Width / 4
list5.ColumnHeaders(3).Width = List2.Width / 4

TAX_TOT = "0.00"
UNT_TOT = "0.000"
ADD_TOT = "0.00"
DIS_TOT = "0.00"
AMT_TOT = "0.00"
list1.ColumnHeaders(1).Width = 3800
List2.ColumnHeaders(1).Width = 3800
list3.ColumnHeaders(1).Width = 3800
Label23 = labdisp("idy_ops")
M_LRN_AMT = "0.00"
deliverycharges = "0.00"
M_LRN_IDY = 0
M_FRT_TYPamt = "0.00"
waybill = 0
M_NET_AMT = "0.00"
M_GRS_AMT = "0.00"
M_CRT_DAY = 0
waybill = 0
M_SPD_DED = "0.00"
M_PRD_TOT = "0.00"
tax = "0.00"
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press Ctrl+N to Generate New Opening Voucher ID by System, Press <Enter> to View Existing Vouchers"
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
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

Private Sub list1_Validate(cancel As Boolean)
'Cancel = True
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
    V_SMN_LST
ElseIf Label49.Caption = "Select Debit Account" Then
    v_dbt_lst
ElseIf Label49.Caption = "Select Voucher" Then
    V_INV_LST
ElseIf Label49.Caption = "Select Credit Account" Then
    v_crt_lst
ElseIf Label49.Caption = "Select Product" Then
    V_PRD_LST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If Label49.Caption = "Select Salesman" Then
    V_SMN_LST
ElseIf Label49.Caption = "Select Debit Account" Then
    v_dbt_lst
ElseIf Label49.Caption = "Select Voucher" Then
    V_INV_LST
ElseIf Label49.Caption = "Select Credit Account" Then
    v_crt_lst
ElseIf Label49.Caption = "Select Product" Then
    If Check2.Value = 1 Then
    con.Execute "insert into prt values('p_prdnum')"
    Else
    con.Execute "delete from prt where prt_nme like 'p_prdnum'"
    End If
    V_PRD_LST
End If
'End If
End Sub

Private Sub M_WBL_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Way Bill Date, Press F4 to get View"
End Sub

Private Sub MS_EnterCell()
If MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product ID View"
ElseIf MS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Product Quantity"
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Purchase Rate"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measure (Nos, Kgs, Ltrs, Mtrs, MT)"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Discount Amount"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Purchase Discount %"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Tax Percent"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Value (System Calculated Value)"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Maximum Retail Price (MRP)"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Unit Sale Price (@Billing to Retailer)"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product (SKU) Description"
ElseIf MS.Col = 14 Then
    M_MDI.STS_BAR.Panels(1).Text = "Margin Excluding TAX"
ElseIf MS.Col = 15 Then
    M_MDI.STS_BAR.Panels(1).Text = "Margin Including TAX"
ElseIf MS.Col = 16 Then
    M_MDI.STS_BAR.Panels(1).Text = "Manufacturing Date"
ElseIf MS.Col = 17 Then
    M_MDI.STS_BAR.Panels(1).Text = "BBD-Exp Date"
End If
End Sub

Private Sub MS_Scroll()
If M_FLX_AMT.Visible = True Then
    M_FLX_AMT = ""
    M_FLX_AMT.Visible = False
    MS.SetFocus
ElseIf List6.Visible = True Then
    List6.Visible = False
    MS.SetFocus
ElseIf Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    MS.SetFocus
ElseIf DTPicker1.Visible = True Then
    DTPicker1.Visible = False
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
    If M_FRM_OPS.ActiveControl.name = "M_INV_IDY" Then
        Label49 = "Select Voucher"
        V_INV_LST
    ElseIf M_FRM_OPS.ActiveControl.name = "M_DBT_IDY" Then
        Label49 = "Select Debit Account"
        
        v_dbt_lst
    ElseIf M_FRM_OPS.ActiveControl.name = "M_SMN_IDY" Then
        Label49 = "Select Salesman"
        V_SMN_LST
    ElseIf M_FRM_OPS.ActiveControl.name = "M_CRT_IDY" Then
        Label49 = "Select Credit Account"
        v_crt_lst
    ElseIf M_FRM_OPS.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        V_PRD_LST
    End If
ElseIf Button.KEY = "DAY" Then
'''''    If Button.Value = tbrPressed Then
''''''        Text1.Visible = True
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
If check("prt", "prt_nme", "m_frm_ops") = False Then
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
    clrctr M_FRM_OPS
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
On Error GoTo LBL
If CHECKTB("inl.dbf,inv.dbf,prd.dbf,brd.dbf,pgr.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Len(Trim(M_INV_IDY)) <> 10 Then
REC_SET.Open "select * from inl where inv_idy like 'OPS" & Tincrement(Trim(M_INV_IDY)) & "' and txn_typ like 'OPS'", con, adOpenKeyset
Else
REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'OPS'", con, adOpenKeyset
End If
If REC_SET.RecordCount > 0 Then
    If Len(Trim(M_INV_IDY)) <> 10 Then
    M_INV_IDY = "OPS" & Tincrement(Trim(M_INV_IDY))
    End If
    M_INV_IDY.Enabled = False
    MOD_REC = True
    M_INV_DAT = REC_SET("inv_dat")
    M_DBT_IDY = Trim(REC_SET("cst_idy") & "")
    dlr_idy = M_DBT_IDY
    If Len(Trim(M_DBT_IDY)) = 0 Then Exit Sub
    M_PUR_IDY = Trim(REC_SET("ORD_IDY") & "")
    M_CRT_IDY = Trim(REC_SET("act_idy") & "")
    M_DCH_IDY = Trim(REC_SET("dch_idy") & "")
    M_DCH_DAT = REC_SET("dch_dat")
    M_EXD_AMT = convert(amt(REC_SET("EXD_AMT") & ""))
    M_CRT_DAY = Val(REC_SET("crt_day"))
    M_ROU_OFF = Val(REC_SET("rou_off") & "")
    If Val(REC_SET("rou_off") & "") = 0 Then M_ROU_OFF = "0.00"
    M_SPD_PCG = discount(Val(REC_SET("spd_pcg") & ""))
    M_SPD_AMT = convert(amt(Val(REC_SET("spd_amt") & "")))
    M_CSD_PCG = discount(Val(REC_SET("csd_pcg") & ""))
    M_CSD_AMT = convert(amt(Val(REC_SET("csd_amt") & "")))
    M_TAX_PCG = discount(Val(REC_SET("tax_pcg") & ""))
    M_TAX_AMT = convert(amt(Val(REC_SET("tax_amt") & "")))
    M_WBL_IDY = Trim(REC_SET("WBL_IDY") & "")
    M_FRT_TYP = Trim(REC_SET("FRT_typ") & "")
    M_FRT_AMT = convert(amt(Val(REC_SET("frt_amt") & "")))
    M_WBL_IDY = Trim(REC_SET("wbl_idy") & "")
    M_WBL_DAT = REC_SET("wbl_dat")
    M_PKG_CST = convert(amt(REC_SET("PKG_CSt") & ""))
    M_TTL_CAS = Trim(REC_SET("ttl_cas") & "")
    M_TXN_RMK = Trim(REC_SET("txn_rmk") & "")
    M_TXN_TPT = Trim(REC_SET("txn_tpt") & "")
    M_LRN_IDY = Trim(REC_SET("lrn_idy") & "")
    M_LRN_AMT = convert(amt(REC_SET("lrn_amt") & ""))
    M_LRN_DAT = Trim(REC_SET("lrn_dat") & "")
    M_ROU_IDY = Trim(REC_SET("rou_idy") & "")
    M_AMT_ADD = Trim(REC_SET!txn_rm1) & ""
    M_AMT_LES = Trim(REC_SET!TXN_RM2) & ""
    M_EXD_PCG = discount(Val(REC_SET!EXD_PCG & ""))
    M_EXD_AMT = convert(amt(Val(REC_SET!EXD_AMT & "")))
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Trim(M_CRT_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then M_ACT_NME = REC_SET("act_nme")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from ACT where ACT_idy like '" & Trim(M_DBT_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_SUP_NME = Trim(REC_SET("ACT_nme") & "")
        M_SUP_AD1 = Trim(REC_SET("ACT_ad1") & "")
        M_SUP_AD2 = Trim(REC_SET("ACT_ad2") & "")
        M_SUP_AD3 = Trim(REC_SET("ACT_ad3") & "")
        M_SUP_AD4 = Trim(REC_SET("ACT_ad4") & "")
        Label17 = amt(Trim(REC_SET("dbt_mcl")) & "")
        Label18 = Trim(REC_SET("dbt_mcp") & "")
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ = 'OPS'", con, adOpenKeyset, adLockPessimistic
    MS.Rows = REC_SET.RecordCount + 1
    For A = 1 To MS.Rows - 1
        MS.TextMatrix(A, 0) = Trim(REC_SET("prd_idy"))
        MS.TextMatrix(A, 1) = QTY(Val(REC_SET("prd_qty")))
        MS.TextMatrix(A, 2) = convert(amt(REC_SET("prd_rat") & ""))
        MS.TextMatrix(A, 4) = convert(REC_SET("add_AMT") & "")
        MS.TextMatrix(A, 5) = discount(REC_SET("trd_dis") & "")
        MS.TextMatrix(A, 6) = discount(REC_SET("trd_tax") & "")
        MS.TextMatrix(A, 13) = Trim(REC_SET!BAT_IDY)
        MS.TextMatrix(A, 21) = Trim(REC_SET!stk_typ)
        MS.TextMatrix(A, 22) = convert(REC_SET("LST_RAT") & "")
        MS.TextMatrix(A, 23) = convert(REC_SET("LST_DIS") & "")
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT * FROM BRD WHERE BAT_IDY LIKE '" & Trim(REC_SET!BAT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If REC_SET2.RecordCount > 0 Then
            MS.TextMatrix(A, 8) = convert(Val(REC_SET2!PRD_MRP))
            MS.TextMatrix(A, 9) = convert(Val(REC_SET2!PRD_SDR))
            MS.TextMatrix(A, 16) = Format(REC_SET2!bRD_dom, "dd/MM/yyyy")
            MS.TextMatrix(A, 17) = Format(REC_SET2!brd_doe, "dd/MM/yyyy")
            MS.TextMatrix(A, 24) = convert(Val(REC_SET2!P_INC_DIS))
        End If
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT prd_nme,CON_FAC,SUB_UNT,pgr.grp_idy,pgr.grp_cat,pgr.grp_cls FROM PRD,pgr WHERE prd.grp_idy like pgr.grp_idy and PRD_IDY LIKE '" & REC_SET("PRD_IDY") & "'", con, adOpenStatic
        If REC_SET2.RecordCount > 0 Then
            MS.TextMatrix(A, 11) = Trim(REC_SET2(1))
            MS.TextMatrix(A, 10) = Trim(REC_SET2(0))
            MS.TextMatrix(A, 12) = Val(REC_SET2(2))
            MS.TextMatrix(A, 18) = Trim(REC_SET2!GRP_IDY)
            MS.TextMatrix(A, 19) = Trim(REC_SET2!GRP_CAT)
            MS.TextMatrix(A, 20) = Trim(REC_SET2!GRP_CLS)
        End If
        REC_SET.MoveNext
        If Val(MS.TextMatrix(A, 9)) <> 0 And Val(MS.TextMatrix(A, 2)) <> 0 Then MS.TextMatrix(A, 14) = discount(Round((Val(MS.TextMatrix(A, 9)) / Val(MS.TextMatrix(A, 2)) * 100), 2) - 100)
        If Val(MS.TextMatrix(A, 6)) <> 0 And Val(MS.TextMatrix(A, 9)) <> 0 And Val(MS.TextMatrix(A, 2)) <> 0 Then
            MS.TextMatrix(A, 15) = discount(Round((Val(MS.TextMatrix(A, 2)) * Val(MS.TextMatrix(A, 6)) / 100), 2))
            MS.TextMatrix(A, 15) = discount(Round(((Val(MS.TextMatrix(A, 9)) / (Val(MS.TextMatrix(A, 15)) + Val(MS.TextMatrix(A, 2)))) * 100) - 100, 2))
        ElseIf Val(MS.TextMatrix(A, 9)) <> 0 And Val(MS.TextMatrix(A, 2)) <> 0 Then
            MS.TextMatrix(A, 15) = discount(Round(((Val(MS.TextMatrix(A, 9)) / (Val(MS.TextMatrix(A, 15)) + Val(MS.TextMatrix(A, 2)))) * 100) - 100, 2))
        End If
    Next
    P_DIS_CAL
    GRS_CAL
    M_GRS_AMT = M_PRD_TOT
    SSTab1.Tab = 0
    M_INV_DAT.SetFocus
Else
    V_INV_LST
End If
    If MS.Rows = 1 Then MS.Rows = 2
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



Private Sub list2_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If List2.ListItems.count > 0 Then
        M_CRT_IDY = List2.SelectedItem.ListSubItems.Item(1)
        M_ACT_NME = List2.SelectedItem
        SSTab1.Tab = 1
        MS.SetFocus
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
        P_DIS_PRD ls.SubItems(4), ls.SubItems(6)
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
ElseIf SSTab1.Tab = 2 Then
  '  M_PRD_TOT.Visible = True
    M_ROU_OFF.SetFocus
End If
If list1.Visible = True Then list1.Visible = False
If List2.Visible = True Then List2.Visible = False
If list3.Visible = True Then list3.Visible = False
If list4.Visible = True Then list4.Visible = False
If list5.Visible = True Then list5.Visible = False

Text3.Visible = False
M_FLX_AMT.Visible = False
If SSTab1.Tab = 3 Then
    M_LRN_IDY.SetFocus
End If
If SSTab1.Tab = 4 Then
'Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTab1.Tab = 1 Then MS.SetFocus
    If SSTab1.Tab = 2 Then M_CRT_DAY.SetFocus
    If SSTab1.Tab = 0 Then M_INV_DAT.SetFocus
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub M_CRT_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If

    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_CRT_IDY)) & "' AND act_idy not like '" & Sincrement(M_DBT_IDY) & "%' and ACT_TYP like '' and (GRP_CLF in ('OpeningStockPL','ClsStockPL'))", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_CRT_IDY = Sincrement(M_CRT_IDY)
        If K = 38 Then
        M_DBT_IDY.SetFocus
        Else
        M_ACT_NME = Trim(REC_SET("act_nme"))
        SSTab1.Tab = 1
        MS.SetFocus
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
If Text3.Visible = True Then Text3.Visible = False
If M_FLX_AMT.Visible = True Then M_FLX_AMT.Visible = False
If List6.Visible = True Then List6.Visible = False
End Sub

Private Sub MS_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K > 47 And K < 58 Or K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If MS.Col = 0 And Trim(MS.TextMatrix(MS.Row, 0)) = "" Then
        Text3 = ""
        V_PRD_LST
        Exit Sub
    ElseIf MS.Cols = MS.Col + 1 Or MS.Col = 1 Then
        If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
            MS.AddItem ""
            MS.Col = 0
            MS.Row = MS.Rows - 1
            Exit Sub
        End If
    Else
        MS.Col = MS.Col + 1
    End If
End If
If MS.Row <> 0 Then
    If MS.Col = 0 Then                                      'PRODUCT
        If K <> 27 Then
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
    ElseIf MS.Col = 1 Then                                              'QUANTITY
        If MS.TextMatrix(MS.Row, 0) <> "" Then
           If K > 47 And K < 58 Or K = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.MaxLength = 9
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.Visible = True
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(K)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            ElseIf K = 13 Then
                If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                    MS.AddItem ""
                    MS.Col = 0
                    MS.Row = MS.Rows - 1
                End If
            Else
                K = 0
            End If
        End If
    ElseIf MS.Col = 2 Then                                              'PURCHASE RATE
        If MS.TextMatrix(MS.Row, 0) <> "" Then
           If K > 47 And K < 58 Or K = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.MaxLength = 9
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.Visible = True
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(K)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            Else
                K = 0
            End If
        End If
    ElseIf MS.Col = 5 Or MS.Col = 6 Then                                        'AD%,TAX
        If MS.TextMatrix(MS.Row, 0) <> "" Then
           If K > 47 And K < 58 Or K = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.MaxLength = 9
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.Visible = True
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(K)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            Else
                K = 0
            End If
        End If
    ElseIf MS.Col = 4 Or MS.Col = 3 Then                        'DISCOUNT AMOUNT,UOM
        If Len(Trim(MS.TextMatrix(MS.Row, 0))) = 10 Then
         If K > 47 And K < 58 Or K = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.Visible = True
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(K)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            End If
        End If
    ElseIf MS.Col = 8 Or MS.Col = 9 Then                    'MRP, UNIT RATE
        If Len(Trim(MS.TextMatrix(MS.Row, 0))) = 10 Then
         If K > 47 And K < 58 Or K = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.MaxLength = 9
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.Visible = True
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(K)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            End If
        End If
    ElseIf MS.Col = 16 Or MS.Col = 17 Then
        If K <> 13 And K <> 27 Then
            If CHECKTB("dir.dbf") = 1 Then
                MsgBox UCase(TNAME) & " Missing in " & Me.Caption
                CANCELRECORD
                Exit Sub
            End If
            If querycheck("DIR", "DIR_MSG", "E", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX ='EXP'") = True And Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                DTPicker1 = MS.Text
                DTPicker1.Visible = True
                DTPicker1.Top = MS.Top + MS.CellTop
                DTPicker1.Left = MS.CellLeft + MS.Left
                DTPicker1.Width = MS.CellWidth
                DTPicker1.Height = MS.CellHeight
                DTPicker1.Visible = True
                DTPicker1.SetFocus
            End If
        End If
    ElseIf MS.Col = 21 Then
        If K <> 27 Then
            If K = vbKeySpace Then
            Else
                If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                    List6.Visible = True
                    List6.Top = MS.Top + MS.CellTop
                    List6.Left = MS.CellLeft + MS.Left
                    List6.Width = MS.CellWidth
                    List6.Height = MS.CellHeight
                    List6.Visible = True
                    List6.SetFocus
                    If K <> 13 Then List6 = Chr(K)
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Text3_KeyDown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If

    If MS.Col = 0 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(Text3) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY", con, adOpenKeyset
        If REC_SET.RecordCount = 1 Then
            P_DIS_PRD Sincrement(Text3), Trim(REC_SET!BAT_IDY)
        ElseIf REC_SET.RecordCount > 1 Then
            Text3 = Sincrement(Text3)
            V_PRD_IDY_LST
        Else
            V_PRD_LST
       End If
       End If
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then

ElseIf K = 13 Then
    If CHECKTB("prd.dbf,brd.dbf,inv.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If

    If MS.Col = 1 Then                                                  'QUANTIY
        M_FLX_AMT = QTY(M_FLX_AMT)
        If REC_SET.State = 1 Then REC_SET.Close
        If REC_SET1.State = 1 Then REC_SET1.Close
        If REC_SET2.State = 1 Then REC_SET2.Close
        If REC_SET3.State = 1 Then REC_SET3.Close
        REC_SET.Open "select * from PRD,BRD where BRD.PRD_IDY LIKE PRD.PRD_IDY AND BAT_IDY like '" & Trim(MS.TextMatrix(MS.Row, 13)) & "'", con, adOpenKeyset, adLockPessimistic
        If Trim(M_INV_IDY) <> "" Then
            REC_SET2.Open "select PRD_QTY from inv where BAT_idy like '" & Trim(MS.TextMatrix(MS.Row, 13)) & "' and inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'OPS'", con, adOpenKeyset, adLockPessimistic
            If REC_SET2.RecordCount > 0 Then
                    INVTOT = QTY(REC_SET2(0))
            End If
        End If
        b = InStr(M_FLX_AMT, ".")
        If b <> 0 Then M = Mid(M_FLX_AMT, b + 1, Len(M_FLX_AMT) - b + 1)
        If Len(M) = 1 Then
            M_FLX_AMT = M_FLX_AMT & "00"
        ElseIf Len(M) = 2 Then
            M_FLX_AMT = M_FLX_AMT & "0"
        ElseIf Len(M) = 0 Then
            M_FLX_AMT = M_FLX_AMT & ".000"
        End If
        If b > 0 And Val(REC_SET("sub_unt")) > 0 Then
            If InStr(b + 1, M_FLX_AMT, ".") > 0 Then
                MS.TextMatrix(MS.Row, 3) = "0.000"
                Exit Sub
            End If
            M = Mid(M_FLX_AMT, b + 1, 3)
            If M >= Val(REC_SET("sub_unt")) Then
                MsgBox "Product Units Per Case (Sub Units) Are " & REC_SET("sub_unt"), vbExclamation, "DAS Version 4.0"
                If M_FLX_AMT.Visible = True Then
                    M_FLX_AMT.SetFocus
                Else
                    MS.SetFocus
                End If
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                M_QTY = M_FLX_AMT
            End If
        Else
            M_QTY = M_FLX_AMT
        End If
        If REC_SET.RecordCount > 0 Then
            REMIND = ADDQTY(INVTOT, MS.TextMatrix(MS.Row, 0))
            MS.Text = M_QTY
            M_FLX_AMT.Visible = False
            If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                MS.AddItem ""
                MS.Col = 0
                MS.Row = MS.Rows - 1
            End If
        End If
        MS.SetFocus
        P_DIS_CAL
    ElseIf MS.Col = 2 Then                                          'PURCHASE RATE
        If Val(MS.TextMatrix(MS.Row, 2)) <> Val(M_FLX_AMT) Then createbatch
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 4 Then                                          'DISCOUNT AMOUNT
        MS.TextMatrix(MS.Row, 4) = convert(amt(Val(M_FLX_AMT)))
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 5 Then                                          'AD%
        MS.TextMatrix(MS.Row, 5) = discount(Val(M_FLX_AMT))
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 6 Then                                          'TAX%
        MS.TextMatrix(MS.Row, 6) = discount(Val(M_FLX_AMT))
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 3 Then                                          'UOM
        M_FLX_AMT = UNTCAL(M_FLX_AMT, MS.TextMatrix(MS.Row, 0))
        MS.Col = 1
        M_FLX_AMT_KeyPress 13
    ElseIf MS.Col = 8 Or MS.Col = 9 Then                            'MRP,UNITRATE
        If Val(MS.TextMatrix(MS.Row, MS.Col)) <> Val(M_FLX_AMT) Then createbatch
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    Else
        K = 0
    End If
Else
    K = 0
End If
End Sub

Public Sub P_DIS_PRD(PRDID As String, BATID As String)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
MSROW = MS.Row
Dim REC_SET As New ADODB.Recordset
For A = 1 To MS.Rows - 1
    If Trim(BATID) = MS.TextMatrix(A, 13) And "Salable" = MS.TextMatrix(A, 21) Then
        SSTab1.Tab = 1
        list4.Visible = False
        MsgBox "Product Already Exists In Voucher !", vbExclamation, "DAS Version 4.0"
        Text3.Visible = False
        MS.SetFocus
        MS.Col = 0
        MS.Row = A
        Exit Sub
    End If
Next
REC_SET.Open "select * from PRD,BRD where PRD.PRD_IDY LIKE BRD.PRD_IDY AND BRD.BAT_idy like '" & Trim(BATID) & "'", con, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    list4.Visible = False
    MS.TextMatrix(MS.Row, 0) = REC_SET("PRD_IDY")
    MS.TextMatrix(MS.Row, 1) = "0.000"
    MS.TextMatrix(MS.Row, 3) = "0"
    MS.TextMatrix(MS.Row, 2) = convert(REC_SET("PRD_PDR"))
    MS.TextMatrix(MS.Row, 3) = "0.000"
    MS.TextMatrix(MS.Row, 4) = "0.00"
    MS.TextMatrix(MS.Row, 5) = discount(Val(REC_SET("pur_dis") & ""))
    MS.TextMatrix(MS.Row, 6) = discount(Val(REC_SET("pur_tax") & ""))
    MS.TextMatrix(MS.Row, 7) = "0.00"
    MS.TextMatrix(MS.Row, 8) = convert(REC_SET("PRD_MRP"))
    MS.TextMatrix(MS.Row, 9) = convert(REC_SET("PRD_SDR"))
    MS.TextMatrix(MS.Row, 10) = Trim(REC_SET("PRD_NME"))
    MS.TextMatrix(MS.Row, 11) = Trim(REC_SET("CON_FAC"))
    MS.TextMatrix(MS.Row, 12) = Trim(REC_SET("SUB_UNT"))
    MS.TextMatrix(MS.Row, 13) = Trim(REC_SET("BAT_IDY"))
    MS.TextMatrix(MS.Row, 21) = "Salable"
    MS.TextMatrix(MS.Row, 16) = Format(REC_SET("Brd_dom"), "dd/MM/yyyy")
    MS.TextMatrix(MS.Row, 17) = Format(REC_SET("Brd_doe"), "dd/MM/yyyy")
    MS.TextMatrix(MS.Row, 24) = convert(REC_SET("P_INC_DIS"))
    If IsNull(REC_SET("LST_RAT")) = False Then
        MS.TextMatrix(MS.Row, 22) = Val(REC_SET("LST_RAT"))
    Else
        MS.TextMatrix(MS.Row, 22) = "0.00"
    End If
    If IsNull(REC_SET("LST_DIS")) = False Then
        MS.TextMatrix(MS.Row, 23) = Val(REC_SET("LST_DIS"))
    Else
        MS.TextMatrix(MS.Row, 23) = "0.00"
    End If
    If Val(MS.TextMatrix(MS.Row, 9)) > 0 And Val(MS.TextMatrix(MS.Row, 2)) > 0 Then MS.TextMatrix(MS.Row, 14) = discount(Round((Val(MS.TextMatrix(MS.Row, 9)) / Val(MS.TextMatrix(MS.Row, 2)) * 100), 2) - 100)
    If Val(MS.TextMatrix(MS.Row, 6)) <> 0 And Val(MS.TextMatrix(MS.Row, 9)) > 0 And Val(MS.TextMatrix(MS.Row, 2)) > 0 Then
        MS.TextMatrix(MS.Row, 15) = discount(Round((Val(MS.TextMatrix(MS.Row, 2)) * Val(MS.TextMatrix(MS.Row, 6)) / 100), 2))
        MS.TextMatrix(MS.Row, 15) = discount(Round(((Val(MS.TextMatrix(MS.Row, 9)) / (Val(MS.TextMatrix(MS.Row, 15)) + Val(MS.TextMatrix(MS.Row, 2)))) * 100) - 100, 2))
    ElseIf Val(MS.TextMatrix(MS.Row, 9)) <> 0 And Val(MS.TextMatrix(MS.Row, 2)) <> 0 Then
        MS.TextMatrix(MS.Row, 15) = discount(Round(((Val(MS.TextMatrix(MS.Row, 9)) / (Val(MS.TextMatrix(MS.Row, 15)) + Val(MS.TextMatrix(MS.Row, 2)))) * 100) - 100, 2))
    End If
    MS.TextMatrix(MS.Row, 18) = Trim(REC_SET!GRP_IDY)
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from pgr where grp_idy like '" & Trim(MS.TextMatrix(MS.Row, 18)) & "'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        MS.TextMatrix(MS.Row, 19) = Trim(REC_SET!GRP_CAT)
        MS.TextMatrix(MS.Row, 20) = Trim(REC_SET!GRP_CLS)
    End If
    Text3.Visible = False
    SSTab1.Tab = 1
    MS.Row = MSROW
    MS.Col = 1
    MS.SetFocus
End If
End Sub

Public Sub P_DIS_CAL()
UNT_TOT = 0
TAX_TOT = 0
AMT_TOT = 0
DIS_TOT = 0
ADD_TOT = 0
LST_TOT = 0
M_INC_AMT = 0
For A = 1 To MS.Rows - 1
    'Customer Discount calculating on units and rate
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
        MS.TextMatrix(A, 7) = RATCAL(Val(MS.TextMatrix(A, 1) & ""), Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 12)))
        M_INC_AMT = RATCAL(Val(MS.TextMatrix(A, 1) & ""), Val(MS.TextMatrix(A, 24) & ""), Val(MS.TextMatrix(A, 12)))
        PRD_TOT = Val(MS.TextMatrix(A, 7)) - (Val(MS.TextMatrix(A, 4) + M_INC_AMT)) 'Incentive
        MS.TextMatrix(A, 25) = convert(Val(M_INC_AMT))
        ADDTOT = convert(Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 5)) / 100), 2))
        PRD_TOT = Val(PRD_TOT) - Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 5)) / 100), 2)
        TAXTOT = convert(Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 6)) / 100), 2))
        PRD_TOT = convert(amt(Val(PRD_TOT) + Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 6)) / 100), 2)))
        MS.TextMatrix(A, 3) = QTY(WGTCAL(MS.TextMatrix(A, 1), Val(MS.TextMatrix(A, 11)), Val(MS.TextMatrix(A, 12))))
        LST_TOT = MTS(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 12)))
        LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 12))))
        M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
        'M_CPN_AMT = Val(MS.TextMatrix(A, 1)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        'LST_TOT = Val(MS.TextMatrix(A, 1)) * Val(MS.TextMatrix(A, 22))
        DIS_TOT = convert(Val(DIS_TOT) + Val(MS.TextMatrix(A, 4)))
        TAX_TOT = convert(Val(TAX_TOT) + Val(TAXTOT) + M_CPN_AMT)
        UNT_TOT = QTY(Val(UNT_TOT) + Val(MS.TextMatrix(A, 1)))
        ADD_TOT = convert(Val(ADD_TOT) + Val(ADDTOT))
        AMT_TOT = convert(amt(Val(AMT_TOT) + MS.TextMatrix(A, 7)) + LST_TOT)
        If Val(MS.TextMatrix(MS.Row, 6)) <> 0 And Val(MS.TextMatrix(MS.Row, 9)) <> 0 And Val(MS.TextMatrix(MS.Row, 2)) <> 0 Then
            MS.TextMatrix(MS.Row, 15) = discount(Round((Val(MS.TextMatrix(MS.Row, 2)) * Val(MS.TextMatrix(MS.Row, 6)) / 100), 2))
            MS.TextMatrix(MS.Row, 15) = discount(Round(((Val(MS.TextMatrix(MS.Row, 9)) / (Val(MS.TextMatrix(MS.Row, 15)) + Val(MS.TextMatrix(MS.Row, 2)))) * 100) - 100, 2))
        ElseIf Val(MS.TextMatrix(MS.Row, 9)) <> 0 And Val(MS.TextMatrix(MS.Row, 2)) <> 0 Then
            MS.TextMatrix(MS.Row, 15) = discount(Round(((Val(MS.TextMatrix(MS.Row, 9)) / (Val(MS.TextMatrix(MS.Row, 15)) + Val(MS.TextMatrix(MS.Row, 2)))) * 100) - 100, 2))
        End If
    End If
Next
M_GRS_AMT = convert(amt(AMT_TOT))
M_PRD_TOT = convert(amt(AMT_TOT))
M_NET_AMT = convert(amt(Val(M_GRS_AMT)))
GRS_CAL
End Sub
Public Sub V_PRD_IDY_LST()
If CHECKTB("prt.dbf,prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("prt", "prt_nme", "p_prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD.PRD_IDY LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') ORDER BY BRW_SEQ", con, adOpenKeyset
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(REC_SET!PRD_MRP)
        ls.ListSubItems.Add , , convert(REC_SET!PRD_PDR)
        If Val(REC_SET("PRD_QTY") & "") <> 0 Then
            ls.ListSubItems.Add , , Trim(QTY(REC_SET("PRD_QTY") & ""))
        Else
            ls.ListSubItems.Add , , "0.000"
        End If
        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_IDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_NME") & "")
        ls.ListSubItems.Add , , convert(REC_SET("prd_sdr") & "")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_dom"), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_doe"), "dd/MM/yyyy")
        REC_SET.MoveNext
    Next
    Label49 = "Select Product"
    SSTab1.Tab = 4
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product Not Found, Define Product !", vbExclamation, "DAS Version 4.0"
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

Public Sub V_PRD_LST()
If CHECKTB("prt.dbf,prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

If check("prt", "prt_nme", "p_prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') ORDER BY BRW_SEQ", con, adOpenKeyset
Else
    REC_SET.Open "SELECT * FROM brd,prd WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD_PRN LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') ORDER BY BRW_SEQ", con, adOpenKeyset
End If
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(REC_SET!PRD_MRP)
        ls.ListSubItems.Add , , convert(REC_SET!PRD_PDR)
        If Val(REC_SET("PRD_QTY") & "") <> 0 Then
            ls.ListSubItems.Add , , Trim(QTY(REC_SET("PRD_QTY") & ""))
        Else
            ls.ListSubItems.Add , , "0.000"
        End If
        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_IDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_NME") & "")
        ls.ListSubItems.Add , , convert(REC_SET("prd_sdr") & "")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_dom"), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_doe"), "dd/MM/yyyy")
        REC_SET.MoveNext
    Next
    Label49 = "Select Product"
    SSTab1.Tab = 4
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product Not Found, Define Product !", vbExclamation, "DAS Version 4.0"
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
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

If check("PRT", "PRT_NME", "M_FRM_OPS") = True Then
M_FRM_PRV.Check3.Value = 1
Else
M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "OPS"
M_FRM_PRV.Show 1
End Sub

Public Sub setval()
clrctr M_FRM_OPS
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

Public Sub ROU_OFF()
If CHECKTB("dir.dbf,inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

If querycheck("DIR", "DIR_MSG", "R", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ROU'") = True Then
Else
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select rou_off from inl where inv_idy like '" & Trim(M_INV_IDY) & "' AND TXN_TYP LIKE 'OPS'", con, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then If (Val(REC_SET(0) & "") = 0 Or Val(M_ROU_OFF) = 0) And MOD_REC = True Then Exit Sub
dot = 0
rof = 0
dot = InStr(Trim(T_FRT_AMT), ".")
If dot > 0 Then
    rof = Mid(Trim(T_FRT_AMT), dot + 1, 2)
    If Len(rof) = 1 Then rof = rof & "0"
    If rof > 0 And rof < 50 Then
        M_ROU_OFF = "-0." & rof
        M_NET_AMT = Mid(Val(T_FRT_AMT), 1, dot) & "00"
    ElseIf rof > 0 And rof >= 50 Then
        rof = 100 - Val(rof)
        If Len(rof) = 1 Then
            M_ROU_OFF = "+0.0" & rof
        ElseIf Len(rof) = 2 Then
            M_ROU_OFF = "+0." & rof
        End If
        T_NET_AMT = Mid(Val(T_FRT_AMT), 1, dot - 1)
        T_NET_AMT = Val(M_NET_AMT) + 1 & ".00"
    End If
Else
    M_ROU_OFF = "0.00"
    T_NET_AMT = convert(amt(Val(M_NET_AMT)))
End If
End Sub

Private Sub M_TXN_TPT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Transporter Name"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TXN_TPT_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
M_TXN_RMK.SetFocus
ElseIf KeyCode = 38 Then
M_WBL_DAT.SetFocus
End If
End Sub

Private Sub M_TXN_TPT_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub
Private Sub M_TXN_RMK_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
saverecord
ElseIf KeyCode = 38 Then
M_TXN_TPT.SetFocus
End If
End Sub

Private Sub M_TXN_RMK_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub
Private Sub M_WBL_DAT_KeyDown(K As Integer, s As Integer)
If K = 38 And s = 1 Then
    M_WBL_IDY.SetFocus
ElseIf K = 13 Then
    M_TXN_TPT.SetFocus
End If
End Sub

Private Sub M_WBL_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Way Bill Number"
SendKeys "{home}+{End}"
End Sub

Private Sub M_WBL_IDY_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
M_WBL_DAT.SetFocus
ElseIf KeyCode = 38 Then
M_TTL_CAS.SetFocus
End If
End Sub

Public Sub createbatch()
If MsgBox("Create New Batch Record?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        If CHECKTB("dir.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            CANCELRECORD
            Exit Sub
        End If
    Me.Tag = "Batch"
    If querycheck("DIR", "DIR_MSG", "E", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX ='EXP'") = True And Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then M_FRM_NBC.Width = 7530
    M_FRM_NBC.Tag = "OPS"
    M_FRM_NBC.Show 1
Else
    MS.SetFocus
End If
End Sub
