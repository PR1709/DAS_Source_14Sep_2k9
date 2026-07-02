VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SOR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale - Orders "
   ClientHeight    =   5625
   ClientLeft      =   1455
   ClientTop       =   330
   ClientWidth     =   8820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "M_FRM_SOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   2
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
      TabCaption(0)   =   "&Order Details"
      TabPicture(0)   =   "M_FRM_SOR.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "M_CST_NME"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "M_CST_AD1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "M_CST_AD2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "M_CST_AD3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "M_CST_AD4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "M_SMN_NME"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "M_ACT_NME"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Shape1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Shape2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "GST"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CST"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label0"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "M_INV_DAT"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "M_INV_IDY"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "M_ACT_IDY"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "M_CST_IDY"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "M_SMN_IDY"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "M_VAN_IDY"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "pcs_chk"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "JUST_MADE"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "&Products"
      TabPicture(1)   =   "M_FRM_SOR.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(2)=   "List6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "M_FLX_AMT"
      Tab(1).Control(5)=   "MS"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Taxes and  Levies"
      TabPicture(2)   =   "M_FRM_SOR.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "M_ROU_IDY"
      Tab(2).Control(1)=   "Command2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(3)=   "Command1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "M_DUE_DAT"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "M_CRT_DAY"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame4"
      Tab(2).Control(7)=   "LabelAAS"
      Tab(2).Control(8)=   "Label28"
      Tab(2).Control(9)=   "Label4"
      Tab(2).Control(10)=   "Label9"
      Tab(2).Control(11)=   "Label8"
      Tab(2).Control(12)=   "Label7"
      Tab(2).Control(13)=   "Label5"
      Tab(2).Control(14)=   "Label42"
      Tab(2).Control(15)=   "Label24"
      Tab(2).Control(16)=   "Label25"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "&Shipment Details"
      TabPicture(3)   =   "M_FRM_SOR.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label54"
      Tab(3).Control(1)=   "M_TXN_RMK"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame5"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "View"
      TabPicture(4)   =   "M_FRM_SOR.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LIST8"
      Tab(4).Control(1)=   "Frame1"
      Tab(4).Control(2)=   "list5"
      Tab(4).Control(3)=   "list2"
      Tab(4).Control(4)=   "list1"
      Tab(4).Control(5)=   "list3"
      Tab(4).Control(6)=   "list4"
      Tab(4).Control(7)=   "V_LST_CAP"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Fr&ee and Takebacks"
      TabPicture(5)   =   "M_FRM_SOR.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text5"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Text4"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Picture2"
      Tab(5).Control(3)=   "List7"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "MSF"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Sc&hemes"
      TabPicture(6)   =   "M_FRM_SOR.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "SC"
      Tab(6).Control(1)=   "TXT_SCH"
      Tab(6).Control(2)=   "Frame6"
      Tab(6).Control(3)=   "Frame2"
      Tab(6).Control(4)=   "SCHLST"
      Tab(6).Control(5)=   "MSS"
      Tab(6).ControlCount=   6
      Begin VB.CommandButton JUST_MADE 
         Caption         =   "&Just Made"
         Height          =   330
         Left            =   2640
         TabIndex        =   166
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox M_ROU_IDY 
         Height          =   315
         Left            =   -73320
         TabIndex        =   160
         Text            =   "M_ROU_IDY"
         Top             =   1320
         Width           =   1650
      End
      Begin MSFlexGridLib.MSFlexGrid SC 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   159
         Top             =   3960
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         _Version        =   393216
         Rows            =   0
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
      End
      Begin VB.TextBox TXT_SCH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   -72720
         TabIndex        =   156
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Height          =   1320
         Left            =   -68280
         TabIndex        =   153
         Top             =   2650
         Width           =   2055
         Begin VB.TextBox SCH_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   155
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Labe18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Scheme Amount"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2400
         Left            =   -67560
         TabIndex        =   148
         Top             =   270
         Width           =   1335
         Begin VB.CommandButton Command6 
            Caption         =   "&Valid Schms"
            Height          =   400
            Left            =   120
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Apply to &Bill"
            Height          =   400
            Left            =   120
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   1820
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Refresh"
            Height          =   400
            Left            =   120
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Select A&ll"
            Height          =   400
            Left            =   120
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   780
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView SCHLST 
         Height          =   2295
         Left            =   -74940
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Scheme Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Start Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "End Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Free&&Takeback"
         Height          =   330
         Left            =   -74520
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Height          =   4700
         Left            =   -70200
         TabIndex        =   106
         Top             =   315
         Width           =   3375
         Begin VB.TextBox G_FRE_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   3840
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
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   2400
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
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   2040
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
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox T_FRE_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   126
            Top             =   3840
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
            TabIndex        =   125
            Top             =   1320
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
            TabIndex        =   124
            Top             =   4200
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
            TabIndex        =   123
            Top             =   960
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
            TabIndex        =   122
            Top             =   600
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
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   3120
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
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   2760
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
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   3480
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
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   2400
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
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   2040
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
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox M_ROU_OFF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   4200
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
            TabIndex        =   114
            Top             =   1320
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
            TabIndex        =   113
            Top             =   960
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
            TabIndex        =   112
            Top             =   240
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
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   3120
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
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   2760
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
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   3480
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
            TabIndex        =   108
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox G_PRD_DIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lbl18 
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
            TabIndex        =   141
            Top             =   4200
            Width           =   225
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
            TabIndex        =   140
            Top             =   315
            Width           =   120
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
            TabIndex        =   139
            Top             =   1388
            Width           =   120
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
            TabIndex        =   138
            Top             =   615
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
            Index           =   1
            Left            =   1600
            TabIndex        =   137
            Top             =   1005
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
            TabIndex        =   136
            Top             =   2445
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
            Index           =   3
            Left            =   1600
            TabIndex        =   135
            Top             =   2805
            Width           =   105
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
            TabIndex        =   134
            Top             =   3525
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
            TabIndex        =   133
            Top             =   1695
            Width           =   75
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
            TabIndex        =   132
            Top             =   2055
            Width           =   75
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
            TabIndex        =   131
            Top             =   3135
            Width           =   105
         End
         Begin VB.Label Label 
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
            Left            =   1605
            TabIndex        =   130
            Top             =   3855
            Width           =   75
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Apply Schemes"
         Height          =   330
         Left            =   -74520
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   4685
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -69360
         MaxLength       =   10
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   -74940
         ScaleHeight     =   195
         ScaleWidth      =   8640
         TabIndex        =   94
         Top             =   4800
         Width           =   8700
         Begin VB.Label F_UNT_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   1200
            TabIndex        =   102
            Top             =   0
            Width           =   45
         End
         Begin VB.Label F_AMT_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3240
            TabIndex        =   101
            Top             =   0
            Width           =   45
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4080
            TabIndex        =   100
            Top             =   0
            Width           =   45
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   45
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   98
            Top             =   0
            Width           =   45
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   97
            Top             =   0
            Width           =   45
         End
         Begin VB.Label F_PRD_NME 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   1080
            TabIndex        =   96
            Top             =   0
            Width           =   45
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty."
            Height          =   195
            Left            =   2640
            TabIndex        =   95
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.ListBox List7 
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
         ItemData        =   "M_FRM_SOR.frx":0506
         Left            =   -73920
         List            =   "M_FRM_SOR.frx":051F
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox pcs_chk 
         Height          =   255
         Left            =   8280
         TabIndex        =   89
         Top             =   1440
         Width           =   255
      End
      Begin MSComctlLib.ListView LIST8 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   88
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
            Text            =   "Challan ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Route"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Reference"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox M_VAN_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   163
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox M_SMN_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -73440
         TabIndex        =   81
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74186755
         CurrentDate     =   37634
      End
      Begin VB.Frame Frame5 
         Caption         =   "Transport Receipt Details"
         Height          =   2175
         Left            =   -73800
         TabIndex        =   64
         Top             =   960
         Width           =   5655
         Begin VB.TextBox M_WBL_IDY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox M_TXN_TPT 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1620
            Width           =   4335
         End
         Begin VB.TextBox M_TTL_CAS 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4305
            MaxLength       =   5
            TabIndex        =   71
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
            TabIndex        =   66
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
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   765
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker M_LRN_DAT 
            Height          =   330
            Left            =   4305
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74186755
            CurrentDate     =   37097
            MaxDate         =   42369
            MinDate         =   36161
         End
         Begin MSComCtl2.DTPicker M_WBL_DAT 
            Height          =   330
            Left            =   4305
            TabIndex        =   76
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
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   74
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
            TabIndex        =   72
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   -74940
         ScaleHeight     =   195
         ScaleWidth      =   8640
         TabIndex        =   33
         Top             =   4800
         Width           =   8700
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty."
            Height          =   195
            Left            =   2640
            TabIndex        =   41
            Top             =   0
            Width           =   285
         End
         Begin VB.Label PRD_NME 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   1080
            TabIndex        =   40
            Top             =   0
            Width           =   45
         End
         Begin VB.Label AMT_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   45
         End
         Begin VB.Label TAX_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   45
         End
         Begin VB.Label ADD_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   45
         End
         Begin VB.Label DIS_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4080
            TabIndex        =   36
            Top             =   0
            Width           =   45
         End
         Begin VB.Label UOM_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3240
            TabIndex        =   35
            Top             =   0
            Width           =   45
         End
         Begin VB.Label UNT_TOT 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   1200
            TabIndex        =   34
            Top             =   0
            Width           =   45
         End
      End
      Begin VB.TextBox M_TXN_RMK 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72580
         MaxLength       =   70
         TabIndex        =   31
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
         TabIndex        =   23
         Top             =   4560
         Width           =   8295
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
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
         ItemData        =   "M_FRM_SOR.frx":0560
         Left            =   -72960
         List            =   "M_FRM_SOR.frx":0570
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox M_CRT_DAY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73320
         MaxLength       =   10
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73680
         MaxLength       =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox M_CST_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   164
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox M_ACT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4515
         Visible         =   0   'False
         Width           =   1095
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
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   4920
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74186755
         CurrentDate     =   38721
      End
      Begin MSMask.MaskEdBox M_FLX_AMT 
         Height          =   375
         Left            =   -69240
         TabIndex        =   12
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
         TabIndex        =   13
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
      Begin MSComctlLib.ListView list5 
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
            Text            =   "Invoice  ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Retailer Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Place"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Van Chln XRf"
            Object.Width           =   2383
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
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
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74640
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Salesman Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2293
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
            Text            =   "Retailer Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Market Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Route"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Salesman Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   2400
         Left            =   -74520
         TabIndex        =   42
         Top             =   1800
         Width           =   4095
         Begin VB.TextBox M_AMT_LES 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   240
            MaxLength       =   70
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   2000
            Width           =   1695
         End
         Begin VB.TextBox M_AMT_ADD 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   240
            MaxLength       =   70
            TabIndex        =   79
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   56
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
            TabIndex        =   55
            Top             =   2070
            Width           =   720
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sur - Charge on Tax % "
            Height          =   195
            Left            =   435
            TabIndex        =   52
            Top             =   1350
            Width           =   1635
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tax % "
            Height          =   195
            Left            =   1575
            TabIndex        =   51
            Top             =   975
            Width           =   480
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Discount % "
            Height          =   195
            Left            =   1215
            TabIndex        =   50
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Additonal Discount % "
            Height          =   195
            Left            =   510
            TabIndex        =   49
            Top             =   630
            Width           =   1545
         End
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   87
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product Name"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Quantity"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MRP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Sale Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Spl. Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Short Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "UPC"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ID"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Batidy"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Batch"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   4395
         Left            =   -74940
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   360
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   20
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
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
      Begin MSFlexGridLib.MSFlexGrid MSS 
         Height          =   2235
         Left            =   -74940
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2760
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
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
      Begin VB.Label LabelAAS 
         AutoSize        =   -1  'True
         Caption         =   "Schemes Applied-AutoMode"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -73080
         TabIndex        =   165
         Top             =   4755
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Order ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1575
         TabIndex        =   162
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label0 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Van Ref No."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   6450
         TabIndex        =   161
         Top             =   675
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label CST 
         Caption         =   "CST"
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label GST 
         Caption         =   "GST"
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free and Takebacks (-)"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -72105
         TabIndex        =   143
         Top             =   4245
         Width           =   1665
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Route"
         Height          =   195
         Left            =   -73920
         TabIndex        =   90
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1845
         TabIndex        =   83
         Top             =   4028
         Width           =   690
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Amount"
         Height          =   195
         Left            =   -71430
         TabIndex        =   62
         Top             =   1620
         Width           =   990
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Tax"
         Height          =   195
         Left            =   -71310
         TabIndex        =   61
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Discount"
         Height          =   195
         Left            =   -71670
         TabIndex        =   60
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Value"
         Height          =   195
         Left            =   -71445
         TabIndex        =   59
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73390
         TabIndex        =   32
         Top             =   3435
         Width           =   630
      End
      Begin VB.Label V_LST_CAP 
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
         TabIndex        =   26
         Top             =   480
         Width           =   8295
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1935
         Left            =   4935
         Top             =   1935
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   4920
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   4920
         X2              =   8520
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   4920
         X2              =   8520
         Y1              =   2295
         Y2              =   2295
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
         Left            =   4920
         TabIndex        =   20
         Top             =   4515
         Visible         =   0   'False
         Width           =   3615
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
         Left            =   4920
         TabIndex        =   19
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Label M_CST_AD4 
         BackColor       =   &H00800000&
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
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5040
         TabIndex        =   18
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label M_CST_AD3 
         BackColor       =   &H00800000&
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
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5040
         TabIndex        =   17
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label M_CST_AD2 
         BackColor       =   &H00800000&
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
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5040
         TabIndex        =   16
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label M_CST_AD1 
         BackColor       =   &H00800000&
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
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5040
         TabIndex        =   15
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label M_CST_NME 
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
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5040
         TabIndex        =   14
         Top             =   1950
         Width           =   3375
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R/Off"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -70845
         TabIndex        =   10
         Top             =   4605
         Width           =   405
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Days"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -74250
         TabIndex        =   9
         Top             =   675
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
         TabIndex        =   8
         Top             =   1035
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2370
         TabIndex        =   7
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( Credit ) Sales Account"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   870
         TabIndex        =   6
         Top             =   4583
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4320
         TabIndex        =   4
         Top             =   675
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1965
         TabIndex        =   3
         Top             =   1995
         Width           =   570
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
            Picture         =   "M_FRM_SOR.frx":0580
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":0AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":0E8C
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":11E0
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":1818
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":1B6C
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":1FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":22C0
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":26CC
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":27D8
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":2B2C
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":2F40
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOR.frx":348C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   84
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
         NumButtons      =   32
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pre"
            Object.ToolTipText     =   "Click Here for System Generated Order Number"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "crt"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "brw"
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
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EMP_SHL"
            Object.ToolTipText     =   "Click to Check Post Empties After SAVE"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Freeze Van Reference No. of Last Invoice"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "smn1"
            Object.ToolTipText     =   "Freeze Retailers for a Salesman"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAY"
            Object.ToolTipText     =   "Freeze Retailers based on Market Visit Day"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AUS"
            Object.ToolTipText     =   "Check to Auto Apply Schemes"
            ImageIndex      =   10
            Style           =   1
         EndProperty
      EndProperty
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   225
         ItemData        =   "M_FRM_SOR.frx":38E0
         Left            =   6840
         List            =   "M_FRM_SOR.frx":38E7
         TabIndex        =   146
         Top             =   80
         Width           =   615
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         Height          =   225
         ItemData        =   "M_FRM_SOR.frx":38F3
         Left            =   7560
         List            =   "M_FRM_SOR.frx":38F5
         TabIndex        =   145
         TabStop         =   0   'False
         ToolTipText     =   "Select Route"
         Top             =   80
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   340
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ListBox text1 
         Appearance      =   0  'Flat
         Height          =   225
         ItemData        =   "M_FRM_SOR.frx":38F7
         Left            =   5880
         List            =   "M_FRM_SOR.frx":3910
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   80
         Visible         =   0   'False
         Width           =   765
      End
   End
End
Attribute VB_Name = "M_FRM_SOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, bslab, MSROW, INVIDY, INC, dlr_idy, ZIPCODE, M_DBT_RAP
Dim SYS_NUM As Boolean
Dim APPLYCSR As Boolean
Dim APPSCH As Boolean
Dim M_AUS_OK As Integer
Dim CHGQTY
Dim PFA As Double
Dim SFA As Double
Dim M_LST_RAT As Double
Dim M_LST_DIS As Double
Dim M_CPN_AMT As Double
Dim M_INC_AMT As Double
Dim rsrdy As New ADODB.Recordset
Dim VAN_IDY As String
Dim M_HVO_FLG As String
Dim M_HVO_CHK As String
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
T_PRD_DIS = convert(Val(AMT_TOT) - (Val(DIS_TOT) + Val(ADD_TOT)))

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

G_FRE_AMT = convert(Abs(Val(F_AMT_TOT)))
T_FRE_AMT = convert(Val(T_FRT_AMT) - Val(G_FRE_AMT))

ROU_OFF
T_NET_AMT = convert(Val(T_FRE_AMT) + Val(M_ROU_OFF))
End Sub
Private Sub saverecord()
On Error GoTo LBL
If CHECKTB("act.dbf,smn.dbf,mar.dbf,loc.dbf,cat.dbf,prd.dbf,brd.dbf,inl.dbf,inv.dbf,lnv.dbf,pnl.dbf,pnv.dbf,irt.dbf,pgr.dbf,dir.dbf,bnv.dbf,xnv.dbf,bnl.dbf,xnl.dbf,cbk.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim RSINL As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
Dim REC_IRT As New ADODB.Recordset
Dim REC_PNL As New ADODB.Recordset
Dim SKU_RS As New ADODB.Recordset
Dim VSC_RS As New ADODB.Recordset
Dim ISC_RS As New ADODB.Recordset

If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
   ' M_FRM_ACT.Tag = ""
    Exit Sub
Else
    CheckShutDownStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'")
    If Trim(CheckShutDownStatus) = "IMPROPER SHUTDOWN" Then
        MsgBox ("Please reset the Improper Shutdown flag in System-Shutdown Status"), vbCritical, "DAS Version 4.0"
        CANCELRECORD
        Exit Sub
    End If
End If
'CHK_FTD = Val(Trim(GetValue("DIR_ALS", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'")))
'If Val(CHK_FTD) > 0 Then
'    If M_INV_DAT > DateAdd("d", Val(CHK_FTD), Date) Then
'        MsgBox ("Cannot Save Txn! Verify Date! Future Date Encountered!"), vbInformation, "DAS Version 4.0"
'        CANCELRECORD
'        Exit Sub
'    End If
'End If
    
If SYS_NUM = False Then
    If MOD_REC = False Then
        If RSINL.State = 1 Then RSINL.Close
        RSINL.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SSO'", con, adOpenKeyset, adLockPessimistic
        If Not RSINL.EOF Then
            MsgBox "Cannot Save, Sale Order WITH THIS NUMBER Already Exists!", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 0
            CANCELRECORD
'            M_INV_IDY.SetFocus
            Exit Sub
        End If
    End If
End If
If Toolbar1.Buttons(11).Value = tbrUnpressed Then
    If Len(Trim(M_INV_IDY)) = 0 Then
        MsgBox "Sale Order Number or ID Required!, Cannot Save this Sale Entry!", vbCritical, "DAS Version 4.0"
        CANCELRECORD
        Exit Sub
    End If
End If

If REC_SET1.State = 1 Then REC_SET1.Close
REC_SET1.Open "SELECT * FROM VAN WHERE INV_IDY = '" & Trim(M_VAN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If REC_SET1.RecordCount > 0 Then
    If Trim(REC_SET1("VAN_CLS")) = "C" Then
        MsgBox "Reference Van Challan Status is CLOSED for this Sale Order, Cannot Save !", vbCritical, "DAS Version 4.0"
        DayCls_PrintBill (M_INV_IDY)
        CANCELRECORD
        Exit Sub
    End If
End If

If check("ACT", "(ACT_TYP LIKE 'SAL' OR ACT_TYP LIKE 'BOT') AND ACT_idy", Trim(M_CST_IDY)) = False Then
    MsgBox "Party Not Found ,Define Party !", vbInformation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_CST_IDY.SetFocus
    Exit Sub
End If

If check("SMN", "SMN_IDY", Trim(M_SMN_IDY)) = False Then
    MsgBox "Salesman Not Found ,Define Salesman !", vbInformation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_SMN_IDY.SetFocus
    Exit Sub
End If

'If check("act", "(grp_clf like 'Sales' and ACT_TYP LIKE '') AND act_idy", Trim(M_ACT_IDY)) = False Then
'    MsgBox "Sales Account Not Found, Define Sales  Account!", vbInformation, "DAS Version 4.0"
'    M_ACT_IDY.SetFocus
'    SSTab1.Tab = 0
'    Exit Sub
'End If

If Val(T_NET_AMT) < 0 Then
    MsgBox "Invalid Sale Order Net Amount, Cannot Save Entry!", vbCritical, "DAS Version 4.0"
    SSTab1.Tab = 2
    M_EXD_AMT.SetFocus
    Exit Sub
End If

' Day Closure Routine Added by PNR and PKB as Addon Control

Dim RS_DAYCLS As New ADODB.Recordset
If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
RS_DAYCLS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) like 'DAY' .and. alltrim(DIR_ALS) like '" & M_INV_DAT & "'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
If RS_DAYCLS.RecordCount > 0 Then
   If Trim(RS_DAYCLS!DIR_MSG) = "Y" Then
    MsgBox "Day Closure Effective ! Cannot Alter - Create Transaction FOR " & RS_DAYCLS("DIR_ALS"), vbCritical, "DAS Version 4.0"
    ' Print Bill Even If Day Closure is Effective - Added by PNR
    DayCls_PrintBill (M_INV_IDY)
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
CHK_DAYS = GetValue("DIR_ALS", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX = 'DCD'")

CHK_DATE = LST_DCD + CHK_DAYS
If M_INV_DAT <= CHK_DATE Then
Else
    MsgBox "Day Closure Tolerance Days " & Trim(str(CHK_DAYS)) & " Crossed! Cannot Alter - Create Transaction !", vbCritical, "DAS Version 4.0"
    CANCELRECORD
    Exit Sub
End If

' End of Day Closure Tolerance Days Check
If (Trim(CHGQTY) = "Yes Schemes" Or Trim(CHGQTY) = "Yes Schemes Qty") And APPSCH = False Then
If MsgBox("Old Schemes are Cleared! Re-apply Schemes ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
Command1_Click
Exit Sub
End If
End If
b = 0
C = 0
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 2)) > 0 Then
        b = 1
    End If
Next

For d = 1 To MSF.Rows - 1
    If Val(MSF.TextMatrix(d, 2)) > 0 Then
        C = 1
    End If
Next

If b = 0 And C = 1 Then
    MsgBox "Sale Products Entries Are Not Made! Cannot Save Sale Order Entry !", vbCritical, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_INV_DAT.SetFocus
    Exit Sub
End If

If check("dir", "dir_idy like 'SYS' AND DIR_IDX LIKE 'PSW' AND DIR_MSG", "Y") = True Then
    For A = 1 To MS.Rows - 1
        If Trim(MS.TextMatrix(A, 0)) <> "" Then
            If pcs_chk.Value = 1 Then
                If check("BRD", "prd_pdr > " & Val(MS.TextMatrix(A, 5)) & " and BAT_idy ", Trim(MS.TextMatrix(A, 15))) = True Then
                    MsgBox "Cannot Save Sale Order! Sale - Order Rate is Less Than Purchase Rate", vbExclamation, "DAS Version 4.0"
                    MS.SetFocus
                    MS.Row = A
                    SSTab1.Tab = 1
                    Exit Sub
                End If
            Else
                PURRAT = Round(Val(GetValue("PRD_PDR", "BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 15)) & "'")) / Val(MS.TextMatrix(A, 13)), 2)
                If Val(PURRAT) > Val(MS.TextMatrix(A, 5)) Then
                    MsgBox "Cannot Save Sale Order! Sale - Order Rate is Less Than Purchase Rate", vbExclamation, "DAS Version 4.0"
                    MS.SetFocus
                    MS.Row = A
                    SSTab1.Tab = 1
                    Exit Sub
                End If
            End If
        End If
    Next
End If
If Len(Trim(M_ROU_IDY)) = 0 Then
    MsgBox "Route Input Required!", vbInformation, "DAS Version 4.0"
    SSTab1.Tab = 2
    M_ROU_IDY.SetFocus
    Exit Sub
End If
If MOD_REC = True Then
End If
If Len(Trim(M_VAN_IDY)) = 10 Then
   'Do not Check for Stock Availability if VAN CHALLAN Reference is available while billing - Modified by Radhika and Ramanesh
   'Check added by PNR after Prince Chopras request on this
'''''    A = STOCKCHECK
'''''    If Val(A) <> 0 Then
'''''        MsgBox "Stock Out of Range, Check Stock!", vbInformation, "DAS Version 4.0"
'''''        SSTab1.Tab = 6
'''''        MSS.Row = Val(A)
'''''        MSS.SetFocus
'''''        Exit Sub
'''''    End If
Else
    ' Do not Check for Stock Availability while saving since this is Order - Modified by PNR as per PC update
'    A = STOCKCHECK
'    If Val(A) <> 0 Then
'        MsgBox "Stock Out of Range, Check Stock!", vbInformation, "DAS Version 4.0"
'        SSTab1.Tab = 6
'        MSS.Row = Val(A)
'        MSS.SetFocus
'        Exit Sub
'    End If
End If
GRS_CAL

'Records Save Process Starts - PNR

If MOD_REC = True Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SSO'", con, adOpenKeyset, adLockPessimistic
    For A = 1 To REC_SET.RecordCount
        If (Trim(REC_SET!stk_typ) = "Free" Or Trim(REC_SET!stk_typ) = "Salable") And stk_upd = "Y" Then
            con.Execute "update prd where prd_idy like '" & Trim(REC_SET("prd_idy")) & "' set cur_lev = " & ADDQTY(Val(REC_SET("prd_qty")), Trim(REC_SET("prd_idy"))), d
            con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET("prd_idy")) & "' set PRD_PCS = PRD_PCS + " & MTS(Val(REC_SET("prd_qty")), Val(REC_SET("SUB_UNT"))), d
            con.Execute "update Brd where BAT_idy like '" & Trim(REC_SET("BAT_idy")) & "' set PRD_QTY = " & ADDBRDQTY(Val(REC_SET("PRD_qty")), Trim(REC_SET("BAT_idy"))), d
        ElseIf (Trim(REC_SET!stk_typ) = "Damage" Or Trim(REC_SET!stk_typ) = "TakeBack" Or Trim(REC_SET!stk_typ) = "Breakage" Or Trim(REC_SET!stk_typ) = "Leakage") And stk_upd = "Y" Then
            con.Execute "update prd where prd_idy like '" & Trim(REC_SET("prd_idy")) & "' set cur_lev = " & CONQTY(Val(REC_SET("prd_qty")), Trim(REC_SET("prd_idy")))
            con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET("prd_idy")) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(REC_SET("prd_qty")), Val(REC_SET("SUB_UNT")))
            con.Execute "update Brd where BAT_idy like '" & Trim(REC_SET("BAT_idy")) & "' set PRD_QTY = " & CONBRDQTY(Val(REC_SET("PRD_qty")), Trim(REC_SET("BAT_idy")))
        End If
        REC_SET.MoveNext
    Next
    
    TxnMoveBackup "INV", "BNV", Trim(M_INV_IDY), "SSO"
    TxnMoveBackup "LNV", "CNV", Trim(M_INV_IDY), ""
    TxnMoveBackup "INL", "BNL", Trim(M_INV_IDY), "SSO"
    
    con.Execute "DELETE FROM INV WHERE INV_IDY LIKE '" & Trim(M_INV_IDY) & "' AND TXN_TYP LIKE 'SSO'"
    con.Execute "DELETE FROM LNV WHERE INV_IDY LIKE '" & Trim(M_INV_IDY) & "'"
    
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SSO'", con, adOpenKeyset, adLockPessimistic
    INC = Trim(M_INV_IDY)
Else
    REC_SET.Open "select * from inl where inv_idy =''", con, adOpenKeyset, adLockPessimistic
    REC_SET.AddNew
    If Toolbar1.Buttons(11).Value = tbrPressed Then
        REC_SET("inv_idy") = "SSO" & T7increment("IDY_SSO")
    Else
        REC_SET("inv_idy") = M_INV_IDY
    End If
    INC = REC_SET("inv_idy")
    REC_SET("dch_idy") = M_DCH_IDY & ""
    INVIDY = INC
End If
' *
REC_SET("inv_dat") = M_INV_DAT
REC_SET("ord_idy") = Trim(M_VAN_IDY & "")
VAN_IDY = Trim(M_VAN_IDY)
REC_SET("cst_idy") = Trim(M_CST_IDY & "")
REC_SET("smn_idy") = Trim(M_SMN_IDY & "")
REC_SET("act_idy") = Trim(M_ACT_IDY & "")
REC_SET("dch_dat") = M_DCH_DAT
REC_SET("csd_pcg") = discount(M_CSD_PCG)
REC_SET("csd_amt") = Val((M_CSD_AMT))
REC_SET("spd_pcg") = discount(M_SPD_PCG)
REC_SET("spd_amt") = Val((M_SPD_AMT))
REC_SET("tax_pcg") = discount(M_TAX_PCG)
REC_SET("tax_amt") = Val((M_TAX_AMT))
REC_SET("exd_pcg") = Val(M_EXD_PCG)             'SURCHARGE PCG
REC_SET("EXD_AMT") = Val(Val(M_EXD_AMT))             'SURCHARGE AMT
REC_SET("pkg_cst") = Val(Val(M_PKG_CST))             'AMOUNT (+)
REC_SET("FRT_AMT") = Val(Val(M_FRT_AMT & ""))        'AMOUNT (-)
REC_SET("LRN_IDY") = Trim(M_LRN_IDY)
REC_SET("LRN_AMT") = Val(M_LRN_AMT & "")
REC_SET("LRN_DAT") = M_LRN_DAT
REC_SET("WBL_IDY") = Trim(M_WBL_IDY & "")
REC_SET("WBL_DAT") = M_WBL_DAT
REC_SET("TXN_RMK") = Trim(M_TXN_RMK)
REC_SET("TTL_CAS") = Val(M_TTL_CAS)
REC_SET("txn_typ") = "SSO"
REC_SET("grs_amt") = Val((Val(T_GRS_AMT)))
REC_SET("net_amt") = Val(T_NET_AMT)
REC_SET("crt_day") = Val(M_CRT_DAY)
REC_SET("due_dat") = DateAdd("d", Val(m_crt_dat), M_INV_DAT)
REC_SET("rou_off") = Val(M_ROU_OFF)
REC_SET("ROU_IDY") = Trim(M_ROU_IDY)
REC_SET("TXN_TPT") = Trim(M_TXN_TPT)
REC_SET("TXN_RM1") = Trim(M_AMT_ADD)
REC_SET("TXN_RM2") = Trim(M_AMT_LES)
REC_SET!USR_IDY = Trim(USERNAME)
REC_SET!SRN_AMT = Val(G_FRE_AMT)
REC_SET!sys_dat = Date
REC_SET!sys_tim = Format(Time, "HH:MM:SS")
REC_SET("SCH_CHK") = "Y"
REC_SET!vch_idy = ""
REC_SET!Sch_Idy = ""
REC_SET.Update
'IRT UPDATION   && IRT will not be updated based on order capture modified by PNR
'If Val(T_NET_AMT) >= 0 Then
'    REC_IRT.Open "select * from IRT WHERE INV_IDY LIKE '" & Trim(INC) & "'", CON, adOpenKeyset, adLockPessimistic
'    If REC_IRT.RecordCount = 0 Then
'        REC_IRT.AddNew
'    End If
'    REC_IRT("inv_idy") = INC
'    REC_IRT("inv_dat") = M_INV_DAT
'    REC_IRT("PTY_idy") = Trim(M_CST_IDY & "")
'    REC_IRT("smn_idy") = Trim(M_SMN_IDY & "")
'    REC_IRT("txn_typ") = "SSO"
'    REC_IRT("net_amt") = Val(T_NET_AMT)
'    REC_IRT("due_dat") = Date
'    REC_IRT("act_idy") = ""
'    REC_IRT.Update
'End If
'END OF IRT.DBF UPDATION
'PRINT DATABASE LOGIC BEGIN
If MOD_REC = True Then
    If CHECKTB("pnl.dbf,pnv.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    TxnMoveBackup "PNV", "XNV", Trim(INC), "SSO"
    TxnMoveBackup "PNL", "XNL", Trim(INC), "SSO"
    
    PNL.Execute "delete from pnv where inv_idy like '" & Trim(INC) & "'"
    PNL.Execute "delete from pnl where inv_idy like '" & Trim(INC) & "'"
End If
REC_PNL.Open "SELECT * FROM PNL where inv_idy = ''", PNL, adOpenKeyset, adLockPessimistic
REC_PNL.AddNew
REC_PNL!INV_IDY = INC
REC_PNL("inv_dat") = M_INV_DAT
REC_PNL("ord_idy") = Trim(M_VAN_IDY & "")
REC_PNL("cst_idy") = Trim(M_CST_IDY & "")
REC_PNL("cst_nme") = Trim(M_CST_NME)
REC_PNL("cst_ad1") = Trim(M_CST_AD1)
REC_PNL("cst_ad2") = Trim(M_CST_AD2)
REC_PNL("cst_ad3") = Trim(M_CST_AD3)
REC_PNL("cst_ad4") = Trim(M_CST_AD4)
REC_PNL("cst_zip") = Trim(ZIPCODE)
REC_PNL("Act_idy") = Trim(M_ACT_IDY & "")
REC_PNL("dch_idy") = REC_SET("dch_idy")
REC_PNL("csd_pcg") = Val(M_CSD_PCG)
REC_PNL("csd_amt") = Val(M_CSD_AMT)
REC_PNL("spd_pcg") = Val(M_SPD_PCG)
REC_PNL("spd_amt") = Val(amt(M_SPD_AMT))
REC_PNL("tax_pcg") = Val(M_TAX_PCG)
REC_PNL("tax_amt") = Val(amt(M_TAX_AMT))
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
REC_PNL("txn_typ") = "SSO"
REC_PNL("grs_amt") = Val(T_GRS_AMT)
REC_PNL("net_amt") = Val(T_NET_AMT)
REC_PNL("crt_day") = Val(M_CRT_DAY)
REC_PNL("due_dat") = Date
REC_PNL("rou_off") = Val(M_ROU_OFF)
REC_PNL("ROU_IDY") = Trim(M_ROU_IDY)
REC_PNL("TXN_TPT") = Trim(M_TXN_TPT)
REC_PNL("TXN_RM1") = Trim(M_AMT_ADD)
REC_PNL("TXN_RM2") = Trim(M_AMT_LES)
REC_PNL!usr_nme = Trim(USERNAME)
REC_PNL!sys_dat = Date
REC_PNL!sys_tim = Format(Time, "HH:MM:SS")
REC_PNL("SCH_CHK") = "Y"
REC_PNL!SRN_AMT = Val(G_FRE_AMT)
REC_PNL!vch_idy = ""
REC_PNL!Sch_Idy = ""
REC_PNL!cst_gst = CST
REC_PNL!cst_cst = GST
REC_PNL!MAR_IDY = GetValue("mar_idy", "act where act_idy = '" & Trim(M_CST_IDY) & "'")

REC_PNL!mar_nme = Trim(GetValue("mar_nme", "mar where mar_idy = '" & Trim(REC_PNL("mar_idy")) & "'"))
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
REC_PNL!MAR_DOP = Date
REC_PNL.Update
REC_PNL.Close
'PRINT DATABASE LOGIC END

If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inv where inv_idy =''", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" And Val(MS.TextMatrix(A, 2)) > 0 And MS.TextMatrix(A, 1) <> "" Then
        REC_SET.AddNew
        REC_SET("inv_idy") = Trim(INC)
        REC_SET("inv_dat") = Trim(M_INV_DAT)
        REC_SET("prd_idy") = MS.TextMatrix(A, 0)
        REC_SET("STK_TYP") = Trim(MS.TextMatrix(A, 1))
        REC_SET("RAT_APP") = Trim(MS.TextMatrix(A, 4))
        If pcs_chk.Value = 1 Then
            REC_SET("prd_qty") = Val(MS.TextMatrix(A, 2))
            REC_SET("prd_rat") = Val((Val(MS.TextMatrix(A, 5))))
            LST_TOT = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 13))))
            M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
            'LST_TOT = Val(MS.TextMatrix(A, 2)) * Val(MS.TextMatrix(A, 22))
            'M_CPN_AMT = Val(MS.TextMatrix(A, 2)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        Else
            REC_SET("PRD_QTY") = STM(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            REC_SET("prd_rat") = Val(Val(MS.TextMatrix(A, 5))) * Val(MS.TextMatrix(A, 13))
            LST_TOT = Val(MS.TextMatrix(A, 2))
            LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 13))))
            M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
            'LST_TOT = (Val(MS.TextMatrix(A, 2)) / Val(MS.TextMatrix(A, 13))) * Val(MS.TextMatrix(A, 22))
            'M_CPN_AMT = (Val(MS.TextMatrix(A, 2)) / Val(MS.TextMatrix(A, 13))) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        End If
        REC_SET("ADD_AMT") = Val((Val(MS.TextMatrix(A, 6))))
        REC_SET("trd_dis") = Val(MS.TextMatrix(A, 7))
        REC_SET("LST_RAT") = Val(MS.TextMatrix(A, 22))
        REC_SET("LST_DIS") = Val(MS.TextMatrix(A, 23))
        REC_SET("CPN_AMT") = Round(M_CPN_AMT, 2)
        'Val (MS.TextMatrix(A, 2)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        REC_SET("TRD_TAX") = Val(MS.TextMatrix(A, 8))
        '+ (Val(MS.TextMatrix(A, 2)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100))
        If pcs_chk.Value = 1 Then
            INCAMT = RATCAL(Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 24) & ""), Val(MS.TextMatrix(A, 13)))
            PRDNET = RATCAL(Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 5) & ""), Val(MS.TextMatrix(A, 13)))
        Else
            PRDNET = Val(MS.TextMatrix(A, 2) & "") * Val(MS.TextMatrix(A, 24) & "")
            PRDNET = Val(MS.TextMatrix(A, 2) & "") * Val(MS.TextMatrix(A, 5) & "")
        End If
        REC_SET("DIS_AMT") = ((PRDNET - REC_SET("ADD_AMT")) * Val(MS.TextMatrix(A, 7)) / 100)
'        M_CPN_AMT = Val(MS.TextMatrix(A, 2)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
'        LST_TOT = Val(MS.TextMatrix(A, 2)) * Val(MS.TextMatrix(A, 22))
        If Trim(GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'DPV'")) = "T" Then
            TRDTAX = (PRDNET - (Val(INCAMT))) * Val(MS.TextMatrix(A, 8)) / 100
        Else
            TRDTAX = (PRDNET - (Val(REC_SET("ADD_AMT")) + Val(REC_SET("DIS_AMT")) + Val(INCAMT))) * Val(MS.TextMatrix(A, 8)) / 100
        End If

'        TRDTAX = (PRDNET - (Val(REC_SET("ADD_AMT")) + Val(REC_SET("DIS_AMT")) + Val(INCAMT))) * Val(MS.TextMatrix(A, 8)) / 100
        TRDTAX = TRDTAX + M_CPN_AMT
        REC_SET("TAX_AMT") = Round(Val(TRDTAX & ""), 2)
        REC_SET("txn_typ") = "SSO"
        REC_SET!BAT_IDY = Trim(MS.TextMatrix(A, 15))
        REC_SET!PRD_UOM = Round(Val((MS.TextMatrix(A, 3))), 2)
        REC_SET!SUB_UNT = Val(MS.TextMatrix(A, 13))
        If pcs_chk.Value = 1 Then
            REC_SET!PRD_UNT = Int(Val(MS.TextMatrix(A, 2)))
            REC_SET!PRD_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13))) Mod Val(MS.TextMatrix(A, 13))
            REC_SET!TOT_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
        Else
            REC_SET!PRD_UNT = Int(Val(MS.TextMatrix(A, 2) / MS.TextMatrix(A, 13)))
            REC_SET!PRD_PCS = Val(MS.TextMatrix(A, 2)) Mod Val(MS.TextMatrix(A, 13))
            REC_SET!TOT_PCS = Val(MS.TextMatrix(A, 2))
        End If
        REC_SET("COL_AMT") = Val((MS.TextMatrix(A, 9)) + Round(LST_TOT, 2))
        REC_SET("INC_TTL") = Val(INCAMT)
        If Len(Trim(M_VAN_IDY)) = 0 Then
            REC_SET!stk_upd = "Y"
        Else
            REC_SET!stk_upd = "N"
        End If
        REC_SET!ORD_IDY = M_VAN_IDY
        If APPSCH = True And Trim(MSS.Rows > 1) <> "" Then
            REC_SET!FST_FLG = "Y"
        Else
            REC_SET!FST_FLG = "N"
        End If
        REC_SET.Update
        If Len(Trim(M_VAN_IDY)) = 0 Then
            If pcs_chk.Value = 1 Then
                con.Execute "update PRD where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "' set CUR_LEV = " & CONQTY(Trim(MS.TextMatrix(A, 2)), Trim(MS.TextMatrix(A, 0))), d
                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13))), d
                con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 15)) & "' set prd_qty = " & CONBRDQTY(Trim(MS.TextMatrix(A, 2)), Trim(MS.TextMatrix(A, 15))), d
            Else
                con.Execute "update PRD where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "' set CUR_LEV = " & CONQTY(QTY(STM(MS.TextMatrix(A, 2), Val(MS.TextMatrix(A, 13)))), Trim(MS.TextMatrix(A, 0)))
                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS - " & Val(MS.TextMatrix(A, 2))
                con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MS.TextMatrix(A, 15)) & "' set prd_qty = " & CONBRDQTY(STM(MS.TextMatrix(A, 2), Val(MS.TextMatrix(A, 13))), Trim(MS.TextMatrix(A, 15)))
            End If
        End If
    End If
Next

'FREE inv UPDATE BEGIN
'If REC_SET.State = 1 Then REC_SET.Close
'REC_SET.Open "select * from inv", CON, adOpenKeyset, adlockpessimistic
For A = 1 To MSF.Rows - 1
    If Trim(MSF.TextMatrix(A, 0)) <> "" And Val(MSF.TextMatrix(A, 2)) > 0 And MSF.TextMatrix(A, 1) <> "" Then
        REC_SET.AddNew
        REC_SET("inv_idy") = Trim(INC)
        REC_SET("inv_dat") = Trim(M_INV_DAT)
        REC_SET("prd_idy") = MSF.TextMatrix(A, 0)
        REC_SET("STK_TYP") = Trim(MSF.TextMatrix(A, 1))
        REC_SET("TRD_TAX") = Val(MSF.TextMatrix(A, 16))
        'Crate Rental and TAX on crate rental is updated for free goods - changed by PNR
        REC_SET("LST_RAT") = Val(MSF.TextMatrix(A, 17))
        REC_SET("LST_DIS") = Val(MSF.TextMatrix(A, 18))
        
        LST_TOT = MTS(Val(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9)))
        LST_TOT = (LST_TOT * (Val(MSF.TextMatrix(A, 17)) / Val(MSF.TextMatrix(A, 9))))
        M_CPN_AMT = ((LST_TOT * Val(MSF.TextMatrix(A, 18))) / 100)
        
        'Tax is calculated for free goods - changed by PNR
        M_TAX_AMT = (Val(MSF.TextMatrix(A, 19)) / Val(MSF.TextMatrix(A, 9))) * (MTS(Val(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9))))
        M_TAX_AMT = ((M_TAX_AMT * Val(MSF.TextMatrix(A, 16))) / 100)
        
        REC_SET("TAX_AMT") = 0# 'Tax is reset again to zero
        REC_SET("CPN_AMT") = Round(M_CPN_AMT, 2) 'Tax on Crate Rental is stored
        
        If pcs_chk.Value = 1 Then
            REC_SET("prd_qty") = Val(MSF.TextMatrix(A, 2))
            REC_SET("prd_rat") = Val((Val(MSF.TextMatrix(A, 4))))
        Else
            REC_SET("PRD_QTY") = STM(Val(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9)))
            REC_SET("prd_rat") = Val(Val(MSF.TextMatrix(A, 4))) * Val(MSF.TextMatrix(A, 9))
        End If
        If pcs_chk.Value = 1 Then
            PRDNET = RATCAL(Val(MSF.TextMatrix(A, 2) & ""), Val(MSF.TextMatrix(A, 4) & ""), Val(MSF.TextMatrix(A, 9)))
        Else
            PRDNET = Val(MSF.TextMatrix(A, 2) & "") * Val(MSF.TextMatrix(A, 4) & "")
        End If
        REC_SET("txn_typ") = "SSO"
        REC_SET("PRD_RAT") = Val(MSF.TextMatrix(A, 4))
        REC_SET!BAT_IDY = Trim(MSF.TextMatrix(A, 11))
        REC_SET!PRD_UOM = Val((MSF.TextMatrix(A, 3)))
        REC_SET!SUB_UNT = Val(MSF.TextMatrix(A, 9))
        If pcs_chk.Value = 1 Then
            REC_SET!PRD_UNT = Int(Val(MSF.TextMatrix(A, 2)))
            REC_SET!PRD_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9))) Mod Val(MSF.TextMatrix(A, 9))
            REC_SET!TOT_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9)))
        Else
            REC_SET!PRD_UNT = Int(Val(MSF.TextMatrix(A, 2) / MSF.TextMatrix(A, 9)))
            REC_SET!PRD_PCS = Val(MSF.TextMatrix(A, 2)) Mod Val(MSF.TextMatrix(A, 9))
            REC_SET!TOT_PCS = Val(MSF.TextMatrix(A, 2))
        End If
        REC_SET("COL_AMT") = Val((MSF.TextMatrix(A, 5)))
        If Len(Trim(M_VAN_IDY)) = 0 Then
            REC_SET!stk_upd = "Y"
        Else
            REC_SET!stk_upd = "N"
        End If
        REC_SET!ORD_IDY = M_VAN_IDY
        REC_SET.Update
        If Len(Trim(M_VAN_IDY)) = 0 Then
            If MSF.TextMatrix(A, 1) = "Free" Then
                If pcs_chk.Value = 1 Then
                    con.Execute "update PRD where prd_idy like '" & Trim(MSF.TextMatrix(A, 0)) & "' set CUR_LEV = " & CONQTY(Trim(MSF.TextMatrix(A, 2)), Trim(MSF.TextMatrix(A, 0))), d
                    con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MSF.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9))), d
                    con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MSF.TextMatrix(A, 11)) & "' set prd_qty = " & CONBRDQTY(Trim(MSF.TextMatrix(A, 2)), Trim(MSF.TextMatrix(A, 11))), d
                Else
                    con.Execute "update PRD where prd_idy like '" & Trim(MSF.TextMatrix(A, 0)) & "' set CUR_LEV = " & CONQTY(QTY(STM(MSF.TextMatrix(A, 2), Val(MSF.TextMatrix(A, 9)))), Trim(MSF.TextMatrix(A, 0)))
                    con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MSF.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS - " & Val(MSF.TextMatrix(A, 2))
                    con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MSF.TextMatrix(A, 11)) & "' set prd_qty = " & CONBRDQTY(STM(MSF.TextMatrix(A, 2), Val(MSF.TextMatrix(A, 9))), Trim(MSF.TextMatrix(A, 11)))
                End If
            ElseIf MSF.TextMatrix(A, 1) = "Damage" Or MSF.TextMatrix(A, 1) = "TakeBack" Or MSF.TextMatrix(A, 1) = "Breakage" Or MSF.TextMatrix(A, 1) = "Leakage" Then
                If pcs_chk.Value = 1 Then
                    con.Execute "update PRD where prd_idy like '" & Trim(MSF.TextMatrix(A, 0)) & "' set CUR_LEV = " & ADDQTY(Trim(MSF.TextMatrix(A, 2)), Trim(MSF.TextMatrix(A, 0)))
                    con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MSF.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS + " & MTS(Val(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9)))
                    con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MSF.TextMatrix(A, 11)) & "' set prd_qty = " & ADDBRDQTY(Trim(MSF.TextMatrix(A, 2)), Trim(MSF.TextMatrix(A, 11)))
                Else
                    con.Execute "update PRD where prd_idy like '" & Trim(MSF.TextMatrix(A, 0)) & "' set CUR_LEV = " & ADDQTY(QTY(STM(MSF.TextMatrix(A, 2), Val(MSF.TextMatrix(A, 9)))), Trim(MSF.TextMatrix(A, 0)))
                    con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MSF.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS + " & Val(MSF.TextMatrix(A, 2))
                    con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MSF.TextMatrix(A, 11)) & "' set prd_qty = " & ADDBRDQTY(STM(MSF.TextMatrix(A, 2), Val(MSF.TextMatrix(A, 9))), Trim(MSF.TextMatrix(A, 11)))
                End If
            End If
        End If
    End If
Next
'FREE UPDATE END

'scheme update begin
'If REC_SET.State = 1 Then REC_SET.Close
'REC_SET.Open "select * from inv", CON, adOpenKeyset, adlockpessimistic
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
'If FS_SOU.FileExists(CURDIR & "\SCH_SKP.TXT") = True Then
'    FS_SOU.DeleteFile (CURDIR & "\SCH_SKP.TXT")
'End If
Set F_OMS = FS_SOU.OpenTextFile(CURDIR & "\SCH_SKP.TXT", 8, True)
With MSS
For A = 1 To MSS.Rows - 1
    If .TextMatrix(A, 0) <> "" And (Val(.TextMatrix(A, 2)) > 0 Or Val(.TextMatrix(A, 4)) > 0) And APPSCH = True Then
'        If SKU_RS.State = 1 Then SKU_RS.Close
'        SKU_RS.Open "SELECT * FROM PRD", CON, adOpenKeyset, adlockpessimistic
'        If SKU_RS.RecordCount > 0 Then SKU_RS.MoveFirst
'           SKU_RS.Find "PRD_IDY='" & .TextMatrix(A, 5) & "'"
'           If SKU_RS.EOF = False Then
'                SCH_LST_RAT = SKU_RS("LST_RAT")
'                SCH_LST_DIS = SKU_RS("LST_DIS")
'           Else
'                SCH_LST_RAT = 0#
'                SCH_LST_DIS = 0#
'           End If
        If Len(Trim(M_VAN_IDY)) = 0 Then
            REC_SET.AddNew
            REC_SET("inv_idy") = Trim(INC)
            REC_SET("inv_dat") = Trim(M_INV_DAT)
            REC_SET("prd_idy") = .TextMatrix(A, 5)
            REC_SET("RAT_APP") = "B"
            REC_SET("LST_RAT") = GetValue("LST_RAT", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
            REC_SET("LST_DIS") = GetValue("LST_DIS", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
            REC_SET("STK_TYP") = "Free"
            REC_SET("PRD_QTY") = STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12)))
            REC_SET("prd_rat") = Val(.TextMatrix(A, 3))
            REC_SET("txn_typ") = "SSO"
            REC_SET!BAT_IDY = Trim(.TextMatrix(A, 7))
            REC_SET!SUB_UNT = Val(.TextMatrix(A, 12))
            REC_SET!PRD_UNT = Int(STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12))))
            REC_SET!PRD_PCS = Val(.TextMatrix(A, 2)) Mod Val(.TextMatrix(A, 12))
            REC_SET!TOT_PCS = Val(.TextMatrix(A, 2))
            REC_SET!PRD_UOM = Round(Val(.TextMatrix(A, 16)), 2)
            REC_SET!Sch_Idy = Trim(.TextMatrix(A, 8))
            REC_SET("COL_AMT") = Val((.TextMatrix(A, 4)))
            REC_SET!ORD_IDY = M_VAN_IDY
            If Len(Trim(M_VAN_IDY)) = 0 Then
                REC_SET!stk_upd = "Y"
            Else
                REC_SET!stk_upd = "N"
            End If
            REC_SET.Update
            
            'Code Added by PNR for Scheme Free Purchase Value and Sale Value Insertion into LNV
            PFA = 0#
            SFA = 0#
            P_LPP = GetValue("U_PUR_LPP", "BRD WHERE BAT_IDY LIKE '" & Trim(MSS.TextMatrix(A, 7)) & "'")
            S_LPP = GetValue("U_SAL_LPP", "BRD WHERE BAT_IDY LIKE '" & Trim(MSS.TextMatrix(A, 7)) & "'")
            PFA = Val(.TextMatrix(A, 2)) * Val(P_LPP)
            SFA = Val(.TextMatrix(A, 2)) * Val(S_LPP)
            'CON.Execute "insert into lnv values('" & Trim(INC) & "','" & Trim(.TextMatrix(A, 8)) & "','" & Trim(.TextMatrix(A, 7)) & "'," & Val(.TextMatrix(A, 2)) & "," & Val((.TextMatrix(A, 4))) & ",0,'FRE'," & Val(PFA) & "," & Val(SFA) & ")"
            con.Execute "insert into lnv values('" & Trim(INC) & "','" & Trim(.TextMatrix(A, 8)) & "','" & Trim(.TextMatrix(A, 7)) & "'," & Val(.TextMatrix(A, 2)) & "," & Val((.TextMatrix(A, 4))) & ",0,'FRE'," & Val(PFA) & "," & Val(SFA) & "," & "' ',' ',0,' '" & ")"
            If Len(Trim(M_VAN_IDY)) = 0 Then
                con.Execute "update PRD where prd_idy like '" & Trim(.TextMatrix(A, 5)) & "' set CUR_LEV = " & CONQTY(QTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12)))), Trim(.TextMatrix(A, 5))), d
                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(.TextMatrix(A, 5)) & "' set PRD_PCS = PRD_PCS - " & Val(.TextMatrix(A, 2)), d
                con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(.TextMatrix(A, 7)) & "' set prd_qty = " & CONBRDQTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12))), Trim(.TextMatrix(A, 7))), d
            End If
        Else
            If ISC_RS.State = 1 Then ISC_RS.Close
            ISC_RS.Open "SELECT * FROM INV WHERE ORD_IDY LIKE '" & Trim(M_VAN_IDY) & "' AND BAT_IDY LIKE '" & Trim(.TextMatrix(A, 7)) & "' AND SCH_IDY LIKE '" & Trim(.TextMatrix(A, 8)) & "' AND STK_TYP LIKE 'Free'", con, adOpenKeyset, adLockPessimistic
            If ISC_RS.RecordCount > 0 Then ISC_RS.MoveFirst
            ESC_QTY = 0#
            While Not ISC_RS.EOF
                ESC_QTY = Val(ESC_QTY) + Val(ISC_RS("TOT_PCS"))
                ISC_RS.MoveNext
            Wend
            APP_SCH_QTY = 0#
            AVB_SCH_QTY = 0#
            If VSC_RS.State = 1 Then VSC_RS.Close
            VSC_RS.Open "SELECT * FROM VAN WHERE INV_IDY LIKE '" & Trim(M_VAN_IDY) & "' AND BAT_IDY LIKE '" & Trim(.TextMatrix(A, 7)) & "'", con, adOpenKeyset, adLockPessimistic
            If VSC_RS.RecordCount > 0 Then VSC_RS.MoveFirst
            If VSC_RS.EOF = False Then
                 APP_SCH_QTY = Val(.TextMatrix(A, 2))
                 AVB_SCH_QTY = (MTS(Val(VSC_RS("FRE_QTY")), Val(.TextMatrix(A, 12))) - Val(ESC_QTY))
            '    If (MTS(Val(VSC_RS("FRE_QTY")), Val(.TextMatrix(A, 12))) - Val(ESC_QTY)) >= Val(.TextMatrix(A, 2)) Then
                If (MTS(Val(VSC_RS("FRE_QTY")), Val(.TextMatrix(A, 12))) - Val(ESC_QTY)) > 0 And APP_SCH_QTY > 0 Then
                    APP_SCH_QTY = Val(.TextMatrix(A, 2))
                    AVB_SCH_QTY = (MTS(Val(VSC_RS("FRE_QTY")), Val(.TextMatrix(A, 12))) - Val(ESC_QTY))
                    
                    If Val(APP_SCH_QTY) > Val(AVB_SCH_QTY) Then

                        M_SKP_PRD_NME = GetValue("PRD_NME", "PRD WHERE PRD_IDY LIKE '" & Trim(MSS.TextMatrix(A, 5)) & "'")
                        M_SKP_SCH_NME = GetValue("SCH_NME", "SCH WHERE SCH_IDY LIKE '" & Trim(.TextMatrix(A, 8)) & "'")
                        MsgBox ("Skipped Scheme ID: " & Trim(.TextMatrix(A, 8))) & ", SKU Name: " & M_SKP_PRD_NME & ", Sch Qty: " & Trim(str(Val(AVB_SCH_QTY))) & ", Full Free STOCK Not Available", vbInformation, "DAS Version 4.0"
                        F_OMS.WriteLine Trim(M_VAN_IDY) & " , " & Trim(M_INV_DAT) & " , " & Trim(M_INV_IDY) & " , " & Trim(.TextMatrix(A, 8)) & " , " & Trim(M_SKP_SCH_NME) & " , " & Trim(M_SKP_PRD_NME) & " , " & Trim(str(Val(AVB_SCH_QTY))) & " , " & "Available Qty"
                        SKP_CTR = Val(SKP_CTR) + 1
                        ESC_QTY = 0#

''''''                        While Val(AVB_SCH_QTY) = 0
'''''                            REC_SET.AddNew
'''''                            REC_SET("inv_idy") = Trim(INC)
'''''                            REC_SET("inv_dat") = Trim(M_INV_DAT)
'''''                            REC_SET("prd_idy") = .TextMatrix(A, 5)
'''''                            REC_SET("LST_RAT") = GetValue("LST_RAT", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
'''''                            REC_SET("LST_DIS") = GetValue("LST_DIS", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
'''''                            REC_SET("STK_TYP") = "Free"
'''''                            REC_SET("PRD_QTY") = STM(AVB_SCH_QTY, Val(.TextMatrix(A, 12))) 'STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12)))
'''''                            REC_SET("prd_rat") = Val(.TextMatrix(A, 3))
'''''                            REC_SET("txn_typ") = "SSO"
'''''                            REC_SET!BAT_IDY = Trim(.TextMatrix(A, 7))
'''''                            REC_SET!SUB_UNT = Val(.TextMatrix(A, 12))
'''''                            REC_SET!PRD_UNT = STM(Val(AVB_SCH_QTY), Val(.TextMatrix(A, 12))) 'Int(STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12))))
'''''                            REC_SET!PRD_PCS = Val(AVB_SCH_QTY) Mod Val(.TextMatrix(A, 12)) 'Val(.TextMatrix(A, 2)) Mod Val(.TextMatrix(A, 12))
'''''                            REC_SET!TOT_PCS = Val(AVB_SCH_QTY) 'Val(.TextMatrix(A, 2))
'''''                            REC_SET!PRD_UOM = Round(Val(.TextMatrix(A, 16)), 2)
'''''                            REC_SET!Sch_Idy = Trim(.TextMatrix(A, 8))
'''''                            REC_SET("COL_AMT") = Val((.TextMatrix(A, 4)))
'''''                            REC_SET!ORD_IDY = M_VAN_IDY
'''''                            If Len(Trim(M_VAN_IDY)) = 0 Then
'''''                                REC_SET!stk_upd = "Y"
'''''                            Else
'''''                                REC_SET!stk_upd = "N"
'''''                            End If
'''''                            REC_SET.Update
'''''                            CON.Execute "insert into lnv values('" & Trim(INC) & "','" & Trim(.TextMatrix(A, 8)) & "','" & Trim(.TextMatrix(A, 7)) & "'," & Val(.TextMatrix(A, 2)) & "," & Val((.TextMatrix(A, 4))) & ",0,'FRE')"
'''''                            If Len(Trim(M_VAN_IDY)) = 0 Then
'''''                                CON.Execute "update PRD where prd_idy like '" & Trim(.TextMatrix(A, 5)) & "' set CUR_LEV = " & CONQTY(QTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12)))), Trim(.TextMatrix(A, 5))), d
'''''                                CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(.TextMatrix(A, 5)) & "' set PRD_PCS = PRD_PCS - " & Val(.TextMatrix(A, 2)), d
'''''                                CON.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(.TextMatrix(A, 7)) & "' set prd_qty = " & CONBRDQTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12))), Trim(.TextMatrix(A, 7))), d
'''''                            End If
'''''                          '  ESC_QTY = Val(ESC_QTY) + Val(.TextMatrix(A, 2))
'''''                            AVB_SCH_QTY = Val(AVB_SCH_QTY) - Val(.TextMatrix(A, 2))
''''''                        Wend

                    Else
'                        While Val(APP_SCH_QTY) = 0
                            REC_SET.AddNew
                            REC_SET("inv_idy") = Trim(INC)
                            REC_SET("inv_dat") = Trim(M_INV_DAT)
                            REC_SET("prd_idy") = .TextMatrix(A, 5)
                            REC_SET("RAT_APP") = "B"
                            REC_SET("LST_RAT") = GetValue("LST_RAT", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
                            REC_SET("LST_DIS") = GetValue("LST_DIS", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
                            REC_SET("STK_TYP") = "Free"
                            REC_SET("PRD_QTY") = STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12)))
                            REC_SET("prd_rat") = Val(.TextMatrix(A, 3))
                            REC_SET("txn_typ") = "SSO"
                            REC_SET!BAT_IDY = Trim(.TextMatrix(A, 7))
                            REC_SET!SUB_UNT = Val(.TextMatrix(A, 12))
                            REC_SET!PRD_UNT = Int(STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12))))
                            REC_SET!PRD_PCS = Val(.TextMatrix(A, 2)) Mod Val(.TextMatrix(A, 12))
                            REC_SET!TOT_PCS = Val(.TextMatrix(A, 2))
                            REC_SET!PRD_UOM = Round(Val(.TextMatrix(A, 16)), 2)
                            REC_SET!Sch_Idy = Trim(.TextMatrix(A, 8))
                            REC_SET("COL_AMT") = Val((.TextMatrix(A, 4)))
                            REC_SET!ORD_IDY = M_VAN_IDY
                            If Len(Trim(M_VAN_IDY)) = 0 Then
                                REC_SET!stk_upd = "Y"
                            Else
                                REC_SET!stk_upd = "N"
                            End If
                            REC_SET.Update
                            
                            'Code Added by PNR for Scheme Free Purchase Value and Sale Value Insertion into LNV
                            PFA = 0#
                            SFA = 0#
                            P_LPP = GetValue("U_PUR_LPP", "BRD WHERE BAT_IDY LIKE '" & Trim(MSS.TextMatrix(A, 7)) & "'")
                            S_LPP = GetValue("U_SAL_LPP", "BRD WHERE BAT_IDY LIKE '" & Trim(MSS.TextMatrix(A, 7)) & "'")
                            PFA = Val(.TextMatrix(A, 2)) * Val(P_LPP)
                            SFA = Val(.TextMatrix(A, 2)) * Val(S_LPP)
                            
                            con.Execute "insert into lnv values('" & Trim(INC) & "','" & Trim(.TextMatrix(A, 8)) & "','" & Trim(.TextMatrix(A, 7)) & "'," & Val(.TextMatrix(A, 2)) & "," & Val((.TextMatrix(A, 4))) & ",0,'FRE'," & Val(PFA) & "," & Val(SFA) & ")"
                            
                            If Len(Trim(M_VAN_IDY)) = 0 Then
                                con.Execute "update PRD where prd_idy like '" & Trim(.TextMatrix(A, 5)) & "' set CUR_LEV = " & CONQTY(QTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12)))), Trim(.TextMatrix(A, 5))), d
                                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(.TextMatrix(A, 5)) & "' set PRD_PCS = PRD_PCS - " & Val(.TextMatrix(A, 2)), d
                                con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(.TextMatrix(A, 7)) & "' set prd_qty = " & CONBRDQTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12))), Trim(.TextMatrix(A, 7))), d
                            End If
                           ' ESC_QTY = Val(ESC_QTY) + Val(.TextMatrix(A, 2))
                            APP_SCH_QTY = Val(APP_SCH_QTY) - Val(.TextMatrix(A, 2))
'                        Wend
                    End If
                Else
                    M_SKP_PRD_NME = GetValue("PRD_NME", "PRD WHERE PRD_IDY LIKE '" & Trim(MSS.TextMatrix(A, 5)) & "'")
                    M_SKP_SCH_NME = GetValue("SCH_NME", "SCH WHERE SCH_IDY LIKE '" & Trim(.TextMatrix(A, 8)) & "'")
                    MsgBox ("Skipped Scheme ID: " & Trim(.TextMatrix(A, 8))) & ", SKU Name: " & M_SKP_PRD_NME & ", Sch Qty: " & Trim(str(Val(.TextMatrix(A, 2)))) & ", No STOCK", vbInformation, "DAS Version 4.0"
                    F_OMS.WriteLine Trim(M_VAN_IDY) & " , " & Trim(M_INV_DAT) & " , " & Trim(M_INV_IDY) & " , " & Trim(.TextMatrix(A, 8)) & " , " & Trim(M_SKP_SCH_NME) & " , " & Trim(M_SKP_PRD_NME) & " , " & Trim(str(Val(.TextMatrix(A, 2)))) & " , " & "Short Qty"
                    SKP_CTR = Val(SKP_CTR) + 1
                    ESC_QTY = 0#
                End If
            Else
                    M_SKP_PRD_NME = GetValue("PRD_NME", "PRD WHERE PRD_IDY LIKE '" & Trim(MSS.TextMatrix(A, 5)) & "'")
                    M_SKP_SCH_NME = GetValue("SCH_NME", "SCH WHERE SCH_IDY LIKE '" & Trim(.TextMatrix(A, 8)) & "'")
                    MsgBox ("Skipped Scheme ID: " & Trim(.TextMatrix(A, 8))) & ", No SKU  : " & M_SKP_PRD_NME & " in LOAD OUT !", vbInformation, "DAS Version 4.0"
                    F_OMS.WriteLine Trim(M_VAN_IDY) & " , " & Trim(M_INV_DAT) & " , " & Trim(M_INV_IDY) & " , " & Trim(.TextMatrix(A, 8)) & " , " & Trim(M_SKP_SCH_NME) & " , " & Trim(M_SKP_PRD_NME) & " , " & "0.000" & " , " & "No SKU Ref"
                    SKP_CTR = Val(SKP_CTR) + 1
                    ESC_QTY = 0#
            End If
            F_OMS.WriteBlankLines 1
        End If
    End If
Next

F_OMS.Close
If Len(Trim(M_VAN_IDY)) <> 0 And Val(SKP_CTR) > 0 Then
    MsgBox ("Total Scheme Count : " & Val(SKP_CTR) & " SKIPPED without SAVE!, Check SKU & Free Columns in Van Challan : " & Trim(M_VAN_IDY)), vbInformation, "DAS Version 4.0"
End If
End With
'scheme update end
'SCHRECORDS BEGIN
If APPSCH = True Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "SELECT * FROM LNV", con, adOpenKeyset, adLockPessimistic
    For A = 0 To SC.Rows - 1
'        If Val(SC.TextMatrix(A, 2)) > 0 Or Val(SC.TextMatrix(A, 3)) > 0 Then
            REC_SET.AddNew
            REC_SET!INV_IDY = Trim(INC)
            REC_SET!LEV_IDY = Trim(SC.TextMatrix(A, 0)) 'SCHIDY
            REC_SET!LEV_DES = Trim(SC.TextMatrix(A, 1)) 'BATIDY
            REC_SET!INV_GRS = Val(SC.TextMatrix(A, 2))  'FREEQTY
            REC_SET!LNV_AMT = Val(SC.TextMatrix(A, 3))  'LSSAMT         Again enabled by PNR for scheme amount breakup
'            REC_SET!LNV_AMT = Val(SCH_AMT)             'FRT_AMT        This field will put total scheme amounts value ie, wrong.
            REC_SET!TXN_TYP = Trim(SC.TextMatrix(A, 4))
            REC_SET.Update
'        End If
    Next
End If
'SCHRECORD END

'salable pnv update begin
If REC_PNL.State = 1 Then REC_PNL.Close
REC_PNL.Open "select * from pnv where inv_idy =''", PNL, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" And Val(MS.TextMatrix(A, 2)) > 0 And MS.TextMatrix(A, 1) <> "" Then
        REC_PNL.AddNew
        REC_PNL("inv_idy") = Trim(INC)
        REC_PNL("inv_dat") = Trim(M_INV_DAT)
        REC_PNL("prd_idy") = MS.TextMatrix(A, 0)
        REC_PNL("grp_idy") = Trim(MS.TextMatrix(A, 18))
        REC_PNL("grp_cat") = Trim(MS.TextMatrix(A, 19))
        REC_PNL("grp_cls") = Trim(MS.TextMatrix(A, 20))
        REC_PNL("RAT_APP") = Trim(MS.TextMatrix(A, 4))
        If pcs_chk.Value = 1 Then
            REC_PNL("prd_qty") = Val(MS.TextMatrix(A, 2))
            REC_PNL("prd_rat") = Val((Val(MS.TextMatrix(A, 5))))
            REC_PNL!MRP_RAT = Val(MS.TextMatrix(A, 10))
            LST_TOT = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 13))))
            M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
            'LST_TOT = Val(MS.TextMatrix(A, 2)) * Val(MS.TextMatrix(A, 22))
            'M_CPN_AMT = Val(MS.TextMatrix(A, 2)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        Else
            REC_PNL("PRD_QTY") = STM(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            REC_PNL("prd_rat") = Val(Val(MS.TextMatrix(A, 5))) * Val(MS.TextMatrix(A, 13))
            REC_PNL!MRP_RAT = Val(MS.TextMatrix(A, 10)) * Val(MS.TextMatrix(A, 13))
            LST_TOT = Val(MS.TextMatrix(A, 2))
            LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 13))))
            M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
            'LST_TOT = (Val(MS.TextMatrix(A, 2)) / Val(MS.TextMatrix(A, 13))) * Val(MS.TextMatrix(A, 22))
            'M_CPN_AMT = (Val(MS.TextMatrix(A, 2)) / Val(MS.TextMatrix(A, 13))) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        End If
        REC_PNL("ADD_AMT") = Val((Val(MS.TextMatrix(A, 6))))
        REC_PNL("trd_dis") = Val(MS.TextMatrix(A, 7))
        REC_PNL("trd_tax") = Val(MS.TextMatrix(A, 8))
        REC_PNL("CPN_AMT") = Round(M_CPN_AMT, 2)
        If pcs_chk.Value = 1 Then
            INCAMT = RATCAL(Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 24) & ""), Val(MS.TextMatrix(A, 13)))
            PRDNET = RATCAL(Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 5) & ""), Val(MS.TextMatrix(A, 13)))
        Else
            PRDNET = Val(MS.TextMatrix(A, 2) & "") * Val(MS.TextMatrix(A, 24) & "")
            PRDNET = Val(MS.TextMatrix(A, 2) & "") * Val(MS.TextMatrix(A, 5) & "")
        End If
        REC_PNL("DIS_AMT") = ((PRDNET - REC_PNL("ADD_AMT")) * Val(MS.TextMatrix(A, 7)) / 100)
        If Trim(GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'DPV'")) = "T" Then
            TRDTAX = (PRDNET - (Val(INCAMT))) * Val(MS.TextMatrix(A, 8)) / 100
        Else
            TRDTAX = (PRDNET - (Val(REC_PNL("ADD_AMT")) + Val(REC_PNL("DIS_AMT")) + Val(INCAMT))) * Val(MS.TextMatrix(A, 8)) / 100
        End If
'        TRDTAX = (PRDNET - (Val(REC_PNL("ADD_AMT")) + Val(REC_PNL("DIS_AMT")) + Val(INCAMT))) * Val(MS.TextMatrix(A, 8)) / 100
        REC_PNL("TAX_AMT") = (Val(TRDTAX & "") + Round(M_CPN_AMT, 2))
        REC_PNL("txn_typ") = "SSO"
        REC_PNL!BAT_IDY = Trim(MS.TextMatrix(A, 15))
        REC_PNL!PRD_UOM = Round(Val((MS.TextMatrix(A, 3))), 2)
        REC_PNL!SUB_UNT = Val(MS.TextMatrix(A, 13))
        If pcs_chk.Value = 1 Then
            REC_PNL!PRD_UNT = Int(Val(MS.TextMatrix(A, 2)))
            REC_PNL!PRD_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13))) Mod Val(MS.TextMatrix(A, 13))
            REC_PNL!TOT_PCS = P_SUB_UNT(QTY(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
        Else
            REC_PNL!PRD_UNT = Int(Val(MS.TextMatrix(A, 2) / MS.TextMatrix(A, 13)))
            REC_PNL!PRD_PCS = Val(MS.TextMatrix(A, 2)) Mod Val(MS.TextMatrix(A, 13))
            REC_PNL!TOT_PCS = Val(MS.TextMatrix(A, 2))
        End If
        REC_PNL("COL_AMT") = Val(MS.TextMatrix(A, 9) + Round(LST_TOT, 2))
        REC_PNL("INC_TTL") = Val(INCAMT)
        REC_PNL!stk_upd = "Y"
        REC_PNL("prd_nme") = Trim(MS.TextMatrix(A, 11))
        REC_PNL("ret_qty") = 0
        REC_PNL("iss_qty") = 0
        REC_PNL("tot_qty") = 0
        REC_PNL("rat_app") = Trim(MS.TextMatrix(A, 4)) 'ADDED BY PNR
        REC_PNL("cpn_rat") = 0
        REC_PNL("add_dis") = 0
        REC_PNL("fst_flg") = ""
        REC_PNL("opn_qty") = 0
        REC_PNL("sch_amt") = 0
        REC_PNL("ord_idy") = 0
        REC_PNL("bat_nme") = Trim(MS.TextMatrix(A, 21))
        REC_PNL("stk_typ") = "Salable"
        REC_PNL("sch_idy") = ""
        REC_PNL("LST_RAT") = Val(MS.TextMatrix(A, 22))
        REC_PNL("LST_DIS") = Val(MS.TextMatrix(A, 23))
        REC_PNL("CPN_AMT") = Round(M_CPN_AMT, 2)
        'Val (MS.TextMatrix(A, 2)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        REC_PNL.Update
    End If
Next
'pnv salable update close

'free pnv update close
'If REC_PNL.State = 1 Then REC_PNL.Close
'REC_PNL.Open "select * from pnv", PNL, adOpenKeyset, adlockpessimistic
For A = 1 To MSF.Rows - 1
    If Trim(MSF.TextMatrix(A, 0)) <> "" And Val(MSF.TextMatrix(A, 2)) > 0 And MSF.TextMatrix(A, 1) <> "" Then
        REC_PNL.AddNew
        REC_PNL("inv_idy") = Trim(INC)
        REC_PNL("inv_dat") = Trim(M_INV_DAT)
        REC_PNL("prd_idy") = MSF.TextMatrix(A, 0)
        REC_PNL("grp_idy") = Trim(MSF.TextMatrix(A, 12))
        REC_PNL("grp_cat") = Trim(MSF.TextMatrix(A, 13))
        REC_PNL("grp_cls") = Trim(MSF.TextMatrix(A, 14))
        REC_PNL("RAT_APP") = "B"
        If pcs_chk.Value = 1 Then
            REC_PNL("prd_qty") = Val(MSF.TextMatrix(A, 2))
            REC_PNL("prd_rat") = Val((Val(MSF.TextMatrix(A, 4))))
            REC_PNL!MRP_RAT = Val(MSF.TextMatrix(A, 6))
        Else
            REC_PNL("PRD_QTY") = STM(Val(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9)))
            REC_PNL("prd_rat") = Val(Val(MSF.TextMatrix(A, 4))) * Val(MSF.TextMatrix(A, 9))
            REC_PNL!MRP_RAT = Val(MSF.TextMatrix(A, 6)) * Val(MSF.TextMatrix(A, 9))
        End If
        REC_PNL("ADD_AMT") = 0
        REC_PNL("trd_dis") = 0
        REC_PNL("trd_tax") = 0
        If pcs_chk.Value = 1 Then
            PRDNET = RATCAL(Val(MSF.TextMatrix(A, 2) & ""), Val(MSF.TextMatrix(A, 5) & ""), Val(MSF.TextMatrix(A, 9)))
        Else
            PRDNET = Val(MSF.TextMatrix(A, 2) & "") * Val(MSF.TextMatrix(A, 2) & "")
        End If
        REC_PNL("DIS_AMT") = 0
        REC_PNL("TAX_AMT") = 0
        REC_PNL("txn_typ") = "SSO"
        REC_PNL!BAT_IDY = Trim(MSF.TextMatrix(A, 11))
        REC_PNL!PRD_UOM = Round(Val((MSF.TextMatrix(A, 3))), 2)
        REC_PNL!SUB_UNT = Val(MSF.TextMatrix(A, 9))
        If pcs_chk.Value = 1 Then
            REC_PNL!PRD_UNT = Int(Val(MSF.TextMatrix(A, 2)))
            REC_PNL!PRD_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9))) Mod Val(MSF.TextMatrix(A, 9))
            REC_PNL!TOT_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 2)), Val(MSF.TextMatrix(A, 9)))
        Else
            REC_PNL!PRD_UNT = Int(Val(MSF.TextMatrix(A, 2) / MSF.TextMatrix(A, 9)))
            REC_PNL!PRD_PCS = Val(MSF.TextMatrix(A, 2)) Mod Val(MSF.TextMatrix(A, 9))
            REC_PNL!TOT_PCS = Val(MSF.TextMatrix(A, 2))
        End If
        REC_PNL("COL_AMT") = Val(MSF.TextMatrix(A, 5) + Round(LST_TOT, 2))
        REC_PNL!stk_upd = "Y"
        REC_PNL("prd_nme") = Trim(MSF.TextMatrix(A, 7))
        REC_PNL("lst_rat") = Val(MSF.TextMatrix(A, 17))
        REC_PNL("lst_dis") = Val(MSF.TextMatrix(A, 18))
        REC_PNL("ret_qty") = 0
        REC_PNL("iss_qty") = 0
        REC_PNL("tot_qty") = 0
'        REC_PNL("rat_app") = Trim(MSF.TextMatrix(A, 4)) 'ADDED BY PNR
        REC_PNL("RAT_APP") = "B"
        REC_PNL("cpn_rat") = 0
        REC_PNL("cpn_amt") = 0
        REC_PNL("add_dis") = 0
        REC_PNL("fst_flg") = ""
        REC_PNL("opn_qty") = 0
        REC_PNL("sch_amt") = 0
        REC_PNL("ord_idy") = 0
        REC_PNL("bat_nme") = Trim(MSF.TextMatrix(A, 15))
        REC_PNL("stk_typ") = Trim(MSF.TextMatrix(A, 1))
        REC_PNL("sch_idy") = ""
        REC_PNL.Update
    End If
Next
'FREE UPDATE END

'schemes pnv update close
'If REC_PNL.State = 1 Then REC_PNL.Close
'REC_PNL.Open "select * from pnv", PNL, adOpenKeyset, adlockpessimistic
For A = 1 To MSS.Rows - 1
    If Trim(MSS.TextMatrix(A, 0)) <> "" And (Val(MSS.TextMatrix(A, 2)) > 0 Or Val(MSS.TextMatrix(A, 4)) > 0) And APPSCH = True Then
        REC_PNL.AddNew
        REC_PNL("inv_idy") = Trim(INC)
        REC_PNL("inv_dat") = Trim(M_INV_DAT)
        REC_PNL("prd_idy") = MSS.TextMatrix(A, 5)
        REC_PNL("grp_idy") = Trim(MSS.TextMatrix(A, 13))
        REC_PNL("grp_cat") = GetValue("grp_cat", "pgr where grp_idy like '" & Trim(REC_PNL("grp_idy")) & "'")
        REC_PNL("grp_cls") = GetValue("grp_cls", "pgr where grp_idy like '" & Trim(REC_PNL("grp_idy")) & "'")
        REC_PNL("RAT_APP") = "B" 'ADDED BY PNR
        REC_PNL("PRD_QTY") = STM(Val(MSS.TextMatrix(A, 2)), Val(MSS.TextMatrix(A, 12)))
        REC_PNL("prd_rat") = Val(Val(MSS.TextMatrix(A, 3)))
        REC_PNL!MRP_RAT = Val(MSS.TextMatrix(A, 1))
        REC_PNL("ADD_AMT") = 0
        REC_PNL("trd_dis") = 0
        REC_PNL("trd_tax") = 0
        REC_PNL("DIS_AMT") = 0
        REC_PNL("TAX_AMT") = 0
        REC_PNL("txn_typ") = "SSO"
        REC_PNL!BAT_IDY = Trim(MSS.TextMatrix(A, 7))
        REC_PNL!PRD_UOM = Round(Val(MSS.TextMatrix(A, 16)), 2)
        REC_PNL!SUB_UNT = Val(MSS.TextMatrix(A, 12))
        REC_PNL!PRD_UNT = Int(Val(MSS.TextMatrix(A, 2) / MSS.TextMatrix(A, 12)))
        REC_PNL!PRD_PCS = Val(MSS.TextMatrix(A, 2)) Mod Val(MSS.TextMatrix(A, 12))
        REC_PNL!TOT_PCS = Val(MSS.TextMatrix(A, 2))
        REC_PNL("COL_AMT") = Val(MSS.TextMatrix(A, 4))
        REC_PNL!stk_upd = "Y"
        REC_PNL("prd_nme") = Trim(MSS.TextMatrix(A, 0))
        REC_PNL("lst_rat") = GetValue("LST_RAT", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
        REC_PNL("lst_dis") = GetValue("LST_DIS", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
        REC_PNL("ret_qty") = 0
        REC_PNL("iss_qty") = 0
        REC_PNL("tot_qty") = 0
        REC_PNL("rat_app") = "B" 'ADDED BY PNR
        REC_PNL("cpn_rat") = 0
        REC_PNL("cpn_amt") = 0
        REC_PNL("add_dis") = 0
        REC_PNL("fst_flg") = ""
        REC_PNL("opn_qty") = 0
        REC_PNL("sch_amt") = 0
        REC_PNL("ord_idy") = 0
        REC_PNL("bat_nme") = Trim(MSS.TextMatrix(A, 6))
        REC_PNL("stk_typ") = "Free"
        REC_PNL("sch_idy") = Trim(MSS.TextMatrix(A, 8))
        REC_PNL("INC_TTL") = 0# 'ADDED THIS LINE OF CODE AFTER ODBC ERROR REPORTED BY PRINCE CHOPRA
        REC_PNL.Update
    End If
Next
If Len(M_VAN_IDY) = 0 Then
    'AddShells Trim(INC), "SSO", M_CST_IDY, "Y", M_SMN_IDY, M_ROU_IDY
Else
    'AddShells Trim(INC), "SSO", M_CST_IDY, "N", M_SMN_IDY, M_ROU_IDY
End If
'Schemes UPDATE END
If Toolbar1.Buttons(11).Value = tbrPressed Then
If MOD_REC = False Then
MsgBox "New Sale Order ID is " & INC, vbInformation, "DAS Version 4.0"
End If
End If
Label23 = labdisp("IDY_SSO")
If check("PRT", "PRT_NME", "SAL_SHL") Then LOADSHELLS "SSO", Trim(INC)
If check("PRT", "PRT_NME", "M_FRM_SOR") = True Then
    con.Execute "update dir where dir_idy like 'INV' AND DIR_IDX LIKE 'NUM' set dir_als = '" & INC & "'"
    con.Execute "update dir where dir_idy like 'DOC' AND DIR_IDX LIKE 'TYP' set dir_als = 'SSO'"
    If querycheck("DIR", "DIR_MSG", "D", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
        If REC_PNL.State = 1 Then REC_PNL.Close
        REC_PNL.Open "SELECT * FROM RRI", PNL, adOpenKeyset, adLockPessimistic
        If REC_PNL.RecordCount > 0 Then
            REC_PNL.Move 6
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
        'Shell CURDIR & "\RRUNTIME RRI 7", vbHide
        If Dir(CURDIR & "\pri.exe") = "" Then
            MsgBox "PRI.EXE not found", vbCritical, "DAS Version 4.0"
        Else
            Shell CURDIR & "\pri.exe"
        End If
    Else
        OpenDocument CURDIR & "\DOS_VCH.RP1", "Sale Order", 2, Trim(INC), Trim(INC), 6, CURDIR & "\pnv.cdx,c,inv_idy"
    End If
    CANCELRECORD
End If
REC_SET.Close
'RSINL.Close
REC_SET1.Close
'REC_IRT.Close
REC_PNL.Close
'SKU_RS.Close
'VSC_RS.Close
'ISC_RS.Close

Set REC_SET = Nothing
Set RSINL = Nothing
Set REC_SET1 = Nothing
Set REC_IRT = Nothing
Set REC_PNL = Nothing
Set SKU_RS = Nothing
Set VSC_RS = Nothing
Set ISC_RS = Nothing

CANCELRECORD
Exit Sub
LBL:
MsgBox err.Description, vbError
CANCELRECORD
End Sub

Private Sub CANCELRECORD()
m_smn = Text2
clrctr M_FRM_SOR
CHGQTY = ""
Text2 = m_smn
setval
MS.Rows = 1
MS.AddItem ""
MSF.Rows = 1
MSF.AddItem ""
SSTab1.Tab = 0
APPLYCSR = False
SC.Rows = 0
M_VAN_IDY.Enabled = True
M_CRT_DAY.Locked = False
M_AMT_LES.Locked = False
M_FRT_AMT.Locked = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
LabelAAS.Visible = False
SYS_NUM = False
pcs_chk.Enabled = True
MOD_REC = False
List7 = "Y"
M_INV_IDY.Enabled = True
M_INV_IDY.SetFocus
End Sub

Private Sub newrecord()
CANCELRECORD
If Toolbar1.Buttons(11).Value = tbrPressed Then
    M_INV_IDY.Enabled = False
    M_INV_DAT.SetFocus
Else
    M_INV_IDY.SetFocus
End If
SSTab1.Tab = 0
If Toolbar1.Buttons(19).Value = tbrPressed Then M_VAN_IDY = VAN_IDY
End Sub

Private Sub V_DLR_LST()
If CHECKTB("act.dbf,mar.dbf,smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If List9 <> "All" Then
    qry = " and RDY.rou_idy like '" & Trim(List9) & "%'" & "GROUP BY ACT.ACT_IDY"
Else
    qry = " and RDY.rou_idy like '%'" & "GROUP BY ACT.ACT_IDY"
End If
If Check1.Value = 1 Then
    If Toolbar1.Buttons(31).Value = tbrPressed And Toolbar1.Buttons(30).Value = 1 Then
        'REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SSO') and act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND MAR_VDA LIKE '" & UCase(text1) & "' AND SMN.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'" & (qry), CON, adOpenKeyset, adlockpessimistic
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,rdy where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND rdy.SMN_IDY LIKE SMN.SMN_IDY AND day_idy = " & Text1.ListIndex + 1 & " and cst_typ not in('Van') AND rdy.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    ElseIf Toolbar1.Buttons(31).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,rdy where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_NME like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND day_idy = " & Text1.ListIndex + 1 & " and cst_typ not in('Van') AND DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    ElseIf Toolbar1.Buttons(30).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,RDY where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_NME like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND rdy.SMN_IDY LIKE '" & Trim(Text2) & "' and cst_typ not in('Van') AND DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    Else
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,RDY where (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND cst_typ not in('Van') and DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    End If
ElseIf Check2.Value = 1 Then
    If Toolbar1.Buttons(31).Value = tbrPressed And Toolbar1.Buttons(30).Value = 1 Then
        'REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND MAR_VDA LIKE '" & UCase(text1) & "' AND SMN.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'" & (qry), CON, adOpenKeyset, adlockpessimistic
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,rdy where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND rdy.SMN_IDY LIKE SMN.SMN_IDY AND day_idy = " & Text1.ListIndex + 1 & " AND rdy.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    ElseIf Toolbar1.Buttons(31).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,rdy where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND day_idy = " & Text1.ListIndex + 1 & " AND DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    ElseIf Toolbar1.Buttons(30).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,RDY where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND rdy.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    Else
        REC_SET.Open "select ACT.*,MAR.MAR_NME,RDY.ROU_IDY ROUIDY,smn.smn_nme from act,MAR,smn,RDY where (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND RDY.SMN_IDY LIKE SMN.SMN_IDY AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND DBT_TYP='Y'" & (qry), con, adOpenKeyset, adLockPessimistic
    End If
End If
list3.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list3.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("MAR_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ROUIDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("SMN_NME") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    V_LST_CAP = "Select Retailer"
    list3.Visible = True
    list3.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Retailer Not Found, Define Retailer !", vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub

Private Sub V_SMN_LST()
If CHECKTB("smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If V_LST_CAP.Caption = "Select Salesman" Then
    If Check1.Value = 1 Then
        REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
    ElseIf Check2.Value = 1 Then
        REC_SET.Open "select * from SMN WHERE SMN_AD4 like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
    End If
Else
    REC_SET.Open "select * from SMN", con, adOpenKeyset, adLockPessimistic
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
    MsgBox "Salesman Not Found, Define Salesman !", vbInformation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_SMN_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub

Private Sub V_ACT_LST()
If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where grp_clf like 'Sales ' and ACT_NME LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from act where grp_clf like 'Sales ' and ACT_AD4 LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
End If
List2.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    V_LST_CAP = "Select Account"
    List2.Visible = True
    List2.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Sales Account Not Found, Define Sales Account!", vbCritical, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_ACT_IDY.SetFocus
End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub

Private Sub EXT_REC()
Set Form = Nothing
Unload Me
End Sub
Private Sub Command1_Click()
M_HVO_CHK = GetValue("CST_TYP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'")
M_HVO_FLG = GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'AOM'")
If Trim(M_HVO_CHK) = "HVO-P" Or Trim(M_HVO_CHK) = "HVO-NP" Or Trim(M_HVO_CHK) = "HVO-P-WO" Or Trim(M_HVO_CHK) = "HVO-NP-WO" Then
    If Trim(M_HVO_CHK) = "T" Then
        MsgBox "Global Restrictions Active on Open Market Schemes for this HVO Account!", vbInformation, "DAS Version 4.0"
        SSTab1.TabEnabled(6) = False
        SSTab1.Tab = 2
        M_ROU_OFF.SetFocus
        Exit Sub
    Else
        If Trim(M_HVO_CHK) = "HVO-P-WO" Or Trim(M_HVO_CHK) = "HVO-NP-WO" Then
            MsgBox "This HVO Account is Restricted for Open Market Schemes!", vbInformation, "DAS Version 4.0"
            SSTab1.TabEnabled(6) = False
            SSTab1.Tab = 2
            M_ROU_OFF.SetFocus
            Exit Sub
        Else
            SSTab1.TabEnabled(6) = True
            SSTab1.Tab = 6
            Command6.SetFocus
        End If
    End If
Else
    SSTab1.TabEnabled(6) = True
    SSTab1.Tab = 6
    Command6.SetFocus
End If
End Sub
Private Sub Command2_Click()
SSTab1.Tab = 5
MSF.SetFocus
End Sub
Private Sub Command3_Click()
If Command3.Caption = "Select A&ll" Then
    For A = 1 To SCHLST.ListItems.count
        SCHLST.ListItems.Item(A).Checked = True
    Next
    Command3.Caption = "De-Select A&ll"
Else
    For A = 1 To SCHLST.ListItems.count
        SCHLST.ListItems.Item(A).Checked = False
    Next
    Command3.Caption = "Select A&ll"
End If
SCHLST.SetFocus
End Sub
Private Sub Command4_Click()
Dim rssch As New ADODB.Recordset
Dim ls As ListItem
MSS.Rows = 1
SC.Rows = 0
SCH_AMT = "0.00"
APPSCH = False
For A = 1 To SCHLST.ListItems.count
    If SCHLST.ListItems.Item(A).Checked = True Then
        Set ls = SCHLST.ListItems.Item(A)
        If ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "SKU Level" Then
            ADDSKUQUANTIY Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Bill Level" Then
            ADDBILLQUANTIY Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Multiple SKU Level" Then
            ADDMLTQTYAND Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(3) = "SKU Level" And ls.ListSubItems(4) = "Cash Off" Then
            ADDSKUAMOUNT Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(3) = "SKU Level" And ls.ListSubItems(4) = "Discount Percent" Then
            ADDSKUPCG Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Cash Off" And ls.ListSubItems(3) = "Bill Level" Then
            ADDBILAMT Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Discount Percent" And ls.ListSubItems(3) = "Bill Level" Then
            ADDBILPCG Trim(ls.ListSubItems(5))
        ElseIf (ls.ListSubItems(4) = "Cash Off" Or ls.ListSubItems(4) = "Discount Percent") And ls.ListSubItems(3) = "Multiple SKU Level" Then
            MLTBILPCGAMT Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Brand" Then
            AddBrand Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Pack Size" Then
            AddPackSize Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Pack Type" Then
            AddPackType Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Segment" Then
            AddSegment Trim(ls.ListSubItems(5))
        ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Category" Then
            AddCategory Trim(ls.ListSubItems(5))
        ElseIf (ls.ListSubItems(4) = "Cash Off" Or ls.ListSubItems(4) = "Discount Percent") And (ls.ListSubItems(3) = "Pack Size" Or ls.ListSubItems(3) = "Pack Type" Or ls.ListSubItems(3) = "Brand" Or ls.ListSubItems(3) = "Category" Or ls.ListSubItems(3) = "Segment") Then
            AddBillAmtGroupType Trim(ls.ListSubItems(5))
        End If
    End If
Next
MSS.SetFocus
End Sub
Private Sub Command5_Click()
APPSCH = True
If Val(SCH_AMT) > 0 Then
    M_AMT_LES = "Scheme Amount"
    M_FRT_AMT = SCH_AMT
    M_AMT_LES.Locked = True
    M_FRT_AMT.Locked = True
End If
GRS_CAL
SSTab1.Tab = 2
M_ROU_OFF.SetFocus
End Sub
Private Sub Command6_Click()
'''Dim HVO_REC_SET As New ADODB.Recordset
'''If HVO_REC_SET.State = 1 Then HVO_REC_SET.Close
'''HVO_REC_SET.Open "SELECT * FROM ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "' AND (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') AND CST_TYP NOT IN ('Van') AND DBT_TYP='Y'", CON, adOpenKeyset, adlockpessimistic
'''If Trim(HVO_REC_SET!CST_TYP) = "HVO-P" Or Trim(HVO_REC_SET!CST_TYP) = "HVO-NP" Then
'''    MsgBox "This is A High Volume Outlet (HVO) - Open Market Schemes are Not Applicable!", vbCritical, "DAS Version 4.0"
'''    GRS_CAL
'''    SSTab1.Tab = 2
'''    M_ROU_OFF.SetFocus
'''    Exit Sub
'''End If
Dim ls As ListItem
MSS.Rows = 1
SC.Rows = 0
SCH_AMT = 0
Command3.Caption = "Select A&ll"
VLD_FRE_QTY
VLD_BIL_AMT
VLD_MLT_QTY
VLD_MLT_QTY_OR
VLD_MLT_QTY_OR_MIX
VLD_BRAND
VLD_PackSize
VLD_PackType
VLD_Segment
VLD_Category
'VLD_MLT_UOM
'VLD_MLT_UOM_OR
If SCHLST.ListItems.count > 0 Then
    SCHLST.SetFocus
End If
End Sub
Private Sub Command7_Click()
M_FRM_SIN.Show 1
CANCELRECORD
End Sub
Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If DateDiff("d", Format(DTPicker1, "dd/MM/yyyy"), Format(MS.Text, "dd/MM/yyyy")) <> 0 Then
        DTPicker1.Visible = False
        createbatch
    Else
        DTPicker1.Visible = False
        MS.SetFocus
    End If
End If
End Sub
Private Sub JUST_MADE_Click()
Dim INL_JMD As New ADODB.Recordset
If INL_JMD.State = 1 Then INL_JMD.Close
INL_JMD.Open "SELECT * FROM INL WHERE TXN_TYP = 'SSO'", con, adOpenKeyset, adLockPessimistic
If INL_JMD.RecordCount > 0 Then
    INL_JMD.MoveLast
    CANCELRECORD
    M_INV_IDY = Trim(INL_JMD("INV_IDY"))
Else
    CANCELRECORD
End If
INL_JMD.Close
End Sub
Private Sub List7_keypress(K As Integer)
If K = 13 Then
    For A = 1 To MSF.Rows - 1
        If (MSF.TextMatrix(MSF.Row, 11) = MSF.TextMatrix(A, 11)) And Trim(List7) = MSF.TextMatrix(A, 1) Then
            MsgBox "Product - SKU Already Exists with this SKU Type!", vbExclamation, "DAS Version 4.0"
            List7.Visible = False
            MSF.Text = ""
            MSF.SetFocus
            MSF.Row = A
            Exit Sub
        End If
    Next
    If List7 = "Free" Or List7 = "Replc." Then
        MSF.Text = List7
        MSF.TextMatrix(MSF.Row, 4) = "0.00"
        If pcs_chk.Value = 1 Then
            MSF.TextMatrix(MSF.Row, 2) = "0.000"
        Else
            MSF.TextMatrix(MSF.Row, 2) = "0"
        End If
    ElseIf List7 = "TakeBack" Or List7 = "Damage" Or List7 = "Shortage" Or List7 = "Breakage" Or List7 = "Leakage" Then
        MSF.Text = List7
        ' Changed by PKB/PNR for TakeBack Tax Calculation
        Dim Prdtax As New ADODB.Recordset
        Dim Tkbtax As Integer
        If Prdtax.State = 1 Then Prdtax.Close
        Prdtax.Open "select * from PRD where PRD_IDY like '" & MSF.TextMatrix(MSF.Row, 0) & "'", con, adOpenKeyset, adLockPessimistic
        Tkbtax = Val(Trim(Prdtax("SAL_TAX")))
        ' Changed by PKB/PNR for TakeBack Tax Calculation
        If pcs_chk.Value = 1 Then
            MSF.TextMatrix(MSF.Row, 2) = "0.000"
            MSF.TextMatrix(MSF.Row, 4) = convert(Val(GetValue("prd_sdr", "brd where bat_idy like '" & Trim(MSF.TextMatrix(MSF.Row, 11)) & "'")))
            If APPLYCSR = True Then MSF.TextMatrix(MSF.Row, 4) = convert(Val(GetValue("DIS_AMT", "csr where B_BAT_IDY like '" & Trim(MSF.TextMatrix(MSF.Row, 11)) & "' and act_idy like '" & Trim(M_CST_IDY) & "'")))
        Else
            MSF.TextMatrix(MSF.Row, 2) = "0"
            MSF.TextMatrix(MSF.Row, 4) = convert(Val(GetValue("prd_sdr", "brd where bat_idy like '" & Trim(MSF.TextMatrix(MSF.Row, 11)) & "'")) / Val(MSF.TextMatrix(MSF.Row, 9)))
            If APPLYCSR = True Then MSF.TextMatrix(MSF.Row, 4) = convert(Val(GetValue("DIS_AMT", "csr where B_BAT_IDY like '" & Trim(MSF.TextMatrix(MSF.Row, 11)) & "' and act_idy like '" & Trim(M_CST_IDY) & "'")) / Val(MSF.TextMatrix(MSF.Row, 9)))
        End If
'        MSF.TextMatrix(MSF.Row, 4) = convert(Val(MSF.TextMatrix(MSF.Row, 4)) - (MSF.TextMatrix(MSF.Row, 4) * 2)) 'Changed by PNR
        Dim Chg_Cal As Integer
        Chg_Cal = convert(Val(MSF.TextMatrix(MSF.Row, 4)) - (MSF.TextMatrix(MSF.Row, 4) * 2))
        'MSF.TextMatrix(MSF.Row, 4) = MSF.TextMatrix(MSF.Row, 4) - ((MSF.TextMatrix(MSF.Row, 4) * Tkbtaxpcg) / 100) 'Changed by PNR
        MSF.TextMatrix(MSF.Row, 4) = convert(MSF.TextMatrix(MSF.Row, 4) + ((MSF.TextMatrix(MSF.Row, 4) * Tkbtax) / 100)) 'Changed by PKB/PNR Once Again
    End If
    List7.Visible = False
    MSF.Col = 2
    MSF.SetFocus
End If
End Sub
Private Sub LIST8_DblClick()
LIST8_KeyPress 13
End Sub
Private Sub LIST8_KeyPress(K As Integer)
If K = 13 Then
    M_VAN_IDY = Mid(LIST8.SelectedItem, 4, 7)
    m_van_idy_keypress 13
    M_VAN_IDY.Enabled = False
    LIST8.Visible = False
    SSTab1.Tab = 0
    M_CST_IDY.SetFocus
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
Private Sub M_DUE_DAT_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
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
Private Sub M_ROU_IDY_Change()
If Len(Trim(M_ROU_IDY.Text)) = 15 Then SLT
End Sub
Private Sub M_ROU_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select or Type Route Name"
End Sub
Private Sub M_ROU_IDY_keydown(K As Integer, s As Integer)
If K = 13 Then
    M_CSD_PCG.SetFocus
ElseIf K = 38 And s = 1 Then
    M_CRT_DAY.SetFocus
End If
End Sub
Private Sub M_ROU_OFF_Validate(cancel As Boolean)
If Val(M_ROU_OFF) >= 1 Then
    M_ROU_OFF = "0.00"
Else
    M_ROU_OFF = amt(M_ROU_OFF)
End If
End Sub

Private Sub M_SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Salesman"
End Sub

Private Sub M_SMN_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim RSSMN As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 38 Then
    RSSMN.Open "select * from smn where smn_idy like '" & Sincrement(M_SMN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If RSSMN.RecordCount > 0 Then
        M_SMN_IDY = Sincrement(M_SMN_IDY)
        M_SMN_NME = Trim(RSSMN!SMN_NME)
        If MOD_REC = False Then
            If M_VAN_IDY.Visible = False Then
                If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AYR'") = True Then
                    rid = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
                Else
                    rid = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
                End If
                For K = 0 To M_ROU_IDY.ListCount - 1
                    If Trim(M_ROU_IDY.List(K)) = Trim(rid) Then
                        M_ROU_IDY.ListIndex = K
                    End If
                Next
            End If
        End If
        If KeyCode = 38 Then
            M_CST_IDY.SetFocus
        Else
            SSTab1.Tab = 1
            MS.SetFocus
            'M_ACT_IDY.SetFocus
        End If
    Else
        V_LST_CAP = "Select Salesman"
        V_SMN_LST
    End If
End If
End Sub

Private Sub M_SMN_IDY_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
KeyAscii = CHGCASE(KeyAscii)
End Sub

Private Sub m_van_idy_keypress(KeyAscii As Integer)
If CHECKTB("van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii = 13 Then
Dim rs As New ADODB.Recordset
    If Len(Trim(M_VAN_IDY)) = 0 Then
        M_VAN_LST
    Else
        If querycheck("van", "inv_idy", M_VAN_IDY, "AND VAN_CLS='S' AND VAN_CHK='V'") = True Then
            If MOD_REC = False Then
                rs.Open "select * from van where inv_idy = '" & Trim(M_VAN_IDY) & "' AND VAN_CLS='S' AND VAN_CHK='V'", con, adOpenKeyset, adLockPessimistic
                If rs.RecordCount > 0 Then
                    If IsNull(rs!LIN_DAT) = False Then
                        M_INV_DAT = rs!LIN_DAT
                    Else
                        M_INV_DAT = rs!INV_DAT
                    End If
                    M_SMN_IDY = rs!SMN_IDY & ""
                    M_SMN_NME = GetValue("SMN_NME", "SMN WHERE SMN_IDY = '" & Trim(M_SMN_IDY) & "'")
                    If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBR'") = True Then
                        M_ROU_IDY = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
                    Else
                        M_ROU_IDY = Trim(rs!rou_idy & "")
                    End If
                End If
                M_CST_IDY.SetFocus
                Exit Sub
            End If
        End If

        rs.Open "select * from van where inv_idy = 'RSI" & Tincrement(M_VAN_IDY) & "' AND VAN_CLS='S' AND VAN_CHK='V'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            M_VAN_IDY = "RSI" & Tincrement(M_VAN_IDY)
            M_VAN_IDY.Enabled = False
            If MOD_REC = False Then
                If IsNull(rs!LIN_DAT) = False Then
                    M_INV_DAT = rs!LIN_DAT
                Else
                    M_INV_DAT = rs!INV_DAT
                End If
                M_SMN_IDY = rs!SMN_IDY & ""
                M_SMN_NME = GetValue("SMN_NME", "SMN WHERE SMN_IDY = '" & Trim(M_SMN_IDY) & "'")
                If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBR'") = True Then
                    M_ROU_IDY = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
                Else
                    M_ROU_IDY = Trim(rs!rou_idy & "")
                End If
            End If
            M_CST_IDY.SetFocus
        Else
            M_VAN_LST
        End If
    End If
End If
End Sub

Private Sub M_WBL_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
End Sub

Private Sub MS_RowColChange()
PRD_NME = MS.TextMatrix(MS.Row, 11)
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
        M_ROU_IDY.SetFocus
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

Private Sub M_ROU_OFF_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Zero to retain paise values"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_ROU_OFF_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
If Val(M_ROU_OFF) = 0 Then
    T_NET_AMT = T_FRE_AMT
    M_ROU_OFF = "0.00"
ElseIf M_ROU_OFF >= 1 Then
    M_ROU_OFF = "0.00"
    GRS_CAL
Else
    M_ROU_OFF = amt(M_ROU_OFF)
    GRS_CAL
End If
    SSTab1.Tab = 3
    M_LRN_IDY.SetFocus
ElseIf K = 38 Then
    M_FRT_AMT.SetFocus
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
    Dim REC_SET As New ADODB.Recordset
    List6.Visible = False
    MS.Text = List6
    List6.Visible = False
    If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from prd,BRD where PRD.PRD_IDY LIKE BRD.PRD_IDY AND BRD.BAT_idy like '" & MS.TextMatrix(MS.Row, 15) & "'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        If APPLYCSR = False Then
            If Len(Trim(M_VAN_IDY)) <> 0 Then
                If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBV'") = True Then
                    If pcs_chk.Value = 1 Then
                        MS.TextMatrix(MS.Row, 4) = "B"
                        MS.TextMatrix(MS.Row, 5) = amt(REC_SET("PRD_SDR") & "")
                    Else
                        MS.TextMatrix(MS.Row, 4) = "B"
                        MS.TextMatrix(MS.Row, 5) = convert(Round(Val(REC_SET("PRD_SDR") & "") / MS.TextMatrix(MS.Row, 13), 2))
                    End If
                Else
                    If pcs_chk.Value = 1 Then
                        If List6 = "A" Then
                            MS.TextMatrix(MS.Row, 5) = amt(REC_SET("BAS_RAT") & "")
                        ElseIf List6 = "B" Then
                            MS.TextMatrix(MS.Row, 5) = amt(REC_SET("PRD_SDR") & "")
                        ElseIf List6 = "C" Then
                            MS.TextMatrix(MS.Row, 5) = amt(REC_SET("PRD_MRP") & "")
                        Else
                            MS.TextMatrix(MS.Row, 5) = "0.00"
                        End If
                    Else
                        If List6 = "A" Then
                            MS.TextMatrix(MS.Row, 5) = convert(Round(Val(REC_SET("BAS_RAT") & "") / MS.TextMatrix(MS.Row, 13), 2))
                        ElseIf List6 = "B" Then
                            MS.TextMatrix(MS.Row, 5) = convert(Round(Val(REC_SET("PRD_SDR") & "") / MS.TextMatrix(MS.Row, 13), 2))
                        ElseIf List6 = "C" Then
                            MS.TextMatrix(MS.Row, 5) = convert(Round(Val(REC_SET("PRD_MRP") & "") / MS.TextMatrix(MS.Row, 13), 2))
                        Else
                            MS.TextMatrix(MS.Row, 5) = "0.00"
                        End If
                        MS.TextMatrix(MS.Row, 22) = amt(REC_SET("LST_RAT") & "")
                        MS.TextMatrix(MS.Row, 23) = amt(REC_SET("LST_DIS") & "")
                    End If
                End If
            Else
                If pcs_chk.Value = 1 Then
                    If List6 = "A" Then
                        MS.TextMatrix(MS.Row, 5) = amt(REC_SET("BAS_RAT") & "")
                    ElseIf List6 = "B" Then
                        MS.TextMatrix(MS.Row, 5) = amt(REC_SET("PRD_SDR") & "")
                    ElseIf List6 = "C" Then
                        MS.TextMatrix(MS.Row, 5) = amt(REC_SET("PRD_MRP") & "")
                    Else
                        MS.TextMatrix(MS.Row, 5) = "0.00"
                    End If
                Else
                    If List6 = "A" Then
                        MS.TextMatrix(MS.Row, 5) = convert(Round(Val(REC_SET("BAS_RAT") & "") / MS.TextMatrix(MS.Row, 13), 2))
                    ElseIf List6 = "B" Then
                        MS.TextMatrix(MS.Row, 5) = convert(Round(Val(REC_SET("PRD_SDR") & "") / MS.TextMatrix(MS.Row, 13), 2))
                    ElseIf List6 = "C" Then
                        MS.TextMatrix(MS.Row, 5) = convert(Round(Val(REC_SET("PRD_MRP") & "") / MS.TextMatrix(MS.Row, 13), 2))
                    Else
                        MS.TextMatrix(MS.Row, 5) = "0.00"
                    End If
                    MS.TextMatrix(MS.Row, 22) = amt(REC_SET("LST_RAT") & "")
                    MS.TextMatrix(MS.Row, 23) = amt(REC_SET("LST_DIS") & "")
                End If
            End If
        End If
    End If
    P_DIS_CAL
    MS.Col = MS.Col + 1
    MS.SetFocus
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
If K = 39 Or K = 41 Then K = 0
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
M_TAX_AMT = convert(PERCENT(Val(T_SPD_AMT), Val(M_TAX_PCG)))
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
    If M_TAX_PCG > 0 Then M_TAX_AMT = convert(PERCENT(Val(T_SPD_AMT), Val(M_TAX_PCG)))
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
        MsgBox "Credit Days Of Bill Can't Go Beyond 365 Days !", vbCritical, "DAS Version 4.0"
        M_CRT_DAY.SetFocus
    Else
        If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AAS'") = True Then
        M_HVO_CHK = GetValue("CST_TYP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'")
        M_HVO_FLG = GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'AOM'")
        If Trim(M_HVO_CHK) = "HVO-P" Or Trim(M_HVO_CHK) = "HVO-NP" Or Trim(M_HVO_CHK) = "HVO-P-WO" Or Trim(M_HVO_CHK) = "HVO-NP-WO" Then
            If Trim(M_HVO_FLG) = "T" Then
                MsgBox "Global Restrictions Active on Open Market Schemes for this HVO Account!", vbInformation, "DAS Version 4.0"
                SSTab1.TabEnabled(6) = False
                SSTab1.Tab = 2
                M_ROU_IDY.SetFocus
                Exit Sub
            Else
                If Trim(M_HVO_CHK) = "HVO-P-WO" Or Trim(M_HVO_CHK) = "HVO-NP-WO" Then
                    MsgBox "This HVO Account is Restricted for Open Market Schemes!", vbInformation, "DAS Version 4.0"
                    SSTab1.TabEnabled(6) = False
                    SSTab1.Tab = 2
                    M_ROU_IDY.SetFocus
                    Exit Sub
                Else
                    AutoApplyScheme
                    MsgBox "Schemes Auto-Apply Mode! Applied Successfully All Valid Schemes for this Bill!", vbInformation, "DAS Version 4.0"
                    LabelAAS.Visible = True
                End If
            End If
        Else
            AutoApplyScheme
            MsgBox "Schemes Auto-Apply Mode! Applied Successfully All Valid Schemes for this Bill!", vbInformation, "DAS Version 4.0"
            LabelAAS.Visible = True
        End If
        End If
        If M_AUS_OK = 1 Then
        M_HVO_CHK = GetValue("CST_TYP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'")
        M_HVO_FLG = GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'AOM'")
        If Trim(M_HVO_CHK) = "HVO-P" Or Trim(M_HVO_CHK) = "HVO-NP" Or Trim(M_HVO_CHK) = "HVO-NP-WO" Or Trim(M_HVO_CHK) = "HVO-P-WO" Then
            If Trim(M_HVO_FLG) = "T" Then
                MsgBox "Global Restrictions Active on Open Market Schemes for this HVO Account!", vbInformation, "DAS Version 4.0"
                SSTab1.TabEnabled(6) = False
                SSTab1.Tab = 2
                M_ROU_IDY.SetFocus
                Exit Sub
            Else
                If Trim(M_HVO_CHK) = "HVO-P-WO" Or Trim(M_HVO_CHK) = "HVO-NP-WO" Then
                    MsgBox "This HVO Account is Restricted for Open Market Schemes!", vbInformation, "DAS Version 4.0"
                    SSTab1.TabEnabled(6) = False
                    SSTab1.Tab = 2
                    M_ROU_IDY.SetFocus
                    Exit Sub
                Else
                    AutoApplyScheme
                    MsgBox "Schemes Auto-Apply Mode! Applied Successfully All Valid Schemes for this Bill!", vbInformation, "DAS Version 4.0"
                    LabelAAS.Visible = True
                End If
            End If
        Else
            AutoApplyScheme
            MsgBox "Schemes Auto-Apply Mode! Applied Successfully All Valid Schemes for this Bill!", vbInformation, "DAS Version 4.0"
            LabelAAS.Visible = True
        End If
        End If
        M_ROU_IDY.SetFocus
        M_DUE_DAT = Format(DateAdd("d", Val(Trim(M_CRT_DAY)), M_INV_DAT.Value), "DD/MM/YYYY")
    End If
ElseIf KeyCode = 38 Then
    If Val(M_CRT_DAY) >= 365 Then
        MsgBox "Credit Days Of Bill Can't Go Beyond 365 Days !", vbCritical, "DAS Version 4.0"
        M_CRT_DAY.SetFocus
    Else
        SSTab1.Tab = 0
        M_DUE_DAT = Format(DateAdd("d", Val(Trim(M_CRT_DAY)), M_INV_DAT.Value), "DD/MM/YYYY")
'        M_ACT_IDY.SetFocus
        M_CST_IDY.SetFocus
    End If
End If
End Sub

Private Sub M_CRT_DAY_Keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub M_ACT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Sales  Account ID or Press <Enter> to View Existing Accounts"
End Sub

Private Sub M_ACT_IDY_KeyPress(K As Integer)
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
    If M_VAN_IDY.Visible = True Then
        If M_VAN_IDY.Enabled = True Then
            M_VAN_IDY.SetFocus
        Else
            M_CST_IDY.SetFocus
        End If
    Else
        M_CST_IDY.SetFocus
    End If
End If
End Sub
Private Sub M_CST_IDY_GotFocus()
Dim Chk_PCart As New ADODB.Recordset
If Chk_PCart.State = 1 Then Chk_PCart.Close
Chk_PCart.Open "select * from INV where INV_IDY like '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If Chk_PCart.RecordCount > 0 Then
    If Chk_PCart("ISS_QTY") > 0 Or Chk_PCart("OPN_QTY") > 0 Or Chk_PCart("TOT_QTY") > 0 Then
       MsgBox "This is a PushCart Sale Order!", vbCritical, "DAS Version 4.0"
       CANCELRECORD
       Exit Sub
    End If
End If
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Party"
End Sub

Private Sub M_CST_IDY_keydown(K As Integer, s As Integer)
If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Len(Trim(M_CST_IDY)) = 10 Then
    M_CST_IDY = Mid(Trim(M_CST_IDY), 6, 5)
    Else
    End If
    REC_SET.Open "select * from act where SUBSTR(act_idy,6,5) like '" & preincrement(Trim(M_CST_IDY)) & "' AND (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') AND CST_TYP NOT IN ('Van') AND DBT_TYP='Y'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        M_CST_IDY = Trim(REC_SET!ACT_IDY)
        If Len(Trim(M_CST_IDY)) = 0 Then Exit Sub
        M_CST_NME = Trim(REC_SET("act_nme") & "")
        M_CST_AD1 = Trim(REC_SET("act_ad1") & "")
        M_CST_AD2 = Trim(REC_SET("act_ad2") & "")
        M_CST_AD3 = Trim(REC_SET("act_ad3") & "")
        M_CST_AD4 = Trim(REC_SET("act_ad4") & "")
        CST = Trim(REC_SET!ACT_CST)
        GST = Trim(REC_SET!ACT_GST)
        M_DBT_MCL = amt(Trim(REC_SET("dbt_mcl")) & "")
        M_DBT_MCP = Trim(REC_SET("dbt_mcp") & "")
        SMN = Split(GetSalesman(M_CST_IDY, M_INV_DAT), "|")
        If MOD_REC = False And UBound(SMN) > 0 Then
            M_SMN_IDY = Trim(SMN(0))
            M_SMN_NME = Trim(SMN(1))
        End If
        If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBR'") = True Then
            M_ROU_IDY = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
        Else
            If MOD_REC = False And Len(Trim(M_VAN_IDY)) = 0 Then M_ROU_IDY = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
        End If
        If MOD_REC = False Then M_CSD_PCG = Val(Trim(REC_SET("dbt_dis")))
        If REC_SET1.State = 1 Then REC_SET1.Close
        'REC_SET1.Open "select * from smn where smn_idy like '" & Trim(REC_SET("smn_idy")) & "'", CON, adOpenKeyset, adlockpessimistic
        'If REC_SET1.RecordCount > 0 And MOD_REC = False Then M_SMN_NME = Trim(REC_SET1("smn_nme"))
        SSTab1.Tab = 0
        M_DBT_RAP = Trim(REC_SET("dbt_rap"))
        CHK_HVO_APP = Trim(GetValue("CST_TYP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
        If Trim(CHK_HVO_APP) = "HVO-NP" Or Trim(CHK_HVO_APP) = "HVO-NP-WO" Then
            'If MOD_REC = False Then
            If check("CSR", "ACT_IDY", Sincrement(M_CST_IDY)) = True Then If MsgBox("Apply HVO Scheme Discounts and Institutional Rates?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then APPLYCSR = True
        Else
            APPLYCSR = False
        End If
        If K = 38 Then
            M_INV_DAT.SetFocus
        Else
            M_SMN_IDY.SetFocus
        End If
    Else
        V_DLR_LST
    End If
ElseIf K = 45 Then
  M_FRM_ACT.Tag = "sact"
  M_FRM_ACT.Show 1
  M_SMN_IDY.SetFocus
End If

End Sub

Private Sub M_CST_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Form_Activate()
checkdongle
Toolbar1.Buttons(11).Value = tbrPressed
Toolbar1.Buttons(11).Enabled = True
If Me.Tag <> "Batch" And Me.Tag <> "Valid" Then
    SSTab1.Tab = 0
    List7 = "Y"
    If M_INV_IDY.Enabled = True Then
        M_INV_IDY.SetFocus
    ElseIf M_FRM_ACT.Tag = SACT Then 'CODE ADDED BY PKB
    M_FRM_SOR.M_CST_IDY.SetFocus
    Else
        M_INV_DAT.SetFocus
    End If
    Label23 = labdisp("IDY_SSO")
    SSTab1.TabEnabled(4) = False
    setval
    DTPicker1.MinDate = "1 / 1 / 1999"
    DTPicker1.MaxDate = "1 / 1 / 2015"
    SETYEAR M_FRM_SOR
End If
End Sub

Private Sub V_INV_LST()
If CHECKTB("act.dbf,inl.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim RS_DAYCLS As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
DateDayCls = Date
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MDC'") = True Then
    If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
    RS_DAYCLS.Open "SELECT MAX(DIR_DAT) FROM DIR WHERE DIR_IDY = 'DAY' .AND. TRIM(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
    If RS_DAYCLS.EOF = False Then
        RS_DAYCLS.MoveFirst
        DateDayCls = RS_DAYCLS(0)
    End If
End If

If Toolbar1.Buttons(11).Value = tbrPressed Then
    If Check1.Value = 1 Then
        If DateDayCls = Date Then
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_nme like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO'", con, adOpenKeyset, adLockPessimistic
        Else
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_nme like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO' and INL.INV_DAT > CTOD('" & DateDayCls & "')", con, adOpenKeyset, adLockPessimistic
        End If
    Else
        If DateDayCls = Date Then
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_ad4 like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO'", con, adOpenKeyset, adLockPessimistic
        Else
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_ad4 like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO' and INL.INV_DAT > CTOD('" & DateDayCls & "')", con, adOpenKeyset, adLockPessimistic
        End If
    End If
Else
    If Len(Trim(M_INV_IDY)) <> 0 Then
        If DateDayCls = Date Then
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_nme like '" & Trim(M_INV_IDY) & "' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO'", con, adOpenKeyset, adLockPessimistic
        Else
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_nme like '" & Trim(M_INV_IDY) & "' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO' and INL.INV_DAT > CTOD('" & DateDayCls & "')", con, adOpenKeyset, adLockPessimistic
        End If
    Else
        If DateDayCls = Date Then
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_nme like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO'", con, adOpenKeyset, adLockPessimistic
        Else
            REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4,ord_idy from act,inl where act_nme like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SSO' and INL.INV_DAT > CTOD('" & DateDayCls & "')", con, adOpenKeyset, adLockPessimistic
        End If
    End If
End If
list5.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ORD_IDY") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    V_LST_CAP = "Select Sale Order"
    list5.Visible = True
    list5.SetFocus
Else
    If Toolbar1.Buttons(11).Value = tbrPressed Then
        If SSTab1.Tab = 0 Then
            MsgBox "Sale Order Number Not Found!", vbExclamation, "DAS Version 4.0"
            SSTab1.Tab = 0
            M_INV_IDY.Enabled = True
            SendKeys "{HOME}+{END}"
            M_INV_IDY.SetFocus
        End If
    Else
        If Len(Trim(M_INV_IDY)) <> 0 Then
            If MsgBox("Create New Sale Order?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                M_INV_IDY.Enabled = False
                If Toolbar1.Buttons(19).Value = tbrPressed Then M_VAN_IDY = VAN_IDY
                M_INV_DAT.SetFocus
            Else
                M_INV_IDY.SetFocus
                SendKeys "{home}+{end}"
            End If
        Else
            MsgBox "Sale Orders Not Found !", vbInformation, "DAS Version 4.0"
            SSTab1.Tab = 0
            M_INV_IDY.SetFocus
        End If
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

If list1.Visible = False And List2.Visible = False And list3.Visible = False And list4.Visible = False And list5.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then EXT_REC
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_SOR.ActiveControl.name = "M_INV_IDY" Then
            V_INV_LST
        ElseIf M_FRM_SOR.ActiveControl.name = "M_CST_IDY" Then
            V_DLR_LST
        ElseIf M_FRM_SOR.ActiveControl.name = "M_SMN_IDY" Then
            V_LST_CAP = "Select Salesman"
            V_SMN_LST
        ElseIf M_FRM_SOR.ActiveControl.name = "M_ACT_IDY" Then
            V_ACT_LST
        ElseIf M_FRM_SOR.ActiveControl.name = "Text3" Then
            list4.Tag = ""
            V_PRD_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_SOR.ActiveControl Is TextBox Then M_FRM_SOR.ActiveControl.Text = FUNKEY(KeyCode)
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
    SYS_NUM = True
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
        M_ACT_IDY.SetFocus
    ElseIf list3.Visible = True Then
        list3.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    ElseIf LIST8.Visible = True Then
        LIST8.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_VAN_IDY.SetFocus
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
        If Text3.Visible = True Then
            Text3.SetFocus
        Else
            Check2.Caption = "Place"
            SSTab1.Tab = 1
            MS.Row = MS.Rows - 1
            If list4.Tag = "MSF" Then
                SSTab1.Tab = 5
                MSF.SetFocus
            Else
                SSTab1.Tab = 1
                MS.SetFocus
            End If
        End If
    ElseIf SSTab1.Tab = 1 Then
        SSTab1.Tab = 2
        Text3.Visible = False
        M_FLX_AMT.Visible = False
        P_DIS_CAL
        M_CSD_PCG_keydown 13, 1
        M_SPD_PCG_keydown 13, 1
        M_TAX_PCG_keydown 13, 1
        M_EXD_PCG_KeyDown 13, 1
        GRS_CAL
        M_CRT_DAY.SetFocus
    ElseIf SSTab1.Tab = 5 Then
        SSTab1.Tab = 2
        Text3.Visible = False
        M_FLX_AMT.Visible = False
        P_FRE_CAL
        G_FRE_AMT = Val(F_AMT_TOT)
        GRS_CAL
        M_ROU_OFF.SetFocus
    ElseIf SSTab1.Tab = 6 Then
        SSTab1.Tab = 2
        TXT_SCH.Visible = False
        GRS_CAL
        M_ROU_OFF.SetFocus
    Else
        If M_FRM_SOR.ActiveControl.name = "M_INV_IDY" Then
            Unload Me
            Set Form = Nothing
        Else
            saverecord
        End If
    End If
End If
End Sub
Private Sub Form_Load()
checkdongle
MOD_REC = False
CHGQTY = "No Scheme"
SSTab1.TabEnabled(6) = False
LOADROUTES
With MS
.ColWidth(0) = 980
.ColWidth(1) = 0
.ColWidth(2) = 800
.ColWidth(3) = 800
.ColWidth(4) = 600
.ColWidth(5) = 900
.ColWidth(6) = 800
.ColWidth(7) = 700
.ColWidth(8) = 700
.ColWidth(9) = 1500
.ColWidth(10) = 900
.ColWidth(11) = 3000
.ColWidth(12) = 2000
.ColWidth(13) = 0
.ColWidth(14) = 0
.ColWidth(15) = 0
.ColWidth(16) = 1200
.ColWidth(17) = 1200
.ColWidth(18) = 0
.ColWidth(19) = 0
.ColWidth(20) = 0
.ColWidth(21) = 1600


.TextMatrix(0, 0) = "Product ID"
SCl MS, 0
.TextMatrix(0, 1) = "Stock Type"
.TextMatrix(0, 2) = "Quantity"
SCl MS, 2
.ColAlignment(2) = 7
.TextMatrix(0, 3) = "UOM"
.ColAlignment(3) = 7
SCl MS, 3
.TextMatrix(0, 4) = "RateID"
SCl MS, 4
.ColAlignment(5) = 7
.TextMatrix(0, 5) = "Rate"
SCl MS, 5
.ColAlignment(5) = 7
.TextMatrix(0, 6) = "Dis.Amt"
SCl MS, 6
.ColAlignment(6) = 7
.TextMatrix(0, 7) = "CD%"
SCl MS, 7
.ColAlignment(7) = 7
.TextMatrix(0, 8) = "Tax%"
SCl MS, 8
.ColAlignment(8) = 7
.TextMatrix(0, 9) = "Amount"
.ColAlignment(9) = 7
.TextMatrix(0, 10) = "MRP"
.ColAlignment(10) = 7
.TextMatrix(0, 11) = "Product Name"
.TextMatrix(0, 12) = "Short Name"
.TextMatrix(0, 13) = "UPC"
.ColAlignment(13) = 7
.TextMatrix(0, 14) = "uom"
.TextMatrix(0, 15) = "BATIDY"
.ColAlignment(16) = 1
.TextMatrix(0, 16) = "Mfg.Date"
.ColAlignment(17) = 1
.TextMatrix(0, 17) = "BBD-Exp.Date"
.TextMatrix(0, 21) = "Batch"
.ColAlignment(22) = 7
.TextMatrix(0, 22) = "List Rate"
.ColAlignment(23) = 7
.TextMatrix(0, 23) = "List %+"
.ColAlignment(24) = 7
.TextMatrix(0, 24) = "Inc'tve"
.ColAlignment(25) = 7
.TextMatrix(0, 25) = "Inc'TTL"
.Col = 0
.Col = 3
End With

With MSF
.TextMatrix(0, 0) = "Product ID"
SCl MSF, 0
.TextMatrix(0, 1) = "Stock Type"
SCl MSF, 1
.TextMatrix(0, 2) = "Quantity"
SCl MSF, 2
.ColAlignment(2) = 7
.TextMatrix(0, 3) = "UOM"
SCl MSF, 3
.ColAlignment(3) = 7
.TextMatrix(0, 4) = "Rate"
.ColAlignment(4) = 7
.TextMatrix(0, 5) = "Amount"
.ColAlignment(5) = 7
.TextMatrix(0, 6) = "MRP"
.ColAlignment(6) = 7
.TextMatrix(0, 7) = "Product Name"
.TextMatrix(0, 8) = "Short Name"
.TextMatrix(0, 9) = "UPC"
.ColAlignment(9) = 7
.TextMatrix(0, 10) = "uom"
.TextMatrix(0, 11) = "BATIDY"
.TextMatrix(0, 12) = "GRPIDY"
.TextMatrix(0, 13) = "GRPCAT"
.TextMatrix(0, 14) = "GRPCLS"
.TextMatrix(0, 15) = "Batch"
.TextMatrix(0, 16) = "Tax %"
.ColAlignment(16) = 7
.TextMatrix(0, 17) = "List Rate"
.ColAlignment(17) = 7
.TextMatrix(0, 18) = "List %+"
.ColAlignment(18) = 7
.TextMatrix(0, 19) = "Rate"
.ColAlignment(19) = 7
.ColWidth(0) = 980
.ColWidth(1) = 1200
.ColWidth(2) = 800
.ColWidth(3) = 800
.ColWidth(4) = 900
.ColWidth(5) = 1500
.ColWidth(6) = 800
.ColWidth(7) = 3000
.ColWidth(8) = 2000
.ColWidth(9) = 1000
.ColWidth(10) = 0
.ColWidth(11) = 0
.ColWidth(12) = 0
.ColWidth(13) = 0
.ColWidth(14) = 0
End With
With MSS
    .TextMatrix(0, 0) = "Product Name"
    .TextMatrix(0, 1) = "MRP"
    .TextMatrix(0, 2) = "Quantity"
    .TextMatrix(0, 3) = "Rate"
    .TextMatrix(0, 4) = "Amount"
    .TextMatrix(0, 5) = "Product ID"
    .TextMatrix(0, 6) = "Batch"
    .TextMatrix(0, 7) = "batidy"
    .TextMatrix(0, 8) = "subunts"
    .TextMatrix(0, 9) = "grpid"
    .TextMatrix(0, 10) = "grpcat"
    .TextMatrix(0, 11) = "grpclf"
    .TextMatrix(0, 16) = "PRDUOM"   ' MODIFIED BY RADHIKA
    .ColWidth(0) = 2000
    .ColWidth(16) = 0
    .ColAlignment(1) = 7
    .ColWidth(1) = 1200
    .ColAlignment(2) = 7
    .ColWidth(2) = 800
    .ColAlignment(3) = 7
    .ColWidth(3) = 800
    .ColAlignment(4) = 7
    .ColWidth(4) = 800
    .ColWidth(5) = 1050
    .ColWidth(6) = 800
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 0
End With
AddData M_ROU_IDY, Me, "ROU_IDY", "RDY", "", "GROUP BY ROU_IDY"
UNT_TOT.Left = Label2.Left + Label2.Width + 400
MS.Col = 9
AMT_TOT.Left = (((MS.CellWidth + MS.CellLeft + MS.Left) - AMT_TOT.Width) - Picture1.Left) - 200
MS.Col = 8
TAX_TOT.Left = ((MS.CellWidth + MS.CellLeft + MS.Left) - TAX_TOT.Width) - Picture1.Left
MS.Col = 7
ADD_TOT.Left = ((MS.CellWidth + MS.CellLeft + MS.Left) - ADD_TOT.Width) - Picture1.Left
MS.Col = 6
DIS_TOT.Left = ((MS.CellWidth + MS.CellLeft + MS.Left) - DIS_TOT.Width) - Picture1.Left
PRD_NME.Left = 10

F_UNT_TOT.Left = Label12.Left + Label12.Width + 400
MSF.Col = 5
F_AMT_TOT.Left = (((MSF.CellWidth + MSF.CellLeft + MSF.Left) - F_AMT_TOT.Width) - Picture2.Left) - 200
F_PRD_NME.Left = 10

MS.Col = 0
M_INV_DAT = Date
M_LRN_DAT = Date
M_DCH_DAT = Date
waydate = Date
'list5.ColumnHeaders(1).Width = LIST2.Width / 4 'CHANGED/COMMENTED BY PNR
'list5.ColumnHeaders(2).Width = LIST2.Width / 4
'list5.ColumnHeaders(3).Width = LIST2.Width / 4
'list5.ColumnHeaders(4).Width = LIST2.Width / 4
List9.Selected(0) = True
M_DBT_MCL = "0.00"
M_DBT_MCP = "0"
Label19 = "0.00"
Label20 = "0.00"
list1.ColumnHeaders(1).Width = 3800
List2.ColumnHeaders(1).Width = 3800
list3.ColumnHeaders(1).Width = 3800
M_PRV_NUM = labdisp("IDY_SSO")
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
cd = "0.00"
units = "0"
M_PRD_AMT = "0.00"
tax = "0.00"
If CHECKTB("dir.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        Exit Sub
End If
If querycheck("DIR", "DIR_MSG", "N", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'NUM'") = True Then
    Toolbar1.Buttons(11).Value = tbrPressed
End If
If querycheck("DIR", "DIR_MSG", "C", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PCS'") = True Then
    pcs_chk.Value = 1
End If
pcs_chk.Visible = False
' Van Reference Always Enabled to False in Sales Order Module
'If querycheck("DIR", "DIR_MSG", "V", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'VAN'") = True Then
    Label0.Visible = False
    M_VAN_IDY.Visible = False
'End If
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press Ctrl+N to Generate New Sale Order ID by System, Press <Enter> to View Existing Sale Order"
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Or K = 13 Then
ElseIf K > 96 And K < 123 Then
K = K - 32
ElseIf K > 64 And K < 91 Then
Else
K = 0
End If
If M_INV_IDY = "" And K = 27 Then Unload Me
Set Form = Nothing
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
        If V_LST_CAP.Caption = "Select Salesman" Then
            M_SMN_IDY = list1.SelectedItem.ListSubItems.Item(1)
            If MOD_REC = False Then
            If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AYR'") = True Then
                rid = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
            Else
                rid = GETROUTE(M_CST_IDY, M_SMN_IDY, M_INV_DAT)
            End If
            For CL = 0 To M_ROU_IDY.ListCount - 1
                If Trim(M_ROU_IDY.List(CL)) = rid Then
                    M_ROU_IDY.ListIndex = CL
                    Exit For
                End If
            Next
            End If
            M_SMN_NME = list1.SelectedItem
            SSTab1.Tab = 0
            list1.Visible = False
            SSTab1.Tab = 1
            MS.SetFocus
            'M_ACT_IDY.SetFocus
        Else
            Text2 = list1.SelectedItem.ListSubItems.Item(1)
            SSTab1.Tab = 0
            list1.Visible = False
            M_INV_DAT.SetFocus
        End If
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
    SYS_NUM = True
    DISPDETAILS
End If
End If
End Sub


Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
If V_LST_CAP.Caption = "Select Salesman" Then
    V_LST_CAP = "Select Salesman"
    V_SMN_LST
ElseIf V_LST_CAP.Caption = "Select Party" Then
    V_DLR_LST
ElseIf V_LST_CAP.Caption = "Select Sale Order" Then
    V_INV_LST
ElseIf V_LST_CAP.Caption = "Select Account" Then
    V_ACT_LST
ElseIf V_LST_CAP.Caption = "Select Product" Then
    V_PRD_LST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If V_LST_CAP.Caption = "Select Salesman" Then
    V_LST_CAP = "Select Salesman"
    V_SMN_LST
ElseIf V_LST_CAP.Caption = "Select Party" Then
    V_DLR_LST
ElseIf V_LST_CAP.Caption = "Select Sale Order" Then
    V_INV_LST
ElseIf V_LST_CAP.Caption = "Select Account" Then
    V_ACT_LST
ElseIf V_LST_CAP.Caption = "Select Product" Then
    If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If

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
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Product Quantity"
ElseIf MSF.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Rate ID"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "SaleOrd Rate"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measure (Nos, Kgs, Ltrs, Mtrs, MT)"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Discount Amount"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Cash Discount Percentage"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Tax Percentage"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Amount"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf MS.Col = 11 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Name"
ElseIf MS.Col = 12 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
ElseIf MS.Col = 13 Then
    M_MDI.STS_BAR.Panels(1).Text = "Sub Units"
ElseIf MS.Col = 14 Then
    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measurement"
ElseIf MS.Col = 16 Then
    M_MDI.STS_BAR.Panels(1).Text = "Manufacturing Date"
ElseIf MS.Col = 17 Then
    M_MDI.STS_BAR.Panels(1).Text = "BBD-Exp Date"
ElseIf MS.Col = 21 Then
    M_MDI.STS_BAR.Panels(1).Text = "Batch"
ElseIf MS.Col = 22 Then
    M_MDI.STS_BAR.Panels(1).Text = "SS"
ElseIf MS.Col = 23 Then
    M_MDI.STS_BAR.Panels(1).Text = "List %+"
'ElseIf MS.Col = 24 Then
'    M_MDI_STS_BAR.Panels(1).Text = "Incentive"
'ElseIf MS.Col = 25 Then
'    M_MDI_STS_BAR.Panels(1).Text = "Incentive Total"
End If
End Sub

Private Sub SCH_CHK_kEYPRESS(K As Integer)
If K = 13 Then M_INV_DAT.SetFocus
End Sub

Private Sub M_INV_IDY_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
If Len(Trim(M_INV_IDY)) = 0 Then
    V_INV_LST
Else
    inV = M_INV_IDY
    m_smn = Text2
    clrctr M_FRM_SOR
    Text2 = m_smn
    M_CST_NME = ""
    M_CST_AD1 = ""
    M_CST_AD2 = ""
    M_CST_AD3 = ""
    M_CST_AD4 = ""
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
If CHECKTB("inl.dbf,dir.dbf,smn.dbf,act.dbf,inv.dbf,prd.dbf,pgr.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
Dim msg As Boolean
msg = False
If REC_SET.State = 1 Then REC_SET.Close
If SYS_NUM = True Then
    REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SSO'", con, adOpenKeyset
    GoTo NEXTLINE
End If
If Toolbar1.Buttons(11).Value = tbrPressed Then
    If Len(Trim(M_INV_IDY)) <> 10 Then
        REC_SET.Open "select * from inl where inv_idy like 'SSO" & Tincrement(Trim(M_INV_IDY)) & "' and txn_typ like 'SSO'", con, adOpenKeyset
    Else
        REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SSO'", con, adOpenKeyset
    End If
Else
    REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SSO'", con, adOpenKeyset
End If
NEXTLINE:
If REC_SET.RecordCount > 0 Then
    SSTab1.Tab = 0
    MsgBox "Please reselect the HVO Retailer in case of Spot Discount", vbInformation, "DAS Version 4.0"
    APPLYCSR = True
    If Trim(REC_SET!TXN_RM2) = "Scheme Amount" Then
        msg = True
        MsgBox "Scheme Sale Order! Re-apply Schemes for any changes made to this Sale Order!", vbInformation, "DAS Version 4.0"
        CHGQTY = "Yes Schemes"
    End If
    If Toolbar1.Buttons(11).Value = tbrPressed And SYS_NUM = False Then If Len(Trim(M_INV_IDY)) <> 10 Then M_INV_IDY = "SSO" & Tincrement(Trim(M_INV_IDY))
    If querycheck("dir", "dir_msg", "M", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MOD'") = True Then
        If DateDiff("D", REC_SET!INV_DAT, Date) <> 0 Then
            If Me.Tag = "Valid" Then Me.Tag = ""
            SYS_PSW.Tag = "Check Date"
            SYS_PSW.Show 1
            If Me.Tag = "Valid" Then
            Else
                CANCELRECORD
                Exit Sub
            End If
        End If
    End If
    SYS_NUM = False
    M_INV_IDY.Enabled = False
    MOD_REC = True
'    On Error Resume Next
    M_INV_DAT = REC_SET("INV_DAT")
    M_CST_IDY = Trim(REC_SET("cst_idy") & "")
    M_SMN_IDY = Trim(REC_SET("smn_idy") & "")
    dlr_idy = M_CST_IDY
    If Len(Trim(M_CST_IDY)) = 0 Then Exit Sub
    M_VAN_IDY = Trim(REC_SET("ORD_IDY") & "")

    If CHECKTB("dir.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        Exit Sub
    End If
    If querycheck("DIR", "DIR_MSG", "F", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'REF'") = True Then
        M_VAN_IDY.Enabled = True
    Else
        M_VAN_IDY.Enabled = False
    End If
    M_ACT_IDY = Trim(REC_SET("act_idy") & "")
    M_CRT_DAY.Locked = True
    M_DCH_IDY = Trim(REC_SET("dch_idy") & "")
    M_DCH_DAT = REC_SET("dch_dat")
    M_EXD_AMT = convert(amt(REC_SET("EXD_AMT") & ""))
    M_CRT_DAY = Val(REC_SET("crt_day") & "")
    M_DUE_DAT = Format(REC_SET!due_dat, "dd/mm/yyyy")
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
    M_WBL_IDY = Trim(REC_SET("wbl_idy") & "")
    M_WBL_DAT = REC_SET("wbl_dat")
    If Trim(REC_SET!TXN_RM2) <> "Scheme Amount" Then
        M_AMT_LES = Trim(REC_SET!TXN_RM2) & ""
        M_FRT_AMT = convert(amt(Val(REC_SET("frt_amt") & "")))
    End If
    M_PKG_CST = convert(amt(REC_SET("PKG_CST") & ""))
    M_TTL_CAS = Trim(REC_SET("ttl_cas") & "")
    M_TXN_RMK = Trim(REC_SET("txn_rmk") & "")
    M_TXN_TPT = Trim(REC_SET("txn_tpt") & "")
    M_LRN_IDY = Trim(REC_SET("lrn_idy") & "")
    M_LRN_AMT = convert(amt(REC_SET("lrn_amt") & ""))
    M_LRN_DAT = Trim(REC_SET("lrn_dat") & "")
    rid = Trim(REC_SET("rou_idy") & "")
    For K = 0 To M_ROU_IDY.ListCount - 1
        If Trim(M_ROU_IDY.List(K)) = Trim(rid) Then
            M_ROU_IDY.ListIndex = K
        End If
    Next
    M_EXD_PCG = discount(REC_SET!EXD_PCG & "")
    M_EXD_AMT = convert(amt(REC_SET!EXD_AMT & ""))
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from smn where smn_idy like '" & Trim(M_SMN_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then M_SMN_NME = REC_SET("smn_nme")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Trim(M_ACT_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then M_ACT_NME = REC_SET("act_nme")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from ACT where ACT_idy like '" & Trim(M_CST_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_CST_NME = Trim(REC_SET("ACT_nme") & "")
        M_CST_AD1 = Trim(REC_SET("ACT_ad1") & "")
        M_CST_AD2 = Trim(REC_SET("ACT_ad2") & "")
        M_CST_AD3 = Trim(REC_SET("ACT_ad3") & "")
        M_CST_AD4 = Trim(REC_SET("ACT_ad4") & "")
        CST = Trim(REC_SET("ACT_CST") & "") 'Added by PNR
        GST = Trim(REC_SET("ACT_GST") & "") 'Added by PNR
        ZIPCODE = Trim(REC_SET("act_zip"))
        Label18 = Trim(REC_SET("dbt_mcp") & "")
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ = 'SSO' AND LEN(TRIM(SCH_IDY))<>0 ", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        If msg = False Then
            MsgBox "Scheme Sale Order! Re-apply Schemes for any changes made to this Sale Order!", vbInformation, "DAS Version 4.0"
        End If
        CHGQTY = "Yes Schemes"
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ = 'SSO' and stk_typ like 'Salable'", con, adOpenKeyset, adLockPessimistic
    MS.Rows = REC_SET.RecordCount + 1
    
    For A = 1 To MS.Rows - 1
        MS.TextMatrix(A, 0) = Trim(REC_SET("prd_idy"))
        MS.TextMatrix(A, 1) = Trim(REC_SET("STK_TYP"))
        MS.TextMatrix(A, 4) = Trim(REC_SET("RAT_APP") & "")
        If pcs_chk.Value = 1 Then
            MS.TextMatrix(A, 2) = QTY(Val(REC_SET("prd_qty")))
            MS.TextMatrix(A, 5) = convert(amt(REC_SET("prd_rat") & ""))
        Else
            MS.TextMatrix(A, 2) = MTS(Val(REC_SET("prd_qty")), Val(REC_SET("SUB_UNT")))
            MS.TextMatrix(A, 5) = convert(Val(REC_SET("prd_rat") & "") / Val(REC_SET("SUB_UNT")))
        End If
        MS.TextMatrix(A, 6) = convert(REC_SET("add_AMT") & "")
        MS.TextMatrix(A, 7) = discount(REC_SET("trd_dis") & "")
        MS.TextMatrix(A, 8) = discount(REC_SET("trd_tax") & "")
        MS.TextMatrix(A, 15) = Trim(REC_SET!BAT_IDY)
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT * FROM BRD WHERE BAT_IDY LIKE '" & Trim(REC_SET!BAT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If REC_SET2.RecordCount > 0 Then
            If pcs_chk.Value = 1 Then
                MS.TextMatrix(A, 10) = convert(Val(REC_SET2!PRD_MRP))
            Else
                MS.TextMatrix(A, 10) = convert(Val(REC_SET2!PRD_MRP) / Val(REC_SET!SUB_UNT))
            End If
            MS.TextMatrix(A, 16) = Format(REC_SET2!bRD_dom, "dd/MM/yyyy")
            MS.TextMatrix(A, 17) = Format(REC_SET2!brd_doe, "dd/MM/yyyy")
            MS.TextMatrix(A, 21) = Trim(REC_SET2!Bat_nme)
            MS.TextMatrix(A, 22) = convert(REC_SET("LST_RAT") & "")
            MS.TextMatrix(A, 23) = convert(REC_SET("LST_DIS") & "")
            MS.TextMatrix(A, 24) = convert(REC_SET2("S_INC_DIS") & "")
'            MS.TextMatrix(A, 24) = convert(REC_SET("CPN_AMT") & "")
        End If
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT prd_nme,CON_FAC,SUB_UNT,pgr.GRP_IDY,GRP_CAT,GRP_CLS,PRD_PRN FROM PRD,pgr WHERE prd.grp_idy Like pgr.grp_idy and PRD_IDY LIKE '" & REC_SET("PRD_IDY") & "'", con, adOpenKeyset, adLockPessimistic
        If REC_SET2.RecordCount > 0 Then
            MS.TextMatrix(A, 11) = Trim(REC_SET2("PRD_NME"))
            MS.TextMatrix(A, 12) = Trim(REC_SET2("PRD_PRN"))
            MS.TextMatrix(A, 13) = Val(REC_SET2("SUB_UNT"))
            MS.TextMatrix(A, 14) = Val(REC_SET2("CON_FAC"))
            MS.TextMatrix(A, 18) = Trim(REC_SET2!GRP_IDY)
            MS.TextMatrix(A, 19) = Trim(REC_SET2!GRP_CAT)
            MS.TextMatrix(A, 20) = Trim(REC_SET2!GRP_CLS)
        End If
        REC_SET.MoveNext
    Next
    With MSF
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ = 'SSO' and sch_idy IS NULL and stk_typ in('Free','Damage','Replc.','TakeBack', 'Shortage','Breakage','Leakage')", con, adOpenKeyset, adLockPessimistic
    .Rows = REC_SET.RecordCount + 1
    For A = 1 To .Rows - 1
        .TextMatrix(A, 0) = Trim(REC_SET("prd_idy"))
        .TextMatrix(A, 1) = Trim(REC_SET("STK_TYP"))
        .TextMatrix(A, 3) = QTY(REC_SET!PRD_UOM)
        If pcs_chk.Value = 1 Then
            .TextMatrix(A, 2) = QTY(Val(REC_SET("prd_qty")))
            .TextMatrix(A, 4) = convert(amt(REC_SET("prd_rat") & ""))
        Else
            .TextMatrix(A, 2) = MTS(Val(REC_SET("prd_qty")), Val(REC_SET("SUB_UNT")))
            If IsNull(REC_SET!TAX_AMT) = True Then
                .TextMatrix(A, 4) = convert(Val(REC_SET("prd_rat") & "") / Val(REC_SET("SUB_UNT")))  'Changed by PNR
            Else
                .TextMatrix(A, 4) = convert(Val(REC_SET("prd_rat") & "") / Val(REC_SET("SUB_UNT")) + Val(REC_SET("TAX_AMT"))) 'Changed by PNR
            End If
        End If
        .TextMatrix(A, 5) = convert(REC_SET!COL_AMT)
        .TextMatrix(A, 11) = Trim(REC_SET!BAT_IDY)
        .TextMatrix(A, 16) = convert(0#)   'convert(REC_SET!TRD_TAX)
        .TextMatrix(A, 17) = convert(0#)   'convert(REC_SET!LST_RAT)
        .TextMatrix(A, 18) = convert(0#)   'convert(REC_SET!LST_DIS)
        .TextMatrix(A, 19) = convert(0#)   'convert(REC_SET!PRD_RAT)
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT * FROM BRD WHERE BAT_IDY LIKE '" & Trim(REC_SET!BAT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If REC_SET2.RecordCount > 0 Then
            .TextMatrix(A, 15) = Trim(REC_SET2!Bat_nme)
            If pcs_chk.Value = 1 Then
                MSF.TextMatrix(A, 6) = convert(Val(REC_SET2!PRD_MRP))
            Else
                MSF.TextMatrix(A, 6) = convert(Val(REC_SET2!PRD_MRP) / Val(REC_SET!SUB_UNT))
            End If
        End If
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT prd_nme,CON_FAC,SUB_UNT,pgr.GRP_IDY,GRP_CAT,GRP_CLS,PRD_PRN FROM PRD,pgr WHERE prd.grp_idy Like pgr.grp_idy and PRD_IDY LIKE '" & REC_SET("PRD_IDY") & "'", con, adOpenStatic
        If REC_SET2.RecordCount > 0 Then
            .TextMatrix(A, 7) = Trim(REC_SET2("PRD_NME"))
            .TextMatrix(A, 8) = Trim(REC_SET2("PRD_PRN"))
            .TextMatrix(A, 9) = Val(REC_SET2("SUB_UNT"))
            .TextMatrix(A, 10) = Val(REC_SET2("CON_FAC"))
            .TextMatrix(A, 12) = Trim(REC_SET2!GRP_IDY)
            .TextMatrix(A, 13) = Trim(REC_SET2!GRP_CAT)
            .TextMatrix(A, 14) = Trim(REC_SET2!GRP_CLS)
        End If
        REC_SET.MoveNext
    Next
    End With
    pcs_chk.Enabled = False
    P_DIS_CAL
    P_FRE_CAL
    GRS_CAL
    M_CST_IDY.SetFocus
Else
    V_INV_LST
End If
    If MS.Rows = 1 Then MS.Rows = 2
Exit Sub
LBL:
If err.Number = 35773 Then
    MsgBox "Invalid Sale Order Selected for this Financial Year", vbInformation, "DAS Version 4.0"
    CANCELRECORD
Else
    MsgBox err.Description
End If
End Sub

Private Sub list2_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If List2.ListItems.count > 0 Then
        M_ACT_IDY = List2.SelectedItem.ListSubItems.Item(1)
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
        M_CST_IDY = list3.SelectedItem.ListSubItems(1)
        M_CST_IDY_keydown 13, 0
        SSTab1.Tab = 0
        list3.Visible = False
        M_SMN_IDY.SetFocus
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
        If list4.Tag = "MSF" Then
            P_DIS_FPRD ls.ListSubItems(8)
        Else
            P_DIS_PRD ls.SubItems(8)
        End If
    End If
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

Private Sub MSF_EnterCell()
If MSF.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product ID View"
ElseIf MSF.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Stock Type"
ElseIf MSF.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Product Quantity"
ElseIf MSF.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Rate"
ElseIf MSF.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measure (Nos, Kgs, Ltrs, Mtrs, MT)"
ElseIf MSF.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Amount"
ElseIf MSF.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf MSF.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Name"
ElseIf MSF.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
ElseIf MSF.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Sub Units"
ElseIf MSF.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measurement"
ElseIf MSF.Col = 15 Then
    M_MDI.STS_BAR.Panels(1).Text = "Batch"
ElseIf MSF.Col = 16 Then
    M_MDI.STS_BAR.Panels(1).Text = "Tax %"
End If
End Sub

Private Sub MSF_keypress(K As Integer)
K = CHGCASE(K)
If K = 13 Then
    If MSF.Col = 0 And Trim(MSF.TextMatrix(MSF.Row, 0)) = "" Then
        Text4 = ""
        list4.Tag = "MSF"
        V_PRD_LST
        Exit Sub
    ElseIf (MSF.Cols = MSF.Col + 1) Or MSF.Col = 2 Then
        If MSF.TextMatrix(MSF.Rows - 1, 0) <> "" Then
            MSF.AddItem ""
            MSF.Col = 0
            MSF.Row = MSF.Rows - 1
            Exit Sub
        ElseIf MSF.Col = 2 Then
            MSF.Col = MSF.Col + 1
        End If
    ElseIf MSF.Col = 0 Then
        MSF.Col = 1
    Else
        MSF.Col = MSF.Col + 1
    End If
End If
If MSF.Col = 0 Then
    If K = 13 And MSF.TextMatrix(MSF.Row, 0) = "" Then
        Text4 = ""
        list4.Tag = "MSF"
        V_PRD_LST
    Else
        If K <> 27 Then
            If MSF.TextMatrix(MSF.Row, 0) = "" Then
                Text4.Visible = True
                Text4.Width = MSF.CellWidth
                Text4.Top = MSF.Top + MSF.CellTop
                Text4.Height = MSF.RowHeight(MSF.Row)
                Text4.Left = MSF.Left + MSF.CellLeft
                Text4.Text = Chr(K)
                Text4.SelStart = 1
                Text4.Visible = True
                Text4.SetFocus
            End If
        End If
    End If
ElseIf MSF.Col = 1 Then
    If MSF.TextMatrix(MSF.Row, 0) <> "" Then
        If K = 13 Then
        ElseIf K <> 27 Then
            List7.Visible = True
            List7.Top = MSF.Top + MSF.CellTop
            List7.Left = MSF.CellLeft + MSF.Left
            List7.Width = MSF.CellWidth
            List7.Height = MSF.RowHeight(MSF.Row)
            List7.Selected(0) = True
            List7.Visible = True
            List7.SetFocus
        End If
    End If
ElseIf MSF.Col = 2 Or MSF.Col = 3 Then
    If Len(Trim(MSF.TextMatrix(MSF.Row, 0))) = 10 Then
        If K = 13 Then
            If MSF.Row = MSF.Rows - 1 And MSF.TextMatrix(MSF.Row, 0) <> "" Then
                MSF.AddItem ""
                MSF.Row = MSF.Rows - 1
                MSF.Col = 0
            Else
            End If
        Else
            K = CHKNUM(K)
            If K <> 0 Then
                Text5.Visible = True
                Text5.Width = MSF.CellWidth
                Text5.Top = MSF.Top + MSF.CellTop
                Text5.Height = MSF.RowHeight(MSF.Row)
                Text5.Left = MSF.Left + MSF.CellLeft
                Text5.Text = Chr(K)
                Text5.SelStart = 1
                Text5.Visible = True
                Text5.SetFocus
            End If
        End If
    End If
ElseIf MSF.Col = 3 Then
ElseIf MSF.Col = 4 Then
End If
End Sub

Private Sub MSF_LeaveCell()
If Text4.Visible = True Then Text4.Visible = False
If List7.Visible = True Then List7.Visible = False
If Text5.Visible = True Then Text5.Visible = False
End Sub

Private Sub MSF_RowColChange()
F_PRD_NME = Trim(MSF.TextMatrix(MSF.Row, 7))
End Sub

Private Sub MSF_Scroll()
If Text4.Visible = True Then
    Text4 = ""
    Text4.Visible = False
    MS.SetFocus
ElseIf List7.Visible = True Then
    List7.Visible = False
    MS.SetFocus
ElseIf Text5.Visible = True Then
    Text5 = ""
    Text5.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub MSS_KeyPress(K As Integer)
K = CHGCASE(K)
With MSS
    If .Col = 2 Or .Col = 4 Then
        K = CHKNUM(K)
        If K <> 0 Then
            TXT_SCH.Width = .ColWidth(2)
            TXT_SCH.Height = .RowHeight(.Row)
            TXT_SCH.Left = .CellLeft + .Left
            TXT_SCH.Top = .CellTop + .Top
            TXT_SCH.Text = Chr(K)
            TXT_SCH.Visible = True
            TXT_SCH.SelStart = Len(TXT_SCH)
            TXT_SCH.Visible = True
            TXT_SCH.SetFocus
        End If
    End If
End With
End Sub

Private Sub MSS_LeaveCell()
TXT_SCH.Visible = False
End Sub

Private Sub MSS_Scroll()
If TXT_SCH.Visible = True Then
    TXT_SCH = ""
    TXT_SCH.Visible = False
    MSS.SetFocus
End If
End Sub
Private Sub SCH_AMT_GotFocus()
SCH_AMT.Locked = False
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Edit Scheme Amount - if Required"
End Sub
Private Sub SCH_AMT_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    SCH_AMT = convert(amt(SCH_AMT))
    Command5.SetFocus
ElseIf K = 38 Then
    SCH_AMT = convert(amt(SCH_AMT))
    Command5.SetFocus
End If
End Sub
Private Sub SCH_AMT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    If M_INV_IDY.Enabled = True Then
        M_INV_IDY.SetFocus
    Else
        M_INV_DAT.SetFocus
    End If
ElseIf SSTab1.Tab = 1 Then
    MS.Col = 0
    MS.Row = 1
    MS.SetFocus
ElseIf SSTab1.Tab = 2 Then
    M_CRT_DAY.SetFocus
ElseIf SSTab1.Tab = 5 Then
    MSF.Col = 0
    If MSF.Rows = 1 Then MSF.AddItem ""
    MSF.Row = 1
    MSF.SetFocus
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

Private Sub SSTab1_GotFocus()
If SSTab1.Tab = 2 Then
    P_DIS_CAL
    GRS_CAL
End If
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTab1.Tab = 1 Then MS.SetFocus
    If SSTab1.Tab = 2 Then M_CSD_PCG.SetFocus
    If SSTab1.Tab = 0 Then M_INV_DAT.SetFocus
End If
End Sub


Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub M_ACT_IDY_keydown(K As Integer, s As Integer)
If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_ACT_IDY)) & "' AND GRP_CLF LIKE 'Sales' AND ACT_TYP LIKE ''", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_ACT_IDY = Sincrement(M_ACT_IDY)
        If K = 38 Then
        M_SMN_IDY.SetFocus
        Else
        M_ACT_NME = Trim(REC_SET("act_nme"))
        SSTab1.Tab = 1
        MS.SetFocus
        End If
    Else
        V_ACT_LST
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
        list4.Tag = ""
        V_PRD_LST
        Exit Sub
    ElseIf MS.Cols = MS.Col + 1 Then
        If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
            MS.AddItem ""
            MS.Col = 0
            MS.Row = MS.Rows - 1
            Exit Sub
        End If
    ElseIf MS.Col = 0 Then
        MS.Col = 2
    Else
        MS.Col = MS.Col + 1
    End If
End If
If MS.Row <> 0 Then
    If MS.Col = 0 Then
        If K = 13 Then
            If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                MS.Col = 2
                MS.SetFocus
            Else
                list4.Tag = ""
                V_PRD_LST
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
    ElseIf MS.Col = 2 Then
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
    ElseIf MS.Col = 6 Then
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
    ElseIf MS.Col = 5 Or MS.Col = 3 Then
        If Len(Trim(MS.TextMatrix(MS.Row, 0))) = 10 Then
            If (MS.TextMatrix(MS.Row, 0) <> "" And MS.TextMatrix(MS.Row, 4) = "D") Or MS.Col = 3 Then
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
        End If
    ElseIf MS.Col = 4 Then
        If MS.TextMatrix(MS.Row, 0) <> "" Then
            If K = 13 Then
                'MS.Col = 9
            ElseIf K <> 27 Then
                List6.Visible = True
                List6.Top = MS.Top + MS.CellTop
                List6.Left = MS.CellLeft + MS.Left
                List6.Width = MS.CellWidth
                List6.Height = MS.RowHeight(MS.Row)
                List6.Selected(0) = True
                List6 = Chr(K)
                List6.Visible = True
                List6.SetFocus
            End If
        End If
    ElseIf MS.Col = 6 Or MS.Col = 7 Or MS.Col = 8 Then
        If K > 47 And K < 58 Or K = 46 Then
            If MS.TextMatrix(MS.Row, 0) <> "" Then
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
    
    ElseIf MS.Col = 10 Or MS.Col = 11 Then
        If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
            MS.AddItem ""
            MS.Col = 0
            MS.Row = MS.Rows - 1
        End If
    End If
End If
End Sub

Private Sub Text3_KeyDown(K As Integer, s As Integer)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If MS.Col = 0 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(Text3) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", con, adOpenKeyset
        If REC_SET.RecordCount = 1 Then
            P_DIS_PRD Trim(REC_SET!BAT_IDY)
        ElseIf REC_SET.RecordCount > 1 Then
            Text3 = Sincrement(Text3)
            list4.Tag = ""
            V_PRD_IDY_LST
        Else
            list4.Tag = ""
            V_PRD_LST
        End If
    End If
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(K As Integer)
If CHECKTB("inv.dbf,dir.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset, REC_VAN As New ADODB.Recordset, REC_VAN_INV As New ADODB.Recordset
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then

ElseIf K = 13 Then
    If MS.Col = 2 Then
        M_FLX_AMT = Val(M_FLX_AMT)
        If CHGQTY = "Yes Schemes" Then
            CHGQTY = "Yes Schemes " & "Qty"
        End If
        If REC_SET1.State = 1 Then REC_SET1.Close
        If REC_SET2.State = 1 Then REC_SET2.Close
        If REC_SET3.State = 1 Then REC_SET3.Close
        If Trim(M_INV_IDY) <> "" Then
            REC_SET2.Open "select SUM(TOT_PCS),SUB_UNT from inv where BAT_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 15)) & "' and inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ LIKE 'SSO' and stk_typ in('Salable','Free') GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
            If REC_SET2.RecordCount > 0 Then INVTOT = QTY(Val(STM(Val(REC_SET2(0)), REC_SET2!SUB_UNT)))
        End If
        If pcs_chk.Value = 1 Then
            STK = Val(GETFRESTK(MS.TextMatrix(MS.Row, 15), "Free"))
            STK = Val(STK) + Val(MTS(Val(M_FLX_AMT), Val(MS.TextMatrix(MS.Row, 13))))
        End If
        If pcs_chk.Value = 0 Then
            M_FLX_AMT = Int(amt(M_FLX_AMT))
            STK = Val(GETFRESTK(MS.TextMatrix(MS.Row, 15), "Free"))
            STK = STK + Val(M_FLX_AMT)
            pcs = Int((Val(STK)))
            GoTo PiecesTrack
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
        If b > 0 And Val(MS.TextMatrix(MS.Row, 13)) > 0 Then
            If InStr(b + 1, M_FLX_AMT, ".") > 0 Then
                MS.TextMatrix(MS.Row, 3) = "0.000"
                Exit Sub
            End If
            M = Mid(M_FLX_AMT, b + 1, 3)
            If M >= Val(MS.TextMatrix(MS.Row, 13)) Then
                MsgBox "Product Sub Units Are " & Val(MS.TextMatrix(MS.Row, 13)), vbExclamation, "DAS Version 4.0"
                M_FLX_AMT.Visible = True
                M_FLX_AMT.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                M_QTY = Val(STK)
            End If
        Else
            M_QTY = Val(STK)
        End If
PiecesTrack:
        If pcs_chk.Value = 0 Then M_QTY = pcs
            REMIND = MTS(Val(ADDBRDQTY(QTY(Val(INVTOT)), Trim(MS.TextMatrix(MS.Row, 15)))), Val(MS.TextMatrix(MS.Row, 13)))
            CURQTY = Val(M_QTY)
            If querycheck("DIR", "DIR_MSG", "U", " and dir_idy like 'SYS' AND DIR_IDX LIKE 'STK'") = False Then
                If Len(Trim(M_VAN_IDY)) = 0 Then
                    If Val(REMIND) - Val(CURQTY) < 0 Then
                        MsgBox "Stock Out Of Range, Check Stock !", vbCritical, "DAS Version 4.0"
'                        If M_FLX_AMT.Visible = True Then
'                            M_FLX_AMT.SetFocus
'                        Else
'                            MS.SetFocus
'                        End If
'                        SendKeys "{home}+{end}"
'                        Exit Sub
                     End If
                Else
                    If REC_VAN.State = 1 Then REC_VAN.Close
                    REC_VAN.Open "SELECT * FROM VAN WHERE INV_IDY LIKE '" & Trim(M_VAN_IDY) & "' AND BAT_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 15)) & "'", con, adOpenKeyset, adLockPessimistic
                    If REC_VAN.RecordCount > 0 Then
                        REC_VAN.MoveFirst
                        While Not REC_VAN.EOF
                            V_NET_LOD = Val(V_NET_LOD) + (MTS(Val(REC_VAN("TOT_QTY")), Val(MS.TextMatrix(MS.Row, 13))) - (MTS(Val(REC_VAN("RET_QTY")), Val(MS.TextMatrix(MS.Row, 13))) + MTS(Val(REC_VAN("LKG_QTY")), Val(MS.TextMatrix(MS.Row, 13))) + MTS(Val(REC_VAN("BRK_QTY")), Val(MS.TextMatrix(MS.Row, 13)))))
                            V_FRE_LOD = Val(V_FRE_LOD) + (MTS(Val(REC_VAN("FRE_QTY")), Val(MS.TextMatrix(MS.Row, 13))))
                            REC_VAN.MoveNext
                        Wend
                    End If
                    If REC_VAN_INV.State = 1 Then REC_VAN_INV.Close
                    REC_VAN_INV.Open "SELECT * FROM INV WHERE ORD_IDY LIKE '" & Trim(M_VAN_IDY) & "' AND BAT_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 15)) & "' AND STK_TYP LIKE 'Salable'", con, adOpenKeyset, adLockPessimistic
                    If REC_VAN_INV.RecordCount > 0 Then
                        REC_VAN_INV.MoveFirst
                        While Not REC_VAN_INV.EOF
                            B_NET_LOD = Val(B_NET_LOD) + Val(REC_VAN_INV("TOT_PCS"))
                            REC_VAN_INV.MoveNext
                        Wend
                    End If
                    C_NET_LOD = STM(Val(V_NET_LOD) - (Val(B_NET_LOD) + Val(V_FRE_LOD)), Val(MS.TextMatrix(MS.Row, 13)))
                    If M_FLX_AMT <> 0 Then
                        If (Val(V_NET_LOD) - (Val(B_NET_LOD) + Val(V_FRE_LOD))) - Val(CURQTY) < 0 Then
                            MsgBox "Net Load Stock Balance for this Sale Order is : " & str(Val(C_NET_LOD)) & ", Check Net Load Stock !", vbExclamation, "DAS Version 4.0"
                            If M_FLX_AMT.Visible = True Then
                                M_FLX_AMT.SetFocus
                            Else
                                MS.SetFocus
                            End If
                            SendKeys "{home}+{end}"
                            Exit Sub
                        End If
                    End If
                End If
            End If
            MS.Text = M_FLX_AMT
            M_FLX_AMT.Visible = False
            If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                MS.AddItem ""
                MS.Col = 0
                MS.Row = MS.Rows - 1
            End If
    MS.SetFocus
    P_DIS_CAL
    ElseIf MS.Col = 5 Then                                     'rate
        MS.TextMatrix(MS.Row, 5) = convert(amt(Val(M_FLX_AMT)))
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 6 Then                                     'DISCOUNT AMOUNT
        MS.TextMatrix(MS.Row, 6) = convert(amt(Val(M_FLX_AMT)))
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 7 Then                                          'AD%
        MS.TextMatrix(MS.Row, 7) = discount(Val(M_FLX_AMT))
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 8 Then                                          'TAX%
        MS.TextMatrix(MS.Row, 8) = discount(Val(M_FLX_AMT))
        P_DIS_CAL
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 3 Then                                          'UOM
        M_FLX_AMT = UNTCAL(M_FLX_AMT, MS.TextMatrix(MS.Row, 0))
        MS.Col = 2
        M_FLX_AMT_KeyPress 13
    Else
        K = 0
    End If
Else
    K = 0
End If
End Sub
Public Sub P_DIS_PRD(BRDID As String)
If CHECKTB("prd.dbf,brd.dbf,pgr.dbf,csr.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
MSROW = MS.Row
Dim REC_SET As New ADODB.Recordset
Dim CHK_VAN As New ADODB.Recordset
STYPE = "Salable"
For A = 1 To MS.Rows - 1
    If BRDID = MS.TextMatrix(A, 15) Then
        SSTab1.Tab = 1
        list4.Visible = False
        MsgBox "Product - SKU Already Exists In Voucher !", vbExclamation, "DAS Version 4.0"
        Text3.Visible = False
        MS.SetFocus
        MS.Col = 0
        MS.Row = A
        Exit Sub
    End If
Next
'* Check for Non Existence of any particular SKU in VAN Challan Ref Provided for this INVOICE
If Len(Trim(M_VAN_IDY)) > 0 Then
    If CHK_VAN.State = 1 Then CHK_VAN.Close
    CHK_VAN.Open "SELECT * FROM VAN where VAN.INV_IDY = '" & M_VAN_IDY & "' AND VAN.BAT_IDY = '" & BRDID & "'", con, adOpenKeyset, adLockPessimistic
        If CHK_VAN.EOF = True Then
           MS.CellForeColor = vbRed
        End If
End If
'*
REC_SET.Open "select * from PRD,brd,PGR where PRD.GRP_IDY = PGR.GRP_IDY AND prd.prd_idy = brd.prd_idy and brd.bat_idy = '" & Trim(BRDID) & "'", con, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    list4.Visible = False
    With MS
    .TextMatrix(.Row, 0) = REC_SET("PRD_IDY")
    .TextMatrix(.Row, 1) = STYPE
    If pcs_chk.Value = 0 Then
        .TextMatrix(.Row, 2) = "0"
    Else
        .TextMatrix(.Row, 2) = "0.000"
    End If
    .TextMatrix(.Row, 3) = "0.000"
    .TextMatrix(.Row, 4) = M_DBT_RAP
CHK_HVO_APP = Trim(GetValue("CST_TYP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
If Trim(CHK_HVO_APP) = "HVO-P" Or Trim(CHK_HVO_APP) = "HVO-P-WO" Then
    APPLYCSR = False
End If
If APPLYCSR = False Then
    If Trim(M_DBT_RAP) = "A" Then
        If pcs_chk.Value = 1 Then
            .TextMatrix(.Row, 5) = convert(amt(REC_SET("bas_rat") & ""))
        Else
            .TextMatrix(.Row, 5) = convert(Round(Val(REC_SET("bas_rat") & "") / REC_SET("sub_unt")))
        End If
    ElseIf Trim(M_DBT_RAP) = "C" Then
        If pcs_chk.Value = 1 Then
            .TextMatrix(.Row, 5) = convert(amt(REC_SET("prd_mrp") & ""))
        Else
            .TextMatrix(.Row, 5) = convert(Round(Val(REC_SET("prd_mrp") & "") / REC_SET("sub_unt"), 2))
        End If
    Else
        If pcs_chk.Value = 1 Then
            .TextMatrix(.Row, 5) = convert(amt(REC_SET("prd_sdr") & ""))
        Else
            .TextMatrix(.Row, 5) = convert(Round(Val(REC_SET("prd_sdr") & "") / REC_SET("sub_unt"), 2))
        End If
    End If
Else
    If pcs_chk.Value = 1 Then           'Additional Code for HVO Scheme Discounts from CSR Table, One for Rate, One for Discount % by PNR
        CSR_SCH_IDY = Trim(GetValue("SCH_IDY", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
        If Len(Trim(GetValue("SCH_IDY", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))) > 0 Then
            If Trim(GetValue("DIS_TYP", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'")) = "OnSpot" Then
                .TextMatrix(.Row, 6) = convert(GetValue("DIS_AMT", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
                .TextMatrix(.Row, 7) = convert(GetValue("DIS_PCG", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
                CSR_SCH_IDY = Trim(GetValue("SCH_IDY", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
            End If
        End If
        .TextMatrix(.Row, 5) = convert(GetValue("B_PRD_RAT", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
        If Val(.TextMatrix(.Row, 5)) = 0 Then .TextMatrix(.Row, 5) = convert(amt(REC_SET("prd_sdr") & ""))
    Else
        If Len(Trim(GetValue("SCH_IDY", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))) > 0 Then
            If Trim(GetValue("DIS_TYP", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'")) = "OnSpot" Then
                .TextMatrix(.Row, 6) = convert(GetValue("DIS_AMT", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
                .TextMatrix(.Row, 7) = convert(GetValue("DIS_PCG", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
                CSR_SCH_IDY = Trim(GetValue("SCH_IDY", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
            End If
        End If
        .TextMatrix(.Row, 5) = convert(Val(GetValue("B_PRD_RAT", "CSR WHERE B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'")) / Val(REC_SET("sub_unt")))
        If Val(.TextMatrix(.Row, 5)) = 0 Then .TextMatrix(.Row, 5) = convert(Round(Val(REC_SET("prd_sdr") & "") / REC_SET("sub_unt"), 2))
    End If
End If
    If Val(.TextMatrix(.Row, 6)) = 0# Then
        .TextMatrix(.Row, 6) = "0.00"
    End If
    If Val(.TextMatrix(.Row, 7)) = 0 Then
        .TextMatrix(.Row, 7) = discount(REC_SET("SAL_DIS") & "")
    End If
    .TextMatrix(.Row, 8) = discount(REC_SET("SAL_TAX") & "")
    .TextMatrix(.Row, 9) = "0.00"
    .TextMatrix(.Row, 22) = "0.00"
    .TextMatrix(.Row, 23) = "0.00"
    If pcs_chk.Value = 1 Then
        .TextMatrix(.Row, 10) = convert(amt(REC_SET("prd_mrp") & ""))
    Else
        .TextMatrix(.Row, 10) = convert(Round(Val(REC_SET("prd_mrp") & "") / REC_SET("sub_unt"), 2))
    End If
    .TextMatrix(.Row, 11) = Trim(REC_SET("prd_nme"))
    .TextMatrix(.Row, 12) = Trim(REC_SET("prd_prn"))
    .TextMatrix(.Row, 13) = Val(REC_SET("sub_unt"))
    .TextMatrix(.Row, 14) = Trim(REC_SET("con_fac"))
    .TextMatrix(.Row, 15) = REC_SET("BAT_IDY")
    .TextMatrix(.Row, 16) = Format(REC_SET!bRD_dom, "DD/MM/YYYY")
    .TextMatrix(.Row, 17) = Format(REC_SET!brd_doe, "DD/MM/YYYY")
    .TextMatrix(.Row, 18) = REC_SET("grp_IDY")
    .TextMatrix(.Row, 19) = Trim(REC_SET!GRP_CAT)
    .TextMatrix(.Row, 20) = Trim(REC_SET!GRP_CLS)
    .TextMatrix(.Row, 21) = Trim(REC_SET!Bat_nme)
    .TextMatrix(.Row, 24) = convert(Val(REC_SET!S_INC_DIS))
    If IsNull(REC_SET("LST_RAT")) = False Then
        .TextMatrix(.Row, 22) = Val(REC_SET("LST_RAT"))
    End If
    If IsNull(REC_SET("LST_DIS")) = False Then
        .TextMatrix(.Row, 23) = Val(REC_SET("LST_DIS"))
    End If
    list4.Tag = ""
    Text3.Visible = False
    SSTab1.Tab = 1
    If REC_SET("BRD_DOE") <= Date Then 'BBD Check
        MsgBox "SKU Batch BBD - Expiry Date is " & Format(REC_SET("BRD_DOE"), "dd/mmm/yyyy"), vbCritical, "DAS Version 4.0"
    End If
    .Row = MSROW
    .Col = 2
    .SetFocus
    End With
    pcs_chk.Enabled = False
End If
End Sub
Public Sub P_DIS_FPRD(BRDID As String)
If CHECKTB("prd.dbf,brd.dbf,pgr.dbf,csr.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
MSROW = MSF.Row
Dim REC_SET As New ADODB.Recordset
REC_SET.Open "select * from PRD,brd,PGR where PRD.GRP_IDY = PGR.GRP_IDY AND prd.prd_idy = brd.prd_idy and brd.bat_idy = '" & Trim(BRDID) & "'", con, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    list4.Visible = False
    With MSF
    .TextMatrix(.Row, 0) = REC_SET("PRD_IDY")
    .TextMatrix(.Row, 1) = ""
    If pcs_chk.Value = 0 Then
        .TextMatrix(.Row, 2) = "0"
    Else
        .TextMatrix(.Row, 2) = "0.000"
    End If
    .TextMatrix(.Row, 3) = "0.000"
If APPLYCSR = False Then
    If pcs_chk.Value = 1 Then
        .TextMatrix(.Row, 4) = convert(amt(REC_SET("prd_sdr") & ""))
    Else
        .TextMatrix(.Row, 4) = convert(Round(Val(REC_SET("prd_sdr") & "") / REC_SET("sub_unt"), 2))
    End If
Else
    If pcs_chk.Value = 1 Then
        .TextMatrix(.Row, 4) = convert(GetValue("PRD_RAT", "CSR where B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
    Else
        .TextMatrix(.Row, 4) = convert(Val(GetValue("PRD_RAT", "CSR where B_BAT_IDY = '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'")) / Val(REC_SET("sub_unt")))
    End If
End If
    .TextMatrix(.Row, 5) = "0.00"
    If pcs_chk.Value = 1 Then
        .TextMatrix(.Row, 6) = convert(amt(REC_SET("prd_mrp") & ""))
    Else
        .TextMatrix(.Row, 6) = convert(Round(Val(REC_SET("prd_mrp") & "") / REC_SET("sub_unt"), 2))
    End If
    .TextMatrix(.Row, 7) = Trim(REC_SET("prd_nme"))
    .TextMatrix(.Row, 8) = Trim(REC_SET("prd_prn"))
    .TextMatrix(.Row, 9) = Val(REC_SET("sub_unt"))
    .TextMatrix(.Row, 10) = Trim(REC_SET("con_fac"))
    .TextMatrix(.Row, 11) = REC_SET("BAT_IDY")
    .TextMatrix(.Row, 12) = REC_SET("grp_IDY")
    .TextMatrix(.Row, 13) = Trim(REC_SET!GRP_CAT)
    .TextMatrix(.Row, 14) = Trim(REC_SET!GRP_CLS)
    .TextMatrix(.Row, 15) = Trim(REC_SET!Bat_nme)
    .TextMatrix(.Row, 16) = convert(Val(REC_SET!SAL_TAX))
    .TextMatrix(.Row, 17) = convert(Val(REC_SET!LST_RAT))
    .TextMatrix(.Row, 18) = convert(Val(REC_SET!LST_DIS))
    .TextMatrix(.Row, 19) = convert(Val(REC_SET!PRD_SDR))
    list4.Tag = ""
    Text4.Visible = False
    SSTab1.Tab = 5
    .Row = MSROW
    .Col = 1
    MSF_keypress Asc("F")
    List7.Visible = True
    List7.SetFocus
    End With
    pcs_chk.Enabled = False
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
        If pcs_chk.Value = 1 Then
            M_INC_AMT = RATCAL(Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 24) & ""), Val(MS.TextMatrix(A, 13)))
            MS.TextMatrix(A, 9) = RATCAL(Val(MS.TextMatrix(A, 2) & ""), Val(MS.TextMatrix(A, 5) & ""), Val(MS.TextMatrix(A, 13)))
            MS.TextMatrix(A, 3) = QTY(WGTCAL(MS.TextMatrix(A, 2), Val(MS.TextMatrix(A, 14)), Val(MS.TextMatrix(A, 13))))
            LST_TOT = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 13))))
            M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
        Else
            M_INC_AMT = convert(Val(MS.TextMatrix(A, 2) & "") * Val(MS.TextMatrix(A, 24)))
            MS.TextMatrix(A, 9) = convert(Val(MS.TextMatrix(A, 2) & "") * Val(MS.TextMatrix(A, 5)))
            MS.TextMatrix(A, 3) = QTY(Val(MS.TextMatrix(A, 2)) * convert(Val(MS.TextMatrix(A, 14)) / Val(MS.TextMatrix(A, 13))))
            LST_TOT = Val(MS.TextMatrix(A, 2))
            LST_TOT = (LST_TOT * (Val(MS.TextMatrix(A, 22)) / Val(MS.TextMatrix(A, 13))))
            M_CPN_AMT = ((LST_TOT * Val(MS.TextMatrix(A, 23))) / 100)
            'LST_TOT = (Val(MS.TextMatrix(A, 2)) / Val(MS.TextMatrix(A, 2))) * Val(MS.TextMatrix(A, 22))
            'M_CPN_AMT = (Val(MS.TextMatrix(A, 2)) / Val(MS.TextMatrix(A, 13))) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        End If
        If Trim(GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'DPV'")) = "T" Then
            PRD_TOT = Val(MS.TextMatrix(A, 9)) - Val(M_INC_AMT)
        Else
            PRD_TOT = Val(MS.TextMatrix(A, 9)) - (Val(MS.TextMatrix(A, 6) + Val(M_INC_AMT)))
            ADDTOT = convert(Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 7)) / 100), 2))
            PRD_TOT = Val(PRD_TOT) - Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 7)) / 100), 2)
        End If
        MS.TextMatrix(A, 25) = convert(Val(M_INC_AMT))
 
        'M_CPN_AMT = Val(MS.TextMatrix(A, 2)) * (Val(MS.TextMatrix(A, 22)) * Val(MS.TextMatrix(A, 23)) / 100)
        'LST_TOT = Val(MS.TextMatrix(A, 2)) * Val(MS.TextMatrix(A, 22))
        TAXTOT = convert(Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 8)) / 100), 2))
        PRD_TOT = convert(amt(Val(PRD_TOT) + Round(Val(Val(PRD_TOT) * Val(MS.TextMatrix(A, 8)) / 100), 2)))
        If pcs_chk.Value = 0 Then
            CHK_HVO_APP = Trim(GetValue("CST_TYP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
            If Trim(CHK_HVO_APP) = "HVO-NP" Or Trim(CHK_HVO_APP) = "HVO-NP-WO" Then
                If APPLYCSR = True Then
                    DIS_PER_CASE = 0#
                    DIS_PER_CASE = convert(GetValue("DIS_AMT", "CSR WHERE B_BAT_IDY = '" & Trim(MS.TextMatrix(A, 15)) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
                    DIS_PER_PIECE = Val(DIS_PER_CASE) / Val(MS.TextMatrix(A, 13))
                    DIS_TOT = convert(Val(DIS_TOT) + (DIS_PER_PIECE * Val(MS.TextMatrix(A, 2))))
                    MS.TextMatrix(A, 6) = amt(Val((DIS_PER_PIECE * Val(MS.TextMatrix(A, 2)))))
                    MS.TextMatrix(A, 6) = amt(Round(MS.TextMatrix(A, 6), 2))
                End If
            End If
        Else
            CHK_HVO_APP = Trim(GetValue("CST_TYP", "ACT WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
            If Trim(CHK_HVO_APP) = "HVO-NP" Or Trim(CHK_HVO_APP) = "HVO-NP-WO" Then
                If APPLYCSR = True Then
                    DIS_PER_CASE = 0#
                    DIS_PER_CASE = convert(GetValue("DIS_AMT", "CSR WHERE B_BAT_IDY = '" & Trim(MS.TextMatrix(A, 15)) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'"))
                    DIS_PER_PIECE = Val(DIS_PER_CASE) / Val(MS.TextMatrix(A, 13))
                    DIS_TOT = convert(Val(DIS_TOT) + Val(DIS_PER_PIECE) * MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13))))
                    MS.TextMatrix(A, 6) = amt(Val(DIS_PER_PIECE) * MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13))))
                    MS.TextMatrix(A, 6) = amt(Round(MS.TextMatrix(A, 6), 2))
                End If
            End If
        End If
        TAX_TOT = convert(Val(TAX_TOT) + Val(TAXTOT) + M_CPN_AMT)
        If pcs_chk.Value = 1 Then
            UNT_TOT = QTY(Val(UNT_TOT) + Val(MS.TextMatrix(A, 2)))
        Else
            UNT_TOT = Val(UNT_TOT) + Val(MS.TextMatrix(A, 2))
        End If
        ADD_TOT = convert(Val(ADD_TOT) + Val(ADDTOT))
        AMT_TOT = convert(amt(Val(AMT_TOT) + Val(MS.TextMatrix(A, 9))) + LST_TOT)
    End If
Next
M_GRS_AMT = convert(amt(AMT_TOT))
M_PRD_TOT = convert(amt(AMT_TOT))
M_NET_AMT = convert(amt(Val(M_GRS_AMT)))
GRS_CAL
End Sub
Public Sub P_FRE_CAL()
F_UNT_TOT = 0
F_AMT_TOT = 0
For A = 1 To MSF.Rows - 1
    'Customer Discount calculating on units and rate
    If Trim(MSF.TextMatrix(A, 0)) <> "" Then
        If pcs_chk.Value = 1 Then
            MSF.TextMatrix(A, 5) = RATCAL(Val(MSF.TextMatrix(A, 2) & ""), Val(MSF.TextMatrix(A, 4) & ""), Val(MSF.TextMatrix(A, 9) & ""))
        Else
            MSF.TextMatrix(A, 5) = convert(Val(MSF.TextMatrix(A, 2) & "") * Val(MSF.TextMatrix(A, 4) & ""))
        End If
        If pcs_chk.Value = 1 Then
            F_UNT_TOT = QTY(Val(F_UNT_TOT) + Val(MSF.TextMatrix(A, 2)))
        Else
            F_UNT_TOT = Val(F_UNT_TOT) + Val(MSF.TextMatrix(A, 2))
        End If
        F_AMT_TOT = convert(amt(Val(F_AMT_TOT) + Val(MSF.TextMatrix(A, 5))))
    End If
Next
End Sub
Public Sub V_PRD_LST()
If CHECKTB("prd.dbf,brd.dbf,prt.dbf,dir.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
If check("prt", "prt_nme", "prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If querycheck("dir", "dir_msg", "U", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'STK'") = True Then
    qry = ""
Else
    If Len(Trim(M_VAN_IDY)) = 10 Then
        qry = ""
    Else
        qry = " AND PRD_QTY >0 "
    End If
End If

If list4.Tag = "MSF" Then   'Checking whether the Grid is Free & TakeBack
    If Len(Trim(M_VAN_IDY)) = 10 Then
        'REC_SET.Open "SELECT * FROM PRD,PGR,BRD,VAN WHERE PRD.GRP_IDY LIKE PGR.GRP_IDY AND PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.BAT_IDY LIKE BRD.BAT_IDY AND PRD.PRD_IDY LIKE '" & Sincrement(Text3) & "' AND VAN.INV_IDY LIKE '" & Trim(M_VAN_IDY) & "'", CON, adOpenKeyset
        REC_SET.Open "SELECT * FROM PRD,BRD,VAN WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.BAT_IDY LIKE BRD.BAT_IDY AND PRD_NME LIKE '" & Trim(Text4) & "%' AND VAN.INV_IDY LIKE '" & Trim(M_VAN_IDY) & "' AND STK_TYP IN('Salable')" & qry & "order by brw_seq", con, adOpenKeyset
    Else
        If Check1.Value = 1 Then    'Name Based Selection
            REC_SET.Open "SELECT * FROM PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text4) & "%' AND STK_TYP IN('Free','Salable','Promotional Item') order by brw_seq", con, adOpenKeyset
        Else
            REC_SET.Open "SELECT * FROM PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND PRD_PRN LIKE '" & Trim(Text4) & "%' AND STK_TYP IN('Free','Salable','Promotional Item') order by brw_seq", con, adOpenKeyset
        End If
    End If
Else
    If Check1.Value = 1 Then
        If Val(Text3.Text) > 0 Then
        MSROW = MS.Row  'Direct Selection by pressing the SKU Code
        If Len(Trim(M_VAN_IDY)) = 10 Then
            REC_SET.Open "SELECT * FROM PRD,PGR,BRD,VAN WHERE PRD.GRP_IDY LIKE PGR.GRP_IDY AND PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.BAT_IDY LIKE BRD.BAT_IDY AND PRD.PRD_IDY LIKE '" & Sincrement(Text3) & "' AND VAN.INV_IDY LIKE '" & Trim(M_VAN_IDY) & "'", con, adOpenKeyset
        Else
            REC_SET.Open "select * from PRD,brd,PGR where PRD.GRP_IDY LIKE PGR.GRP_IDY AND prd.prd_idy like brd.prd_idy and PRD.PRD_idy like '" & Sincrement(Text3) & "'", con, adOpenKeyset
        End If
        'REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(Text3) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') ", CON, adOpenKeyset
        '----------------------------------------------------
        If REC_SET.RecordCount > 0 Then
                list4.Visible = False
                'P_DIS_PRD Trim(REC_SET!BAT_IDY) 'BY PNR - CODE SIGNIFICANCE TO BE IDENTIFIED...
                With MS
                    .TextMatrix(.Row, 0) = REC_SET("PRD_IDY")
                    .TextMatrix(.Row, 1) = STYPE
                    If pcs_chk.Value = 0 Then
                        .TextMatrix(.Row, 2) = "0"
                    Else
                        .TextMatrix(.Row, 2) = "0.000"
                    End If
                    .TextMatrix(.Row, 3) = "0.000"
                    .TextMatrix(.Row, 4) = M_DBT_RAP
                    If APPLYCSR = False Then
                            If Trim(M_DBT_RAP) = "A" Then
                                If pcs_chk.Value = 1 Then
                                    .TextMatrix(.Row, 5) = convert(amt(REC_SET("bas_rat") & ""))
                                Else
                                    .TextMatrix(.Row, 5) = convert(Round(Val(REC_SET("bas_rat") & "") / REC_SET("sub_unt")))
                                End If
                            ElseIf Trim(M_DBT_RAP) = "C" Then
                                If pcs_chk.Value = 1 Then
                                    .TextMatrix(.Row, 5) = convert(amt(REC_SET("prd_mrp") & ""))
                                Else
                                .TextMatrix(.Row, 5) = convert(Round(Val(REC_SET("prd_mrp") & "") / REC_SET("sub_unt"), 2))
                            End If
                    Else
                            If pcs_chk.Value = 1 Then
                                .TextMatrix(.Row, 5) = convert(amt(REC_SET("prd_sdr") & ""))
                            Else
                                .TextMatrix(.Row, 5) = convert(Round(Val(REC_SET("prd_sdr") & "") / REC_SET("sub_unt"), 2))
                            End If
                    End If
                    Else
                    If pcs_chk.Value = 1 Then
                        .TextMatrix(.Row, 5) = convert(GetValue("DIS_AMT", "CSR where B_BAT_IDY LIKE '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'"))
                    Else
                        .TextMatrix(.Row, 5) = convert(Val(GetValue("DIS_AMT", "CSR where B_BAT_IDY LIKE '" & Trim(REC_SET("bat_idy")) & "' AND ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'")) / Val(REC_SET("sub_unt")))
                    End If
                    End If
                    .TextMatrix(.Row, 6) = "0.00"
                    .TextMatrix(.Row, 7) = discount(REC_SET("sal_dis") & "")
                    .TextMatrix(.Row, 8) = discount(REC_SET("SAL_TAX") & "")
                    .TextMatrix(.Row, 9) = "0.00"
                    If pcs_chk.Value = 1 Then
                        .TextMatrix(.Row, 10) = convert(amt(REC_SET("prd_mrp") & ""))
                    Else
                        .TextMatrix(.Row, 10) = convert(Round(Val(REC_SET("prd_mrp") & "") / REC_SET("sub_unt"), 2))
                    End If
                    .TextMatrix(.Row, 11) = Trim(REC_SET("prd_nme"))
                    .TextMatrix(.Row, 12) = Trim(REC_SET("prd_prn"))
                    .TextMatrix(.Row, 13) = Val(REC_SET("sub_unt"))
                    .TextMatrix(.Row, 14) = Trim(REC_SET("con_fac"))
                    .TextMatrix(.Row, 15) = REC_SET("BAT_IDY")
                    .TextMatrix(.Row, 16) = Format(REC_SET!bRD_dom, "DD/MM/YYYY")
                    .TextMatrix(.Row, 17) = Format(REC_SET!brd_doe, "DD/MM/YYYY")
                    .TextMatrix(.Row, 18) = REC_SET("grp_IDY")
                    .TextMatrix(.Row, 19) = Trim(REC_SET!GRP_CAT)
                    .TextMatrix(.Row, 20) = Trim(REC_SET!GRP_CLS)
                    .TextMatrix(.Row, 21) = Trim(REC_SET!Bat_nme)
                    list4.Tag = ""
                    Text3.Visible = False
                    SSTab1.Tab = 1
                    .Row = MSROW
                    .Col = 2
                    .SetFocus
            End With
            pcs_chk.Enabled = False
        Else
            MsgBox "Product, SKU Not Available for Entry, Check Van Challan!", vbInformation, "DAS Version 4.0"
            MS.Row = MSROW
            MS.Col = 0
            MS.SetFocus
        End If
        REC_SET.Close
        Exit Sub
        '----------------------------------------------------
        Else    'Filtered Selection by pressing the SKU Alphabet lookup in  PRD_NME field
            If Len(Trim(M_VAN_IDY)) = 10 Then
                REC_SET.Open "SELECT * FROM PRD,BRD,VAN WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.BAT_IDY LIKE BRD.BAT_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' AND VAN.INV_IDY LIKE '" & Trim(M_VAN_IDY) & "' AND STK_TYP IN('Salable')" & qry & "order by brw_seq", con, adOpenKeyset
            Else
                REC_SET.Open "SELECT * FROM PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' AND STK_TYP IN('Salable')" & qry & "order by brw_seq", con, adOpenKeyset
            End If
        End If
    Else        'Filtered Selection by pressing the SKU Alphabet lookup in ShortName field
        If Len(Trim(M_VAN_IDY)) = 10 Then
            REC_SET.Open "SELECT * FROM PRD,BRD,VAN WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.BAT_IDY LIKE BRD.BAT_IDY AND PRD_PRN LIKE '" & Trim(Text3) & "%' AND VAN.INV_IDY LIKE '" & Trim(M_VAN_IDY) & "' AND STK_TYP IN('Salable')" & qry & "order by brw_seq", con, adOpenKeyset
        Else
            REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND PRD_PRN LIKE '" & Trim(Text3) & "%' AND STK_TYP IN('Salable')" & qry & "order by brw_seq", con, adOpenKeyset
        End If
    End If
End If
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    REC_SET.MoveFirst
    If Trim(M_VAN_IDY) = "" Then
        For A = 1 To REC_SET.RecordCount
            Set ls = list4.ListItems.Add(, , Trim(REC_SET("PRD_NME")))
            ls.ListSubItems.Add , , QTY(REC_SET("prd_qty") & "")
            ls.ListSubItems.Add , , convert(amt(Val(REC_SET("PRD_MRP") & "")))
            ls.ListSubItems.Add , , convert(amt(Val(REC_SET("PRD_SDR") & "")))
            ls.ListSubItems.Add , , convert(amt(Val(REC_SET("BAS_RAT") & "")))
            ls.ListSubItems.Add , , Trim(REC_SET("PRD_PRN") & "")
            ls.ListSubItems.Add , , Val(REC_SET("SUB_UNT") & "")
            ls.ListSubItems.Add , , Trim(REC_SET("PRD_IDY") & "")
            ls.ListSubItems.Add , , REC_SET!BAT_IDY
            ls.ListSubItems.Add , , Trim(REC_SET!Bat_nme)
            REC_SET.MoveNext
        Next
    Else
        For A = 1 To REC_SET.RecordCount
            NET_LOD_QTY = QTY(MTS(Val(REC_SET("TOT_QTY")), Val(REC_SET("SUB_UNT"))) - (MTS(Val(REC_SET("RET_QTY")), Val(REC_SET("SUB_UNT"))) + MTS(Val(REC_SET("LKG_QTY")), Val(REC_SET("SUB_UNT"))) + MTS(Val(REC_SET("BRK_QTY")), Val(REC_SET("SUB_UNT")))))
            C_NET_LOD_QTY = STM(Val(NET_LOD_QTY), Val(REC_SET("SUB_UNT")))
            list4.ColumnHeaders(2).Text = "Net Load"
            Set ls = list4.ListItems.Add(, , Trim(REC_SET("PRD_NME")))
            ls.ListSubItems.Add , , QTY(Val(C_NET_LOD_QTY))
    '        ls.ListSubItems.Add , , QTY(REC_SET("prd_qty") & "")
            ls.ListSubItems.Add , , convert(amt(Val(REC_SET("PRD_MRP") & "")))
            ls.ListSubItems.Add , , convert(amt(Val(REC_SET("PRD_SDR") & "")))
            ls.ListSubItems.Add , , convert(amt(Val(REC_SET("BAS_RAT") & "")))
            ls.ListSubItems.Add , , Trim(REC_SET("PRD_PRN") & "")
            ls.ListSubItems.Add , , Val(REC_SET("SUB_UNT") & "")
            ls.ListSubItems.Add , , Trim(REC_SET("PRD_IDY") & "")
            ls.ListSubItems.Add , , REC_SET!BAT_IDY
            ls.ListSubItems.Add , , Trim(REC_SET!Bat_nme)
            REC_SET.MoveNext
        Next
    End If
    V_LST_CAP = "Select Product - SKU"
    SSTab1.Tab = 4
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product - SKU Not Found, Define Product, SKU !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 1
    If Text3.Visible = True Then
        Text3.SetFocus
    Else
        Check2.Caption = "Place"
        MS.SetFocus
    End If
    End If
    If SSTab1.Tab = 5 Then
        MsgBox "Product - SKU Not Found, Define Product, SKU !", vbExclamation, "DAS Version 4.0"
        MSF.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub
Public Sub V_PRD_IDY_LST()
If CHECKTB("prd.dbf,brd.dbf,prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
If check("prt", "prt_nme", "prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If list4.Tag = "MSF" Then
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_IDY LIKE '" & Sincrement(Text4) & "%' AND PRD.STK_TYP IN('Free','Salable','Promotional Item') and BRD.prd_qty >0 order by brw_seq", con, adOpenKeyset
Else
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_IDY LIKE '" & Sincrement(Text3) & "%' AND PRD.STK_TYP IN('Salable') and BRD.prd_qty >0 order by brw_seq", con, adOpenKeyset
End If
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , QTY(REC_SET("prd_qty") & "")
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("prd_mrp") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("prd_sdr") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("bas_rat") & "")))
        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
        ls.ListSubItems.Add , , Val(REC_SET("sub_unt") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET!BAT_IDY & "")
        REC_SET.MoveNext
    Next
    V_LST_CAP = "Select Product - SKU"
    SSTab1.Tab = 4
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
        MsgBox "Product - SKU Not Found, Define Product, SKU !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 1
        If Text3.Visible = True Then
            Text3.SetFocus
        Else
            Check2.Caption = "Place"
            MS.SetFocus
        End If
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub
Public Sub Printrecord()
If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_SOR") = True Then
    M_FRM_PRV.Check3.Value = 1
Else
    M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "SOR"
M_FRM_PRV.Show 1
End Sub
Public Sub setval()
On Error Resume Next
CST = ""
GST = ""
m_smn = Text2
clrctr M_FRM_SOR
APPSCH = False
SCHLST.ListItems.CLEAR
MSS.Rows = 1
Text2 = m_smn
'M_ROU_IDY.Selected(0) = True
M_CST_NME = ""
M_CST_AD1 = ""
M_CST_AD2 = ""
M_CST_AD3 = ""
M_CST_AD4 = ""
'Label17 = "0.00"
Label18 = "0"
Label19 = "0.00"
Label20 = "0.00"
M_GRS_AMT = "0.00"
M_NET_AMT = "0.00"
M_ROU_OFF = "0.00"
M_SMN_NME = ""
M_ACT_NME = ""
PRD_NME = ""
M_SPD_DED = "0.00"
cd = "0.00"
units = "0.00"
tax = "0.00"
AMT_TOT = "0.00"
UNT_TOT = "0.00"
TAX_TOT = "0.00"
DIS_TOT = "0.00"
ADD_TOT = "0.00"
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
F_PRD_NME = ""
F_AMT_TOT = "0.00"
F_UNT_TOT = "0"
M_LST_RAT = "0.00"
M_LST_DIS = "0.00"
M_CPN_AMT = "0.00"
'M_ROU_IDY = ""
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
REC_SET.Open "select rou_off from inl where inv_idy like '" & Trim(M_INV_IDY) & "' AND TXN_TYP LIKE 'SSO'", con, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then If (Val(REC_SET(0) & "") = 0 Or Val(M_ROU_OFF) = 0) And MOD_REC = True Then Exit Sub
dot = 0
rof = 0
dot = InStr(Trim(T_FRE_AMT), ".")
If dot > 0 Then
    rof = Mid(Trim(T_FRE_AMT), dot + 1, 2)
    If Len(rof) = 1 Then rof = "0" & rof '& "0"
    If Val(rof) > 0 And Val(rof) < 50 Then
        M_ROU_OFF = "-0." & rof
        M_NET_AMT = Mid(Val(T_FRT_AMT), 1, dot) & "00"
    ElseIf Val(rof) > 0 And Val(rof) >= 50 Then
        rof = 100 - Val(rof)
        If Len(rof) = 1 Then
            M_ROU_OFF = "+0.0" & rof
        ElseIf Len(rof) = 2 Then
            M_ROU_OFF = "+0." & rof
        End If
        T_NET_AMT = Mid(Val(T_FRE_AMT), 1, dot - 1)
        T_NET_AMT = Val(M_NET_AMT) + 1 & ".00"
    Else
        M_ROU_OFF = "0.00"
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
If K = 39 Or K = 41 Then K = 0
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
If MsgBox("Create New Price n Batch Record?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    If CHECKTB("dir.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    Me.Tag = "Batch"
    If querycheck("DIR", "DIR_MSG", "E", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'EXP'") = True Then M_FRM_NBC.Width = 7530
    M_FRM_NBC.Tag = "SSO"
    M_FRM_NBC.Show 1
    MS.SetFocus
Else
    MS.SetFocus
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rsprd As New ADODB.Recordset
If MSF.Col = 0 Then
    If KeyCode = 13 Then
        rsprd.Open "select * from prd,brd where prd.prd_idy like brd.prd_idy and prd.prd_idy like '" & Sincrement(Text4) & "'", con, adOpenKeyset, adLockPessimistic
        If rsprd.RecordCount = 1 Then
            P_DIS_FPRD Trim(rsprd!BAT_IDY)
        ElseIf rsprd.RecordCount > 1 Then
            list4.Tag = "MSF"
            V_PRD_IDY_LST
        Else
            list4.Tag = "MSF"
            V_PRD_LST
        End If
    End If
End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
K = CHGCASE(K)
End Sub

Private Sub TEXT5_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("inv.dbf,brd.dbf,dir.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim RSINV As New ADODB.Recordset
If K = 13 Then
With MSF
    If MSF.Col = 2 Then
        If pcs_chk.Value = 1 Then
            Text5 = QTY(Val(Text5))
            b = InStr(Trim(Text5), ".")
            If b <> 0 Then M = Mid(Text5, b + 1, Len(Text5) - b + 1)
            If Val(M) >= Val(.TextMatrix(.Row, 9)) Then
                MsgBox "Product SubUnits are " & Val(.TextMatrix(.Row, 9)), vbExclamation, "DAS Version 4.0"
                Text5.Visible = True
                Text5.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
            STK = Val(Text5)
        Else
            Text5 = Int(Val(Text5))
            STK = STM(Int(Val(Text5)), .TextMatrix(.Row, 9))
        End If
        RSINV.Open "select sum(tot_pcs) from inv where bat_idy like '" & Trim(.TextMatrix(.Row, 11)) & "' and inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SSO' and stk_typ in('Free','Salable')", con, adOpenKeyset, adLockPessimistic
        If RSINV.RecordCount > 0 Then
            INVTOT = Val(RSINV(0))
        End If
        If .Col = 2 Then
            salqty = GETSALSTK(.TextMatrix(.Row, 11))
            If pcs_chk.Value = 1 Then
                CURQTY = (Val(salqty) + MTS(Val(STK), .TextMatrix(.Row, 9)))
            Else
                CURQTY = (Val(salqty) + MTS(Val(STK), .TextMatrix(.Row, 9)))
            End If
            BRDQTY = Val(GetValue("prd_qty", "brd where bat_idy like '" & Trim(.TextMatrix(.Row, 11)) & "'"))
            BRDQTY = MTS(Val(BRDQTY), .TextMatrix(.Row, 9))
            If MSF.TextMatrix(MSF.Row, 1) = "Free" Or MSF.TextMatrix(MSF.Row, 1) = "Replc." Then
                If querycheck("DIR", "DIR_MSG", "U", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'STK'") = False Then
                    If Len(Trim(M_VAN_IDY)) = 10 Then
                        'No Stock Check is Required if VAN Challan Reference is there! - Modified by Radhika & Ramanesh
                    Else
                        If (Val(INVTOT) + Val(BRDQTY)) - Val(CURQTY) < 0 Then
                            MsgBox "Product - SKU Out of Stock, Check Available Stock!", vbInformation, "DAS Version 4.0"
                            If Text5.Visible = True Then
                                SendKeys "{home}+{end}"
                                Text5.Visible = True
                                Text5.SetFocus
                            Else
                                MSF.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
            If pcs_chk.Value = 1 Then
                MSF.Text = QTY(Text5)
            Else
                MSF.Text = Val(Text5)
            End If
            If pcs_chk.Value = 1 Then
                MSF.TextMatrix(MSF.Row, 3) = QTY(WGTCAL(QTY(Text5), .TextMatrix(.Row, 10), .TextMatrix(.Row, 9)))
            Else
                MSF.TextMatrix(MSF.Row, 3) = QTY(Val(Text5) * (Val(.TextMatrix(.Row, 10)) / Val(.TextMatrix(.Row, 9))))
            End If
            Text5.Visible = False
            If .TextMatrix(.Rows - 1, 0) <> "" Then
                .AddItem ""
                .Col = 0
                .Row = .Rows - 1
                .SetFocus
                P_FRE_CAL
                Exit Sub
            End If
            .SetFocus
        End If
    ElseIf MSF.Col = 3 Then
        If pcs_chk.Value = 1 Then
            Text5 = UNTCAL(QTY(Text5), MSF.TextMatrix(MSF.Row, 0))
        Else
            Text5 = UNTCAL(QTY(Text5), MSF.TextMatrix(MSF.Row, 0))
        End If
        MSF.Col = 2
        TEXT5_KeyDown 13, 1
    End If
    P_FRE_CAL
End With
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
K = CHKNUM(K)
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
ElseIf Button.KEY = "AUS" Then
    If Button.Value = tbrPressed Then
        M_AUS_OK = 1
    Else
        M_AUS_OK = 0
    End If
ElseIf Button.KEY = "crt" Then
CREATE
ElseIf Button.KEY = "smn" Then
V_LST_CAP = "Select  Salesman"
V_SMN_LST
ElseIf Button.KEY = "EMP_SHL" Then
EMP_SHL "SAL_SHL"
ElseIf Button.KEY = "ed" Then
    If M_FRM_SOR.ActiveControl.name = "M_INV_IDY" Then
        V_LST_CAP = "Select SaleOrder"
        V_INV_LST
    ElseIf M_FRM_SOR.ActiveControl.name = "M_CST_IDY" Then
        V_LST_CAP = "Select Retailer"
        V_DLR_LST
    ElseIf M_FRM_SOR.ActiveControl.name = "M_SMN_IDY" Then
        V_LST_CAP = "Select  Salesman"
        V_SMN_LST
    ElseIf M_FRM_SOR.ActiveControl.name = "M_ACT_IDY" Then
        V_LST_CAP = "Select Account"
        V_ACT_LST
    ElseIf M_FRM_SOR.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        list4.Tag = ""
        V_PRD_LST
    End If
ElseIf Button.KEY = "DAY" Then
    If Button.Value = tbrPressed Then
        Text1.Visible = True
        Text1 = Format(Date, "dddd")
        Text1.BackColor = vbBlue
        Text1.ForeColor = vbWhite
    Else
        Text1 = ""
        Text1.Visible = False
    End If
ElseIf Button.KEY = "smn1" Then
    If Button.Value = tbrPressed Then
        Text2.Visible = True
        list1.Tag = "Salesman"
        V_LST_CAP = "Select  Salesman"
        V_SMN_LST
        Text2.BackColor = vbBlue
        Text2.ForeColor = vbWhite
    Else
        Text2 = ""
        Text2.Visible = False
    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Picture1.Visible = True
If check("prt", "prt_nme", "M_FRM_SOR") = False Then
Check3.Value = 0
Else
Check3.Value = 1
End If
Check3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Public Sub M_VAN_LST()
If CHECKTB("van.dbf,smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rsVan As New ADODB.Recordset
rsVan.Open "SELECT * FROM VAN,SMN WHERE VAN.SMN_IDY LIKE SMN.SMN_IDY AND VAN_CLS='S' AND VAN_CHK = 'V' GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
If rsVan.RecordCount > 0 Then
    LIST8.ListItems.CLEAR
    For A = 1 To rsVan.RecordCount
        Set ls = LIST8.ListItems.Add(, , Trim(rsVan!INV_IDY))
        ls.ListSubItems.Add , , Format(rsVan!INV_DAT, "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(rsVan("SMN_NME") & "")
        'ls.ListSubItems.Add , , Trim(rsVan("SMN_AD4") & "")
        ls.ListSubItems.Add , , Trim(rsVan("ROU_IDY") & "")
        ls.ListSubItems.Add , , Trim(rsVan("MNL_REF") & "")
        rsVan.MoveNext
    Next
    SSTab1.Tab = 4
    V_LST_CAP = "Select Van Reference"
    LIST8.Visible = True
    LIST8.SetFocus
Else
    MsgBox "Van Reference Not found!", vbExclamation
    SSTab1.Tab = 0
    SendKeys "{home}+{end}"
    M_VAN_IDY.SetFocus
End If
rsVan.Close
Set rsVan = Nothing
End Sub

Public Function GETSALSTK(BRDID As String) As String
With MS
For A = 1 To .Rows - 1
If .TextMatrix(A, 15) = Trim(BRDID) Then
    If pcs_chk.Value = 1 Then
        GETSALSTK = MTS(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 13)))
    Else
        GETSALSTK = Val(.TextMatrix(A, 2))
    End If
    Exit Function
End If
Next
End With
End Function
Public Function CHECKBATCH(BAT_IDY As String) As Double
For A = 1 To MS.Rows - 1
    If (Trim(BAT_IDY) = Trim(MS.TextMatrix(A, 15))) And Val(MS.TextMatrix(A, 2)) > 0 Then
        If pcs_chk.Value = 1 Then
            CHECKBATCH = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
        Else
            CHECKBATCH = Val(MS.TextMatrix(A, 2))
        End If
        Exit Function
    End If
Next
End Function
Public Function GETFRESTK(BRDID As String, STKTYP As String) As String
With MSF
For A = 1 To .Rows - 1
    If .TextMatrix(A, 11) = Trim(BRDID) And STKTYP = .TextMatrix(A, 1) Then
        If pcs_chk.Value = 1 Then
            GETFRESTK = MTS(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 9)))
        Else
            GETFRESTK = Val(.TextMatrix(A, 2))
        End If
    End If
Next
End With
End Function

Public Function CHECKBATCHValue(BAT_IDY As String) As Double
For A = 1 To MS.Rows - 1
    If (Trim(BAT_IDY) = Trim(MS.TextMatrix(A, 15))) And Val(MS.TextMatrix(A, 2)) > 0 Then
        If pcs_chk.Value = 1 Then
            CHECKBATCHValue = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13))) * Round(Val(MS.TextMatrix(A, 5)) / Val(MS.TextMatrix(A, 13)), 2)
        Else
            CHECKBATCHValue = Val(MS.TextMatrix(A, 2)) * Val(MS.TextMatrix(A, 5))
        End If
        Exit Function
    End If
Next
End Function
Public Function chkidy(schid) As Boolean
Dim ls As ListItem
Set ls = SCHLST.FINDITEM(Trim(schid), 1)
If ls Is Nothing Then
    chkidy = False
Else
    chkidy = True
End If
End Function


Public Function CHECKUOM(BATID As String) As Double
For A = 1 To MS.Rows - 1
    If (Trim(BATID) = Trim(MS.TextMatrix(A, 15))) And Val(MS.TextMatrix(A, 3)) > 0 Then
        CHECKUOM = Val(MS.TextMatrix(A, 3))
        Exit Function
    End If
Next
End Function

Public Sub ADDSKUQUANTIY(SCHIDY As String)
If CHECKTB("sch.dbf,prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
With MSS
rs.Open "SELECT PRD.PRD_IDY,BAT_NME,PRD_NME,SCH.BAT_IDY AS BAT_IDY,PRD_MRP,B_PRD_QTY,GRP_IDY,PRD_SDR,TXN_TYP,F_PRD_IDY,SUB_UNT,PRO_RAT,PRD.CON_FAC AS CONFAC FROM SCH,BRD,PRD WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND SCH.BAT_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
        rs.filter = "F_PRD_idy = '" & Trim(MS.TextMatrix(A, 15)) & "' and txn_typ like 'FRE'"
        RSTMP.filter = "BAT_IDY = '" & Trim(MS.TextMatrix(A, 15)) & "' and txn_typ like 'BAS'"
        If rs.RecordCount > 0 And RSTMP.RecordCount > 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = Trim(rs!PRD_NME)
            .TextMatrix(.Rows - 1, 1) = convert(rs!PRD_MRP)
            If rs!PRO_RAT = "Y" Then
                If Val(CHECKBATCH(RSTMP!BAT_IDY)) < Val(RSTMP!B_PRD_QTY) Then
                Exit Sub
                End If
                APPQTY = Int(Val(CHECKBATCH(RSTMP!BAT_IDY)) / Val(RSTMP!B_PRD_QTY))
                APP1QTY = Val(APPQTY) * Val(RSTMP!B_PRD_QTY)
                APPQTY = Val(APPQTY) * Val(rs!B_PRD_QTY)
                ACTQTY = Val(CHECKBATCH(RSTMP!BAT_IDY)) - Val(APP1QTY)
'                ACTQTY = Int(Val(ACTQTY) / Int(Val(RSTMP!B_PRD_QTY) / Val(rs!B_PRD_QTY)))
                ACTQTY = (ACTQTY / RSTMP!B_PRD_QTY) / Val(rs!B_PRD_QTY) ' CHANGED BY PKB INSPITE OF ABOVE COMMENTED LINE
                'ACTQTY = (ACTQTY / rs!B_PRD_QTY) * RSTMP!B_PRD_QTY ' CHANGED BY PKB INSPITE OF ABOVE COMMENTED LINE
                ACTQTY = ACTQTY + Val(APPQTY)
                .TextMatrix(.Rows - 1, 2) = ACTQTY
                .TextMatrix(.Rows - 1, 15) = ACTQTY
                SC.AddItem ""
                SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
                SC.TextMatrix(SC.Rows - 1, 1) = Trim(RSTMP!BAT_IDY)
                SC.TextMatrix(SC.Rows - 1, 2) = Val(APP1QTY) + Val(ACTQTY1)
                SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
            Else
                .TextMatrix(.Rows - 1, 2) = Int(Val(CHECKBATCH(RSTMP!BAT_IDY)) / Val(RSTMP!B_PRD_QTY)) * Val(rs!B_PRD_QTY)
                .TextMatrix(.Rows - 1, 15) = Int(Val(CHECKBATCH(RSTMP!BAT_IDY)) / Val(RSTMP!B_PRD_QTY)) * Val(rs!B_PRD_QTY)
                SC.AddItem ""
                SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
                SC.TextMatrix(SC.Rows - 1, 1) = Trim(RSTMP!BAT_IDY)
                SC.TextMatrix(SC.Rows - 1, 2) = Int(Val(CHECKBATCH(RSTMP!BAT_IDY)) / Val(RSTMP!B_PRD_QTY)) * Val(RSTMP!B_PRD_QTY)
                SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
            End If
            
            .TextMatrix(.Rows - 1, 3) = convert(rs!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rs(0)) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rs(1)) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rs(3)) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = NOT APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(RSTMP!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rs!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rs!GRP_IDY) 'product group
'            If pcs_chk.Value = 1 Then
'                .TextMatrix(.Rows - 1, 16) = QTY(WGTCAL(.TextMatrix(.Rows - 1, 11), Val(RS!CONFAC), Val(.TextMatrix(.Rows - 1, 12))))
'            Else
'                .TextMatrix(.Rows - 1, 16) = QTY(Val(MS.TextMatrix(.Rows - 1, 11)) * convert(Val(RS!CONFAC) / Val(.TextMatrix(.Rows - 1, 12))))
'            End If
            .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
        End If
        rs.filter = ""
    End If
Next
rs.Close
End With
End Sub
Public Sub ADDBILLQUANTIY(SCHIDY As String)
If CHECKTB("sch.dbf,prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
With MSS
rs.Open "SELECT PRD.PRD_IDY,BAT_NME,PRD_NME,SCH.BAT_IDY AS BAT_IDY,PRD_MRP,B_PRD_QTY,GRP_IDY,PRD_SDR,TXN_TYP,F_PRD_IDY,SUB_UNT,B_BIL_AMT,PRD.CON_FAC AS CONFAC FROM SCH,BRD,PRD WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND SCH.BAT_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = Trim(rs!PRD_NME)
    .TextMatrix(.Rows - 1, 1) = convert(rs!PRD_MRP)
    .TextMatrix(.Rows - 1, 2) = Int(Val(T_GRS_AMT) / Val(rs!b_BIL_AMT)) * Val(rs!B_PRD_QTY)
    .TextMatrix(.Rows - 1, 15) = Int(Val(T_GRS_AMT) / Val(rs!b_BIL_AMT)) * Val(rs!B_PRD_QTY)
    .TextMatrix(.Rows - 1, 3) = convert(rs!PRD_SDR)
    .TextMatrix(.Rows - 1, 4) = "0.00"
    .TextMatrix(.Rows - 1, 5) = Trim(rs(0)) 'PRD_IDY
    .TextMatrix(.Rows - 1, 6) = Trim(rs(1)) 'BAT_NME
    .TextMatrix(.Rows - 1, 7) = Trim(rs(3)) 'BAT_IDY
    .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
    '.TextMatrix(.Rows - 1, 9) = NOT APPLICABLE 'TXN_TYP
    .TextMatrix(.Rows - 1, 10) = Val(rs!b_BIL_AMT) ' BAS AMOUNT
    .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
    .TextMatrix(.Rows - 1, 12) = Val(rs!SUB_UNT) 'subunts
    .TextMatrix(.Rows - 1, 13) = Trim(rs!GRP_IDY) 'product group
    .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
    rs.MoveNext
Next
rs.Close
End With
End Sub
Public Sub ADDMLTQTYAND(SCHIDY As String)
If CHECKTB("sch.dbf,brd.dbf,prd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
rs.Open "SELECT PRD.PRD_IDY,BAT_NME,PRD_NME,SCH.BAT_IDY AS BAT_IDY,PRD_MRP,B_PRD_QTY,GRP_IDY,PRD_SDR,TXN_TYP,F_PRD_IDY,SUB_UNT,F_CON_APP,B_CON_APP,PRD.CON_FAC AS CONFAC FROM SCH,BRD,PRD WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND SCH.BAT_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
rs.filter = "B_CON_APP = 'AND'"
'rs.filter = ""
If rs.RecordCount > 0 Then
    RSTMP.filter = "B_CON_APP='AND'"
    rs.filter = "F_CON_APP = 'AND'"
    CTR = 1
    If RSTMP.BOF = False Then RSTMP.MoveFirst
    For A = 1 To 2
        If RSTMP.EOF = True Then
            RSTMP.MoveFirst
            CTR = CTR + 1
        End If
        If (CHECKBATCH(RSTMP!BAT_IDY) - (CTR * RSTMP!B_PRD_QTY)) < 0 Then Exit For
        RSTMP.MoveNext
        A = A - 1
    Next
    
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            .TextMatrix(.Rows - 1, 0) = Trim(rs!PRD_NME)
            .TextMatrix(.Rows - 1, 1) = convert(rs!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR - 1)
    

            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 3) = convert(rs!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rs(0)) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rs(1)) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rs(3)) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rs!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rs!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
            
        rs.MoveNext
        End With
    Next
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        SC.AddItem ""
        SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
        SC.TextMatrix(SC.Rows - 1, 1) = Trim(RSTMP!BAT_IDY)
        SC.TextMatrix(SC.Rows - 1, 2) = Val(RSTMP!B_PRD_QTY) * (CTR - 1)
        SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
        RSTMP.MoveNext
    Next
    
    chk = False
    rs.filter = ""
        
    For A = 1 To rs.RecordCount
        If Trim(rs!F_CON_APP & "") = "ORO" Then
            With MSS
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = Trim(rs!PRD_NME)
            .TextMatrix(.Rows - 1, 1) = convert(rs!PRD_MRP)
            'If chk = False Then
                .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY & "") * (CTR - 1)
                .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY & "") * (CTR - 1)
                SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
                'chk = True
            'Else
                '.TextMatrix(.Rows - 1, 2) = 0
                '.TextMatrix(.Rows - 1, 15) = 0
            'End If
            .TextMatrix(.Rows - 1, 3) = convert(rs!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rs(0)) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rs(1)) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rs(3)) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            .TextMatrix(.Rows - 1, 9) = "AND"
            '.TextMatrix(.Rows - 1, 10) = Val(RS!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rs!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rs!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = "ORO" 'f_CON_APP
            If rs.EOF = False Then
            .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
            End If
            End With
        End If
        rs.MoveNext
    Next
    'For A = 1 To MSS.Rows - 1
     '   If MSS.TextMatrix(A, 8) = SCHIDY And chk = True Then MSS.TextMatrix(A, 11) = SCHQTY
'    Next
Else
    rs.filter = "B_CON_APP = 'ORO'"
    Set RSTMP = rs.Clone
    RSTMP.filter = "F_CON_APP = 'AND'"
    If rs.RecordCount > 0 And RSTMP.RecordCount > 0 Then
        MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
        For A = 1 To rs.RecordCount
            If CHECKBATCH(rs!BAT_IDY) >= Val(MAXVALUE) Then
                SC.AddItem ""
                SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
                SC.TextMatrix(SC.Rows - 1, 1) = Trim(rs!BAT_IDY)
                SC.TextMatrix(SC.Rows - 1, 2) = Int(CHECKBATCH(rs!BAT_IDY) / Val(MAXVALUE)) * Val(MAXVALUE)
                SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
                RSTMP.MoveFirst
                For b = 1 To RSTMP.RecordCount
                    With MSS
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = Trim(RSTMP!PRD_NME)
                    .TextMatrix(.Rows - 1, 1) = convert(RSTMP!PRD_MRP)
                    If Val(RSTMP!B_PRD_QTY) > 0 Then
                        .TextMatrix(.Rows - 1, 2) = Val(RSTMP!B_PRD_QTY) * Int(CHECKBATCH(rs!BAT_IDY) / Val(MAXVALUE))
                        .TextMatrix(.Rows - 1, 15) = Val(RSTMP!B_PRD_QTY) * Int(CHECKBATCH(rs!BAT_IDY) / Val(MAXVALUE))
                    Else
                        .TextMatrix(.Rows - 1, 2) = 0
                        .TextMatrix(.Rows - 1, 15) = 0
                    End If
                    .TextMatrix(.Rows - 1, 3) = convert(RSTMP!PRD_SDR)
                    .TextMatrix(.Rows - 1, 4) = "0.00"
                    .TextMatrix(.Rows - 1, 5) = Trim(RSTMP(0)) 'PRD_IDY
                    .TextMatrix(.Rows - 1, 6) = Trim(RSTMP(1)) 'BAT_NME
                    .TextMatrix(.Rows - 1, 7) = Trim(RSTMP(3)) 'BAT_IDY
                    .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
                    '.TextMatrix(.Rows - 1, 9) = NOT APPLICABLE 'TXN_TYP
                    .TextMatrix(.Rows - 1, 10) = Val(RSTMP!B_PRD_QTY) ' BAS QUANTITY
                    .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
                    .TextMatrix(.Rows - 1, 12) = Val(RSTMP!SUB_UNT) 'subunts
                    .TextMatrix(.Rows - 1, 13) = Trim(RSTMP!GRP_IDY) 'product group
                    .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
                    End With
                    RSTMP.MoveNext
                Next
            End If
            rs.MoveNext
        Next
    Else
        rs.filter = "B_CON_APP = 'ORO'"
        RSTMP.filter = "F_CON_APP = 'ORO'"
        If rs.RecordCount > 0 And RSTMP.RecordCount > 0 Then
            rs.MoveFirst
            RSTMP.MoveFirst
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
            FREMAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'FRE'")
            With MSS
            For A = 1 To rs.RecordCount
                STK = Val(STK) + Val(GETSALSTK(rs!BAT_IDY))
                'If Val(GETSALSTK(Trim(rs!BAT_IDY))) >= Val(MAXVALUE) Then 'ORIGINALLY WRITTEN CODE
                If Val(STK) >= Val(GETSALSTK(rs!BAT_IDY)) Then 'CORRECTED CODE OF PNR
                    SC.AddItem ""
                    SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
                    SC.TextMatrix(SC.Rows - 1, 1) = Trim(rs!BAT_IDY)
                    'SC.TextMatrix(SC.Rows - 1, 2) = Int(CHECKBATCH(rs!BAT_IDY) / Val(MAXVALUE)) * Val(MAXVALUE) ' ORIGINALLY WRITTEN CODE
                    SC.TextMatrix(SC.Rows - 1, 2) = Val(GETSALSTK(Trim(rs!BAT_IDY))) 'CORRECTED CODE OF PNR
                    SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
                End If
                rs.MoveNext
            Next
            STK = Int(Val(STK) / Val(MAXVALUE))
            
            For b = 1 To RSTMP.RecordCount
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = Trim(RSTMP!PRD_NME)
                .TextMatrix(.Rows - 1, 1) = convert(RSTMP!PRD_MRP)
                If b = 1 Then
                    .TextMatrix(.Rows - 1, 2) = Val(FREMAXVALUE & "") * STK
                    .TextMatrix(.Rows - 1, 15) = Val(FREMAXVALUE & "") * STK
                Else
                    .TextMatrix(.Rows - 1, 2) = 0
                    .TextMatrix(.Rows - 1, 15) = 0
                End If
                .TextMatrix(.Rows - 1, 3) = convert(RSTMP!PRD_SDR)
                .TextMatrix(.Rows - 1, 4) = "0.00"
                .TextMatrix(.Rows - 1, 5) = Trim(RSTMP("PRD_IDY")) 'PRD_IDY
                .TextMatrix(.Rows - 1, 6) = Trim(RSTMP(1)) 'BAT_NME
                .TextMatrix(.Rows - 1, 7) = Trim(RSTMP(3)) 'BAT_IDY
                .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
                .TextMatrix(.Rows - 1, 9) = "ORO" 'BASE CONDITION
                .TextMatrix(.Rows - 1, 10) = Val(RSTMP!B_PRD_QTY) ' BAS QUANTITY
                .TextMatrix(.Rows - 1, 11) = Val(FREMAXVALUE) * STK
                .TextMatrix(.Rows - 1, 12) = Val(RSTMP!SUB_UNT) 'subunts
                .TextMatrix(.Rows - 1, 13) = Trim(RSTMP!GRP_IDY) 'product group
                .TextMatrix(.Rows - 1, 14) = "ORO" 'FREE CONDITION
                'CORRECTED WHEN Ramanesh sir asked from Bombay
                If rs.EOF = False Then
                .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
                End If
                RSTMP.MoveNext
            Next
            End With
        End If
    End If
End If
rs.Close
ADDSKUMIX SCHIDY
End Sub
Private Sub ADDSKUMIX(SCHIDY As String)
If CHECKTB("sch.dbf,brd.dbf,prd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
rs.Open "SELECT PRD.PRD_IDY,BAT_NME,PRD_NME,SCH.BAT_IDY AS BAT_IDY,PRD_MRP,B_PRD_QTY,GRP_IDY,PRD_SDR,TXN_TYP,F_PRD_IDY,SUB_UNT,F_CON_APP,B_CON_APP,PRD.CON_FAC AS CONFAC FROM SCH,BRD,PRD WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND SCH.BAT_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
rs.filter = "B_CON_APP = 'ORM'"
RSTMP.filter = "F_CON_APP = 'AND'"
MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'")

If rs.RecordCount > 0 And RSTMP.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        STK = Val(STK) + CHECKBATCH(rs!BAT_IDY)
        rs.MoveNext
    Next
    If Int(STK / MAXVALUE) > 0 Then
        rs.MoveFirst
        For A = 1 To rs.RecordCount
            SC.AddItem ""
            SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
            SC.TextMatrix(SC.Rows - 1, 1) = Trim(rs!BAT_IDY)
            SC.TextMatrix(SC.Rows - 1, 2) = Int(CHECKBATCH(rs!BAT_IDY))
            SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
            rs.MoveNext
        Next
    End If
    If rs.RecordCount > 0 Then rs.MoveFirst
    For A = 1 To RSTMP.RecordCount
        With MSS
        .AddItem ""
            .TextMatrix(.Rows - 1, 0) = Trim(RSTMP!PRD_NME)
            .TextMatrix(.Rows - 1, 1) = convert(RSTMP!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(RSTMP!B_PRD_QTY) * Int(STK / MAXVALUE)
            .TextMatrix(.Rows - 1, 15) = Val(RSTMP!B_PRD_QTY) * Int(STK / MAXVALUE)
            .TextMatrix(.Rows - 1, 3) = convert(RSTMP!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(RSTMP(0)) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(RSTMP(1)) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(RSTMP(3)) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            .TextMatrix(.Rows - 1, 9) = "NotAPPLICABLE" 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(RSTMP!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(RSTMP!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(RSTMP!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(RSTMP!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
        RSTMP.MoveNext
        End With
    Next
Else
    rs.filter = "B_CON_APP = 'ORM'"
    RSTMP.filter = "F_CON_APP = 'ORO'"
    If rs.RecordCount > 0 And RSTMP.RecordCount > 0 Then
        MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
        FREMAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'FRE'")
        With MSS
        For A = 1 To rs.RecordCount
            STK = Val(STK) + Val(GETSALSTK(rs!BAT_IDY))
            rs.MoveNext
        Next
        If Int(STK / MAXVALUE) > 0 Then
            rs.MoveFirst
            For A = 1 To rs.RecordCount
                SC.AddItem ""
                SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
                SC.TextMatrix(SC.Rows - 1, 1) = Trim(rs!BAT_IDY)
                SC.TextMatrix(SC.Rows - 1, 2) = Int(CHECKBATCH(rs!BAT_IDY))
                SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
                rs.MoveNext
            Next
        End If
        STK = Int(Val(STK) / Val(MAXVALUE))
        If rs.RecordCount > 0 Then rs.MoveFirst
        For b = 1 To RSTMP.RecordCount
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = Trim(RSTMP!PRD_NME)
            .TextMatrix(.Rows - 1, 1) = convert(RSTMP!PRD_MRP)
            If b = 1 Then
                .TextMatrix(.Rows - 1, 2) = Val(FREMAXVALUE & "") * STK
                .TextMatrix(.Rows - 1, 15) = Val(FREMAXVALUE & "") * STK
            Else
                .TextMatrix(.Rows - 1, 2) = 0
                .TextMatrix(.Rows - 1, 15) = 0
            End If
            .TextMatrix(.Rows - 1, 3) = convert(RSTMP!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(RSTMP("PRD_IDY")) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(RSTMP(1)) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(RSTMP(3)) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            .TextMatrix(.Rows - 1, 9) = "ORO" 'BASE CONDITION
            .TextMatrix(.Rows - 1, 10) = Val(RSTMP!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(FREMAXVALUE) * STK
            .TextMatrix(.Rows - 1, 12) = Val(RSTMP!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(RSTMP!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = "ORO" 'FREE CONDITION
            .TextMatrix(.Rows - 1, 16) = (Val(.TextMatrix(.Rows - 1, 2)) * Val(Trim(rs!CONFAC))) / (Val(.TextMatrix(.Rows - 1, 12)))
            RSTMP.MoveNext
        Next
        End With
    End If
End If
End Sub
Public Sub ADDSKUAMOUNT(SCHIDY As String)
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
rs.filter = "TXN_TYP='BAS'"
salqty = CHECKBATCH(rs!BAT_IDY)
BATIDY = Trim(rs!BAT_IDY)
SCHQTY = Val(rs!B_PRD_QTY)
rs.filter = "TXN_TYP='CAS'"
If rs.RecordCount > 0 Then
    SCHAMT = Val(rs!B_DIS_AMT & "")
    If rs!PRO_RAT = "Y" Then
        If Val(salqty) >= Val(SCHQTY) Then
            SCAMT = Val(Round((Val(SCHAMT) / Val(SCHQTY)) * Val(salqty), 2))
            SCH_AMT = convert(Val(SCH_AMT) + Val(SCAMT))
        End If
    Else
        SCAMT = Round(Int(Val(salqty) / Val(SCHQTY)) * Val(SCHAMT))
        SCH_AMT = convert(Val(SCH_AMT) + Val(SCAMT))
    End If
    SC.AddItem ""
    SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
    SC.TextMatrix(SC.Rows - 1, 1) = BATIDY 'FREPRDIDY
    SC.TextMatrix(SC.Rows - 1, 3) = Val(SCAMT) 'LESS AMT
    SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
End If
End Sub
Public Sub ADDSKUPCG(SCHIDY As String)
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
rs.filter = "TXN_TYP='BAS'"
BATIDY = rs!BAT_IDY
salqty = CHECKBATCHValue(rs!BAT_IDY)
SCHQTY = Val(rs!B_PRD_QTY)
rs.filter = "TXN_TYP='PCG'"
If rs.RecordCount > 0 Then
    SCHPCG = Val(rs!B_DIS_PCG & "")
    SCHAMT = PERCENT(Val(salqty), Val(rs!B_DIS_PCG & ""))
    'If rs!PRO_RAT = "Y" Then
        'If Val(salqty) >= Val(SCHQTY) Then
        '    SCAMT = Val(Round((Val(SCHAMT) / Val(SCHQTY)) * Val(salqty), 2))
        '    SCH_AMT = convert(Val(SCH_AMT) + Val(SCAMT))
        'End If
    'Else
    '    SCAMT = Round(Int(Val(salqty) / Val(SCHQTY)) * Val(SCHAMT), 2)
'    '    SCH_AMT = convert(Val(SCH_AMT) + Val(SCAMT))
        SCH_AMT = convert(Val(SCH_AMT) + Val(SCHAMT))
    'End If
    SC.AddItem ""
    SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
    SC.TextMatrix(SC.Rows - 1, 1) = BATIDY 'FREPRDIDY
    SC.TextMatrix(SC.Rows - 1, 3) = Val(SCHAMT) 'LESS AMT
    SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
End If
End Sub

Public Sub ADDBILAMT(SCHIDY As String)
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
SCHAMT = Val(rs!B_DIS_AMT & "")
BILAMT = Val(rs!b_BIL_AMT & "")
SALAMT = Val(T_GRS_AMT)
If rs!PRO_RAT = "Y" Then
    If Val(SALAMT) >= Val(BILAMT) Then
        SCAMT = Round(Int(Val(SALAMT) / Val(BILAMT)) * Val(SCHAMT))
        BILMT = Val(SALAMT) - convert((Int(Val(SALAMT) / Val(BILAMT)) * Val(BILAMT)))
        BILMT = convert((Val(BILMT) / Val(BILAMT)) * Val(SCHAMT))
        SCH_AMT = convert(Val(SCH_AMT) + (Val(SCAMT) + Val(BILMT)))
    End If
Else
    SCAMT = Round(Int(Val(SALAMT) / Val(BILAMT)) * Val(SCHAMT))
    SCH_AMT = convert(Val(SCH_AMT) + Val(SCAMT))
End If
SC.AddItem ""
SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
SC.TextMatrix(SC.Rows - 1, 3) = Val(SCAMT) + Val(BILMT) 'LESS AMT
SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
End Sub

Public Sub ADDBILPCG(SCHIDY As String)
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
SCHAMT = Val(rs!B_DIS_PCG & "")
SALAMT = Val(T_GRS_AMT)
SCH_AMT = convert(Val(SCH_AMT) + Val(PERCENT(Val(T_GRS_AMT), Val(SCHAMT))))
SC.AddItem ""
SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
SC.TextMatrix(SC.Rows - 1, 3) = Val(PERCENT(Val(T_GRS_AMT), Val(SCHAMT))) 'LESS AMT
SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
End Sub

Public Sub MLTBILPCGAMT(SCHIDY As String)
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
LSCH_AMT = 0#
'SCH_AMT = 0#
'SAMT = 0#
'SCAMT = 0#
'SCHAMT = 0#
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
rs.filter = "TXN_TYP='CAS'"
If rs.RecordCount > 0 Then
    disamt = Val(rs!B_DIS_AMT & "")
    PRO = Trim(rs!PRO_RAT)
    MAXQTY = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'")
Else
    rs.filter = "TXN_TYP='PCG'"
    If rs.RecordCount > 0 Then dispcg = Val(rs!B_DIS_PCG & "")
End If
rs.filter = "TXN_TYP='BAS'"
If rs!B_CON_APP = "AND" Then
    rs.filter = "TXN_TYP='BAS'"
    CTR = 1
    PRDVAL = 0
    For A = 1 To 2
        If rs.EOF = True Then
            rs.MoveFirst
            CTR = CTR + 1
        End If
        If (CHECKBATCH(rs!BAT_IDY) - (CTR * rs!B_PRD_QTY)) < 0 Then Exit For
        rs.MoveNext
        A = A - 1
    Next
    If rs.BOF = False Then rs.MoveFirst
    For A = 1 To rs.RecordCount
        PRDVAL = Val(PRDVAL) + Val(GETSCHVAL(rs!BAT_IDY))
        rs.MoveNext
    Next
    SAMT = Val(SCH_AMT)
    SAMT = Val(LSCH_AMT)
    If Val(dispcg) > 0 Then
        SCH_AMT = Val(SCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg))
        LSCH_AMT = Val(LSCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg))
    Else
        SCH_AMT = convert(Val(SCH_AMT) + (Val(disamt) * (CTR - 1)))
        LSCH_AMT = convert(Val(LSCH_AMT) + (Val(disamt) * (CTR - 1)))
    End If
    SC.AddItem ""
    SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
    SC.TextMatrix(SC.Rows - 1, 3) = Val(LSCH_AMT) 'Val(SAMT) - Val(SCH_AMT)
    SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
ElseIf rs!B_CON_APP = "ORO" Then
    rs.filter = "B_CON_APP='ORO'"
    SCAMT = SCHAMT
    For A = 1 To rs.RecordCount
        b = CHECKBATCH(Trim(rs!BAT_IDY))
        PRDVAL = Val(PRDVAL) + Val(GETSCHVAL(rs!BAT_IDY))
        If Val(dispcg) > 0 Then
            SCH_AMT = Val(SCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg))
            LSCH_AMT = Val(LSCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg))
        Else
            If Val(b) = Val(MAXQTY) Then
                SCH_AMT = convert(Val(SCH_AMT) + Val(disamt))
                LSCH_AMT = convert(Val(LSCH_AMT) + Val(disamt))
            ElseIf Val(b) > Val(MAXQTY) Then
                SCH_AMT = convert(Val(SCH_AMT) + (Int(Val(b) / Val(MAXQTY)) * (disamt)))
                LSCH_AMT = convert(Val(LSCH_AMT) + (Int(Val(b) / Val(MAXQTY)) * (disamt)))
            End If
        End If
        rs.MoveNext
    Next
    SC.AddItem ""
    SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
    SC.TextMatrix(SC.Rows - 1, 3) = Val(LSCH_AMT) '(Val(SCAMT) - Val(SCH_AMT))
    SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
ElseIf rs!B_CON_APP = "ORM" Then
    rs.filter = "B_CON_APP='ORM'"
    SCAMT = SCHAMT
    For A = 1 To rs.RecordCount
        b = CHECKBATCHValue(Trim(rs!BAT_IDY))
        PRDVAL = Val(PRDVAL) + Val(GETSCHVAL(rs!BAT_IDY))
        If Val(dispcg) > 0 Then
            SCH_AMT = Val(SCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg))
            LSCH_AMT = Val(LSCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg))
        Else
            If Val(b) = Val(MAXQTY) Then
                SCH_AMT = convert(Val(SCH_AMT) + Val(disamt))
                LSCH_AMT = convert(Val(LSCH_AMT) + Val(disamt))
            ElseIf Val(b) > Val(MAXQTY) Then
                PNR_SUT = GetValue("SKU_SBU", "BRD WHERE BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")
                PNR_DRT = GetValue("PRD_SDR", "BRD WHERE BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")
                MAXQTY = STM(Val(MAXQTY), Val(PNR_SUT))
'                SCH_AMT = convert(Val(SCH_AMT) + (Int(Val(b) / Val(MAXQTY)) * (disamt)))
                SCH_AMT = convert(Val(SCH_AMT) + ((Int(Val(b)) / Int(PNR_DRT)) * disamt))
                LSCH_AMT = convert(Val(LSCH_AMT) + ((Int(Val(b)) / Int(PNR_DRT)) * disamt))
            End If
        End If
        rs.MoveNext
    Next
    SC.AddItem ""
    SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
    If Val(SCAMT) >= Val(SCH_AMT) Then
        SC.TextMatrix(SC.Rows - 1, 3) = (Val(SCAMT) - Val(LSCH_AMT))
    Else
        SC.TextMatrix(SC.Rows - 1, 3) = Val(LSCH_AMT)
    End If
    SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
End If
End Sub
Public Function GETSCHVAL(BAT_IDY) As String
For A = 1 To MS.Rows - 1
    If (Trim(BAT_IDY) = Trim(MS.TextMatrix(A, 15))) And Val(MS.TextMatrix(A, 2)) > 0 Then
        GETSCHVAL = Val(MS.TextMatrix(A, 9))
        Exit Function
    End If
Next
End Function

Private Sub TXT_SCH_Keypress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K <> 13 Then K = CHKNUM(K)
With MSS
If K = 13 Then
    If MSS.Col = 2 Then
        If Trim(.TextMatrix(.Row, 14)) = "ORO" And Trim(.TextMatrix(.Row, 9)) = "ORO" Then
            If (Val(.TextMatrix(.Row, 11)) + Val(MSS.Text)) - (Val(TXT_SCH) + Val(SCHEMEQTY(Trim(.TextMatrix(.Row, 8))))) < 0 Then
                TXT_SCH.Visible = False
                MSS.SetFocus
            Else
                MSS.Text = TXT_SCH
                TXT_SCH.Visible = False
                MSS.SetFocus
            End If
        ElseIf Trim(.TextMatrix(.Row, 14)) = "ORO" And Trim(.TextMatrix(.Row, 9)) = "ORM" Then
            If (Val(.TextMatrix(.Row, 11)) + Val(MSS.Text)) - (Val(TXT_SCH) + Val(SCHEMEQTY(Trim(.TextMatrix(.Row, 8))))) < 0 Then
                TXT_SCH.Visible = False
                MSS.SetFocus
            Else
                MSS.Text = TXT_SCH
                TXT_SCH.Visible = False
                MSS.SetFocus
            End If
        ElseIf Trim(.TextMatrix(.Row, 14)) = "ORO" And Trim(.TextMatrix(.Row, 9)) = "AND" Then
            If (getSCHEMEQTY(Trim(.TextMatrix(.Row, 8))) + Val(MSS.Text)) - (Val(TXT_SCH) + Val(SCHEMEQTY(Trim(.TextMatrix(.Row, 8))))) < 0 Then
                TXT_SCH.Visible = False
                MSS.SetFocus
            Else
                MSS.Text = TXT_SCH
                TXT_SCH.Visible = False
                MSS.SetFocus
            End If
        Else
            If Val(MSS.TextMatrix(MSS.Row, 15)) >= Val(TXT_SCH) Then
                MSS.Text = TXT_SCH
                TXT_SCH.Visible = False
                MSS.SetFocus
            Else
                MsgBox "Invalid Scheme Quantity !", vbExclamation, "DAS Version 4.0"
                TXT_SCH.Visible = True
                TXT_SCH.SetFocus
                SendKeys "{home}+{end}"
            End If
        End If
    ElseIf MSS.Col = 4 Then
            SCH_AMT = Val(SCH_AMT) - Val(MSS.Text)
            MSS.Text = convert(TXT_SCH)
            SCH_AMT = convert(Val(SCH_AMT) + Val(MSS.Text))
            TXT_SCH.Visible = False
            MSS.SetFocus
    End If
End If
End With
End Sub

Public Function STOCKCHECK() As Integer
If CHECKTB("brd.dbf,inv.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Function
End If
With MSS
For A = 1 To .Rows - 1
    FREQTY = GETFRESTK(Trim(.TextMatrix(A, 7)), "Free")
    salqty = GETSALSTK(Trim(.TextMatrix(A, 7)))
    SCHQTY = Val(.TextMatrix(A, 2))
    BRDQTY = GetValue("PRD_QTY", "BRD WHERE BAT_IDY LIKE '" & Trim(.TextMatrix(A, 7)) & "'")
    BRDQTY = MTS(Val(BRDQTY), .TextMatrix(A, 12))
    If MOD_REC = True Then d = GetValue("sum(tot_pcs)", "inv where inv_idy like '" & Trim(M_INV_IDY) & "' and stk_typ in ('Salable','Free') and bat_idy like '" & Trim(.TextMatrix(A, 7)) & "'")
    If (Val(BRDQTY) + Val(d)) - (Val(FREQTY) + Val(salqty) + Val(SCHQTY)) < 0 Then
        STOCKCHECK = A
        Exit Function
    End If
Next
    STOCKCHECK = 0
End With
End Function
Public Function SCHEMEQTY(SCHIDY) As Long
For A = 1 To MSS.Rows - 1
    If Trim(MSS.TextMatrix(A, 8)) = Trim(SCHIDY) Then SCHQTY = Val(SCHQTY) + Val(MSS.TextMatrix(A, 2))
Next
SCHEMEQTY = Val(SCHQTY)
End Function
Public Function getSCHEMEQTY(SCHIDY) As Long
For A = 1 To MSS.Rows - 1
    If Trim(MSS.TextMatrix(A, 8)) = Trim(SCHIDY) And Val(MSS.TextMatrix(A, 11)) > 0 Then SCHQTY = Val(MSS.TextMatrix(A, 11))
Next
getSCHEMEQTY = Val(SCHQTY)
End Function
Private Sub LOADROUTES()
If rsrdy.State = 1 Then rsrdy.Close
rsrdy.Open "SELECT ROU_IDY FROM RDY GROUP BY ROU_IDY", con
List9.AddItem "All"
While Not rsrdy.EOF
    List9.AddItem rsrdy(0)
    rsrdy.MoveNext
Wend
End Sub

Public Sub VLD_FRE_QTY()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim Sch_Idy
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'SKU Level' and TXN_TYP LIKE 'BAS'", con, adOpenKeyset, adLockPessimistic
SCHLST.ListItems.CLEAR
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 2)) > 0 Then
        rs.filter = "bat_idy ='" & Trim(MS.TextMatrix(A, 15)) & "'"
        For b = 1 To rs.RecordCount
            If chkidy(Trim(rs!Sch_Idy)) = False Then
                SUBUNT = Val(MS.TextMatrix(A, 2))
                If pcs_chk.Value = 1 Then SUBUNT = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
                If Val(SUBUNT) >= Val(rs!B_PRD_QTY) Then
                    If check("SID", "SCH_IDY", rs!Sch_Idy) = False Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(rs!Sch_Nme))
                        ls.ListSubItems.Add , , Format(rs!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(rs!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(rs!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(rs!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(rs!Sch_Idy)
                    Else
                        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'")
                        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'")
                        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'")
                        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                            Set ls = SCHLST.ListItems.Add(, , Trim(rs!Sch_Nme))
                            ls.ListSubItems.Add , , Format(rs!sch_bdt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Format(rs!sch_edt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Trim(rs!SCH_TYP)
                            ls.ListSubItems.Add , , Trim(rs!f_sch_typ)
                            ls.ListSubItems.Add , , Trim(rs!Sch_Idy)
                        End If
                    End If
                End If
            End If
            rs.MoveNext
        Next
        rs.filter = ""
    End If
Next
rs.Close
Set rs = Nothing
''RS1.Close
Set RS1 = Nothing
SCHLST.SetFocus
End Sub

Public Sub VLD_BIL_AMT()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT b_bil_amt,sch_idy,sch_nme,sch_bdt,sch_edt,sch_typ,f_sch_typ FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Bill Level' and b_bil_amt <=" & Val(T_GRS_AMT) & "", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If chkidy(rs!Sch_Idy) = False Then
        If check("SID", "SCH_IDY", rs!Sch_Idy) = False Then
            Set ls = SCHLST.ListItems.Add(, , Trim(rs!Sch_Nme))
            ls.ListSubItems.Add , , Format(rs!sch_bdt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Format(rs!sch_edt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Trim(rs!SCH_TYP)
            ls.ListSubItems.Add , , Trim(rs!f_sch_typ)
            ls.ListSubItems.Add , , Trim(rs!Sch_Idy)
        Else
            smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'")
            ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'")
            CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'")
            If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                Set ls = SCHLST.ListItems.Add(, , Trim(rs!Sch_Nme))
                ls.ListSubItems.Add , , Format(rs!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(rs!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(rs!SCH_TYP)
                ls.ListSubItems.Add , , Trim(rs!f_sch_typ)
                ls.ListSubItems.Add , , Trim(rs!Sch_Idy)
            End If
        End If
    End If
    rs.MoveNext
Next
'rs.Close
Set rs = Nothing
End Sub

Public Sub VLD_MLT_QTY()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim RSS As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT SCH_IDY FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Multiple SKU Level' and b_con_app ='AND' AND TXN_TYP LIKE 'BAS' GROUP by sch_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If RSS.State = 1 Then RSS.Close
    RSS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(rs(0)) & "' AND TXN_TYP LIKE 'BAS'", con, adOpenKeyset
    For b = 1 To RSS.RecordCount
        If CHECKBATCH(RSS!BAT_IDY) < Val(RSS!B_PRD_QTY) Then GoTo SKIPSCH
        RSS.MoveNext
    Next
    RSS.MoveFirst
    If check("SID", "SCH_IDY", RSS!Sch_Idy) = False Then
        Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
        ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
        ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
        ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
        ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
        ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
    Else
        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
            Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
            ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
            ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
            ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
        End If
    End If
SKIPSCH:
    rs.MoveNext
Next
'rs.Close
Set rs = Nothing
End Sub

Public Sub VLD_MLT_QTY_OR()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim RSS As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT SCH_IDY FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Multiple SKU Level' and b_con_app ='ORO' AND TXN_TYP LIKE 'BAS' GROUP by sch_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If RSS.State = 1 Then RSS.Close
    C = 0
    RSS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(rs(0)) & "' AND TXN_TYP LIKE 'BAS'", con, adOpenKeyset
    For b = 1 To RSS.RecordCount
        C = Val(C) + CHECKBATCH(RSS!BAT_IDY)
        RSS.MoveNext
    Next
    If Val(C) >= GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'AND TXN_TYP LIKE 'BAS'") Then
        RSS.MoveFirst
        If check("SID", "SCH_IDY", RSS!Sch_Idy) = False Then
            Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
            ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
            ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
            ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
        Else
            smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
                ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
                ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
                ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
            End If
        End If
    End If
    rs.MoveNext
Next
'rs.Close
Set rs = Nothing
End Sub

Public Sub VLD_MLT_QTY_OR_MIX()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim RSS As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT SCH_IDY FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Multiple SKU Level' and b_con_app ='ORM' AND TXN_TYP LIKE 'BAS' GROUP by sch_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If RSS.State = 1 Then RSS.Close
    C = 0
    RSS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(rs(0)) & "' AND TXN_TYP LIKE 'BAS'", con, adOpenKeyset
    For b = 1 To RSS.RecordCount
        C = Val(C) + CHECKBATCH(RSS!BAT_IDY)
        RSS.MoveNext
    Next
    If Val(C) >= GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "'AND TXN_TYP LIKE 'BAS'") Then
        RSS.MoveFirst
        If check("SID", "SCH_IDY", RSS!Sch_Idy) = False Then
            Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
            ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
            ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
            ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
        Else
            smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
                ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
                ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
                ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
            End If
        End If
    End If
    rs.MoveNext
Next
'rs.Close
'RSS.Close
Set rs = Nothing
Set RSS = Nothing
End Sub

Public Sub VLD_MLT_UOM()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim RSS As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT SCH_IDY FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'UOM Product Group' and b_con_app ='AND' AND TXN_TYP LIKE 'BAS' GROUP by sch_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If RSS.State = 1 Then RSS.Close
    RSS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(rs(0)) & "' AND TXN_TYP LIKE 'BAS'", con, adOpenKeyset
    For b = 1 To RSS.RecordCount
        If CHECKUOM(RSS!BAT_IDY) < Val(RSS!B_PRD_QTY) Then GoTo SKIPSCH
        RSS.MoveNext
    Next
    RSS.MoveFirst
    If check("SID", "SCH_IDY", RSS!Sch_Idy) = False Then
        Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
        ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
        ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
        ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
        ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
        ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
    Else
        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
            Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
            ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
            ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
            ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
        End If
    End If
SKIPSCH:
    rs.MoveNext
Next
'rs.Close
'RSS.Close
Set rs = Nothing
Set RSS = Nothing
End Sub

Public Sub VLD_MLT_UOM_OR()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim RSS As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT SCH_IDY FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'UOM Product Group' and b_con_app ='ORO' AND TXN_TYP LIKE 'BAS' GROUP by sch_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If RSS.State = 1 Then RSS.Close
    RSS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(rs(0)) & "' AND TXN_TYP LIKE 'BAS'", con, adOpenKeyset
    For b = 1 To RSS.RecordCount
        C = Val(C) + CHECKUOM(RSS!BAT_IDY)
        RSS.MoveNext
    Next
If Val(C) >= GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' and txn_typ ='BAS'") Then
RSS.MoveFirst
        If check("SID", "SCH_IDY", RSS!Sch_Idy) = False Then
            Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
            ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
            ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
            ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
        Else
            smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(RSS!Sch_Idy) & "'")
            If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                Set ls = SCHLST.ListItems.Add(, , Trim(RSS!Sch_Nme))
                ls.ListSubItems.Add , , Format(RSS!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(RSS!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(RSS!SCH_TYP)
                ls.ListSubItems.Add , , Trim(RSS!f_sch_typ)
                ls.ListSubItems.Add , , Trim(RSS!Sch_Idy)
            End If
        End If
    End If
rs.MoveNext
Next
'rs.Close
'RSS.Close
Set rs = Nothing
Set RSS = Nothing
End Sub

Public Sub VLD_BRAND()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim Sch_Idy
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim REC_SET  As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Brand' GROUP BY SCH_IDY", con, adOpenKeyset, adLockPessimistic
For C = 1 To REC_SET.RecordCount
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Brand' and TXN_TYP LIKE 'BAS' AND SCH_IDY = '" & Trim(REC_SET!Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If rs!B_CON_APP = "AND" Then
            For A = 1 To rs.RecordCount
                GRPVAL = GETGRPVALUE("Brand", Trim(rs!grp_typ & ""))
                If GRPVAL >= Val(rs!B_PRD_QTY & "") Then
                            
                Else
                    GoTo M
                End If
                rs.MoveNext
            Next
            
        If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
            Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
            ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
            ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
            ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
            ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
        Else
            smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
            ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
            CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
            If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
            End If
        End If
    ElseIf rs!B_CON_APP = "ORO" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            For A = 1 To rs.RecordCount
                If GETGRPVALUE("Brand", Trim(rs!grp_typ & "")) >= Val(MAXVALUE) Then
                    If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    Else
                        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                            Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                            ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                            ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                            ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                        End If
                    End If
                    Exit For
                End If
                rs.MoveNext
            Next
        ElseIf rs!B_CON_APP = "ORM" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            stkcount = 0
            For A = 1 To rs.RecordCount
                stkcount = Val(stkcount) + GETGRPVALUE("Brand", Trim(rs!grp_typ & ""))
                If Val(stkcount) >= Val(MAXVALUE) Then
                    If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    Else
                        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                            Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                            ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                            ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                            ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                        End If
                    End If
                    Exit For
                End If
                rs.MoveNext
            Next
        End If
    End If
M:
    REC_SET.MoveNext
Next
REC_SET.Close
'rs.Close
Set REC_SET = Nothing
Set rs = Nothing
End Sub

Public Sub VLD_PackSize()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim Sch_Idy
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Pack Size' GROUP BY SCH_IDY", con, adOpenKeyset, adLockPessimistic
For C = 1 To REC_SET.RecordCount
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Pack Size' and TXN_TYP LIKE 'BAS' AND SCH_IDY = '" & Trim(REC_SET!Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If rs!B_CON_APP = "AND" Then
            For A = 1 To rs.RecordCount
                GRPVAL = GETGRPVALUE("Pack Size", Trim(rs!grp_typ & ""))
                If GRPVAL >= Val(rs!B_PRD_QTY & "") Then
                            
                Else
                    GoTo M
                End If
                rs.MoveNext
            Next
            If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
            Else
                smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                End If
            End If
        ElseIf rs!B_CON_APP = "ORO" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            For A = 1 To rs.RecordCount
                If GETGRPVALUE("Pack Size", Trim(rs!grp_typ & "")) >= Val(MAXVALUE) Then
                If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                Else
                    smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    End If
                End If
                Exit For
                End If
                rs.MoveNext
            Next
        ElseIf rs!B_CON_APP = "ORM" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            stkcount = 0
            For A = 1 To rs.RecordCount
                stkcount = Val(stkcount) + GETGRPVALUE("Pack Size", Trim(rs!grp_typ & ""))
                If Val(stkcount) >= Val(MAXVALUE) Then
                If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                Else
                    smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    End If
                End If
                Exit For
                End If
                rs.MoveNext
            Next
        End If
    End If
M:
    REC_SET.MoveNext
Next
'rs.Close
'REC_SET.Close
Set rs = Nothing
Set REC_SET = Nothing
End Sub

Public Sub VLD_PackType()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim Sch_Idy
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Pack Type' GROUP BY SCH_IDY", con, adOpenKeyset, adLockPessimistic
For C = 1 To REC_SET.RecordCount
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Pack Type' and TXN_TYP LIKE 'BAS' AND SCH_IDY = '" & Trim(REC_SET!Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If rs!B_CON_APP = "AND" Then
            For A = 1 To rs.RecordCount
                GRPVAL = GETGRPVALUE("Pack Type", Trim(rs!grp_typ & ""))
                If GRPVAL >= Val(rs!B_PRD_QTY & "") Then
                            
                Else
                    GoTo M
                End If
                rs.MoveNext
            Next
            If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
            Else
                smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                End If
            End If
        ElseIf rs!B_CON_APP = "ORO" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            For A = 1 To rs.RecordCount
                If GETGRPVALUE("Pack Type", Trim(rs!grp_typ & "")) >= Val(MAXVALUE) Then
                    If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    Else
                        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                            Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                            ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                            ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                            ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                        End If
                    End If
                    Exit For
                End If
                rs.MoveNext
            Next
        ElseIf rs!B_CON_APP = "ORM" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            stkcount = 0
            For A = 1 To rs.RecordCount
                stkcount = Val(stkcount) + GETGRPVALUE("Pack Type", Trim(rs!grp_typ & ""))
                If Val(stkcount) >= Val(MAXVALUE) Then
                    If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    Else
                        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                            Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                            ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                            ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                            ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                        End If
                    End If
                    Exit For
                End If
                rs.MoveNext
            Next
        End If
    End If
M:
    REC_SET.MoveNext
Next
'rs.Close
'REC_SET.Close
Set rs = Nothing
Set REC_SET = Nothing
End Sub

Public Sub VLD_Segment()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim Sch_Idy
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Segment' GROUP BY SCH_IDY", con, adOpenKeyset, adLockPessimistic
For C = 1 To REC_SET.RecordCount
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Segment' and TXN_TYP LIKE 'BAS' AND SCH_IDY = '" & Trim(REC_SET!Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If rs!B_CON_APP = "AND" Then
            For A = 1 To rs.RecordCount
                GRPVAL = GETGRPVALUE("Segment", Trim(rs!grp_typ & ""))
                If GRPVAL >= Val(rs!B_PRD_QTY & "") Then
                            
                Else
                    GoTo M
                End If
                rs.MoveNext
            Next
            If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
            Else
                smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                End If
            End If
        ElseIf rs!B_CON_APP = "ORO" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            For A = 1 To rs.RecordCount
                If GETGRPVALUE("Segment", Trim(rs!grp_typ & "")) >= Val(MAXVALUE) Then
                    If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    Else
                        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                            Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                            ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                            ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                            ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                        End If
                    End If
                    Exit For
                End If
                rs.MoveNext
            Next
        ElseIf rs!B_CON_APP = "ORM" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            stkcount = 0
            For A = 1 To rs.RecordCount
                stkcount = Val(stkcount) + GETGRPVALUE("Segment", Trim(rs!grp_typ & ""))
                If Val(stkcount) >= Val(MAXVALUE) Then
                    If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    Else
                        smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                        If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                            Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                            ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                            ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                            ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                            ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                        End If
                    End If
                Exit For
                End If
                rs.MoveNext
            Next
        End If
    End If
M:
    REC_SET.MoveNext
Next
'rs.Close
'REC_SET.Close
Set rs = Nothing
Set REC_SET = Nothing
End Sub

Public Sub VLD_Category()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim Sch_Idy
Dim ls As ListItem
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Category' GROUP BY SCH_IDY", con, adOpenKeyset, adLockPessimistic
For C = 1 To REC_SET.RecordCount
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM SCH where sch_sts ='Y' AND SCH_BDT <= CTOD('" & M_INV_DAT & "') AND SCH_EDT >=ctod('" & M_INV_DAT & "') AND SCH_TYP LIKE 'Category' and TXN_TYP LIKE 'BAS' AND SCH_IDY = '" & Trim(REC_SET!Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If rs!B_CON_APP = "AND" Then
            For A = 1 To rs.RecordCount
                GRPVAL = GETGRPVALUE("Category", Trim(rs!grp_typ & ""))
                If GRPVAL >= Val(rs!B_PRD_QTY & "") Then
                            
                Else
                    GoTo M
                End If
                rs.MoveNext
            Next
            If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
            Else
                smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                End If
            End If
        ElseIf rs!B_CON_APP = "ORO" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            For A = 1 To rs.RecordCount
                If GETGRPVALUE("Category", Trim(rs!grp_typ & "")) >= Val(MAXVALUE) Then
                If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                Else
                    smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    End If
                End If
                Exit For
                End If
                rs.MoveNext
            Next
        ElseIf rs!B_CON_APP = "ORM" Then
            MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(rs!Sch_Idy) & "' AND TXN_TYP LIKE 'BAS'")
            stkcount = 0
            For A = 1 To rs.RecordCount
                stkcount = Val(stkcount) + GETGRPVALUE("Category", Trim(rs!grp_typ & ""))
                If Val(stkcount) >= Val(MAXVALUE) Then
                If check("SID", "SCH_IDY", REC_SET!Sch_Idy) = False Then
                    Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                    ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                    ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                    ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                    ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                Else
                    smnid = GetValue("SMN_IDY", "SID WHERE SMN_IDY LIKE '" & Trim(M_SMN_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    ROUID = GetValue("ROU_IDY", "SID WHERE ROU_IDY LIKE '" & Trim(M_ROU_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    CSTID = GetValue("CST_IDY", "SID WHERE CST_IDY LIKE '" & Trim(M_CST_IDY) & "' AND SCH_IDY LIKE '" & Trim(REC_SET!Sch_Idy) & "'")
                    If Len(Trim(smnid)) > 0 Or Len(Trim(ROUID)) > 0 Or Len(Trim(CSTID)) > 0 Then
                        Set ls = SCHLST.ListItems.Add(, , Trim(REC_SET!Sch_Nme))
                        ls.ListSubItems.Add , , Format(REC_SET!sch_bdt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Format(REC_SET!sch_edt, "dd/mmm/yyyy")
                        ls.ListSubItems.Add , , Trim(REC_SET!SCH_TYP)
                        ls.ListSubItems.Add , , Trim(REC_SET!f_sch_typ)
                        ls.ListSubItems.Add , , Trim(REC_SET!Sch_Idy)
                    End If
                End If
        Exit For
                End If
                rs.MoveNext
            Next
        End If
    End If
M:
    REC_SET.MoveNext
Next
'rs.Close
'REC_SET.Close
Set rs = Nothing
Set REC_SET = Nothing
End Sub

Public Function GETGRPVALUE(TYP As String, NME As String) As Long
For A = 1 To MS.Rows - 1
    If TYP = "Brand" Then
        If Trim(NME) = Trim(GetValue("brd_nme", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then
            If pcs_chk.Value = 1 Then
                stkcount = Val(stkcount) + MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(stkcount) + Val(MS.TextMatrix(A, 2))
            End If
        End If
    ElseIf TYP = "Pack Size" Then
        If Trim(NME) = Trim(GetValue("prd_pck", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then
            If pcs_chk.Value = 1 Then
                stkcount = Val(stkcount) + MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(stkcount) + Val(MS.TextMatrix(A, 2))
            End If
        End If
    ElseIf TYP = "Pack Type" Then
        If Trim(NME) = Trim(GetValue("prd_pkt", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then
            If pcs_chk.Value = 1 Then
                stkcount = Val(stkcount) + MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(stkcount) + Val(MS.TextMatrix(A, 2))
            End If
        End If
    ElseIf TYP = "Segment" Then
        If Trim(NME) = Trim(MS.TextMatrix(A, 20)) Then
            If pcs_chk.Value = 1 Then
                stkcount = Val(stkcount) + MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(stkcount) + Val(MS.TextMatrix(A, 2))
            End If
        End If
    ElseIf TYP = "Category" Then
        If Trim(NME) = Trim(MS.TextMatrix(A, 19)) Then
            If pcs_chk.Value = 1 Then
                stkcount = Val(stkcount) + MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(stkcount) + Val(MS.TextMatrix(A, 2))
            End If
        End If
    End If
Next
GETGRPVALUE = stkcount
End Function

Public Sub AddBrand(SCHIDY As String)
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='AND'"
If RSTMP.RecordCount > 0 Then
    rs.filter = "TXN_TYP='FRE'"
    CTR = 1
    If RSTMP.BOF = False Then RSTMP.MoveFirst
    For A = 1 To 2
        If RSTMP.EOF = True Then
            RSTMP.MoveFirst
            CTR = CTR + 1
        End If
        b = GETGRPVALUE("Brand", Trim(RSTMP!grp_typ))
        If (b - (CTR * RSTMP!B_PRD_QTY)) < 0 Then Exit For
        RSTMP.MoveNext
        A = A - 1
    Next
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR - 1)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If

    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        ADDSCITEMS "Brand", Trim(RSTMP!grp_typ), SCHIDY
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORO'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Brand", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           CTR = CTR + Int(INC / MAXVALUE)
        End If
        RSTMP.MoveNext
    Next
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If

    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Brand", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           ADDSCITEMS "Brand", Trim(RSTMP!grp_typ), SCHIDY
        End If
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORM'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = Val(INC) + GETGRPVALUE("Brand", Trim(RSTMP!grp_typ))
        RSTMP.MoveNext
    Next
    CTR = Int(INC / MAXVALUE)
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = Trim(RS!B_CON_APP) 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    If A > 1 Then
        For A = 1 To RSTMP.RecordCount
            ADDSCITEMS "Brand", Trim(RSTMP!grp_typ), SCHIDY
            RSTMP.MoveNext
        Next
    End If
End If
End Sub

Public Sub AddPackSize(SCHIDY As String)
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='AND'"
If RSTMP.RecordCount > 0 Then
    rs.filter = "TXN_TYP='FRE'"
    CTR = 1
    If RSTMP.BOF = False Then RSTMP.MoveFirst
    For A = 1 To 2
        If RSTMP.EOF = True Then
            RSTMP.MoveFirst
            CTR = CTR + 1
        End If
        b = GETGRPVALUE("Pack Size", Trim(RSTMP!grp_typ))
        If (b - (CTR * RSTMP!B_PRD_QTY)) < 0 Then Exit For
        RSTMP.MoveNext
        A = A - 1
    Next
    
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR - 1)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If

    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
        For A = 1 To RSTMP.RecordCount
            ADDSCITEMS "Pack Size", Trim(RSTMP!grp_typ), SCHIDY
            RSTMP.MoveNext
        Next
    End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORO'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Pack Size", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           CTR = CTR + Int(INC / MAXVALUE)
        End If
        RSTMP.MoveNext
    Next
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            If rs!F_CON_APP = "ORO" Then
                For A = 1 To MSS.Rows - 1
                    If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
                Next
            End If
        End If

    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Pack Size", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           ADDSCITEMS "Pack Size", Trim(RSTMP!grp_typ), SCHIDY
        End If
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORM'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = Val(INC) + GETGRPVALUE("Pack Size", Trim(RSTMP!grp_typ))
        RSTMP.MoveNext
    Next
    If MAXVALUE = 0 Then
        CTR = 1
    Else
        CTR = Int(INC / MAXVALUE)
    End If

'    CTR = Int(INC / MAXVALUE)
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    If A > 1 Then
        For A = 1 To RSTMP.RecordCount
            ADDSCITEMS "Pack Size", Trim(RSTMP!grp_typ), SCHIDY
            RSTMP.MoveNext
        Next
    End If
End If
End Sub

Public Sub AddPackType(SCHIDY As String)
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='AND'"
If RSTMP.RecordCount > 0 Then
    rs.filter = "TXN_TYP='FRE'"
    CTR = 1
    If RSTMP.BOF = False Then RSTMP.MoveFirst
    For A = 1 To 2
        If RSTMP.EOF = True Then
            RSTMP.MoveFirst
            CTR = CTR + 1
        End If
        b = GETGRPVALUE("Pack Type", Trim(RSTMP!grp_typ))
        If (b - (CTR * RSTMP!B_PRD_QTY)) < 0 Then Exit For
        RSTMP.MoveNext
        A = A - 1
    Next
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR - 1)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        ADDSCITEMS "Pack Type", Trim(RSTMP!grp_typ), SCHIDY
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORO'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Pack Type", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           CTR = CTR + Int(INC / MAXVALUE)
        End If
        RSTMP.MoveNext
    Next
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Pack Type", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           ADDSCITEMS "Pack Type", Trim(RSTMP!grp_typ), SCHIDY
        End If
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORM'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = Val(INC) + GETGRPVALUE("Pack Type", Trim(RSTMP!grp_typ))
        RSTMP.MoveNext
    Next
    CTR = Int(INC / MAXVALUE)
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    If A > 1 Then
        For A = 1 To RSTMP.RecordCount
            ADDSCITEMS "Pack Type", Trim(RSTMP!grp_typ), SCHIDY
            RSTMP.MoveNext
        Next
    End If
End If
End Sub

Public Sub AddSegment(SCHIDY As String)
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='AND'"
If RSTMP.RecordCount > 0 Then
    rs.filter = "TXN_TYP='FRE'"
    CTR = 1
    If RSTMP.BOF = False Then RSTMP.MoveFirst
    For A = 1 To 2
        If RSTMP.EOF = True Then
            RSTMP.MoveFirst
            CTR = CTR + 1
        End If
        b = GETGRPVALUE("Segment", Trim(RSTMP!grp_typ))
        If (b - (CTR * RSTMP!B_PRD_QTY)) < 0 Then Exit For
        RSTMP.MoveNext
        A = A - 1
    Next
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD,PGR WHER BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR - 1)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        ADDSCITEMS "Segment", Trim(RSTMP!grp_typ), SCHIDY
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORO'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Segment", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           CTR = CTR + Int(INC / MAXVALUE)
        End If
        RSTMP.MoveNext
    Next
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Segment", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           ADDSCITEMS "Segment", Trim(RSTMP!grp_typ), SCHIDY
        End If
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORM'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = Val(INC) + GETGRPVALUE("Segment", Trim(RSTMP!grp_typ))
        RSTMP.MoveNext
    Next
    CTR = Int(INC / MAXVALUE)
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
            .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
            rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then MSS.TextMatrix(A, 11) = SCHQTY
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    If A > 1 Then
        For A = 1 To RSTMP.RecordCount
            ADDSCITEMS "Segment", Trim(RSTMP!grp_typ), SCHIDY
            RSTMP.MoveNext
        Next
    End If
End If
End Sub

Public Sub AddCategory(SCHIDY As String)
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='AND'"
If RSTMP.RecordCount > 0 Then
    rs.filter = "TXN_TYP='FRE'"
    CTR = 1
    If RSTMP.BOF = False Then RSTMP.MoveFirst
    For A = 1 To 2
        If RSTMP.EOF = True Then
            RSTMP.MoveFirst
            CTR = CTR + 1
        End If
        b = GETGRPVALUE("Category", Trim(RSTMP!grp_typ))
        If (b - (CTR * RSTMP!B_PRD_QTY)) < 0 Then Exit For
        RSTMP.MoveNext
        A = A - 1
    Next
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD,PGR WHER BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR - 1)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR - 1)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP)
        rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then
                MSS.TextMatrix(A, 11) = SCHQTY
                MSS.TextMatrix(A, 15) = SCHQTY
                End If
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    For A = 1 To RSTMP.RecordCount
        ADDSCITEMS "Category", Trim(RSTMP!grp_typ), SCHIDY
        RSTMP.MoveNext
    Next
End If

RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORO'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = GETGRPVALUE("Category", Trim(RSTMP!grp_typ))
        If Val(INC) >= Val(MAXVALUE) Then
           CTR = CTR + Int(INC / MAXVALUE)
        End If
        RSTMP.MoveNext
    Next
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
        .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP) 'product group
        rs.MoveNext
        End With
    Next
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then
                    MSS.TextMatrix(A, 11) = SCHQTY
                    MSS.TextMatrix(A, 15) = SCHQTY
                End If
            Next
        End If
    End If
    
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
        For A = 1 To RSTMP.RecordCount
            INC = GETGRPVALUE("Category", Trim(RSTMP!grp_typ))
            If Val(INC) >= Val(MAXVALUE) Then
               ADDSCITEMS "Category", Trim(RSTMP!grp_typ), SCHIDY
            End If
            RSTMP.MoveNext
        Next
    End If
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORM'"
If RSTMP.RecordCount > 0 Then
    CTR = 0
    INC = 0
    MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
    For A = 1 To RSTMP.RecordCount
        INC = Val(INC) + GETGRPVALUE("Category", Trim(RSTMP!grp_typ))
        RSTMP.MoveNext
    Next
    CTR = Int(INC / MAXVALUE)
    rs.filter = "TXN_TYP='FRE'"
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "SELECT * FROM PRD,BRD WHER BRD.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        With MSS
            .AddItem ""
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            rsprd.Find "BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"
            .TextMatrix(.Rows - 1, 0) = rsprd!PRD_NME
            .TextMatrix(.Rows - 1, 1) = convert(rsprd!PRD_MRP)
            .TextMatrix(.Rows - 1, 2) = Val(rs!B_PRD_QTY) * (CTR)
            If .TextMatrix(.Rows - 1, 2) > 0 Then SCHQTY = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 15) = Val(rs!B_PRD_QTY) * (CTR)
            .TextMatrix(.Rows - 1, 3) = convert(rsprd!PRD_SDR)
            .TextMatrix(.Rows - 1, 4) = "0.00"
            .TextMatrix(.Rows - 1, 5) = Trim(rsprd!prd_idy) 'PRD_IDY
            .TextMatrix(.Rows - 1, 6) = Trim(rsprd!Bat_nme) 'BAT_NME
            .TextMatrix(.Rows - 1, 7) = Trim(rsprd!BAT_IDY) 'BAT_IDY
            .TextMatrix(.Rows - 1, 8) = Trim(SCHIDY) 'SCH_IDY
            'To be corrected by KRB
            '.TextMatrix(.Rows - 1, 9) = RSTMP!B_CON_APP 'Not APPLICABLE 'TXN_TYP
            .TextMatrix(.Rows - 1, 10) = Val(rs!B_PRD_QTY) ' BAS QUANTITY
            .TextMatrix(.Rows - 1, 11) = Val(rs!B_PRD_QTY) ' Free Quantity
            .TextMatrix(.Rows - 1, 12) = Val(rsprd!SUB_UNT) 'subunts
            .TextMatrix(.Rows - 1, 13) = Trim(rsprd!GRP_IDY) 'product group
            .TextMatrix(.Rows - 1, 14) = Trim(rs!F_CON_APP) 'product group
            rs.MoveNext
        End With
    Next
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If rs!F_CON_APP = "ORO" Then
            For A = 1 To MSS.Rows - 1
                If MSS.TextMatrix(A, 8) = SCHIDY Then
                    MSS.TextMatrix(A, 11) = SCHQTY
                    MSS.TextMatrix(A, 15) = SCHQTY
                End If
            Next
        End If
    End If
    If RSTMP.RecordCount > 0 Then RSTMP.MoveFirst
    If A > 1 Then
        For A = 1 To RSTMP.RecordCount
            ADDSCITEMS "Category", Trim(RSTMP!grp_typ), SCHIDY
            RSTMP.MoveNext
        Next
    End If
End If
End Sub
Public Sub ADDSCITEMS(TYP As String, NME As String, SCHIDY As String)
stkcount = 0
For A = 1 To MS.Rows - 1
    If TYP = "Brand" Then
        If Trim(NME) = Trim(GetValue("brd_nme", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then
            If pcs_chk.Value = 1 Then
                stkcount = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(MS.TextMatrix(A, 2))
            End If
            SC.AddItem ""
            SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
            SC.TextMatrix(SC.Rows - 1, 1) = Trim(MS.TextMatrix(A, 15))
            SC.TextMatrix(SC.Rows - 1, 2) = stkcount
            SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
        End If
    ElseIf TYP = "Pack Size" Then
        If Trim(NME) = Trim(GetValue("prd_pck", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then
            If pcs_chk.Value = 1 Then
                stkcount = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(MS.TextMatrix(A, 2))
            End If
            SC.AddItem ""
            SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
            SC.TextMatrix(SC.Rows - 1, 1) = Trim(MS.TextMatrix(A, 15))
            SC.TextMatrix(SC.Rows - 1, 2) = stkcount
            SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
        End If
    ElseIf TYP = "Pack Type" Then
        If Trim(NME) = Trim(GetValue("prd_pkt", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then
            If pcs_chk.Value = 1 Then
                stkcount = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(MS.TextMatrix(A, 2))
            End If
            SC.AddItem ""
            SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
            SC.TextMatrix(SC.Rows - 1, 1) = Trim(MS.TextMatrix(A, 15))
            SC.TextMatrix(SC.Rows - 1, 2) = stkcount
            SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
        End If
    ElseIf TYP = "Segment" Then
        If Trim(NME) = Trim(MS.TextMatrix(A, 20)) Then
            If pcs_chk.Value = 1 Then
                stkcount = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(MS.TextMatrix(A, 2))
            End If
            SC.AddItem ""
            SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
            SC.TextMatrix(SC.Rows - 1, 1) = Trim(MS.TextMatrix(A, 15))
            SC.TextMatrix(SC.Rows - 1, 2) = stkcount
            SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
        End If
    ElseIf TYP = "Category" Then
        If Trim(NME) = Trim(MS.TextMatrix(A, 19)) Then
            If pcs_chk.Value = 1 Then
                stkcount = MTS(Val(MS.TextMatrix(A, 2)), Val(MS.TextMatrix(A, 13)))
            Else
                stkcount = Val(MS.TextMatrix(A, 2))
            End If
            SC.AddItem ""
            SC.TextMatrix(SC.Rows - 1, 0) = Trim(SCHIDY)
            SC.TextMatrix(SC.Rows - 1, 1) = Trim(MS.TextMatrix(A, 15))
            SC.TextMatrix(SC.Rows - 1, 2) = stkcount
            SC.TextMatrix(SC.Rows - 1, 4) = "BAS"
        End If
    End If
Next
End Sub

Public Sub AddBillAmtGroupType(SCHIDY As String)
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim RSTMP As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim GrpType As String
rs.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "'", con, adOpenKeyset, adLockPessimistic
Set RSTMP = rs.Clone
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='AND'"
If rs.EOF = False Then GrpType = Trim(rs!SCH_TYP)
If RSTMP.RecordCount > 0 Then
    disamt = Val(rs!B_DIS_AMT & "")
    dispcg = Val(rs!B_DIS_PCG & "")
    rs.filter = "TXN_TYP='FRE'"
    CTR = 1
    If RSTMP.BOF = False Then RSTMP.MoveFirst
    For A = 1 To 2
        If RSTMP.EOF = True Then
            RSTMP.MoveFirst
            CTR = CTR + 1
        End If
        b = GETGRPVALUE(GrpType, Trim(RSTMP!grp_typ))
        If (b - (CTR * RSTMP!B_PRD_QTY)) < 0 Then Exit For
        RSTMP.MoveNext
        A = A - 1
    Next
    
    SAMT = Val(SCH_AMT)
    If RSTMP.EOF = True Then RSTMP.MoveFirst
    Do While RSTMP.EOF = False
        PRDVAL = PRDVAL + GETSCHGRPVAL(GrpType, Trim(RSTMP!grp_typ))
        RSTMP.MoveNext
    Loop
    
    If Val(dispcg) > 0 Then
        SCH_AMT = convert(Val(SCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg)))
    Else
        SCH_AMT = convert(Val(SCH_AMT) + (Val(disamt) * (CTR - 1)))
    End If

    SC.AddItem ""
    SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
    SC.TextMatrix(SC.Rows - 1, 3) = Val(SCH_AMT) - Val(SAMT)
    SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
Else
RSTMP.filter = "txn_typ='BAS' AND B_CON_APP='ORO'"
    If RSTMP.EOF = False Then
        disamt = Val(rs!B_DIS_AMT & "")
        dispcg = Val(rs!B_DIS_PCG & "")
        CTR = 0
        INC = 0
        MAXVALUE = GetValue("MAX(B_PRD_QTY)", "SCH WHERE SCH_IDY LIKE '" & Trim(SCHIDY) & "' AND TXN_TYP LIKE 'BAS'")
        For A = 1 To RSTMP.RecordCount
            INC = GETGRPVALUE(GrpType, Trim(RSTMP!grp_typ))
            If Val(INC) >= Val(MAXVALUE) Then
               CTR = CTR + Int(INC / MAXVALUE)
            End If
            RSTMP.MoveNext
        Next
        
        SAMT = Val(SCH_AMT)
        If RSTMP.EOF = True Then RSTMP.MoveFirst
        Do While RSTMP.EOF = False
            PRDVAL = PRDVAL + GETSCHGRPVAL(GrpType, Trim(RSTMP!grp_typ))
            RSTMP.MoveNext
        Loop
        
        If Val(dispcg) > 0 Then
            SCH_AMT = convert(Val(SCH_AMT) + PERCENT(Val(PRDVAL), Val(dispcg)))
        Else
            SCH_AMT = convert(Val(SCH_AMT) + (Val(disamt) * (CTR)))
        End If
        
        SC.AddItem ""
        SC.TextMatrix(SC.Rows - 1, 0) = SCHIDY 'SCHEMEIDY
        SC.TextMatrix(SC.Rows - 1, 3) = Val(SCH_AMT) - Val(SAMT)
        SC.TextMatrix(SC.Rows - 1, 4) = "FRE"
        
    End If
End If

End Sub

Public Function GETSCHGRPVAL(MType As String, mValue As String) As Double
For A = 1 To MS.Rows - 1
    If MType = "Brand" Then
        If Trim(mValue) = Trim(GetValue("brd_nme", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then prdvalue = prdvalue + Val(MS.TextMatrix(A, 9))
    ElseIf MType = "Pack Size" Then
        If Trim(mValue) = Trim(GetValue("prd_pck", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then prdvalue = prdvalue + Val(MS.TextMatrix(A, 9))
    ElseIf MType = "Pack Type" Then
        If Trim(mValue) = Trim(GetValue("prd_pkt", "prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'")) Then prdvalue = prdvalue + Val(MS.TextMatrix(A, 9))
    ElseIf MType = "Segment" Then
        If Trim(mValue) = Trim(MS.TextMatrix(A, 20)) Then prdvalue = prdvalue + Val(MS.TextMatrix(A, 9))
    ElseIf MType = "Category" Then
        If Trim(mValue) = Trim(MS.TextMatrix(A, 19)) Then prdvalue = prdvalue + Val(MS.TextMatrix(A, 9))
    End If
Next
GETSCHGRPVAL = prdvalue
End Function
Public Sub AutoApplyScheme()
    MSS.Rows = 1
    SC.Rows = 0
    SCH_AMT = 0
    Command3.Caption = "Select A&ll"
    VLD_FRE_QTY
    VLD_BIL_AMT
    VLD_MLT_QTY
    VLD_MLT_QTY_OR
    VLD_MLT_QTY_OR_MIX
    VLD_BRAND
    VLD_PackSize
    VLD_PackType
    VLD_Segment
    VLD_Category
    MSS.Rows = 1
    SC.Rows = 0
    SCH_AMT = "0.00"
    APPSCH = False
    For A = 1 To SCHLST.ListItems.count
        SCHLST.ListItems.Item(A).Checked = True
    Next
    For A = 1 To SCHLST.ListItems.count
        If SCHLST.ListItems.Item(A).Checked = True Then
            Set ls = SCHLST.ListItems.Item(A)
            If ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "SKU Level" Then
                ADDSKUQUANTIY Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Bill Level" Then
                ADDBILLQUANTIY Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Multiple SKU Level" Then
                ADDMLTQTYAND Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(3) = "SKU Level" And ls.ListSubItems(4) = "Cash Off" Then
                ADDSKUAMOUNT Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(3) = "SKU Level" And ls.ListSubItems(4) = "Discount Percent" Then
                ADDSKUPCG Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Cash Off" And ls.ListSubItems(3) = "Bill Level" Then
                ADDBILAMT Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Discount Percent" And ls.ListSubItems(3) = "Bill Level" Then
                ADDBILPCG Trim(ls.ListSubItems(5))
            ElseIf (ls.ListSubItems(4) = "Cash Off" Or ls.ListSubItems(4) = "Discount Percent") And ls.ListSubItems(3) = "Multiple SKU Level" Then
                MLTBILPCGAMT Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Brand" Then
                AddBrand Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Pack Size" Then
                AddPackSize Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Pack Type" Then
                AddPackType Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Segment" Then
                AddSegment Trim(ls.ListSubItems(5))
            ElseIf ls.ListSubItems(4) = "Free Quantity" And ls.ListSubItems(3) = "Category" Then
                AddCategory Trim(ls.ListSubItems(5))
            ElseIf (ls.ListSubItems(4) = "Cash Off" Or ls.ListSubItems(4) = "Discount Percent") And (ls.ListSubItems(3) = "Pack Size" Or ls.ListSubItems(3) = "Pack Type" Or ls.ListSubItems(3) = "Brand" Or ls.ListSubItems(3) = "Category" Or ls.ListSubItems(3) = "Segment") Then
                AddBillAmtGroupType Trim(ls.ListSubItems(5))
            End If
        End If
    Next
    APPSCH = True
    If Val(SCH_AMT) > 0 Then
        M_AMT_LES = "Scheme Amount"
        M_FRT_AMT = SCH_AMT
        M_AMT_LES.Locked = True
        M_FRT_AMT.Locked = True
    End If
    GRS_CAL
    SSTab1.Tab = 2
End Sub
Public Sub DayCls_PrintBill(INC As String)
If check("PRT", "PRT_NME", "M_FRM_SOR") = True Then
con.Execute "update dir where dir_idy like 'INV' AND DIR_IDX LIKE 'NUM' set dir_als = '" & INC & "'"
con.Execute "update dir where dir_idy like 'DOC' AND DIR_IDX LIKE 'TYP' set dir_als = 'SAL'"
If querycheck("DIR", "DIR_MSG", "D", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then
    If REC_PNL.State = 1 Then REC_PNL.Close
    REC_PNL.Open "SELECT * FROM RRI", PNL, adOpenKeyset, adLockPessimistic
    If REC_PNL.RecordCount > 0 Then
        REC_PNL.Move 6
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
    'Shell CURDIR & "\RRUNTIME RRI 7", vbHide
    If Dir(CURDIR & "\pri.exe") = "" Then
        MsgBox "PRI.EXE not found", vbExclamation, "DAS Version 4.0"
    Else
        Shell CURDIR & "\pri.exe"
    End If
Else
    OpenDocument CURDIR & "\DOS_VCH.RP1", "Sale Order", 2, Trim(INC), Trim(INC), 6, CURDIR & "\pnv.cdx,c,inv_idy"
End If
CANCELRECORD
End If
End Sub
