VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schemes - Trade"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "M_FRM_SCH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5100
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Scheme Details"
      TabPicture(0)   =   "M_FRM_SCH.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label16"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "EnDate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "StDate"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Sub_Cmb"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Sch_Cmb"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Sch_Nme"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Sch_Idy"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Pro_Chk"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "MERAMT"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "MERQTY"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "MERNO"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Sch_Act"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "List2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Scheme &Definition"
      TabPicture(1)   =   "M_FRM_SCH.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabBase"
      Tab(1).Control(1)=   "TabFree"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "M_FRM_SCH.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lv1"
      Tab(2).Control(1)=   "LV"
      Tab(2).Control(2)=   "lblview"
      Tab(2).ControlCount=   3
      Begin VB.ListBox List2 
         Height          =   255
         ItemData        =   "M_FRM_SCH.frx":0496
         Left            =   1335
         List            =   "M_FRM_SCH.frx":04A0
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2705
         Width           =   495
      End
      Begin VB.ListBox Sch_Act 
         Height          =   255
         ItemData        =   "M_FRM_SCH.frx":04AA
         Left            =   1320
         List            =   "M_FRM_SCH.frx":04B4
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   4435
         Width           =   495
      End
      Begin VB.TextBox MERNO 
         Height          =   330
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3081
         Width           =   1215
      End
      Begin VB.TextBox MERQTY 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "0.000"
         Top             =   3532
         Width           =   1215
      End
      Begin VB.TextBox MERAMT 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3983
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Height          =   4575
         Left            =   4800
         TabIndex        =   43
         Top             =   360
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "Check &All"
            Height          =   375
            Left            =   1440
            TabIndex        =   57
            Top             =   4125
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_FRM_SCH.frx":04BE
            Left            =   1440
            List            =   "M_FRM_SCH.frx":04CE
            Style           =   2  'Dropdown List
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   300
            Width           =   2175
         End
         Begin MSComctlLib.ListView list1 
            Height          =   3375
            Left            =   120
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   5953
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Type"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specific To"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   435
            TabIndex        =   45
            Top             =   360
            Width           =   810
         End
      End
      Begin MSComctlLib.ListView lv1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6800
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
      Begin VB.CheckBox Pro_Chk 
         Caption         =   "&Pro Rata"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Sch_Idy 
         Height          =   330
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Sch_Nme 
         Height          =   330
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   931
         Width           =   2865
      End
      Begin VB.ComboBox Sch_Cmb 
         Height          =   315
         ItemData        =   "M_FRM_SCH.frx":04F2
         Left            =   1320
         List            =   "M_FRM_SCH.frx":0511
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1382
         Width           =   2175
      End
      Begin VB.ComboBox Sub_Cmb 
         Height          =   315
         ItemData        =   "M_FRM_SCH.frx":0585
         Left            =   1320
         List            =   "M_FRM_SCH.frx":0595
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1818
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker StDate 
         Height          =   330
         Left            =   1320
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2250
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
         Format          =   76152835
         CurrentDate     =   37099
      End
      Begin MSComCtl2.DTPicker EnDate 
         Height          =   330
         Left            =   3480
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2254
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
         Format          =   76152835
         CurrentDate     =   37099
      End
      Begin TabDlg.SSTab TabBase 
         Height          =   4650
         Left            =   -74880
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   8202
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Base SKU Details"
         TabPicture(0)   =   "M_FRM_SCH.frx":05DD
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "TxtBase"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSbase"
         Tab(0).Control(2)=   "Frame1"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Base Value Details"
         TabPicture(1)   =   "M_FRM_SCH.frx":05F9
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "M_FRM_SCH.frx":0615
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame5"
         Tab(2).Control(1)=   "TxtPack"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "MSPack"
         Tab(2).ControlCount=   3
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   -74880
            TabIndex        =   35
            Top             =   360
            Width           =   4095
            Begin VB.ComboBox CmbPCon 
               Height          =   315
               ItemData        =   "M_FRM_SCH.frx":0631
               Left            =   1920
               List            =   "M_FRM_SCH.frx":063E
               Style           =   2  'Dropdown List
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   200
               Width           =   2055
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Condition Applicable"
               Height          =   195
               Left            =   240
               TabIndex        =   37
               Top             =   240
               Width           =   1440
            End
         End
         Begin VB.TextBox TxtPack 
            Height          =   495
            Left            =   -74040
            MaxLength       =   10
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Frame Frame4 
            Caption         =   "Scheme Value"
            Height          =   1575
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   3975
            Begin VB.TextBox TxtGrsAmt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2280
               MaxLength       =   10
               TabIndex        =   31
               TabStop         =   0   'False
               Text            =   "0.00"
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Invoice Gross Amount"
               Height          =   195
               Left            =   480
               TabIndex        =   32
               Top             =   675
               Width           =   1560
            End
         End
         Begin VB.TextBox TxtBase 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   -74040
            MaxLength       =   10
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid MSbase 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1080
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   14737632
            ForeColorSel    =   16711680
            FocusRect       =   0
            Appearance      =   0
         End
         Begin VB.Frame Frame1 
            Height          =   615
            Left            =   -74880
            TabIndex        =   16
            Top             =   360
            Width           =   4095
            Begin VB.ComboBox CmbBase 
               Height          =   315
               ItemData        =   "M_FRM_SCH.frx":0652
               Left            =   1920
               List            =   "M_FRM_SCH.frx":065F
               Style           =   2  'Dropdown List
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   200
               Width           =   2055
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Condition Applicable"
               Height          =   195
               Left            =   240
               TabIndex        =   18
               Top             =   240
               Width           =   1440
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSPack 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1080
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393216
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   14737632
            ForeColorSel    =   16711680
            FocusRect       =   0
            Appearance      =   0
         End
      End
      Begin TabDlg.SSTab TabFree 
         Height          =   4650
         Left            =   -70440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   8202
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Free SKU Details"
         TabPicture(0)   =   "M_FRM_SCH.frx":0673
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSFree"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TxtFree"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Free Value Details"
         TabPicture(1)   =   "M_FRM_SCH.frx":068F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Height          =   1575
            Left            =   -74760
            TabIndex        =   27
            Top             =   1440
            Width           =   3855
            Begin VB.TextBox TxtValue 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1920
               MaxLength       =   10
               TabIndex        =   29
               TabStop         =   0   'False
               Text            =   "0.00"
               Top             =   600
               Width           =   855
            End
            Begin VB.Label LblValue 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cash Off"
               Height          =   195
               Left            =   1080
               TabIndex        =   28
               Top             =   720
               Width           =   615
            End
         End
         Begin VB.TextBox TxtFree 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid MSFree 
            Height          =   3495
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1080
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   -2147483633
            BackColorSel    =   14737632
            ForeColorSel    =   16711680
            FocusRect       =   0
            Appearance      =   0
         End
         Begin VB.Frame Frame2 
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   4095
            Begin VB.ComboBox CmbFree 
               Height          =   315
               ItemData        =   "M_FRM_SCH.frx":06AB
               Left            =   1920
               List            =   "M_FRM_SCH.frx":06B5
               Style           =   2  'Dropdown List
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   200
               Width           =   2055
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Condition Applicable"
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   240
               Width           =   1440
            End
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6800
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "HVO A/c"
         Height          =   195
         Left            =   495
         TabIndex        =   55
         Top             =   2735
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Active"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   690
         TabIndex        =   54
         Top             =   4465
         Width           =   450
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MER Ref No"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   3149
         Width           =   915
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Quantity"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   255
         TabIndex        =   52
         Top             =   3600
         Width           =   885
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Amount"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   300
         TabIndex        =   51
         Top             =   4051
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Free"
         Height          =   195
         Left            =   825
         TabIndex        =   42
         Top             =   1878
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Base"
         Height          =   195
         Left            =   780
         TabIndex        =   41
         Top             =   1442
         Width           =   360
      End
      Begin VB.Label lblview 
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
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   8775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   975
         TabIndex        =   11
         Top             =   548
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   999
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   2322
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   435
         TabIndex        =   8
         Top             =   2322
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   0
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
            Picture         =   "M_FRM_SCH.frx":06C2
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":0C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":0FCE
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":1322
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":195A
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":1CAE
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":210E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":2402
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":280E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":291A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":2C6E
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":3082
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
      Width           =   9075
      _ExtentX        =   16007
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
            Object.ToolTipText     =   "Prefix Scheme ID"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Enabled         =   0   'False
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
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         ScaleHeight     =   375
         ScaleWidth      =   4215
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Press F6 to Shift Between TABS"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1800
            TabIndex        =   40
            Top             =   120
            Width           =   2295
         End
      End
   End
End
Attribute VB_Name = "M_FRM_SCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Command1.Caption = "UnCheck &All" Then
    Command1.Caption = "Check &All"
Else
    Command1.Caption = "UnCheck &All"
End If
If list1.ListItems.count > 0 Then
    For i = 1 To list1.ListItems.count
        list1.ListItems(i).Checked = IIf(Command1.Caption = "UnCheck &All", 1, 0)
    Next
End If
list1.ListItems(1).Selected = True
End Sub

Private Sub List2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(Sch_Idy) = "" Then
        MERNO.SetFocus
    End If
End If
End Sub
Private Sub lv1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If lv1.SortOrder = lvwAscending Then
lv1.SortOrder = lvwDescending
ElseIf lv1.SortOrder = lvwDescending Then
lv1.SortOrder = lvwAscending
End If
lv1.SortKey = CH.Index - 1
End Sub
Private Sub LV_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If LV.SortOrder = lvwAscending Then
LV.SortOrder = lvwDescending
ElseIf LV.SortOrder = lvwDescending Then
LV.SortOrder = lvwAscending
End If
LV.SortKey = CH.Index - 1
End Sub
Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub
Private Sub CmbBase_Click()
If CmbBase.ListIndex = 1 Or CmbBase.ListIndex = 2 Then
    For A = 1 To MSbase.Rows - 1
        MSbase.TextMatrix(A, 2) = ""
    Next
End If
End Sub
Private Sub Cmbpcon_Click()
If CmbPCon.ListIndex = 1 Or CmbPCon.ListIndex = 2 Then
    For A = 1 To MSPack.Rows - 1
        MSPack.TextMatrix(A, 1) = ""
    Next
End If
End Sub
Private Sub CmbFREE_Click()
If CmbFree.ListIndex = 1 Then
    For A = 1 To MSFree.Rows - 1
        MSFree.TextMatrix(A, 2) = ""
    Next
End If
End Sub

Private Sub Combo1_Change()
list1.ListItems.CLEAR
End Sub

Private Sub Combo1_keypress(K As Integer)
If K = 13 Then
    Dim rs As New ADODB.Recordset
    Dim ls As ListItem
    list1.ListItems.CLEAR
    If Combo1.ListIndex > 0 Then
        If rs.State = 1 Then rs.Close
        If Combo1.ListIndex = 1 Then
            rs.Open "SELECT ACT_NME,ACT_IDY,CST_TYP FROM ACT WHERE CST_TYP in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c','HVO-P','HVO-NP') and act_typ = 'SAL'", con, adOpenKeyset, adLockPessimistic
        ElseIf Combo1.ListIndex = 2 Then
            rs.Open "SELECT SMN_NME,SMN_IDY,SMN_AD4 FROM SMN", con, adOpenKeyset, adLockPessimistic
        ElseIf Combo1.ListIndex = 3 Then
            rs.Open "SELECT ROU_IDY,ROU_IDY,SMN_IDY FROM RDY GROUP BY ROU_IDY", con, adOpenKeyset, adLockPessimistic
        End If
        Combo1.Enabled = False
        For A = 1 To rs.RecordCount
            Set ls = list1.ListItems.Add(, , Trim(rs(0)))
            ls.ListSubItems.Add , , Trim(rs(1) & "")
            ls.ListSubItems.Add , , Trim(rs(2) & "")
            rs.MoveNext
        Next
        list1.SetFocus
    Else
        Sch_Act.SetFocus
    End If
Command1.Caption = "Check &All"
End If
End Sub

Private Sub Form_Load()
SSTab1.TabEnabled(2) = False
With MSbase
.ColWidth(3) = 0
.ColWidth(4) = 0
.ColWidth(5) = 0
.ColWidth(0) = 2000
.ColWidth(1) = 950
.ColWidth(2) = 1100
.TextMatrix(0, 0) = "Product(SKU) Name"
.TextMatrix(0, 1) = "MRP"
.TextMatrix(0, 2) = "Qty-Cases"
SCl MSbase, 2
.ColAlignment(0) = 1
.ColAlignment(1) = 7
.ColAlignment(2) = 7
.HighLight = flexHighlightNever
End With
With MSFree
.ColWidth(3) = 0
.ColWidth(4) = 0
.ColWidth(5) = 0
.TextMatrix(0, 0) = "Product(SKU) Name"
.TextMatrix(0, 1) = "MRP"
.TextMatrix(0, 2) = "Qty-Cases"
SCl MSFree, 2
.ColWidth(0) = 2000
.ColWidth(1) = 950
.ColWidth(2) = 1100
.HighLight = flexHighlightNever
.ColAlignment(0) = 1
.ColAlignment(1) = 7
.ColAlignment(2) = 7
End With
With MSPack
    .ColWidth(0) = .Width - 1250
    .ColWidth(1) = 1200
    .TextMatrix(0, 1) = "Qty-Units"
    .ColAlignment(1) = 7
    .ColAlignment(0) = 1
End With
TabBase.TabEnabled(0) = False
TabBase.TabEnabled(1) = False
TabFree.TabEnabled(0) = False
TabFree.TabEnabled(1) = False
Combo1.ListIndex = 0
Frame6.Enabled = False
List2.ListIndex = 1
End Sub
Private Sub Form_Activate()
'On Error Resume Next
SSTab1.Tab = 0
Sch_Idy.SetFocus
Sch_Cmb.ListIndex = 0
Sub_Cmb.ListIndex = 0
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_SCH.ActiveControl Is TextBox Then If K <> 117 Then M_FRM_SCH.ActiveControl.Text = Trim(FUNKEY(K))
    If K = 117 Then
        F6
    End If
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        Unload Me
    ElseIf K = 67 And s = 2 Then
        SchemeCopy
    ElseIf K = 69 And s = 2 Then
        If M_FRM_SCH.ActiveControl.name = "Sch_Idy" Then
            Schemelist
            LV.SetFocus
        End If
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        M_FRM_DSN.Tag = "Schemes List"
        M_FRM_DSN.Show 1
    Else
        If K = 27 And Len(Trim(Sch_Idy)) <> 0 And SSTab1.Tab = 1 Then
            save
        End If
    End If
End Sub

Private Sub List1_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Sch_Act.SetFocus
End If
End Sub

Private Sub LV_DblClick()
lv_KeyPress 13
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lblview = "Select Scheme" Then
        DisplayScheme LV.SelectedItem.ListSubItems.Item(4)
    ElseIf lblview = "Select Product" And LV.Tag = "Base" Then
        P_DIS_PRD LV.SelectedItem.ListSubItems.Item(5), "Base"
    ElseIf lblview = "Select Product" And LV.Tag = "Free" Then
        P_DIS_PRD LV.SelectedItem.ListSubItems.Item(5), "Free"
    End If
End If
End Sub

Private Sub lv1_KeyPress(K As Integer)
For A = 1 To MSPack.Rows - 1
    If Trim(MSPack.TextMatrix(A, 0)) = Trim(lv1.SelectedItem) Then
        MsgBox Sch_Cmb & "Already Exists", vbExclamation
        lv1.Visible = False
        TxtPack.Visible = False
        SSTab1.Tab = 1
        MSPack.SetFocus
        Exit Sub
    End If
Next
MSPack.TextMatrix(MSPack.Row, 0) = Trim(lv1.SelectedItem)
lv1.Visible = False
SSTab1.Tab = 1
TabBase.Tab = 2
MSPack.Col = 1
MSPack.SetFocus
End Sub

Private Sub MERNO_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub MERQTY_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub MERAMT_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub MERNO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    MERQTY.SetFocus
ElseIf KeyCode = 38 Then
    EnDate.SetFocus
End If
End Sub
Private Sub MERQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    MERQTY = QTY(MERQTY)
    MERAMT.SetFocus
ElseIf KeyCode = 38 Then
    MERQTY = QTY(MERQTY)
    MERNO.SetFocus
End If
End Sub
Private Sub MERAMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    MERAMT = convert(MERAMT)
    Frame6.Enabled = True
    If Combo1.Enabled = True Then Combo1.SetFocus
ElseIf KeyCode = 38 Then
    MERAMT = convert(MERAMT)
    MERQTY.SetFocus
End If
End Sub
Private Sub MERQTY_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub MERAMT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub SCH_NME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Sch_Nme)) > 0 Then
        If MOD_REC = False Then
            If Sch_Cmb.Enabled = True Then Sch_Cmb.SetFocus
        Else
            EnDate.SetFocus
        End If
    Else
        MsgBox "Scheme Name Should Not Be Blank, Give Name !", vbExclamation, "DAS Version 4.0"
        Sch_Nme.SetFocus
     End If
ElseIf KeyCode = 38 Then
    If Sch_Idy.Enabled = True Then
        Sch_Idy.SetFocus
    Else
        Sch_Nme.SetFocus
    End If
End If
End Sub

Private Sub MSbase_GotFocus()
MSbase.HighLight = flexHighlightAlways
End Sub

Private Sub MSbase_LeaveCell()
TxtBase.Visible = False
End Sub

Private Sub MSfree_LeaveCell()
TxtFree.Visible = False
End Sub

Private Sub MSbase_lostFocus()
MSbase.HighLight = flexHighlightNever
End Sub

Private Sub MSfree_GotFocus()
MSFree.HighLight = flexHighlightAlways
End Sub

Private Sub MSfree_lostFocus()
MSFree.HighLight = flexHighlightNever
End Sub

Private Sub mspack_GotFocus()
MSPack.HighLight = flexHighlightAlways
End Sub

Private Sub mspack_Keypress(K As Integer)
If K > 96 And K < 123 Then
K = K - 32
ElseIf K > 47 And K < 58 Or K = 46 Then

End If
With MSPack
If K = 39 Or K = 41 Or K = 8 Then K = 0
If K = 13 Then
    If .Col = 0 And Trim(.TextMatrix(.Row, 0)) = "" Then
        TxtPack = ""
        If Sch_Cmb.ListIndex = 4 Then
            V_OTH_LST "prd_pck"
        ElseIf Sch_Cmb.ListIndex = 5 Then
            V_OTH_LST "prd_pkt"
        ElseIf Sch_Cmb.ListIndex = 6 Then
            V_OTH_LST "brd_nme"
        ElseIf Sch_Cmb.ListIndex = 7 Then
            V_OTH_LST "grp_cls"
        ElseIf Sch_Cmb.ListIndex = 8 Then
            V_OTH_LST "grp_cat"
        End If
        Exit Sub
    End If
ElseIf .Col = 0 Then
    If K <> 27 Then
        If K = vbKeySpace Then
        Else
            If Trim(.TextMatrix(.Row, 0)) <> "" Then
            Else
                If K = 0 Then Exit Sub
                TxtPack.Visible = True
                TxtPack.Top = .Top + .CellTop
                TxtPack.Left = .CellLeft + .Left
                TxtPack.Width = .CellWidth
                TxtPack.Height = .CellHeight
                TxtPack.SetFocus
                If K <> 13 Then TxtPack = Chr(K)
                TxtPack.SelStart = Len(TxtPack)
            End If
        End If
    End If
ElseIf .Col = 1 Then
    If Trim(.TextMatrix(.Row, 0)) <> "" Then
        K = CHKNUM(K)
        If K <> 0 Then
            TxtPack.Visible = True
            TxtPack.Top = .Top + .CellTop
            TxtPack.Left = .CellLeft + .Left
            TxtPack.Width = .CellWidth
            TxtPack.Height = .CellHeight
            TxtPack.SetFocus
            If K <> 13 Then TxtPack = Chr(K)
            TxtPack.SelStart = Len(TxtPack)
        End If
    End If
End If
End With
End Sub

Private Sub mspack_LeaveCell()
TxtPack.Visible = False
End Sub

Private Sub mspack_lostFocus()
MSPack.HighLight = flexHighlightNever
End Sub

Private Sub MSbase_Keypress(K As Integer)

If K > 96 And K < 123 Then
K = K - 32
ElseIf K > 47 And K < 58 Or K = 46 Then

End If
With MSbase
If K = 39 Or K = 41 Or K = 8 Then K = 0
If K = 13 Then
    If .Col = 0 And Trim(.TextMatrix(.Row, 0)) = "" Then
        LV.Tag = "Base"
        TxtBase = ""
        V_PRD_LST
        Exit Sub
    End If
ElseIf .Col = 0 Then
    If K <> 27 Then
        If K = vbKeySpace Then
        Else
            If Trim(.TextMatrix(.Row, 0)) <> "" Then
            Else
                If K = 0 Then Exit Sub
                TxtBase.Visible = True
                TxtBase.Top = .Top + .CellTop
                TxtBase.Left = .CellLeft + .Left
                TxtBase.Width = .CellWidth
                TxtBase.Height = .CellHeight
                TxtBase.SetFocus
                If K <> 13 Then TxtBase = Chr(K)
                TxtBase.SelStart = Len(TxtBase)
            End If
        End If
    End If
ElseIf .Col = 2 Then
    If Trim(.TextMatrix(.Row, 0)) <> "" Then
        K = CHKNUM(K)
        If K <> 0 Then
            TxtBase.Visible = True
            TxtBase.Top = .Top + .CellTop
            TxtBase.Left = .CellLeft + .Left
            TxtBase.Width = .CellWidth
            TxtBase.Height = .CellHeight
            TxtBase.SetFocus
            If K <> 13 Then TxtBase = Chr(K)
            TxtBase.SelStart = Len(TxtBase)
        End If
    End If
End If
End With
End Sub

Private Sub msfree_Keypress(K As Integer)
If K > 96 And K < 123 Then
K = K - 32
ElseIf K > 47 And K < 58 Or K = 46 Then

End If
With MSFree
If K = 39 Or K = 41 Or K = 8 Then K = 0
If K = 13 Then
    If .Col = 0 And Trim(.TextMatrix(.Row, 0)) = "" Then
        LV.Tag = "Free"
        TxtFree = ""
        V_PRD_LST
        Exit Sub
    End If
ElseIf .Col = 0 Then
    If K <> 27 Then
        If K = vbKeySpace Then
        Else
            If K = 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, 0)) <> "" Then
            Else
                TxtFree.Visible = True
                TxtFree.Top = .Top + .CellTop
                TxtFree.Left = .CellLeft + .Left
                TxtFree.Width = .CellWidth
                TxtFree.Height = .CellHeight
                TxtFree.SetFocus
                If K <> 13 Then TxtFree = Chr(K)
                TxtFree.SelStart = Len(TxtFree)
            End If
        End If
    End If
ElseIf .Col = 2 Then
    If Trim(.TextMatrix(.Row, 0)) <> "" Then
        K = CHKNUM(K)
        If K <> 0 Then
            TxtFree.Visible = True
            TxtFree.Top = .Top + .CellTop
            TxtFree.Left = .CellLeft + .Left
            TxtFree.Width = .CellWidth
            TxtFree.Height = .CellHeight
            TxtFree.SetFocus
            If K <> 13 Then TxtFree = Chr(K)
            TxtFree.SelStart = Len(TxtFree)
        End If
    End If
End If
End With
End Sub

Private Sub Pro_Chk_keypress(K As Integer)
If K = 13 Then
  StDate.SetFocus
End If
End Sub

Private Sub SCH_ACT_keypress(K As Integer)
If K = 13 Then
    If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
        MSbase.TextMatrix(0, 2) = "Qty-Cases"
        MSFree.TextMatrix(0, 2) = "Qty-Cases"
    Else
        MSbase.TextMatrix(0, 2) = "Qty-Units"
        MSFree.TextMatrix(0, 2) = "Qty-Units"
    End If
    
    sch = Sch_Cmb.ListIndex
    SUBCMB = Sub_Cmb.ListIndex
    SSTab1.Tab = 1
    If sch = 1 Then
        TabBase.Tab = 0
        MSbase.Top = 360
        MSbase.Height = 4215
        If MSbase.Enabled = True Then
            MSbase.SetFocus
            MSbase.Col = 0
        End If
    ElseIf sch = 2 Then
        TabBase.Tab = 1
        If Not Len(Trim(Sch_Idy)) > 0 Then
        TxtGrsAmt.SetFocus
        'Else
       ' save
        End If
    ElseIf sch = 3 Then
        TabBase.Tab = 0
        MSbase.Top = 1080
        MSbase.Height = 3495
        If MSbase.Enabled = True Then
            MSbase.SetFocus
            MSbase.Col = 0
        End If
    ElseIf sch > 3 Then
        TabBase.Tab = 2
        MSPack.TextMatrix(0, 0) = Trim(Sch_Cmb)
        TabBase.TabCaption(2) = Trim(Sch_Cmb) & " Details"
        MSPack.Top = 1080
        MSPack.Height = 3495
        If MSPack.Enabled = True Then MSPack.SetFocus
    End If
    If (SUBCMB = 1 And sch = 1) Or (SUBCMB = 1 And sch = 2) Then
        TabFree.Tab = 0
        MSFree.Top = 360
        MSFree.Height = 4215
    ElseIf SUBCMB = 1 And sch = 3 Then
        TabFree.Tab = 0
        MSFree.Top = 1080
        MSFree.Height = 3495
    ElseIf SUBCMB > 1 Then
        TabFree.Tab = 1
        LblValue = Sub_Cmb
    ElseIf SUBCMB = 1 And sch > 3 Then
        TabFree.Tab = 0
        MSFree.Top = 1080
        MSFree.Height = 3495
    End If
End If
End Sub

Private Sub SCH_CMB_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    If Sch_Cmb.ListIndex = 0 Then
        MsgBox "Select Valid Base Entry", vbExclamation
        Sch_Cmb.SetFocus
    Else
        Sub_Cmb.SetFocus
    End If
ElseIf K = 38 And s = 1 Then
    Sch_Nme.SetFocus
End If
End Sub

Private Sub endate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If DateDiff("D", StDate, EnDate) < 0 Then
        SSTab1.Tab = 0
        MsgBox "End Date Should Be Greater Than Start Date", vbExclamation, "DAS Version 4.0"
        EnDate.SetFocus
        Exit Sub
    Else
        If Trim(Sch_Idy) = "" Then
        List2.SetFocus
'        MERNO.SetFocus
        End If
    End If
End If
End Sub

Private Sub SCH_CMB_Validate(cancel As Boolean)
If Sch_Cmb.ListIndex = 0 Then
    cancel = True
    MsgBox "Select Valid Base Entry", vbExclamation
    Sch_Cmb.SetFocus
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    Picture1.Visible = True
ElseIf SSTab1.Tab = 0 Then
    Picture1.Visible = False
    If Sch_Idy.Enabled = True Then
        Sch_Idy.SetFocus
    Else
        Sch_Nme.SetFocus
    End If
End If
End Sub

Private Sub sub_CMB_Validate(cancel As Boolean)
If Sub_Cmb.ListIndex = 0 Then
    cancel = True
    MsgBox "Select Valid Discount Entry", vbExclamation
    Sub_Cmb.SetFocus
End If
End Sub
Private Sub stdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    EnDate.SetFocus
End If
End Sub

Private Sub SUB_CMB_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    If Sub_Cmb.ListIndex = 0 Then
        MsgBox "Select Valid Discount Entry", vbExclamation
        Sub_Cmb.SetFocus
    Else
        If ((Sub_Cmb.ListIndex = 2 Or Sub_Cmb.ListIndex = 1) And Sch_Cmb.ListIndex = 1) Or (Sub_Cmb.ListIndex = 2 And Sch_Cmb.ListIndex = 2) Or (Sub_Cmb.ListIndex = 2 And Sch_Cmb.ListIndex = 3) Then
            Pro_Chk.Enabled = True
            Pro_Chk.SetFocus
        Else
            StDate.SetFocus
        End If
    End If
End If
If K = 38 And s = 1 Then Sch_Cmb.SetFocus
End Sub


Private Sub SCH_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Scheme Identification"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If M_FRM_SCH.ActiveControl.name = "Sch_Idy" Then
        Unload Me
    ElseIf SSTab1.Tab = 2 And lblview = "Select Scheme" Then
        LV.Visible = False
        SSTab1.Tab = 0
        Sch_Idy.SetFocus
    ElseIf lblview = "Select Product" And SSTab1.Tab = 2 Then
        LV.Visible = False
        If LV.Tag = "Base" Then
            SSTab1.Tab = 1
            MSbase.SetFocus
        ElseIf LV.Tag = "Free" Then
            SSTab1.Tab = 1
            MSFree.SetFocus
            MSFree.Col = 0
        End If
    ElseIf SSTab1.Tab = 2 And Sch_Cmb.ListIndex > 3 Then
        lv1.Visible = False
        SSTab1.Tab = 1
        TabBase.Tab = 2
        MSPack.SetFocus
    Else
        save
    End If
End If
End Sub

Private Sub SCH_IDY_Keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    If check("sch", "sch_idy", "SCH" & Trim(Tincrement(Sch_Idy))) = True Then
        DisplayScheme "SCH" & Trim(Tincrement(Sch_Idy))
        
    Else
        Schemelist
        
    End If
End If
End Sub

Private Sub SCH_IDY_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Schemelist()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'        cancel
        Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim ls As ListItem
If rs.State = 1 Then rs.Close
If MOD_REC = True Then
    rs.Open "select * from SCH GROUP BY SCH_IDY", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from SCH where SCH_NME like '" & Trim(Sch_Idy) & "%' GROUP BY SCH_IDY", con, adOpenKeyset, adLockPessimistic
End If
LV.ColumnHeaders.CLEAR
LV.ColumnHeaders.Add , , "Scheme Name", 3000, 0
LV.ColumnHeaders.Add , , "Status", 1000, 2
LV.ColumnHeaders.Add , , "Start Date", 1500, 0
LV.ColumnHeaders.Add , , "End Date", 1500, 0
LV.ColumnHeaders.Add , , "ID", 1500, 0
LV.ColumnHeaders.Add , , "MER Ref", 1500, 0
If rs.RecordCount > 0 Then
    LV.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = LV.ListItems.Add(, , Trim(rs("SCH_nme")))
        ls.ListSubItems.Add , , Trim(rs("SCH_STS"))
        ls.ListSubItems.Add , , Format(rs("SCH_BDT"), "DD/MM/YYYY")
        ls.ListSubItems.Add , , Format(rs("SCH_EDT"), "DD/MM/YYYY")
        ls.ListSubItems.Add , , Trim(rs!Sch_Idy)
        If IsNull(Trim(rs!MER_NUM)) = True Then
            M_MER_NUM = Space(15)
        Else
            M_MER_NUM = Trim(rs!MER_NUM)
        End If
        ls.ListSubItems.Add , , Trim(M_MER_NUM)
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    lblview = "Select Scheme"
    LV.Visible = True
    LV.ListItems(1).Selected = True
    LV.SetFocus
    'If SCH_IDY <> "" Then SCH_LST.SetFocus
ElseIf rs.RecordCount = 0 Then
    MsgBox "No Schemes Found!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Sch_Idy.Enabled = True
    Sch_Idy.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub SCH_NME_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Scheme, Discount Entry Description"
If Len(Trim(Sch_Idy)) > 0 And Sch_Idy.Enabled = False Then
    Dim rs As New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from sch where sch_idy like '" & Trim(Sch_Idy) & " ' AND sch_sts LIKE 'N' ", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
    Sch_Act.Enabled = False
    Else
    Sch_Act.Enabled = True
    End If
End If
End Sub

Private Sub SCH_NME_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If M_FRM_SCH.ActiveControl.name = "Sch_Idy" Then
        Schemelist
        Sch_Act.SetFocus
    End If
ElseIf Button.KEY = "crt" Then
    SchemeCopy
ElseIf Button.KEY = "p" Then
    M_FRM_DSN.Tag = "Schemes List"
    M_FRM_DSN.Show 1
    cancel
'    Openreport CURDIR & "\lists.RP1", "Schemes List", 0, "", 6
End If
End Sub

Public Sub newrec()
cancel
Sch_Idy.Enabled = False
Sch_Nme.SetFocus
End Sub
Public Sub SchemeCopy()
If Len(Trim(Sch_Idy)) <> 0 Then
    Dim SRC_SCH As New ADODB.Recordset
    Dim DST_SCH As New ADODB.Recordset
    
    Dim SRC_SID As New ADODB.Recordset
    Dim DST_SID As New ADODB.Recordset
    
    If SRC_SID.State = 1 Then SRC_SID.Close
    SRC_SID.Open "SELECT * FROM SID WHERE SCH_IDY LIKE '" & Trim(Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    If SRC_SID.RecordCount > 0 Then SRC_SID.MoveFirst
    
    If DST_SID.State = 1 Then DST_SID.Close
    DST_SID.Open "SELECT * FROM SID WHERE SCH_IDY LIKE '" & Trim(Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    
    If SRC_SCH.State = 1 Then SRC_SCH.Close
    SRC_SCH.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    If SRC_SCH.RecordCount > 0 Then SRC_SCH.MoveFirst
    
    If DST_SCH.State = 1 Then DST_SCH.Close
    DST_SCH.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
    
    
    
    If MsgBox("Confirm Scheme Copy and Save ? ", vbYesNo + vbInformation, "DAS Version 4.0") = vbYes Then
        schid = "SCH" & T7increment("idy_sch")
        While SRC_SCH.EOF = False
            DST_SCH.AddNew
            DST_SCH("SCH_IDY") = Trim(schid)
            DST_SCH("SCH_NME") = Trim(SRC_SCH("SCH_NME"))
            DST_SCH("SCH_TYP") = Trim(SRC_SCH("SCH_TYP"))
            DST_SCH("F_SCH_TYP") = Trim(SRC_SCH("F_SCH_TYP"))
            DST_SCH("BAT_IDY") = Trim(SRC_SCH("BAT_IDY"))
            DST_SCH("B_PRD_IDY") = Trim(SRC_SCH("B_PRD_IDY"))
            DST_SCH("B_PRD_QTY") = SRC_SCH("B_PRD_QTY")
            DST_SCH("E_PRD_QTY") = SRC_SCH("E_PRD_QTY")
            DST_SCH("F_PRD_IDY") = Trim(SRC_SCH("F_PRD_IDY"))
            DST_SCH("F_PRD_QTY") = SRC_SCH("F_PRD_QTY")
            DST_SCH("B_PRD_AMT") = SRC_SCH("B_PRD_AMT")
            DST_SCH("B_DIS_PCG") = SRC_SCH("B_DIS_PCG")
            DST_SCH("B_DIS_AMT") = SRC_SCH("B_DIS_AMT")
            DST_SCH("B_BIL_AMT") = SRC_SCH("B_BIL_AMT")
            DST_SCH("E_BIL_AMT") = SRC_SCH("E_BIL_AMT")
            DST_SCH("SCH_BDT") = Date 'SRC_SCH("SCH_BDT")
            DST_SCH("SCH_EDT") = Date 'SRC_SCH("SCH_EDT")
            DST_SCH("TXN_TYP") = Trim(SRC_SCH("TXN_TYP"))
            DST_SCH("SCH_STS") = Trim("Y") 'SRC_SCH("SCH_STS")
            DST_SCH("USR_IDY") = Trim(SRC_SCH("USR_IDY"))
            DST_SCH("SYS_DAT") = Date 'SRC_SCH("SYS_DAT")
            DST_SCH("SYS_TIM") = Format(Time, "hh:mm:ss")
            DST_SCH("PRO_RAT") = SRC_SCH("PRO_RAT")
            DST_SCH("B_CON_APP") = SRC_SCH("B_CON_APP")
            DST_SCH("F_CON_APP") = SRC_SCH("F_CON_APP")
            DST_SCH("GRP_TYP") = SRC_SCH("GRP_TYP")
            DST_SCH("MER_NUM") = SRC_SCH("MER_NUM")
            DST_SCH("MER_QTY") = SRC_SCH("MER_QTY")
            DST_SCH("MER_AMT") = SRC_SCH("MER_AMT")
            DST_SCH.Update
            SRC_SCH.MoveNext
        Wend
        
        While SRC_SID.EOF = False
            DST_SID.AddNew
            DST_SID("SCH_IDY") = Trim(schid)
            DST_SID("CST_IDY") = Trim(SRC_SID("CST_IDY"))
            DST_SID("SMN_IDY") = Trim(SRC_SID("SMN_IDY"))
            DST_SID("ROU_IDY") = Trim(SRC_SID("ROU_IDY"))
            DST_SID.Update
            SRC_SID.MoveNext
        Wend
        MsgBox "New Scheme ID is " & schid, vbInformation, "DAS Version 4.0"
        cancel
    Else
        cancel
    End If
Else
    MsgBox "Select Existing Scheme for Scheme Cascade - Copy !", vbInformation, "DAS Version 4.0"
    cancel
End If
End Sub

Public Sub save()
If MsgBox("Confirm Scheme Save?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If Len(Trim(Sch_Nme)) = 0 Then
    MsgBox "Scheme Name Should not be blank", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Sch_Nme.SetFocus
    Exit Sub
End If

If Len(Trim(Sch_Idy)) <> 0 Then
    
'    Dim rs As New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from sch where sch_idy like '" & Trim(Sch_Idy) & " ' AND sch_sts LIKE 'N' ", CON, adOpenKeyset, adlockpessimistic
'    If Not rs.RecordCount > 0 Then
        con.Execute "UPDATE SCH SET sch_edt = ctod('" & Trim(EnDate) & "') where sch_idy like '" & Trim(Sch_Idy) & "'", d
        con.Execute "UPDATE SCH SET sch_sts = '" & Sch_Act & "' where sch_idy like '" & Trim(Sch_Idy) & "'", d
        con.Execute "UPDATE SCH SET CMP_DST = '" & List2 & "' where sch_idy like '" & Trim(Sch_Idy) & "'", d
        con.Execute "UPDATE SCH SET MER_NUM = '" & Trim(MERNO) & "' WHERE SCH_IDY LIKE '" & Trim(Sch_Idy) & "'", d
'    Else
'        MsgBox "Scheme Cannot be Made Active Again!", vbCritical, "DAS Version 4.0"
'    End If
    cancel
    Exit Sub
End If
If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
    SAVESKUFREEQUANTITY
ElseIf Sch_Cmb.ListIndex = 3 And Sub_Cmb.ListIndex = 1 Then
    SAVEMULTIPLEFREEQUANTITY
ElseIf Sch_Cmb.ListIndex = 3 And Sub_Cmb.ListIndex > 1 Then
    SAVEMULTIPLECASHORPCG
ElseIf Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex > 1 Then
    saveskucashordiscount
ElseIf Sch_Cmb.ListIndex = 2 Then
    savebillcashordiscount
ElseIf Sch_Cmb.ListIndex > 3 Then
    SAVEGRPTYP
End If
End Sub

Public Sub cancel()
SETYEAROPEN Me
clrctr Me
SSTab1.Tab = 0
Pro_Chk.Value = 0
Pro_Chk.Enabled = False
TabBase.TabCaption(2) = "         "
CmbBase.ListIndex = 0
CmbFree.ListIndex = 0
CmbPCon.ListIndex = 0
Sch_Cmb.Enabled = True
Sub_Cmb.Enabled = True
MSPack.Enabled = True
MSbase.Enabled = True
MSFree.Enabled = True
Sch_Cmb.ListIndex = 0
Sub_Cmb.ListIndex = 0
Frame4.Enabled = True
StDate.Enabled = True
TabBase.Enabled = True
TabFree.Enabled = True
TxtValue = "0.00"
TxtGrsAmt = "0.00"
MSbase.Rows = 1
MSbase.Rows = 2
MSFree.Rows = 1
MSFree.Rows = 2
MSPack.Rows = 1
MSPack.Rows = 2
MERQTY = "0.000"
MERAMT = "0.00"
Sch_Act.Enabled = True
Sch_Act.Selected(0) = True
List2.Enabled = True

Frame6.Enabled = False
list1.ListItems.CLEAR
Combo1.ListIndex = 0
Combo1.Enabled = True

Command1.Caption = "Check &All"
Sch_Idy.Enabled = True
Sch_Idy.SetFocus

End Sub
Public Sub DisplayScheme(Schemeid As String)
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
rs.Open "select * from sch where sch_idy like '" & Trim(Schemeid) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If RS1.State = 1 Then RS1.Close
    RS1.Open "SELECT * FROM SID WHERE SCH_IDY = '" & Trim(Schemeid) & "'", con, adOpenKeyset, adLockPessimistic
    For A = 1 To RS1.RecordCount
        If Len(Trim(RS1!CST_IDY)) > 0 Then
            Combo1.ListIndex = 1
            list1.ListItems.Add , , GetValue("ACT_NME", "ACT WHERE ACT_IDY = '" & Trim(RS1!CST_IDY & "") & "'")
        ElseIf Len(Trim(RS1!SMN_IDY)) > 0 Then
            Combo1.ListIndex = 2
            list1.ListItems.Add , , GetValue("SMN_NME", "SMN WHERE SMN_IDY = '" & Trim(RS1!SMN_IDY & "") & "'")
        ElseIf Len(Trim(RS1!rou_idy)) > 0 Then
            Combo1.ListIndex = 3
            list1.ListItems.Add , , Trim(RS1!rou_idy) & ""
        End If
        RS1.MoveNext
    Next
    For A = 1 To list1.ListItems.count
        list1.ListItems.Item(A).Checked = True
    Next
    Sch_Idy = Trim(Schemeid)
    Sch_Idy.Enabled = False
    Sch_Nme = Trim(rs!Sch_Nme)
    StDate = Trim(rs!sch_bdt)
    EnDate = Trim(rs!sch_edt)
    Sch_Act = Trim(rs!sch_sts)
    If IsNull(rs("CMP_DST")) = False Then
        List2 = Trim(rs!CMP_DST)
    Else
        List2 = 0
    End If
    MERNO = Trim(rs!MER_NUM & "")
    MERQTY = QTY(rs!MER_QTY & "")
    MERAMT = convert(rs!MER_AMT & "")
    For A = 0 To Sch_Cmb.ListCount - 1
        If Trim(Sch_Cmb.List(A)) = Trim(rs!SCH_TYP) Then
            Sch_Cmb.ListIndex = A
        End If
    Next
    For A = 0 To Sub_Cmb.ListCount - 1
        If Trim(Sub_Cmb.List(A)) = Trim(rs!f_sch_typ) Then
            Sub_Cmb.ListIndex = A
        End If
    Next
    SCH_ACT_keypress 13

    Sch_Cmb.Enabled = False
    Sub_Cmb.Enabled = False
    StDate.Enabled = False
    TabBase.Enabled = False
    TabFree.Enabled = False
    MSPack.Enabled = False
    Frame4.Enabled = False
    MSbase.Enabled = False
    MSFree.Enabled = False
    If Sch_Cmb.ListIndex > 3 Then
        dispschemegrt3
    ElseIf Sch_Cmb.ListIndex = 1 Or Sch_Cmb.ListIndex = 3 Then
        DISPSCHEMESKU
    ElseIf Sch_Cmb.ListIndex = 2 And (Sub_Cmb.ListIndex = 2 Or Sub_Cmb.ListIndex = 3) Then
        If Sub_Cmb.ListIndex = 2 Then TxtValue = convert(rs!B_DIS_AMT & "")
        If Sub_Cmb.ListIndex = 3 Then TxtValue = convert(rs!B_DIS_PCG & "")
        If rs.RecordCount > 0 Then TxtGrsAmt = convert(rs!b_BIL_AMT & "")
    ElseIf Sch_Cmb.ListIndex = 2 And Sub_Cmb.ListIndex = 1 Then
        rs.filter = "txn_typ='FRE'"
        If rs.RecordCount > 0 Then TxtGrsAmt = convert(rs!b_BIL_AMT & "")
        If rs!F_CON_APP = "AND" Then
            CmbFree.ListIndex = 0
        ElseIf rs!F_CON_APP = "ORO" Then
            CmbFree.ListIndex = 1
        End If
        MSFree.Rows = rs.RecordCount + 1
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                MSFree.TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")
                MSFree.TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"))
                If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
                    MSFree.TextMatrix(A, 2) = QTY(STM(Val(rs!B_PRD_QTY & ""), GetValue("sub_unt", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")))
                Else
                    MSFree.TextMatrix(A, 2) = Val(rs!B_PRD_QTY & "")
                End If
                rs.MoveNext
            Next
        End If
    End If
    TxtBase.Visible = False
    TxtFree.Visible = False
    TxtPack.Visible = False
End If
End Sub

Public Sub V_PRD_IDY_LST()
If CHECKTB("prd.dbf,brd.dbf,prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
If check("prt", "prt_nme", "prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
LV.ColumnHeaders.CLEAR
LV.ColumnHeaders.Add , , "Name", 2500, 0
LV.ColumnHeaders.Add , , "Quantity", 1500, 1
LV.ColumnHeaders.Add , , "MRP", 1300, 1
LV.ColumnHeaders.Add , , "Sale Rate", 1300, 1
LV.ColumnHeaders.Add , , "Purchase Rate", 1300, 1
LV.ColumnHeaders.Add , , "Batid", 0, 0
LV.ColumnHeaders.Add , , "prdid", 0, 0
LV.ListItems.CLEAR
'REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_IDY LIKE '" & Sincrement(TxtBase) & "%' AND PRD.STK_TYP IN('Salable') and BRD.prd_qty >0 order by brw_seq", CON, adOpenKeyset
REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_IDY LIKE '" & Sincrement(TxtBase) & "%' AND PRD.STK_TYP IN('Salable') order by brw_seq", con, adOpenKeyset 'Modified by Radhika and Ramanesh
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
Set ls = LV.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , QTY(REC_SET("prd_qty") & "")
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("prd_mrp") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("prd_sdr") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("bas_rat") & "")))
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET!prd_idy & "")
        REC_SET.MoveNext
    Next
    lblview = "Select Product"
    SSTab1.Tab = 2
    LV.Visible = True
    LV.SetFocus
Else
    If SSTab1.Tab = 1 Then
        MsgBox "Product Not Found, Define Product !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 1
    End If
End If
End Sub

Public Sub V_PRD_LST()
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim STK As Boolean
Dim REC_SET As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If LV.Tag = "Base" Then
    'REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(TxtBase) & "%'", CON, adOpenKeyset
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 and prd.PRD_NME LIKE '" & Trim(TxtBase) & "%' and stk_typ in('Salable','Free')", con, adOpenKeyset 'Modified By Radhika and Ramanesh
Else
    'REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(TxtFree) & "%'", CON, adOpenKeyset
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 and prd.PRD_NME LIKE '" & Trim(TxtFree) & "%' and stk_typ in('Salable','Free')", con, adOpenKeyset 'Modified by Radhika and Ramanesh
End If
LV.ColumnHeaders.CLEAR
LV.ColumnHeaders.Add , , "Name", 2500, 0
LV.ColumnHeaders.Add , , "Quantity", 1500, 1
LV.ColumnHeaders.Add , , "MRP", 1300, 1
LV.ColumnHeaders.Add , , "Sale Rate", 1300, 1
LV.ColumnHeaders.Add , , "Purchase Rate", 1300, 1
LV.ColumnHeaders.Add , , "Batid", 0, 0
LV.ColumnHeaders.Add , , "prdid", 0, 0
LV.ColumnHeaders.Add , , "Subunit", 0, 0
LV.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = LV.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , QTY(REC_SET("prd_qty") & "")
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("prd_mrp") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("prd_sdr") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("bas_rat") & "")))
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET!prd_idy & "")
        ls.ListSubItems.Add , , Trim(REC_SET!SUB_UNT & "")
        REC_SET.MoveNext
    Next
    lblview = "Select Product"
    SSTab1.Tab = 2
    LV.Visible = True
    LV.SetFocus
Else
    MsgBox "Product Not Found, Define Product !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 1
    If LV.Tag = "Base" Then
        LV.Tag = ""
        MSbase.SetFocus
        TxtBase.Visible = False
    ElseIf LV.Tag = "Free" Then
        LV.Tag = ""
        MSFree.SetFocus
        MSFree.Col = 0
        TxtFree.Visible = False
    End If
    Exit Sub
End If
End Sub
Private Sub txtbase_keypress(K As Integer)
If MSbase.Col = 0 Then
    If K > 96 And K < 123 Then K = K - 32
    If K = 39 Or K = 41 Or K = 37 Then K = 0
Else
    If K <> 13 Then K = CHKNUM(K)
End If
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If MSbase.Col = 0 Then
    If K = 13 Then
        If MSbase.Col = 0 Then
            If REC_SET.State = 1 Then REC_SET.Close
            'REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(TxtBase) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
            REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(TxtBase) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable')", con, adOpenKeyset 'Modified by Radhika and Ramanesh
            If REC_SET.RecordCount = 1 Then
                P_DIS_PRD Trim(REC_SET!BAT_IDY), "Base"
            ElseIf REC_SET.RecordCount > 1 Then
                TxtBase = Sincrement(TxtBase)
                LV.Tag = "Base"
                V_PRD_IDY_LST
            Else
                LV.Tag = "Base"
                V_PRD_LST
            End If
        End If
    End If
ElseIf MSbase.Col = 2 Then
    If K <> 13 Then K = CHKNUM(K)
    If K = 13 Then
        If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
            If CmbBase.ListIndex <> 0 Then
                TxtBase = QTY(TxtBase)
                b = InStr(TxtBase, ".")
                If b > 0 Then
                    If InStr(b + 1, TxtBase, ".") > 0 Then
                        MSbase.TextMatrix(MSbase.Row, 3) = "0.000"
                        Exit Sub
                    End If
                    M = Mid(TxtBase, b + 1, 3)
                    If M >= Val(MSbase.TextMatrix(MSbase.Row, 5)) Then
                        MsgBox "Product Sub Units Are " & Val(MSbase.TextMatrix(MSbase.Row, 5)), vbExclamation, "DAS Version 4.0"
                        TxtBase.Visible = True
                        TxtBase.SetFocus
                        SendKeys "{home}+{end}"
                        Exit Sub
                    End If
                End If
                MSbase.TextMatrix(1, 2) = QTY(Val(TxtBase))
            Else
                TxtBase = QTY(TxtBase)
                b = InStr(TxtBase, ".")
                If b > 0 Then
                    If InStr(b + 1, TxtBase, ".") > 0 Then
                        MSbase.TextMatrix(MSbase.Row, 3) = "0.000"
                        Exit Sub
                    End If
                    M = Mid(TxtBase, b + 1, 3)
                    If M >= Val(MSbase.TextMatrix(MSbase.Row, 5)) Then
                        MsgBox "Product Sub Units Are " & Val(MSbase.TextMatrix(MSbase.Row, 5)), vbExclamation, "DAS Version 4.0"
                        TxtBase.Visible = True
                        TxtBase.SetFocus
                        SendKeys "{home}+{end}"
                        Exit Sub
                    End If
                End If
                MSbase = QTY(Val(TxtBase))
            End If
            TxtBase.Visible = False
            MSbase.SetFocus
            If Sch_Cmb.ListIndex = 3 Then
                If MSbase.TextMatrix(MSbase.Rows - 1, 0) <> "" Then MSbase.AddItem ""
                MSbase.Row = MSbase.Rows - 1
                MSbase.Col = 0
                MSbase.SetFocus
            Else
                If Sub_Cmb.ListIndex = 1 Then
                    If Val(MSbase.Text) <> 0 Then
                        MSFree.SetFocus
                        MSFree.Col = 0
                    End If
                Else
                    SSTab1.Tab = 1
                    TxtValue.SetFocus
                    SLT
                End If
            End If
        ElseIf Sch_Cmb.ListIndex = 3 Then
            TxtBase = Int(Val(TxtBase))
            If CmbBase.ListIndex <> 0 Then
                MSbase.TextMatrix(1, 2) = Val(TxtBase)
            Else
                MSbase = Val(TxtBase)
            End If
            TxtBase.Visible = False
            MSbase.SetFocus
            If Sch_Cmb.ListIndex = 3 Then
                If MSbase.TextMatrix(MSbase.Rows - 1, 0) <> "" Then MSbase.AddItem ""
                MSbase.Row = MSbase.Rows - 1
                MSbase.Col = 0
                MSbase.SetFocus
            Else
                If Sub_Cmb.ListIndex = 1 Then
                    If Val(MSbase.Text) <> 0 Then
                        MSFree.SetFocus
                        MSFree.Col = 0
                    End If
                Else
                    SSTab1.Tab = 1
                    TxtValue.SetFocus
                    SLT
                End If
            End If
        Else
            TxtBase = Int(Val(TxtBase))
            MSbase = Val(TxtBase)
            TxtBase.Visible = False
            MSbase.SetFocus
            If Sch_Cmb.ListIndex = 3 Then
                If MSbase.TextMatrix(MSbase.Rows - 1, 0) <> "" Then MSbase.AddItem ""
                MSbase.Row = MSbase.Rows - 1
                MSbase.Col = 0
                MSbase.SetFocus
            Else
                If Sub_Cmb.ListIndex = 1 Then
                    If Val(MSbase.Text) <> 0 Then
                        MSFree.SetFocus
                        MSFree.Col = 0
                    End If
                Else
                    SSTab1.Tab = 1
                    TxtValue.SetFocus
                    SLT
                End If
            End If
        End If
    End If
End If
End Sub
Private Sub txtpack_keypress(K As Integer)
If MSPack.Col = 0 Then
    If K > 96 And K < 123 Then K = K - 32
    If K = 39 Or K = 41 Or K = 37 Then K = 0
Else
    If K <> 13 Then K = CHKNUM(K)
End If
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If MSPack.Col = 0 Then
    If K = 13 Then
        If MSPack.Col = 0 Then
            If Sch_Cmb.ListIndex = 4 Then
                V_OTH_LST "prd_pck"
            ElseIf Sch_Cmb.ListIndex = 5 Then
                V_OTH_LST "prd_pkt"
            ElseIf Sch_Cmb.ListIndex = 6 Then
                V_OTH_LST "brd_nme"
            ElseIf Sch_Cmb.ListIndex = 7 Then
                V_OTH_LST "grp_cls"
            ElseIf Sch_Cmb.ListIndex = 8 Then
                V_OTH_LST "grp_cat"
            End If
        End If
    End If
ElseIf MSPack.Col = 1 Then
    If K <> 13 Then K = CHKNUM(K)
    If K = 13 Then
        If CmbPCon.ListIndex = 0 Then
            MSPack = Int(Val(TxtPack))
        Else
            MSPack.TextMatrix(1, 1) = Int(Val(TxtPack))
        End If
        TxtPack.Visible = False
        MSPack.SetFocus
        If MSPack.TextMatrix(MSPack.Rows - 1, 0) <> "" Then MSPack.AddItem ""
        MSPack.Row = MSPack.Rows - 1
        MSPack.Col = 0
        MSPack.SetFocus
    End If
End If
End Sub
Private Sub txtFree_keypress(K As Integer)
If MSFree.Col = 0 Then
    If K > 96 And K < 123 Then K = K - 32
    If K = 39 Or K = 41 Or K = 37 Then K = 0
Else
    If K <> 13 Then K = CHKNUM(K)
End If

If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If MSFree.Col = 0 Then
    If K = 13 Then
        If MSFree.Col = 0 Then
            If REC_SET.State = 1 Then REC_SET.Close
            'REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(TxtFree) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
            REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(TxtFree) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable')", con, adOpenKeyset 'Modified by Radhika and Ramanesh
            If REC_SET.RecordCount = 1 Then
                P_DIS_PRD Trim(REC_SET!BAT_IDY), "Free"
            ElseIf REC_SET.RecordCount > 1 Then
                TxtFree = Sincrement(TxtFree)
                LV.Tag = "Free"
                V_PRD_IDY_LST
            Else
                LV.Tag = "Free"
                V_PRD_LST
            End If
        End If
    End If
ElseIf MSFree.Col = 2 Then
    If K <> 13 Then K = CHKNUM(K)
    If K = 13 Then
        If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
            If CmbFree.ListIndex <> 0 Then
                TxtFree = QTY(TxtFree)
                b = InStr(TxtFree, ".")
                If b > 0 Then
                    If InStr(b + 1, TxtFree, ".") > 0 Then
                        MSFree.TextMatrix(MSFree.Row, 3) = "0.000"
                        Exit Sub
                    End If
                    M = Mid(TxtFree, b + 1, 3)
                    If M >= Val(MSFree.TextMatrix(MSFree.Row, 5)) Then
                        MsgBox "Product Sub Units Are " & Val(MSFree.TextMatrix(MSFree.Row, 5)), vbExclamation, "DAS Version 4.0"
                        TxtFree.Visible = True
                        TxtFree.SetFocus
                        SendKeys "{home}+{end}"
                        Exit Sub
                    End If
                End If
                MSFree.TextMatrix(1, 2) = QTY(TxtFree)
            Else
                TxtFree = QTY(TxtFree)
                b = InStr(TxtFree, ".")
                If b > 0 Then
                    If InStr(b + 1, TxtFree, ".") > 0 Then
                        MSFree.TextMatrix(MSFree.Row, 3) = "0.000"
                        Exit Sub
                    End If
                    M = Mid(TxtFree, b + 1, 3)
                    If M >= Val(MSFree.TextMatrix(MSFree.Row, 5)) Then
                        MsgBox "Product Sub Units Are " & Val(MSFree.TextMatrix(MSFree.Row, 5)), vbExclamation, "DAS Version 4.0"
                        TxtFree.Visible = True
                        TxtFree.SetFocus
                        SendKeys "{home}+{end}"
                        Exit Sub
                    End If
                End If
                MSFree = QTY(Val(TxtFree))
            End If
            TxtFree.Visible = False
            If Sch_Cmb.ListIndex >= 3 Then
                If MSFree.TextMatrix(MSFree.Rows - 1, 0) <> "" Then MSFree.AddItem ""
                MSFree.Row = MSFree.Rows - 1
                MSFree.Col = 0
                MSFree.SetFocus
            Else
                If Val(MSFree.Text) <> 0 Then save
            End If
        ElseIf Sch_Cmb.ListIndex >= 3 Then
            If CmbFree.ListIndex <> 0 Then
                    TxtFree = Int(Val(TxtFree))
                    MSFree.TextMatrix(1, 2) = Val(TxtFree)
                Else
                    TxtFree = Int(Val(TxtFree))
                    MSFree = (Val(TxtFree))
                End If
                TxtFree.Visible = False
                If Sch_Cmb.ListIndex >= 3 Then
                    If MSFree.TextMatrix(MSFree.Rows - 1, 0) <> "" Then MSFree.AddItem ""
                    MSFree.Row = MSFree.Rows - 1
                    MSFree.Col = 0
                    MSFree.SetFocus
                Else
                    If Val(MSFree.Text) <> 0 Then save
                End If
        ElseIf Sub_Cmb.ListIndex = 1 And Sch_Cmb.ListIndex = 2 Then
                MSFree = Int(Val(TxtFree))
                If MSFree.TextMatrix(MSFree.Rows - 1, 0) <> "" Then MSFree.AddItem ""
                MSFree.Row = MSFree.Rows - 1
                MSFree.Col = 0
                MSFree.SetFocus
        Else
            MSFree.SetFocus
            MSFree.Col = 0
        End If
    End If
End If
End Sub

Public Sub P_DIS_PRD(BatchID As String, Grid As String)
Dim rs As New ADODB.Recordset
If Grid = "Base" Then
    With MSbase
        For A = 1 To .Rows - 1
            If .TextMatrix(A, 3) = BatchID Then
                SSTab1.Tab = 1
                MsgBox "Product Already Exists", vbExclamation
                LV.Visible = False
                TabBase.Tab = 0
                .Row = A
                .SetFocus
                Exit Sub
            End If
        Next
        If rs.State = 1 Then rs.Close
        'RS.Open "SELECT * FROM prd,BRD WHERE brd.bat_IDY LIKE '" & Trim(BatchID) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
        rs.Open "SELECT * FROM prd,BRD WHERE brd.bat_IDY LIKE '" & Trim(BatchID) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable')", con, adOpenKeyset 'Modified by Radhika and Ramanesh
        If rs.RecordCount > 0 Then
            .TextMatrix(.Row, 0) = Trim(rs!PRD_NME)
           ' .TextMatrix(.Row, 1) = convert((Val(RS!PRD_MRP)) / Val(RS!sub_unt))
            .TextMatrix(.Row, 1) = convert((Val(rs!PRD_MRP)))
            .TextMatrix(.Row, 3) = Trim(rs!BAT_IDY)
            .TextMatrix(.Row, 4) = Trim(rs!prd_idy)
            .TextMatrix(.Row, 5) = Trim(rs!SUB_UNT)
        End If
        SSTab1.Tab = 1
        TxtBase.Visible = False
        .Col = 2
        .SetFocus
    End With
Else
    With MSFree
        For A = 1 To .Rows - 1
            If .TextMatrix(A, 3) = BatchID Then
                SSTab1.Tab = 1
                MsgBox "Product Already Exists", vbExclamation
                LV.Visible = False
                'TabBase.Tab = 2
                .Row = A
                .SetFocus
                Exit Sub
            End If
        Next
        If rs.State = 1 Then rs.Close
        'RS.Open "SELECT * FROM prd,BRD WHERE brd.bat_IDY LIKE '" & Trim(BatchID) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
        rs.Open "SELECT * FROM prd,BRD WHERE brd.bat_IDY LIKE '" & Trim(BatchID) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable')", con, adOpenKeyset 'Modified by Radhika & Ramanesh
        If rs.RecordCount > 0 Then
            .TextMatrix(.Row, 0) = Trim(rs!PRD_NME)
            '.TextMatrix(.Row, 1) = convert((Val(RS!PRD_MRP)) / Val(RS!sub_unt))
            .TextMatrix(.Row, 1) = convert((Val(rs!PRD_MRP)))
            .TextMatrix(.Row, 3) = Trim(rs!BAT_IDY)
            .TextMatrix(.Row, 4) = Trim(rs!prd_idy)
            .TextMatrix(.Row, 5) = Trim(rs!SUB_UNT)
        End If
        SSTab1.Tab = 1
        TxtFree.Visible = False
        .Col = 2
        .SetFocus
    End With
End If
End Sub

Public Sub SAVESKUFREEQUANTITY()
On Error GoTo lab
If Val(MSbase.TextMatrix(1, 2)) = 0 Then
    MsgBox "Enter Valid Quantity", vbExclamation
    SSTab1.Tab = 1
    TabBase.Tab = 0
    MSbase.SetFocus
    MSbase.Col = 2
    MSbase.Row = 1
    Exit Sub
End If
If Val(MSFree.TextMatrix(1, 2)) = 0 Then
    MsgBox "Enter Valid Quantity", vbExclamation
    SSTab1.Tab = 1
    TabFree.Tab = 0
    MSFree.SetFocus
    MSFree.Col = 0
    MSFree.Row = 1
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch", con, adOpenKeyset, adLockPessimistic
con.BeginTrans
schid = "SCH" & T7increment("idy_sch")
SID Trim(schid)
rs.AddNew
rs("sch_idy") = schid
rs("sch_nme") = Trim(Sch_Nme)
rs("sch_typ") = Sch_Cmb
rs("f_sch_typ") = Sub_Cmb
rs("SCH_STS") = Sch_Act
rs("CMP_DST") = List2
rs("SCH_BDT") = StDate
rs("sch_edt") = EnDate
rs!MER_NUM = Trim(MERNO)
rs!MER_QTY = Val(MERQTY)
rs!MER_AMT = Val(MERAMT)
rs("B_PRD_IDY") = Trim(MSbase.TextMatrix(1, 4))
rs("B_PRD_QTY") = MTS(Val(MSbase.TextMatrix(1, 2)), Val(MSbase.TextMatrix(1, 5)))
rs("BAT_IDY") = Trim(MSbase.TextMatrix(1, 3))
rs("txn_typ") = "BAS"
rs("USR_IDY") = USERNAME
rs("SYS_DAT") = Date
rs("SYS_TIM") = Format(Time, "hh:mm:ss")
rs("CMP_DST") = List2
rs.Update

rs.AddNew
rs("sch_idy") = schid
rs("sch_nme") = Trim(Sch_Nme)
rs("sch_typ") = Sch_Cmb
rs("f_sch_typ") = Sub_Cmb
rs("SCH_STS") = Sch_Act
rs("CMP_DST") = List2
rs("SCH_BDT") = StDate
rs("sch_edt") = EnDate
rs!MER_NUM = Trim(MERNO)
rs!MER_QTY = Val(MERQTY)
rs!MER_AMT = Val(MERAMT)
rs("B_PRD_IDY") = Trim(MSFree.TextMatrix(1, 4))
rs("B_PRD_QTY") = MTS(Val(MSFree.TextMatrix(1, 2)), Val(MSFree.TextMatrix(1, 5)))
rs("f_prd_idy") = Trim(MSbase.TextMatrix(1, 3))
rs("BAT_IDY") = Trim(MSFree.TextMatrix(1, 3))
rs("txn_typ") = "FRE"
If Pro_Chk.Value = 1 Then
    rs("pro_rat") = "Y"
Else
    rs("pro_rat") = "N"
End If
rs("USR_IDY") = USERNAME
rs("SYS_DAT") = Date
rs("SYS_TIM") = Format(Time, "hh:mm:ss")
rs.Update
MsgBox "New Scheme ID is " & schid, vbExclamation, "DAS Version 4.0"
con.CommitTrans
cancel
Exit Sub
lab:
MsgBox err.Description
con.RollbackTrans
End Sub

Public Sub SAVEMULTIPLEFREEQUANTITY()
On Error GoTo lab
Dim rs As New ADODB.Recordset, BAS_VAL As Integer
If rs.State = 1 Then rs.Close
rs.Open "select * from sch", con, adOpenKeyset, adLockPessimistic
BAS_VAL = 0
For i = 1 To MSbase.Rows - 1
    If Val(MSbase.TextMatrix(i, 2)) > 0 Then
        BAS_VAL = 1
        Exit For
    End If
Next
If BAS_VAL <> 1 Then
    MsgBox "Cannot Proceed Further! Base Quantity in Units Required!", vbExclamation
    MSbase.SetFocus
    MSbase.Col = 2
    Exit Sub
End If
con.BeginTrans
schid = "SCH" & T7increment("idy_sch")
SID Trim(schid)
If MSFree.Rows < 2 Then
    MsgBox "Select Free Product", vbExclamation
    SSTab1.Tab = 1
    TabBase.Tab = 0
    MSbase.Col = 0
    MSbase.Row = 1
    MSbase.SetFocus
    Exit Sub
End If
If MSFree.Rows < 2 Then
    MsgBox "Select Free Product", vbExclamation
    SSTab1.Tab = 1
    TabFree.Tab = 0
    MSFree.Col = 0
    MSFree.Row = 1
    MSFree.SetFocus
    Exit Sub
End If
For A = 1 To MSbase.Rows - 1
    If MSbase.TextMatrix(A, 0) <> "" Then
        rs.AddNew
        rs("sch_idy") = schid
        rs("sch_nme") = Trim(Sch_Nme)
        rs("sch_typ") = Sch_Cmb
        rs("f_sch_typ") = Sub_Cmb
        rs("SCH_STS") = Sch_Act
        rs("CMP_DST") = List2
        rs("SCH_BDT") = StDate
        rs("sch_edt") = EnDate
        rs!MER_NUM = Trim(MERNO)
        rs!MER_QTY = Val(MERQTY)
        rs!MER_AMT = Val(MERAMT)
        rs("B_PRD_IDY") = Trim(MSbase.TextMatrix(A, 4))
        rs("BAT_IDY") = Trim(MSbase.TextMatrix(A, 3))
        rs("txn_typ") = "BAS"
        rs("USR_IDY") = USERNAME
        If Val(MSbase.TextMatrix(A, 2)) > 0 Then
            rs("B_PRD_QTY") = (Val(MSbase.TextMatrix(A, 2)))
        Else
            rs("B_PRD_QTY") = 0
        End If
        If CmbBase.ListIndex = 0 Then
            rs("b_con_app") = "AND"
        ElseIf CmbBase.ListIndex = 1 Then
            rs("b_con_app") = "ORO"
        ElseIf CmbBase.ListIndex = 2 Then
            rs("b_con_app") = "ORM"
        End If
        rs("SYS_DAT") = Date
        rs("SYS_TIM") = Format(Time, "hh:mm:ss")
        rs.Update
    End If
Next

For A = 1 To MSFree.Rows - 1
    If MSFree.TextMatrix(A, 0) <> "" Then
        rs.AddNew
        rs("sch_idy") = schid
        rs("sch_nme") = Trim(Sch_Nme)
        rs("sch_typ") = Sch_Cmb
        rs("f_sch_typ") = Sub_Cmb
        rs("SCH_STS") = Sch_Act
        rs("CMP_DST") = List2
        rs("SCH_BDT") = StDate
        rs("sch_edt") = EnDate
        rs!MER_NUM = Trim(MERNO)
        rs!MER_QTY = Val(MERQTY)
        rs!MER_AMT = Val(MERAMT)
        rs("B_PRD_IDY") = Trim(MSFree.TextMatrix(A, 4))
        rs("BAT_IDY") = Trim(MSFree.TextMatrix(A, 3))
        If Val(MSFree.TextMatrix(A, 2)) > 0 Then
            rs("B_PRD_QTY") = (Val(MSFree.TextMatrix(A, 2)))
        Else
            rs("B_PRD_QTY") = 0
        End If
        If CmbFree.ListIndex = 0 Then
            rs("F_con_app") = "AND"
        ElseIf CmbFree.ListIndex = 1 Then
            rs("F_con_app") = "ORO"
        End If
        rs("PRO_RAT") = "N"
        rs("txn_typ") = "FRE"
        rs("USR_IDY") = USERNAME
        rs("SYS_DAT") = Date
        rs("SYS_TIM") = Format(Time, "hh:mm:ss")
        rs.Update
    End If
Next
MsgBox "New Scheme ID is " & schid, vbExclamation, "DAS Version 4.0"
con.CommitTrans
cancel
Exit Sub
lab:
MsgBox err.Description
con.RollbackTrans
End Sub

Private Sub TxtGrsAmt_GotFocus()
SLT
End Sub

Private Sub TxtGrsAmt_keypress(K As Integer)
If K = 13 Then
    TxtGrsAmt = amt(TxtGrsAmt)
    If Sub_Cmb.ListIndex > 1 Then
        TxtValue.SetFocus
    ElseIf Sub_Cmb.ListIndex = 1 Then
        TabFree.SetFocus
        MSFree.SetFocus
        MSFree.Col = 0
    End If
End If
End Sub

Private Sub TxtValue_GotFocus()
SLT
End Sub

Private Sub TxtValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    
        If Sub_Cmb.ListIndex = 2 Then
            TxtValue = convert(TxtValue)
        ElseIf Sub_Cmb.ListIndex = 3 Then
            TxtValue = discount(TxtValue)
        End If
    If Val(TxtValue) > 0 Then
        save
    End If
End If
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    KeyAscii = CHKNUM(KeyAscii)
End If
End Sub

Public Sub saveskucashordiscount()
On Error GoTo lab
If Val(MSbase.TextMatrix(1, 2)) = 0 Then
    MsgBox "Enter Valid Quantity", vbExclamation
    SSTab1.Tab = 1
    TabBase.Tab = 0
    MSbase.SetFocus
    MSbase.Col = 2
    MSbase.Row = 1
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch", con, adOpenKeyset, adLockPessimistic
con.BeginTrans
schid = "SCH" & T7increment("idy_sch")
SID Trim(schid)
rs.AddNew
rs("sch_idy") = schid
rs("sch_nme") = Trim(Sch_Nme)
rs("sch_typ") = Sch_Cmb
rs("f_sch_typ") = Sub_Cmb
rs("SCH_STS") = Sch_Act
rs("CMP_DST") = List2
rs("SCH_BDT") = StDate
rs("sch_edt") = EnDate
rs!MER_NUM = Trim(MERNO)
rs!MER_QTY = Val(MERQTY)
rs!MER_AMT = Val(MERAMT)
rs("B_PRD_IDY") = Trim(MSbase.TextMatrix(1, 4))
rs("B_PRD_QTY") = (Val(MSbase.TextMatrix(1, 2) & ""))
rs("BAT_IDY") = Trim(MSbase.TextMatrix(1, 3))
rs("txn_typ") = "BAS"
rs("USR_IDY") = USERNAME
rs("SYS_DAT") = Date
rs("SYS_TIM") = Format(Time, "hh:mm:ss")
rs.Update

rs.AddNew
rs("sch_idy") = schid
rs("sch_nme") = Trim(Sch_Nme)
rs("sch_typ") = Sch_Cmb
rs("f_sch_typ") = Sub_Cmb
rs("SCH_STS") = Sch_Act
rs("CMP_DST") = List2
rs("SCH_BDT") = StDate
rs("sch_edt") = EnDate
rs!MER_NUM = Trim(MERNO)
rs!MER_QTY = Val(MERQTY)
rs!MER_AMT = Val(MERAMT)
If Sub_Cmb.ListIndex = 2 Then rs("B_DIS_AMT") = convert(amt(Val(TxtValue) & ""))
If Sub_Cmb.ListIndex = 3 Then rs("B_DIS_PCG") = discount(Val(TxtValue))
rs("f_prd_idy") = ""
rs("BAT_IDY") = ""
If Sub_Cmb.ListIndex = 2 Then rs("txn_typ") = "CAS"
If Sub_Cmb.ListIndex = 3 Then rs("txn_typ") = "PCG"
If Pro_Chk.Value = 1 Then
    rs("pro_rat") = "Y"
Else
    rs("pro_rat") = "N"
End If
rs("USR_IDY") = USERNAME
rs("SYS_DAT") = Date
rs("SYS_TIM") = Format(Time, "hh:mm:ss")
rs.Update
MsgBox "New Scheme ID is " & schid, vbExclamation, "DAS Version 4.0"
con.CommitTrans
cancel
Exit Sub
lab:
con.RollbackTrans
MsgBox err.Description
cancel
End Sub

Public Sub savemultiplecashordiscount()
On Error GoTo lab
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch", con, adOpenKeyset, adLockPessimistic
con.BeginTrans
schid = "SCH" & T7increment("idy_sch")
For A = 1 To MSbase.Rows - 1
    If MSbase.TextMatrix(A, 0) <> "" Then
        rs.AddNew
        rs("sch_idy") = schid
        rs("sch_nme") = Trim(Sch_Nme)
        rs("sch_typ") = Sch_Cmb
        rs("f_sch_typ") = Sub_Cmb
        rs("SCH_STS") = Sch_Act
        rs("CMP_DST") = List2
        rs("SCH_BDT") = StDate
        rs("sch_edt") = EnDate
        rs!MER_NUM = Trim(MERNO)
        rs!MER_QTY = Val(MERQTY)
        rs!MER_AMT = Val(MERAMT)
        rs("B_PRD_IDY") = Trim(MSbase.TextMatrix(A, 4))
        rs("BAT_IDY") = Trim(MSbase.TextMatrix(A, 3))
        rs("txn_typ") = "BAS"
        rs("USR_IDY") = USERNAME
        If CmbBase.ListIndex = 0 Then
            rs("b_con_app") = "AND"
            rs("B_PRD_QTY") = Val(MSbase.TextMatrix(A, 2))
        ElseIf CmbBase.ListIndex = 1 Then
            rs("b_con_app") = "ORO"
            rs("B_PRD_QTY") = Val(MSbase.TextMatrix(1, 2))
        ElseIf CmbBase.ListIndex = 2 Then
            rs("b_con_app") = "ORM"
            rs("B_PRD_QTY") = Val(MSbase.TextMatrix(1, 2))
        End If
        rs("SYS_DAT") = Date
        rs("SYS_TIM") = Format(Time, "hh:mm:ss")
        rs.Update
    End If
Next
rs.AddNew
rs("sch_idy") = schid
rs("sch_nme") = Trim(Sch_Nme)
rs("sch_typ") = Sch_Cmb
rs("f_sch_typ") = Sub_Cmb
rs("SCH_STS") = Sch_Act
rs("CMP_DST") = List2
rs("SCH_BDT") = StDate
rs("sch_edt") = EnDate
rs!MER_NUM = Trim(MERNO)
rs!MER_QTY = Val(MERQTY)
rs!MER_AMT = Val(MERAMT)
If Sub_Cmb.ListIndex = 2 Then rs("B_DIS_AMT") = convert(amt(Val(TxtValue) & ""))
If Sub_Cmb.ListIndex = 3 Then rs("B_DIS_PCG") = discount(Val(TxtValue))
rs("f_prd_idy") = ""
rs("BAT_IDY") = ""
If Sub_Cmb.ListIndex = 2 Then rs("txn_typ") = "CAS"
If Sub_Cmb.ListIndex = 3 Then rs("txn_typ") = "PCG"
If Pro_Chk.Value = 1 Then
    rs("pro_rat") = "Y"
Else
    rs("pro_rat") = "N"
End If
rs("USR_IDY") = USERNAME
rs("SYS_DAT") = Date
rs("SYS_TIM") = Format(Time, "hh:mm:ss")
rs.Update
MsgBox "New Scheme ID is " & schid, vbExclamation, "DAS Version 4.0"
con.CommitTrans
Exit Sub
lab:
con.RollbackTrans
MsgBox err.Description
cancel
End Sub
Public Sub savebillcashordiscount()
On Error GoTo lab
If Val(TxtGrsAmt) = 0 Then
    MsgBox "Invalid Gross Amount", vbExclamation
    TxtGrsAmt.SetFocus
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch", con, adOpenKeyset, adLockPessimistic
con.BeginTrans
schid = "SCH" & T7increment("idy_sch")
SID Trim(schid)
If MSFree.Rows - 1 = 0 Then MSFree.Rows = 2
For A = 1 To MSFree.Rows - 1
    If MSFree.TextMatrix(A, 0) <> "" Or Sub_Cmb.ListIndex > 1 Then
        rs.AddNew
        rs("sch_idy") = schid
        rs("sch_nme") = Trim(Sch_Nme)
        rs("sch_typ") = Sch_Cmb
        rs("f_sch_typ") = Sub_Cmb
        rs("SCH_STS") = Sch_Act
        rs("CMP_DST") = List2
        rs("SCH_BDT") = StDate
        rs("sch_edt") = EnDate
        rs!MER_NUM = Trim(MERNO)
        rs!MER_QTY = Val(MERQTY)
        rs!MER_AMT = Val(MERAMT)
        rs("B_BIL_AMT") = convert(amt(Val(TxtGrsAmt)))
        rs("BAT_IDY") = ""
        If Sub_Cmb.ListIndex = 1 Then
            rs("BAT_IDY") = Trim(MSFree.TextMatrix(A, 3))
            rs("B_PRD_IDY") = Trim(MSFree.TextMatrix(A, 4)) & ""
            'rs("B_PRD_QTY") = (Val(Trim(MSbase.TextMatrix(A, 2) & "")))
            rs("B_PRD_QTY") = (MSFree.TextMatrix(A, 2) & "") ' CHANGED BY PKB
            rs("txn_typ") = "FRE"
        End If
        If Sub_Cmb.ListIndex = 2 Then rs("B_DIS_AMT") = convert(amt(Val(TxtValue) & ""))
        If Sub_Cmb.ListIndex = 3 Then rs("B_DIS_PCG") = discount(TxtValue & "")
        If Sub_Cmb.ListIndex = 2 Then rs("txn_typ") = "CAS"
        If Sub_Cmb.ListIndex = 3 Then rs("txn_typ") = "PCG"
        If Pro_Chk.Value = 1 Then
            rs("pro_rat") = "Y"
        Else
            rs("pro_rat") = "N"
        End If
        rs("USR_IDY") = USERNAME
        rs("SYS_DAT") = Date
        rs("SYS_TIM") = Format(Time, "hh:mm:ss")
        
    End If
    rs.Update
Next
con.CommitTrans
MsgBox "New Scheme ID is " & schid, vbExclamation, "DAS Version 4.0"
cancel
Exit Sub
lab:
con.RollbackTrans
MsgBox err.Description
cancel
End Sub

Public Sub V_OTH_LST(TYP As String)
If CHECKTB("prd.dbf,pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT " & Trim(TYP) & " FROM PRD,PGR WHERE PGR.GRP_IDY LIKE PRD.GRP_IDY AND " & Trim(TYP) & " LIKE '" & Trim(TxtPack) & "%' GROUP BY " & TYP, con, adOpenKeyset
lv1.ColumnHeaders.CLEAR
lv1.ColumnHeaders.Add , , Sch_Cmb, 2500, 0
lv1.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        If Trim(REC_SET(0)) <> "" Then lv1.ListItems.Add , , Trim(REC_SET(0))
        REC_SET.MoveNext
    Next
    lblview = "Select " & Sch_Cmb
    SSTab1.Tab = 2
    lv1.Visible = True
    lv1.SetFocus
Else
    MsgBox Sch_Cmb & " Not Found, Define " & Sch_Cmb & " !", vbExclamation, "DAS Version 4.0"
    TxtPack.Visible = False
    SSTab1.Tab = 1
    MSPack.SetFocus
    Exit Sub
End If
End Sub

Public Sub SAVEGRPTYP()
On Error GoTo lab
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch", con, adOpenKeyset, adLockPessimistic
con.BeginTrans
schid = "SCH" & T7increment("idy_sch")
SID Trim(schid)
For A = 1 To MSPack.Rows - 1
    If MSPack.TextMatrix(A, 0) <> "" Then
        rs.AddNew
        rs("sch_idy") = schid
        rs("sch_nme") = Trim(Sch_Nme)
        rs("sch_typ") = Sch_Cmb
        rs("f_sch_typ") = Sub_Cmb
        rs("SCH_STS") = Sch_Act
        rs("CMP_DST") = List2
        rs("SCH_BDT") = StDate
        rs("sch_edt") = EnDate
        rs!MER_NUM = Trim(MERNO)
        rs!MER_QTY = Val(MERQTY)
        rs!MER_AMT = Val(MERAMT)
        If Sub_Cmb.ListIndex = 2 Then rs("B_DIS_AMT") = convert(amt(Val(TxtValue) & ""))
        If Sub_Cmb.ListIndex = 3 Then rs("B_DIS_PCG") = discount(TxtValue & "")
        If Sub_Cmb.ListIndex = 2 Then rs("txn_typ") = "CAS"
        If Sub_Cmb.ListIndex = 3 Then rs("txn_typ") = "PCG"
        rs!BAT_IDY = ""
        rs!grp_typ = Trim(MSPack.TextMatrix(A, 0))
        If CmbPCon.ListIndex = 0 Then
            rs!B_CON_APP = "AND"
            rs!B_PRD_QTY = Val(MSPack.TextMatrix(A, 1))
        ElseIf CmbPCon.ListIndex = 1 Then
            rs!B_CON_APP = "ORO"
            rs!B_PRD_QTY = Val(MSPack.TextMatrix(A, 1))
        ElseIf CmbPCon.ListIndex = 2 Then
            rs!B_CON_APP = "ORM"
            rs!B_PRD_QTY = Val(MSPack.TextMatrix(A, 1))
        End If
        rs!TXN_TYP = "BAS"
        rs.Update
    End If
Next
If Sub_Cmb.ListIndex = 1 Then
    For A = 1 To MSFree.Rows - 1
        If MSFree.TextMatrix(A, 0) <> "" Then
            rs.AddNew
            rs("sch_idy") = schid
            rs("sch_nme") = Trim(Sch_Nme)
            rs("sch_typ") = Sch_Cmb
            rs("f_sch_typ") = Sub_Cmb
            rs("SCH_STS") = Sch_Act
            rs("CMP_DST") = List2
            rs("SCH_BDT") = StDate
            rs("sch_edt") = EnDate
            rs!MER_NUM = Trim(MERNO)
            rs!MER_QTY = Val(MERQTY)
            rs!MER_AMT = Val(MERAMT)
            rs("B_PRD_IDY") = Trim(MSFree.TextMatrix(A, 4))
            rs("BAT_IDY") = Trim(MSFree.TextMatrix(A, 3))
            If CmbFree.ListIndex = 0 Then
                rs("F_con_app") = "AND"
                rs("B_PRD_QTY") = Val(MSFree.TextMatrix(A, 2) & "")
            ElseIf CmbFree.ListIndex = 1 Then
                rs("F_con_app") = "ORO"
                rs("B_PRD_QTY") = Val(MSFree.TextMatrix(A, 2) & "")
            End If
            rs("PRO_RAT") = "N"
            rs("txn_typ") = "FRE"
            rs("USR_IDY") = USERNAME
            rs("SYS_DAT") = Date
            rs("SYS_TIM") = Format(Time, "hh:mm:ss")
            rs.Update
        End If
    Next
End If
con.CommitTrans
cancel
Exit Sub
lab:
con.RollbackTrans
MsgBox err.Description
cancel
End Sub

Public Sub dispschemegrt3()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch where sch_idy like '" & Trim(Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs.filter = "txn_typ='BAS'"
    If rs!B_CON_APP = "AND" Then
        CmbPCon.ListIndex = 0
    ElseIf rs!B_CON_APP = "ORO" Then
        CmbPCon.ListIndex = 1
    Else
        CmbPCon.ListIndex = 2
    End If
    MSPack.Rows = rs.RecordCount + 1
    If Sub_Cmb.ListIndex = 2 Then TxtValue = convert(rs!B_DIS_AMT)
    If Sub_Cmb.ListIndex = 3 Then TxtValue = convert(rs!B_DIS_PCG)
    For A = 1 To rs.RecordCount
        MSPack.TextMatrix(A, 0) = Trim(rs!grp_typ)
        MSPack.TextMatrix(A, 1) = Trim(rs!B_PRD_QTY)
        rs.MoveNext
    Next
    If Sub_Cmb.ListIndex = 1 Then
        rs.filter = "txn_typ='FRE'"
        If rs.RecordCount > 0 Then
            If rs!F_CON_APP = "AND" Then
                CmbFree.ListIndex = 0
            ElseIf rs!F_CON_APP = "ORO" Then
                CmbFree.ListIndex = 1
            End If
        End If
        MSFree.Rows = rs.RecordCount + 1
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                MSFree.TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")
                MSFree.TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"))
                MSFree.TextMatrix(A, 2) = Val(rs!B_PRD_QTY & "")
                rs.MoveNext
            Next
        End If
    End If
End If
End Sub

Public Sub F6()
SC = Sch_Cmb.ListIndex
SU = Sub_Cmb.ListIndex
If Len(Trim(Sch_Idy)) = 0 Then
    If (SC = 1 And SU = 1) Or (SC = 3 And SU = 1) Then
        If M_FRM_SCH.ActiveControl.name = "MSbase" Then
            MSFree.SetFocus
            MSFree.Row = MSFree.Rows - 1
            MSFree.Col = 0
        ElseIf M_FRM_SCH.ActiveControl.name = "MSFree" Then
            MSbase.SetFocus
            MSbase.Row = MSbase.Rows - 1
        End If
    ElseIf (SC = 1 Or SC = 3) And (SU = 2 Or SU = 3) Then
        If M_FRM_SCH.ActiveControl.name = "MSbase" Then
            TxtValue.SetFocus
            SLT
        ElseIf M_FRM_SCH.ActiveControl.name = "TxtValue" Then
            MSbase.SetFocus
            MSbase.Row = MSbase.Rows - 1
        End If
    ElseIf SC = 2 Then
        If M_FRM_SCH.ActiveControl.name = "TxtGrsAmt" Then
            If SU = 1 Then
                MSFree.SetFocus
                MSFree.Row = MSFree.Rows - 1
                MSFree.Col = 0
            Else
                TxtValue.SetFocus
                SLT
            End If
        End If
    ElseIf SC > 3 Then
        If SU = 1 Then
            If M_FRM_SCH.ActiveControl.name = "MSPack" Then
                MSFree.SetFocus
                MSFree.Row = MSFree.Rows - 1
                MSFree.Col = 0
            ElseIf M_FRM_SCH.ActiveControl.name = "MSFree" Then
                MSPack.SetFocus
                MSPack.Row = MSPack.Rows - 1
            End If
        ElseIf SU > 1 Then
            If M_FRM_SCH.ActiveControl.name = "MSPack" Then
                TxtValue.SetFocus
                SLT
            ElseIf M_FRM_SCH.ActiveControl.name = "TxtValue" Then
                MSPack.SetFocus
                MSPack.Row = MSPack.Rows - 1
            End If
        End If
    End If
End If
End Sub

Public Sub DISPSCHEMESKU()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch where sch_idy like '" & Trim(Sch_Idy) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs.filter = "txn_typ='BAS'"
    If rs!B_CON_APP = "AND" Then
        CmbBase.ListIndex = 0
    ElseIf rs!B_CON_APP = "ORO" Then
        CmbBase.ListIndex = 1
    ElseIf rs!B_CON_APP = "ORM" Then
        CmbBase.ListIndex = 2
    Else
        MSbase.Top = 360
        MSbase.Height = 4215
        MSFree.Top = 360
        MSFree.Height = 4215
    End If
    With MSbase
        .Rows = rs.RecordCount + 1
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                .TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")
                .TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"))
                If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
                    .TextMatrix(A, 2) = QTY(STM(Val(rs!B_PRD_QTY & ""), GetValue("sub_unt", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")))
                Else
                    .TextMatrix(A, 2) = Val(rs!B_PRD_QTY & "")
                End If
                rs.MoveNext
            Next
        End If
    End With
    
    If Sub_Cmb.ListIndex = 2 Then
        rs.filter = "txn_typ='CAS'"
        If rs.RecordCount > 0 Then TxtValue = convert(rs!B_DIS_AMT)
    ElseIf Sub_Cmb.ListIndex = 3 Then
        rs.filter = "txn_typ='PCG'"
        If rs.RecordCount > 0 Then TxtValue = convert(rs!B_DIS_PCG & "")
    End If
    rs.filter = "txn_typ='FRE'"
        MSFree.Rows = rs.RecordCount + 1
        If rs.RecordCount > 0 Then
            If rs!F_CON_APP = "AND" Then
            CmbFree.ListIndex = 0
        ElseIf rs!F_CON_APP = "ORO" Then
            CmbFree.ListIndex = 1
        End If

        For A = 1 To rs.RecordCount
            MSFree.TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")
            MSFree.TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'"))
            If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
                MSFree.TextMatrix(A, 2) = QTY(STM(Val(rs!B_PRD_QTY & ""), GetValue("sub_unt", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(rs!BAT_IDY) & "'")))
            Else
                MSFree.TextMatrix(A, 2) = Val(rs!B_PRD_QTY & "")
            End If
            rs.MoveNext
        Next
        End If
End If
End Sub

Public Sub SAVEMULTIPLECASHORPCG()
On Error GoTo lab
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from sch", con, adOpenKeyset, adLockPessimistic
schid = "SCH" & T7increment("idy_sch")
SID Trim(schid)
For A = 1 To MSbase.Rows - 1
    If Trim(MSbase.TextMatrix(A, 0)) <> "" Then
        rs.AddNew
        rs("sch_idy") = schid
        rs("sch_nme") = Trim(Sch_Nme)
        rs("sch_typ") = Sch_Cmb
        rs("f_sch_typ") = Sub_Cmb
        rs("SCH_STS") = Sch_Act
        rs("CMP_DST") = List2
        rs("SCH_BDT") = StDate
        rs("sch_edt") = EnDate
        rs("B_PRD_IDY") = Trim(MSbase.TextMatrix(A, 4))
        rs("B_PRD_QTY") = Val(MSbase.TextMatrix(A, 2) & "")
        rs("BAT_IDY") = Trim(MSbase.TextMatrix(A, 3))
        rs("txn_typ") = "BAS"
        rs("USR_IDY") = USERNAME
        rs("SYS_DAT") = Date
        rs("SYS_TIM") = Format(Time, "hh:mm:ss")
        If CmbBase.ListIndex = 0 Then
            rs("b_con_app") = "AND"
        ElseIf CmbBase.ListIndex = 1 Then
            rs("B_CON_APP") = "ORO"
        Else
            rs("B_CON_APP") = "ORM"
        End If
        rs.Update
        savecheck = True
    End If
Next
rs.AddNew
rs("sch_idy") = schid
rs("sch_nme") = Trim(Sch_Nme)
rs("sch_typ") = Sch_Cmb
rs("f_sch_typ") = Sub_Cmb
rs("SCH_STS") = Sch_Act
rs("CMP_DST") = List2
rs("SCH_BDT") = StDate
rs("sch_edt") = EnDate
rs!BAT_IDY = ""
If Sub_Cmb.ListIndex = 2 Then rs("B_DIS_AMT") = Val(TxtValue)
If Sub_Cmb.ListIndex = 3 Then rs("B_DIS_PCG") = discount(Val(TxtValue))
If Sub_Cmb.ListIndex = 2 Then rs("txn_typ") = "CAS"
If Sub_Cmb.ListIndex = 3 Then rs("txn_typ") = "PCG"
If Pro_Chk.Value = 1 Then
rs("pro_rat") = "Y"
Else
rs("pro_rat") = "N"
End If
rs("USR_IDY") = USERNAME
rs("SYS_DAT") = Date
rs("SYS_TIM") = Format(Time, "hh:mm:ss")
rs.Update
MsgBox "New Scheme ID " & schid, vbExclamation
cancel
Exit Sub
lab:
MsgBox err.Description
cancel
End Sub

Public Sub SID(Sch_Idy As String)
Dim rs As New ADODB.Recordset
rs.Open "select * from sid", con, adOpenKeyset, adLockPessimistic
For A = 1 To list1.ListItems.count
    If list1.ListItems.Item(A).Checked = True Then
        rs.AddNew
        rs!Sch_Idy = Sch_Idy
        If Combo1.ListIndex = 1 Then
            rs!CST_IDY = list1.ListItems(A).ListSubItems.Item(1)
            rs!SMN_IDY = ""
            rs!rou_idy = ""
        ElseIf Combo1.ListIndex = 2 Then
            rs!SMN_IDY = list1.ListItems(A).ListSubItems.Item(1)
            rs!CST_IDY = ""
            rs!rou_idy = ""
        ElseIf Combo1.ListIndex = 3 Then
            rs!rou_idy = list1.ListItems(A).ListSubItems.Item(1)
            rs!SMN_IDY = ""
            rs!CST_IDY = ""
        End If
        rs.Update
    End If
Next
End Sub
