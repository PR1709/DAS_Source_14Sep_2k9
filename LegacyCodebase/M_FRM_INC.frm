VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_INC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incentives"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "M_FRM_INC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      TabPicture(0)   =   "M_FRM_INC.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "smnnme"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "EnDate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "StDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Sub_Cmb"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Sch_Cmb"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Sch_Nme"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Sch_Act"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Sch_Idy"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Pro_Chk"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "smnid"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Scheme &Definition"
      TabPicture(1)   =   "M_FRM_INC.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabBase"
      Tab(1).Control(1)=   "TabFree"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "M_FRM_INC.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lv1"
      Tab(2).Control(1)=   "LV"
      Tab(2).Control(2)=   "lblview"
      Tab(2).ControlCount=   3
      Begin VB.TextBox smnid 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin MSComctlLib.ListView lv1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   40
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
         Left            =   4560
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Sch_Idy 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2175
         MaxLength       =   10
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox Sch_Act 
         Height          =   255
         ItemData        =   "M_FRM_INC.frx":0496
         Left            =   2160
         List            =   "M_FRM_INC.frx":04A0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3900
         Width           =   495
      End
      Begin VB.TextBox Sch_Nme 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2175
         MaxLength       =   40
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1575
         Width           =   3345
      End
      Begin VB.ComboBox Sch_Cmb 
         Height          =   315
         ItemData        =   "M_FRM_INC.frx":04AA
         Left            =   2160
         List            =   "M_FRM_INC.frx":04C9
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2085
         Width           =   2175
      End
      Begin VB.ComboBox Sub_Cmb 
         Height          =   315
         ItemData        =   "M_FRM_INC.frx":053D
         Left            =   2160
         List            =   "M_FRM_INC.frx":054D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2445
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker StDate 
         Height          =   330
         Left            =   2175
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2910
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
         CurrentDate     =   37099
      End
      Begin MSComCtl2.DTPicker EnDate 
         Height          =   330
         Left            =   2175
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3390
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
         CurrentDate     =   37099
      End
      Begin TabDlg.SSTab TabBase 
         Height          =   4650
         Left            =   -74880
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   8202
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Base SKU Details"
         TabPicture(0)   =   "M_FRM_INC.frx":0595
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSbase"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TxtBase"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Base Value Details"
         TabPicture(1)   =   "M_FRM_INC.frx":05B1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "M_FRM_INC.frx":05CD
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSPack"
         Tab(2).Control(1)=   "TxtPack"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame5"
         Tab(2).ControlCount=   3
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   -74880
            TabIndex        =   37
            Top             =   360
            Width           =   4095
            Begin VB.ComboBox CmbPCon 
               Height          =   315
               ItemData        =   "M_FRM_INC.frx":05E9
               Left            =   1920
               List            =   "M_FRM_INC.frx":05F6
               Style           =   2  'Dropdown List
               TabIndex        =   38
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
               TabIndex        =   39
               Top             =   240
               Width           =   1440
            End
         End
         Begin VB.TextBox TxtPack 
            Height          =   495
            Left            =   -74040
            MaxLength       =   10
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Frame Frame4 
            Caption         =   "Scheme Value"
            Height          =   1575
            Left            =   -74880
            TabIndex        =   32
            Top             =   1440
            Width           =   3975
            Begin VB.TextBox TxtGrsAmt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2280
               MaxLength       =   10
               TabIndex        =   33
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
               TabIndex        =   34
               Top             =   675
               Width           =   1560
            End
         End
         Begin VB.TextBox TxtBase 
            Height          =   495
            Left            =   960
            MaxLength       =   10
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid MSbase 
            Height          =   3495
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1080
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   14737632
            ForeColorSel    =   16711680
            FocusRect       =   0
            Appearance      =   0
         End
         Begin VB.Frame Frame1 
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   4095
            Begin VB.ComboBox CmbBase 
               Height          =   315
               ItemData        =   "M_FRM_INC.frx":060A
               Left            =   1920
               List            =   "M_FRM_INC.frx":0617
               Style           =   2  'Dropdown List
               TabIndex        =   19
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
               TabIndex        =   20
               Top             =   240
               Width           =   1440
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSPack 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   36
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
         TabIndex        =   15
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
         TabPicture(0)   =   "M_FRM_INC.frx":062B
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSFree"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TxtFree"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Free Value Details"
         TabPicture(1)   =   "M_FRM_INC.frx":0647
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Height          =   1575
            Left            =   -74760
            TabIndex        =   29
            Top             =   1440
            Width           =   3855
            Begin VB.TextBox TxtValue 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1920
               MaxLength       =   10
               TabIndex        =   31
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
               TabIndex        =   30
               Top             =   720
               Width           =   615
            End
         End
         Begin VB.TextBox TxtFree 
            Height          =   495
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid MSFree 
            Height          =   3495
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1080
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   14737632
            ForeColorSel    =   16711680
            FocusRect       =   0
            Appearance      =   0
         End
         Begin VB.Frame Frame2 
            Height          =   615
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   4095
            Begin VB.ComboBox CmbFree 
               Height          =   315
               ItemData        =   "M_FRM_INC.frx":0663
               Left            =   1920
               List            =   "M_FRM_INC.frx":066D
               Style           =   2  'Dropdown List
               TabIndex        =   22
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
               TabIndex        =   23
               Top             =   240
               Width           =   1440
            End
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   860
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
      Begin VB.Label smnnme 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3360
         TabIndex        =   45
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1155
         TabIndex        =   44
         Top             =   1155
         Width           =   690
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
         TabIndex        =   17
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
         Left            =   1695
         TabIndex        =   13
         Top             =   675
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   1650
         Width           =   420
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Active"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1410
         TabIndex        =   11
         Top             =   3930
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1185
         TabIndex        =   10
         Top             =   3465
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1140
         TabIndex        =   9
         Top             =   2985
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
            Picture         =   "M_FRM_INC.frx":067A
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":0F86
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":12DA
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":1912
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":1C66
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":20C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":23BA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":27C6
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":28D2
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":2C26
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC.frx":303A
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
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Salesmen ID "
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
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         ScaleHeight     =   375
         ScaleWidth      =   4215
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Press F6 to Shift Between TABS"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1800
            TabIndex        =   42
            Top             =   120
            Width           =   2295
         End
      End
   End
End
Attribute VB_Name = "M_FRM_INC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

Private Sub Form_Load()
SSTab1.TabEnabled(2) = False
With MSbase
.ColWidth(3) = 0
.ColWidth(4) = 0
.ColWidth(0) = 2000
.ColWidth(1) = 950
.ColWidth(2) = 1100
.TextMatrix(0, 0) = "Product(SKU) Name"
.TextMatrix(0, 1) = "MRP"
.TextMatrix(0, 2) = "Qty-Units"
.ColAlignment(0) = 1
.ColAlignment(1) = 7
.ColAlignment(2) = 7
.HighLight = flexHighlightNever
End With
With MSFree
.ColWidth(3) = 0
.ColWidth(4) = 0
.TextMatrix(0, 0) = "Product(SKU) Name"
.TextMatrix(0, 1) = "MRP"
.TextMatrix(0, 2) = "Qty-Units"
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
SSTab1.Tab = 0
End Sub
Private Sub Form_Activate()
SSTab1.Tab = 0
Sch_Idy.SetFocus
Sch_Cmb.ListIndex = 0
Sub_Cmb.ListIndex = 0
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf M_FRM_INC.ActiveControl Is TextBox Then If K <> 117 Then M_FRM_INC.ActiveControl.Text = Trim(FUNKEY(K))
    If K = 117 Then
        F6
    End If
    If K = 78 And S = 2 Then
        newrec
    ElseIf K = 83 And S = 2 Then
        save
    ElseIf K = 81 And S = 2 Then
        cancel
    ElseIf K = 88 And S = 2 Then
        Unload Me
    ElseIf K = 69 And S = 2 Then
        If M_FRM_INC.ActiveControl.name = "Sch_Idy" Then
            Schemelist
            lv.SetFocus
        End If
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And S = 2 Then
        M_FRM_DSN.Tag = "Schemes List"
        M_FRM_DSN.Show 1
    Else
        If K = 27 And Len(Trim(Sch_Idy)) <> 0 And SSTab1.Tab = 1 Then
            save
        End If
    End If
End Sub

Private Sub LV_DblClick()
lv_KeyPress 13
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lblview = "Select Scheme" Then
        DisplayScheme lv.SelectedItem.ListSubItems.Item(4)
    ElseIf lblview = "Select Product" And lv.Tag = "Base" Then
        P_DIS_PRD lv.SelectedItem.ListSubItems.Item(5), "Base"
    ElseIf lblview = "Select Product" And lv.Tag = "Free" Then
        P_DIS_PRD lv.SelectedItem.ListSubItems.Item(5), "Free"
    End If
End If
End Sub

Private Sub lv1_KeyPress(K As Integer)
For A = 1 To MSPack.Rows - 1
    If Trim(MSPack.TextMatrix(A, 0)) = Trim(lv1.SelectedItem) Then
        MsgBox Sch_Cmb & "Already Exists", vbExclamation
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
        lv.Tag = "Base"
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
        lv.Tag = "Free"
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
    sch = Sch_Cmb.ListIndex
    SUBCMB = Sub_Cmb.ListIndex
    SSTab1.Tab = 1
    If sch = 1 Then
        TabBase.Tab = 0
        MSbase.Top = 360
        MSbase.Height = 4215
        If MSbase.Enabled = True Then MSbase.SetFocus
    ElseIf sch = 2 Then
        TabBase.Tab = 1
        TxtGrsAmt.SetFocus
    ElseIf sch = 3 Then
        TabBase.Tab = 0
        MSbase.Top = 1080
        MSbase.Height = 3495
        If MSbase.Enabled = True Then MSbase.SetFocus
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

Private Sub SCH_CMB_KeyDown(K As Integer, S As Integer)
If K = 13 Then
    If Sch_Cmb.ListIndex = 0 Then
        MsgBox "Select Valid Base Entry", vbExclamation
        Sch_Cmb.SetFocus
    Else
        Sub_Cmb.SetFocus
    End If
ElseIf K = 38 And S = 1 Then
    Sch_Nme.SetFocus
End If
End Sub

Private Sub endate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If DateDiff("D", StDate, EnDate) < 0 Then
        SSTab1.Tab = 0
        MsgBox "End Date Should Be Greater Than Start Date", vbExclamation, "Effmcg"
        EnDate.SetFocus
        Exit Sub
    Else
        Sch_Act.SetFocus
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

Private Sub SUB_CMB_KeyDown(K As Integer, S As Integer)
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
If K = 38 And S = 1 Then Sch_Cmb.SetFocus
End Sub


Private Sub SCH_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Scheme Identification"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If M_FRM_INC.ActiveControl.name = "Sch_Idy" Then
        Unload Me
    ElseIf SSTab1.Tab = 2 And lblview = "Select Scheme" Then
        lv.Visible = False
        SSTab1.Tab = 0
        Sch_Idy.SetFocus
    ElseIf lblview = "Select Product" And SSTab1.Tab = 2 Then
        If lv.Tag = "Base" Then
            SSTab1.Tab = 1
            MSbase.SetFocus
        ElseIf lv.Tag = "Free" Then
            SSTab1.Tab = 1
            MSFree.SetFocus
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

Private Sub SCH_IDY_Keydown(K As Integer, S As Integer)
If K = 13 Or K = 40 Then
    If check("INC", "inc_idy", "INC" & Trim(Tincrement(Sch_Idy))) = True Then
        DisplayScheme "INC" & Trim(Tincrement(Sch_Idy))
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
If CHECKTB("INC.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'        cancel
        Exit Sub
End If
Dim RS As New ADODB.Recordset
Dim ls As ListItem
If RS.State = 1 Then RS.Close
If MOD_REC = True Then
    RS.Open "select * from INC GROUP BY SCH_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    RS.Open "select * from INC where SCH_NME like '" & Trim(Sch_Idy) & "%' GROUP BY SCH_IDY", CON, adOpenKeyset, adLockOptimistic
End If
lv.ColumnHeaders.Clear
lv.ColumnHeaders.Add , , "Scheme Name", 3000, 0
lv.ColumnHeaders.Add , , "Status", 1000, 2
lv.ColumnHeaders.Add , , "Start Date", 1500, 0
lv.ColumnHeaders.Add , , "End Date", 1500, 0
lv.ColumnHeaders.Add , , "ID", 1500, 0
If RS.RecordCount > 0 Then
    lv.ListItems.Clear
    For A = 1 To RS.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(RS("SCH_nme")))
        ls.ListSubItems.Add , , Trim(RS("SCH_STS"))
        ls.ListSubItems.Add , , Format(RS("SCH_BDT"), "DD/MM/YYYY")
        ls.ListSubItems.Add , , Format(RS("SCH_EDT"), "DD/MM/YYYY")
        ls.ListSubItems.Add , , Trim(RS!Sch_Idy)
        RS.MoveNext
    Next
    SSTab1.Tab = 2
    lblview = "Select Scheme"
    lv.Visible = True
    lv.SetFocus
    'If SCH_IDY <> "" Then SCH_LST.SetFocus
ElseIf RS.RecordCount = 0 Then
    MsgBox "No Schemes Found!", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    Sch_Idy.Enabled = True
    Sch_Idy.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub SCH_NME_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Scheme, Discount Entry Description"
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
        MsgBox "Scheme Name Should Not Be Blank, Give Name !", vbExclamation, "Effmcg"
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
    If M_FRM_INC.ActiveControl.name = "Sch_Idy" Then
        Schemelist
        SCH_LST.SetFocus
    End If
ElseIf Button.KEY = "p" Then
    M_FRM_DSN.Tag = "Schemes List"
    M_FRM_DSN.Show 1
'    Openreport CURDIR & "\lists.rp1", "Schemes List", 0, "", 6
End If
End Sub

Public Sub newrec()
cancel
Sch_Idy.Enabled = False
Sch_Nme.SetFocus
End Sub

Public Sub save()
If MsgBox("Confirm Scheme Save?", vbYesNo + vbQuestion) = vbNo Then
    cancel
    Exit Sub
End If
If Len(Trim(Sch_Nme)) = 0 Then
    MsgBox "Scheme Name Should not be blank", vbExclamation
    SSTab1.Tab = 0
    Sch_Nme.SetFocus
    Exit Sub
End If

If Len(Trim(Sch_Idy)) <> 0 Then
    CON.Execute "UPDATE INC SET sch_edt = ctod('" & Trim(EnDate) & "') where sch_idy like '" & Trim(Sch_Idy) & "'", d
    CON.Execute "UPDATE INC SET sch_sts = '" & Sch_Act & "' where sch_idy like '" & Trim(Sch_Idy) & "'", d
    cancel
    Exit Sub
End If
If Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex = 1 Then
    If MSbase.TextMatrix(1, 0) > 0 Then
        MsgBox "Invalid Free Quanity", vbExclamation
        SSTab1.Tab = 1
        TabBase.Tab = 0
        MSbase.Col = 2
        MSbase.SetFocus
    ElseIf MSFree.TextMatrix(1, 0) > 0 Then
        MsgBox "Invalid Free Quanity", vbExclamation
        SSTab1.Tab = 1
        TabFree.Tab = 0
        MSFree.Col = 2
        MSFree.SetFocus
    Else
        SAVESKUFREEQUANTITY
    End If
ElseIf Sch_Cmb.ListIndex = 3 And Sub_Cmb.ListIndex = 1 Then
    SAVEMULTIPLEFREEQUANTITY
ElseIf Sch_Cmb.ListIndex = 1 And Sub_Cmb.ListIndex > 1 Then
    saveskucashordiscount
ElseIf Sch_Cmb.ListIndex = 2 Then
    savebillcashordiscount
ElseIf Sch_Cmb.ListIndex > 3 Then
    SAVEGRPTYP
End If
End Sub

Public Sub cancel()
SETYEAR Me
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
Sch_Act.Selected(0) = True
Sch_Idy.Enabled = True
Sch_Idy.SetFocus
End Sub

Public Sub DisplayScheme(Schemeid As String)
Dim RS As New ADODB.Recordset
RS.Open "select * from INC where sch_idy like '" & Trim(Schemeid) & "'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    Sch_Idy = Trim(Schemeid)
    Sch_Idy.Enabled = False
    Sch_Nme = Trim(RS!Sch_Nme)
    StDate = Trim(RS!sch_bdt)
    EnDate = Trim(RS!sch_edt)
    Sch_Act = Trim(RS!sch_sts)
    For A = 0 To Sch_Cmb.ListCount - 1
        If Trim(Sch_Cmb.list(A)) = Trim(RS!SCH_TYP) Then
            Sch_Cmb.ListIndex = A
        End If
    Next
    For A = 0 To Sub_Cmb.ListCount - 1
        If Trim(Sub_Cmb.list(A)) = Trim(RS!f_sch_typ) Then
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
    MSbase.Enabled = False
    MSFree.Enabled = False
    If Sch_Cmb.ListIndex > 3 Then
        dispschemegrt3
    ElseIf Sch_Cmb.ListIndex = 1 Or Sch_Cmb.ListIndex = 3 Then
        DISPSCHEMESKU
    ElseIf Sch_Cmb.ListIndex = 2 And (Sub_Cmb.ListIndex = 2 Or Sub_Cmb.ListIndex = 3) Then
        TxtGrsAmt = convert(RS!b_BIL_AMT & "")
        If Sub_Cmb = 2 Then TxtValue = convert(RS!B_DIS_AMT & "")
        If Sub_Cmb = 3 Then TxtValue = convert(RS!B_DIS_PCG & "")
    ElseIf Sch_Cmb.ListIndex = 2 And Sub_Cmb.ListIndex = 1 Then
        RS.filter = "txn_typ='FRE'"
        If RS!F_CON_APP = "AND" Then
            CmbFree.ListIndex = 0
        ElseIf RS!F_CON_APP = "ORO" Then
            CmbFree.ListIndex = 1
        End If
        MSFree.Rows = RS.RecordCount + 1
        If RS.RecordCount > 0 Then
            For A = 1 To RS.RecordCount
                MSFree.TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'")
                MSFree.TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'"))
                MSFree.TextMatrix(A, 2) = Val(RS!B_PRD_QTY & "")
                RS.MoveNext
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
lv.ColumnHeaders.Clear
lv.ColumnHeaders.Add , , "Name", 2500, 0
lv.ColumnHeaders.Add , , "Quantity", 1500, 1
lv.ColumnHeaders.Add , , "MRP", 1300, 1
lv.ColumnHeaders.Add , , "Sale Rate", 1300, 1
lv.ColumnHeaders.Add , , "Purchase Rate", 1300, 1
lv.ColumnHeaders.Add , , "Batid", 0, 0
lv.ColumnHeaders.Add , , "prdid", 0, 0
lv.ListItems.Clear
REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_IDY LIKE '" & Sincrement(TxtBase) & "%' AND PRD.STK_TYP IN('Salable') and BRD.prd_qty >0 order by brw_seq", CON, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
Set ls = lv.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , QTY(REC_SET("prd_qty") & "")
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("prd_mrp") & "")))
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("prd_sdr") & "")))
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("bas_rat") & "")))
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET!PRD_IDY & "")
        REC_SET.MoveNext
    Next
    lblview = "Select Product"
    SSTab1.Tab = 2
    lv.Visible = True
    lv.SetFocus
Else
    If SSTab1.Tab = 1 Then
        MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
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
If lv.Tag = "Base" Then
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(TxtBase) & "%'", CON, adOpenKeyset
Else
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(TxtFree) & "%'", CON, adOpenKeyset
End If
lv.ColumnHeaders.Clear
lv.ColumnHeaders.Add , , "Name", 2500, 0
lv.ColumnHeaders.Add , , "Quantity", 1500, 1
lv.ColumnHeaders.Add , , "MRP", 1300, 1
lv.ColumnHeaders.Add , , "Sale Rate", 1300, 1
lv.ColumnHeaders.Add , , "Purchase Rate", 1300, 1
lv.ColumnHeaders.Add , , "Batid", 0, 0
lv.ColumnHeaders.Add , , "prdid", 0, 0
lv.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    
    For A = 1 To REC_SET.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , QTY(REC_SET("prd_qty") & "")
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("prd_mrp") & "")))
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("prd_sdr") & "")))
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("bas_rat") & "")))
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET!PRD_IDY & "")
        REC_SET.MoveNext
    Next
    lblview = "Select Product"
    SSTab1.Tab = 2
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    SSTab1.Tab = 1
    If lv.Tag = "Base" Then
        lv.Tag = ""
        MSbase.SetFocus
        TxtBase.Visible = False
    ElseIf lv.Tag = "Free" Then
        lv.Tag = ""
        MSFree.SetFocus
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
            REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(TxtBase) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
            If REC_SET.RecordCount = 1 Then
                P_DIS_PRD Trim(REC_SET!BAT_IDY), "Base"
            ElseIf REC_SET.RecordCount > 1 Then
                TxtBase = Sincrement(TxtBase)
                lv.Tag = "Base"
                V_PRD_IDY_LST
            Else
                lv.Tag = "Base"
                V_PRD_LST
            End If
        End If
    End If
ElseIf MSbase.Col = 2 Then
    If K <> 13 Then K = CHKNUM(K)
    If K = 13 Then
        If Sch_Cmb.ListIndex = 1 Or Sch_Cmb.ListIndex = 3 Then
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
                        If Val(MSbase.Text) <> 0 Then MSFree.SetFocus
                    Else
                        SSTab1.Tab = 1
                        TxtValue.SetFocus
                        SLT
                    End If
                End If
        Else
            MSbase.SetFocus
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
            MSPack = Val(TxtPack)
        Else
            MSPack.TextMatrix(1, 1) = Val(TxtPack)
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
            REC_SET.Open "SELECT * FROM prd,BRD WHERE PRD.prd_IDY LIKE '" & Sincrement(TxtFree) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
            If REC_SET.RecordCount = 1 Then
                P_DIS_PRD Trim(REC_SET!BAT_IDY), "Free"
            ElseIf REC_SET.RecordCount > 1 Then
                TxtFree = Sincrement(TxtFree)
                lv.Tag = "Free"
                V_PRD_IDY_LST
            Else
                lv.Tag = "Free"
                V_PRD_LST
            End If
        End If
    End If
ElseIf MSFree.Col = 2 Then
    If K <> 13 Then K = CHKNUM(K)
    If K = 13 Then
        If Sch_Cmb.ListIndex = 1 Or Sch_Cmb.ListIndex >= 3 Then
            If CmbFree.ListIndex <> 0 Then
                MSFree.TextMatrix(1, 2) = Val(TxtFree)
            Else
                MSFree = Val(TxtFree)
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
                MSFree = Val(TxtFree)
                If MSFree.TextMatrix(MSFree.Rows - 1, 0) <> "" Then MSFree.AddItem ""
                MSFree.Row = MSFree.Rows - 1
                MSFree.Col = 0
                MSFree.SetFocus
        Else
            MSFree.SetFocus
        End If
    End If
End If
End Sub

Public Sub P_DIS_PRD(BatchID As String, Grid As String)
Dim RS As New ADODB.Recordset
If Grid = "Base" Then
    With MSbase
        For A = 1 To .Rows - 1
            If .TextMatrix(A, 3) = BatchID Then
                SSTab1.Tab = 1
                MsgBox "Product Already Exists", vbExclamation
                TabBase.Tab = 0
                .Row = A
                .SetFocus
                Exit Sub
            End If
        Next
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM prd,BRD WHERE brd.bat_IDY LIKE '" & Trim(BatchID) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
        If RS.RecordCount > 0 Then
            .TextMatrix(.Row, 0) = Trim(RS!PRD_NME)
            .TextMatrix(.Row, 1) = convert(Val(RS!PRD_MRP))
            .TextMatrix(.Row, 3) = Trim(RS!BAT_IDY)
            .TextMatrix(.Row, 4) = Trim(RS!PRD_IDY)
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
                'TabBase.Tab = 2
                .Row = A
                .SetFocus
                Exit Sub
            End If
        Next
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM prd,BRD WHERE brd.bat_IDY LIKE '" & Trim(BatchID) & "' AND BRD.PRD_IDY LIKE PRD.PRD_IDY AND STK_TYP IN('Salable') and prd_qty >0", CON, adOpenKeyset
        If RS.RecordCount > 0 Then
            .TextMatrix(.Row, 0) = Trim(RS!PRD_NME)
            .TextMatrix(.Row, 1) = convert(Val(RS!PRD_MRP))
            .TextMatrix(.Row, 3) = Trim(RS!BAT_IDY)
            .TextMatrix(.Row, 4) = Trim(RS!PRD_IDY)
        End If
        SSTab1.Tab = 1
        TxtFree.Visible = False
        .Col = 2
        .SetFocus
    End With
End If
End Sub

Public Sub SAVESKUFREEQUANTITY()
''ON ERROR GoTo lab
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
    MSFree.Col = 2
    MSFree.Row = 1
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from INC", CON, adOpenKeyset, adLockOptimistic
CON.BeginTrans
schid = "SCH" & T7increment("idy_sch")
RS.AddNew
RS("sch_idy") = schid
RS("sch_nme") = Trim(Sch_Nme)
RS("sch_typ") = Sch_Cmb
RS("f_sch_typ") = Sub_Cmb
RS("SCH_STS") = Sch_Act
RS("SCH_BDT") = StDate
RS("sch_edt") = EnDate
RS("B_PRD_IDY") = Trim(MSbase.TextMatrix(1, 4))
RS("B_PRD_QTY") = Val(MSbase.TextMatrix(1, 2))
RS("BAT_IDY") = Trim(MSbase.TextMatrix(1, 3))
RS("txn_typ") = "BAS"
RS("USR_IDY") = USERNAME
RS("SYS_DAT") = Date
RS("SYS_TIM") = Format(Time, "hh:mm:ss")
RS.Update

RS.AddNew
RS("sch_idy") = schid
RS("sch_nme") = Trim(Sch_Nme)
RS("sch_typ") = Sch_Cmb
RS("f_sch_typ") = Sub_Cmb
RS("SCH_STS") = Sch_Act
RS("SCH_BDT") = StDate
RS("sch_edt") = EnDate
RS("B_PRD_IDY") = Trim(MSFree.TextMatrix(1, 4))
RS("B_PRD_QTY") = Val(MSFree.TextMatrix(1, 2))
RS("f_prd_idy") = Trim(MSbase.TextMatrix(1, 3))
RS("BAT_IDY") = Trim(MSFree.TextMatrix(1, 3))
RS("txn_typ") = "FRE"
If Pro_Chk.Value = 1 Then
    RS("pro_rat") = "Y"
Else
    RS("pro_rat") = "N"
End If
RS("USR_IDY") = USERNAME
RS("SYS_DAT") = Date
RS("SYS_TIM") = Format(Time, "hh:mm:ss")
RS.Update
MsgBox "New Scheme ID is " & schid, vbExclamation, "Effmcg"
CON.CommitTrans
cancel
Exit Sub
lab:
MsgBox err.Description
CON.RollbackTrans
End Sub

Public Sub SAVEMULTIPLEFREEQUANTITY()
''ON ERROR GoTo lab
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from sch", CON, adOpenKeyset, adLockOptimistic
CON.BeginTrans
schid = "SCH" & T7increment("idy_sch")
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
        RS.AddNew
        RS("sch_idy") = schid
        RS("sch_nme") = Trim(Sch_Nme)
        RS("sch_typ") = Sch_Cmb
        RS("f_sch_typ") = Sub_Cmb
        RS("SCH_STS") = Sch_Act
        RS("SCH_BDT") = StDate
        RS("sch_edt") = EnDate
        RS("B_PRD_IDY") = Trim(MSbase.TextMatrix(A, 4))
        RS("BAT_IDY") = Trim(MSbase.TextMatrix(A, 3))
        RS("txn_typ") = "BAS"
        RS("USR_IDY") = USERNAME
        If CmbBase.ListIndex = 0 Then
            RS("b_con_app") = "AND"
            RS("B_PRD_QTY") = Val(MSbase.TextMatrix(A, 2))
        ElseIf CmbBase.ListIndex = 1 Then
            RS("b_con_app") = "ORO"
            RS("B_PRD_QTY") = Val(MSbase.TextMatrix(A, 2))
        ElseIf CmbBase.ListIndex = 2 Then
            RS("b_con_app") = "ORM"
            RS("B_PRD_QTY") = Val(MSbase.TextMatrix(A, 2))
        End If
        RS("SYS_DAT") = Date
        RS("SYS_TIM") = Format(Time, "hh:mm:ss")
        RS.Update
    End If
Next

For A = 1 To MSFree.Rows - 1
    If MSFree.TextMatrix(A, 0) <> "" Then
        RS.AddNew
        RS("sch_idy") = schid
        RS("sch_nme") = Trim(Sch_Nme)
        RS("sch_typ") = Sch_Cmb
        RS("f_sch_typ") = Sub_Cmb
        RS("SCH_STS") = Sch_Act
        RS("SCH_BDT") = StDate
        RS("sch_edt") = EnDate
        RS("B_PRD_IDY") = Trim(MSFree.TextMatrix(A, 4))
        RS("BAT_IDY") = Trim(MSFree.TextMatrix(A, 3))
        If CmbFree.ListIndex = 0 Then
            RS("F_con_app") = "AND"
            RS("B_PRD_QTY") = Val(MSFree.TextMatrix(A, 2))
        ElseIf CmbFree.ListIndex = 1 Then
            RS("F_con_app") = "ORO"
            RS("B_PRD_QTY") = Val(MSFree.TextMatrix(A, 2))
        End If
        RS("PRO_RAT") = "N"
        RS("txn_typ") = "FRE"
        RS("USR_IDY") = USERNAME
        RS("SYS_DAT") = Date
        RS("SYS_TIM") = Format(Time, "hh:mm:ss")
        RS.Update
    End If
Next
MsgBox "New Scheme ID is " & schid, vbExclamation, "Effmcg"
CON.CommitTrans
cancel
Exit Sub
lab:
MsgBox err.Description
CON.RollbackTrans
End Sub

Private Sub TxtGrsAmt_GotFocus()
SLT
End Sub

Private Sub TxtGrsAmt_keypress(K As Integer)
If K = 13 Then
    TxtGrsAmt = AMT(TxtGrsAmt)
    If Sub_Cmb.ListIndex > 1 Then
        TxtValue.SetFocus
    ElseIf Sub_Cmb.ListIndex = 1 Then
        TabFree.SetFocus
        MSFree.SetFocus
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
''ON ERROR GoTo lab
If Val(MSbase.TextMatrix(1, 2)) = 0 Then
    MsgBox "Enter Valid Quantity", vbExclamation
    SSTab1.Tab = 1
    TabBase.Tab = 0
    MSbase.SetFocus
    MSbase.Col = 2
    MSbase.Row = 1
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from sch", CON, adOpenKeyset, adLockOptimistic
CON.BeginTrans
schid = "SCH" & T7increment("idy_sch")
RS.AddNew
RS("sch_idy") = schid
RS("sch_nme") = Trim(Sch_Nme)
RS("sch_typ") = Sch_Cmb
RS("f_sch_typ") = Sub_Cmb
RS("SCH_STS") = Sch_Act
RS("SCH_BDT") = StDate
RS("sch_edt") = EnDate
RS("B_PRD_IDY") = Trim(MSbase.TextMatrix(1, 4))
RS("B_PRD_QTY") = Val(MSbase.TextMatrix(1, 2))
RS("BAT_IDY") = Trim(MSbase.TextMatrix(1, 3))
RS("txn_typ") = "BAS"
RS("USR_IDY") = USERNAME
RS("SYS_DAT") = Date
RS("SYS_TIM") = Format(Time, "hh:mm:ss")
RS.Update

RS.AddNew
RS("sch_idy") = schid
RS("sch_nme") = Trim(Sch_Nme)
RS("sch_typ") = Sch_Cmb
RS("f_sch_typ") = Sub_Cmb
RS("SCH_STS") = Sch_Act
RS("SCH_BDT") = DTPicker1
RS("sch_edt") = DTPicker2
If Sub_Cmb.ListIndex = 2 Then RS("B_DIS_AMT") = convert(AMT(Val(TxtValue) & ""))
If Sub_Cmb.ListIndex = 3 Then RS("B_DIS_PCG") = discount(Val(TxtValue))
RS("f_prd_idy") = ""
RS("BAT_IDY") = ""
If Sub_Cmb.ListIndex = 1 Then RS("txn_typ") = "CAS"
If Sub_Cmb.ListIndex = 2 Then RS("txn_typ") = "PCG"
If Pro_Chk.Value = 1 Then
    RS("pro_rat") = "Y"
Else
    RS("pro_rat") = "N"
End If
RS("USR_IDY") = USERNAME
RS("SYS_DAT") = Date
RS("SYS_TIM") = Format(Time, "hh:mm:ss")
RS.Update
MsgBox "New Scheme ID is " & schid, vbExclamation, "Effmcg"
CON.CommitTrans
cancel
Exit Sub
lab:
CON.RollbackTrans
MsgBox err.Description
cancel
End Sub

Public Sub savemultiplecashordiscount()
'ON ERROR GoTo lab
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from sch", CON, adOpenKeyset, adLockOptimistic
CON.BeginTrans
schid = "SCH" & T7increment("idy_sch")
For A = 1 To MSbase.Rows - 1
    If MSbase.TextMatrix(A, 0) <> "" Then
        RS.AddNew
        RS("sch_idy") = schid
        RS("sch_nme") = Trim(Sch_Nme)
        RS("sch_typ") = Sch_Cmb
        RS("f_sch_typ") = Sub_Cmb
        RS("SCH_STS") = Sch_Act
        RS("SCH_BDT") = StDate
        RS("sch_edt") = EnDate
        RS("B_PRD_IDY") = Trim(MSbase.TextMatrix(A, 4))
        RS("BAT_IDY") = Trim(MSbase.TextMatrix(A, 3))
        RS("txn_typ") = "BAS"
        RS("USR_IDY") = USERNAME
        If CmbBase.ListIndex = 0 Then
            RS("b_con_app") = "AND"
            RS("B_PRD_QTY") = Val(MSbase.TextMatrix(A, 2))
        ElseIf CmbBase.ListIndex = 1 Then
            RS("b_con_app") = "ORO"
            RS("B_PRD_QTY") = Val(MSbase.TextMatrix(1, 2))
        ElseIf CmbBase.ListIndex = 2 Then
            RS("b_con_app") = "ORM"
            RS("B_PRD_QTY") = Val(MSbase.TextMatrix(1, 2))
        End If
        RS("SYS_DAT") = Date
        RS("SYS_TIM") = Format(Time, "hh:mm:ss")
        RS.Update
    End If
Next
RS.AddNew
RS("sch_idy") = schid
RS("sch_nme") = Trim(Sch_Nme)
RS("sch_typ") = Sch_Cmb
RS("f_sch_typ") = Sub_Cmb
RS("SCH_STS") = Sch_Act
RS("SCH_BDT") = DTPicker1
RS("sch_edt") = DTPicker2
If Sub_Cmb.ListIndex = 2 Then RS("B_DIS_AMT") = convert(AMT(Val(TxtValue) & ""))
If Sub_Cmb.ListIndex = 3 Then RS("B_DIS_PCG") = discount(Val(TxtValue))
RS("f_prd_idy") = ""
RS("BAT_IDY") = ""
If Sub_Cmb.ListIndex = 1 Then RS("txn_typ") = "CAS"
If Sub_Cmb.ListIndex = 2 Then RS("txn_typ") = "PCG"
If Pro_Chk.Value = 1 Then
    RS("pro_rat") = "Y"
Else
    RS("pro_rat") = "N"
End If
RS("USR_IDY") = USERNAME
RS("SYS_DAT") = Date
RS("SYS_TIM") = Format(Time, "hh:mm:ss")
RS.Update
MsgBox "New Scheme ID is " & schid, vbExclamation, "Effmcg"
CON.CommitTrans
Exit Sub
lab:
CON.RollbackTrans
MsgBox err.Description
cancel
End Sub
Public Sub savebillcashordiscount()
'ON ERROR GoTo lab
If Val(TxtGrsAmt) = 0 Then
    MsgBox "Invalid Gross Amount", vbExclamation
    TxtGrsAmt.SetFocus
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from sch", CON, adOpenKeyset, adLockOptimistic
CON.BeginTrans
schid = "SCH" & T7increment("idy_sch")
If MSFree.Rows - 1 = 0 Then MSFree.Rows = 2
For A = 1 To MSFree.Rows - 1
    If MSFree.TextMatrix(A, 0) <> "" Or Sub_Cmb.ListIndex > 1 Then
        RS.AddNew
        RS("sch_idy") = schid
        RS("sch_nme") = Trim(Sch_Nme)
        RS("sch_typ") = Sch_Cmb
        RS("f_sch_typ") = Sub_Cmb
        RS("SCH_STS") = Sch_Act
        RS("SCH_BDT") = StDate
        RS("sch_edt") = EnDate
        RS("B_BIL_AMT") = convert(AMT(Val(TxtGrsAmt)))
        RS("BAT_IDY") = ""
        If Sub_Cmb.ListIndex = 1 Then
            RS("BAT_IDY") = Trim(MSFree.TextMatrix(A, 3))
            RS("B_PRD_IDY") = Trim(MSFree.TextMatrix(A, 4)) & ""
            RS("B_PRD_QTY") = Val(MSFree.TextMatrix(A, 2)) & ""
            RS("txn_typ") = "FRE"
        End If
        If Sub_Cmb.ListIndex = 2 Then RS("B_DIS_AMT") = convert(AMT(Val(TxtValue) & ""))
        If Sub_Cmb.ListIndex = 3 Then RS("B_DIS_PCG") = discount(TxtValue & "")
        If Sub_Cmb.ListIndex = 2 Then RS("txn_typ") = "CAS"
        If Sub_Cmb.ListIndex = 3 Then RS("txn_typ") = "PCG"
        If Pro_Chk.Value = 1 Then
            RS("pro_rat") = "Y"
        Else
            RS("pro_rat") = "N"
        End If
        RS("USR_IDY") = USERNAME
        RS("SYS_DAT") = Date
        RS("SYS_TIM") = Format(Time, "hh:mm:ss")
        RS.Update
    End If
Next
CON.CommitTrans
MsgBox "New Scheme ID is " & schid, vbExclamation, "Effmcg"
cancel
Exit Sub
lab:
CON.RollbackTrans
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
REC_SET.Open "SELECT " & Trim(TYP) & " FROM PRD,PGR WHERE PGR.GRP_IDY LIKE PRD.GRP_IDY AND " & Trim(TYP) & " LIKE '" & Trim(TxtPack) & "%' GROUP BY " & TYP, CON, adOpenKeyset
lv1.ColumnHeaders.Clear
lv1.ColumnHeaders.Add , , Sch_Cmb, 2500, 0
lv1.ListItems.Clear
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
    MsgBox Sch_Cmb & " Not Found, Define " & Sch_Cmb & " !", vbExclamation, "Effmcg"
    TxtPack.Visible = False
    SSTab1.Tab = 1
    MSPack.SetFocus
    Exit Sub
End If
End Sub

Public Sub SAVEGRPTYP()
'ON ERROR GoTo lab
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from INC", CON, adOpenKeyset, adLockOptimistic
CON.BeginTrans
schid = "SCH" & T7increment("idy_sch")
For A = 1 To MSPack.Rows - 1
    If MSPack.TextMatrix(A, 0) <> "" Then
        RS.AddNew
        RS("sch_idy") = schid
        RS("sch_nme") = Trim(Sch_Nme)
        RS("sch_typ") = Sch_Cmb
        RS("f_sch_typ") = Sub_Cmb
        RS("SCH_STS") = Sch_Act
        RS("SCH_BDT") = StDate
        RS("sch_edt") = EnDate
        If Sub_Cmb.ListIndex = 2 Then RS("B_DIS_AMT") = convert(AMT(Val(TxtValue) & ""))
        If Sub_Cmb.ListIndex = 3 Then RS("B_DIS_PCG") = discount(TxtValue & "")
        If Sub_Cmb.ListIndex = 2 Then RS("txn_typ") = "CAS"
        If Sub_Cmb.ListIndex = 3 Then RS("txn_typ") = "PCG"
        RS!BAT_IDY = ""
        RS!grp_typ = Trim(MSPack.TextMatrix(A, 0))
        If CmbPCon.ListIndex = 0 Then
            RS!B_CON_APP = "AND"
            RS!B_PRD_QTY = Val(MSPack.TextMatrix(A, 1))
        ElseIf CmbPCon.ListIndex = 1 Then
            RS!B_CON_APP = "ORO"
            RS!B_PRD_QTY = Val(MSPack.TextMatrix(1, 1))
        ElseIf CmbPCon.ListIndex = 2 Then
            RS!B_CON_APP = "ORM"
            RS!B_PRD_QTY = Val(MSPack.TextMatrix(1, 1))
        End If
        RS!TXN_TYP = "BAS"
        RS.Update
    End If
Next
If Sub_Cmb.ListIndex = 1 Then
    For A = 1 To MSFree.Rows - 1
        If MSFree.TextMatrix(A, 0) <> "" Then
            RS.AddNew
            RS("sch_idy") = schid
            RS("sch_nme") = Trim(Sch_Nme)
            RS("sch_typ") = Sch_Cmb
            RS("f_sch_typ") = Sub_Cmb
            RS("SCH_STS") = Sch_Act
            RS("SCH_BDT") = StDate
            RS("sch_edt") = EnDate
            RS("B_PRD_IDY") = Trim(MSFree.TextMatrix(A, 4))
            RS("BAT_IDY") = Trim(MSFree.TextMatrix(A, 3))
            If CmbFree.ListIndex = 0 Then
                RS("F_con_app") = "AND"
                RS("B_PRD_QTY") = Val(MSFree.TextMatrix(A, 2))
            ElseIf CmbFree.ListIndex = 1 Then
                RS("F_con_app") = "ORO"
                RS("B_PRD_QTY") = Val(MSFree.TextMatrix(1, 2))
            End If
            RS("PRO_RAT") = "N"
            RS("txn_typ") = "FRE"
            RS("USR_IDY") = USERNAME
            RS("SYS_DAT") = Date
            RS("SYS_TIM") = Format(Time, "hh:mm:ss")
            RS.Update
        End If
    Next
End If
CON.CommitTrans
cancel
Exit Sub
lab:
CON.RollbackTrans
MsgBox err.Description
cancel
End Sub

Public Sub dispschemegrt3()
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from INC where sch_idy like '" & Trim(Sch_Idy) & "'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    RS.filter = "txn_typ='BAS'"
    If RS!B_CON_APP = "AND" Then
        CmbPCon.ListIndex = 0
    ElseIf RS!B_CON_APP = "ORO" Then
        CmbPCon.ListIndex = 1
    Else
        CmbPCon.ListIndex = 2
    End If
    MSPack.Rows = RS.RecordCount + 1
    If Sub_Cmb.ListIndex = 2 Then TxtValue = convert(RS!B_DIS_AMT)
    If Sub_Cmb.ListIndex = 3 Then TxtValue = convert(RS!B_DIS_PCG)
    For A = 1 To RS.RecordCount
        MSPack.TextMatrix(A, 0) = Trim(RS!grp_typ)
        MSPack.TextMatrix(A, 1) = Trim(RS!B_PRD_QTY)
        RS.MoveNext
    Next
    If Sub_Cmb.ListIndex = 1 Then
        RS.filter = "txn_typ='FRE'"
        If RS.RecordCount > 0 Then
            If RS!F_CON_APP = "AND" Then
                CmbFree.ListIndex = 0
            ElseIf RS!F_CON_APP = "ORO" Then
                CmbFree.ListIndex = 1
            End If
        End If
        MSFree.Rows = RS.RecordCount + 1
        If RS.RecordCount > 0 Then
            For A = 1 To RS.RecordCount
                MSFree.TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'")
                MSFree.TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'"))
                MSFree.TextMatrix(A, 2) = Val(RS!B_PRD_QTY & "")
                RS.MoveNext
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
        If M_FRM_INC.ActiveControl.name = "MSbase" Then
            MSFree.SetFocus
            MSFree.Row = MSFree.Rows - 1
        ElseIf M_FRM_INC.ActiveControl.name = "MSFree" Then
            MSbase.SetFocus
            MSbase.Row = MSbase.Rows - 1
        End If
    ElseIf (SC = 1 Or SC = 3) And (SU = 2 Or SU = 3) Then
        If M_FRM_INC.ActiveControl.name = "MSbase" Then
            TxtValue.SetFocus
            SLT
        ElseIf M_FRM_INC.ActiveControl.name = "TxtValue" Then
            MSbase.SetFocus
            MSbase.Row = MSbase.Rows - 1
        End If
    ElseIf SC = 2 Then
        If M_FRM_INC.ActiveControl.name = "TxtGrsAmt" Then
            If SU = 1 Then
                MSFree.SetFocus
                MSFree.Row = MSFree.Rows - 1
            Else
                TxtValue.SetFocus
                SLT
            End If
        End If
    ElseIf SC > 3 Then
        If SU = 1 Then
            If M_FRM_INC.ActiveControl.name = "MSPack" Then
                MSFree.SetFocus
                MSFree.Row = MSFree.Rows - 1
            ElseIf M_FRM_INC.ActiveControl.name = "MSFree" Then
                MSPack.SetFocus
                MSPack.Row = MSPack.Rows - 1
            End If
        ElseIf SU > 1 Then
            If M_FRM_INC.ActiveControl.name = "MSPack" Then
                TxtValue.SetFocus
                SLT
            ElseIf M_FRM_INC.ActiveControl.name = "TxtValue" Then
                MSPack.SetFocus
                MSPack.Row = MSPack.Rows - 1
            End If
        End If
    End If
End If
End Sub

Public Sub DISPSCHEMESKU()
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from INC where sch_idy like '" & Trim(Sch_Idy) & "'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    RS.filter = "txn_typ='BAS'"
    If RS!B_CON_APP = "AND" Then
        CmbBase.ListIndex = 0
    ElseIf RS!B_CON_APP = "ORO" Then
        CmbBase.ListIndex = 1
    ElseIf RS!B_CON_APP = "ORM" Then
        CmbBase.ListIndex = 2
    Else
        MSbase.Top = 360
        MSbase.Height = 4215
        MSFree.Top = 360
        MSFree.Height = 4215
    End If
    With MSbase
        .Rows = RS.RecordCount + 1
        If RS.RecordCount > 0 Then
            For A = 1 To RS.RecordCount
                .TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'")
                .TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'"))
                .TextMatrix(A, 2) = Val(RS!B_PRD_QTY & "")
                RS.MoveNext
            Next
        End If
    End With
    
    If Sub_Cmb.ListIndex = 2 Then
        RS.filter = "txn_typ='PCG'"
        If RS.RecordCount > 0 Then TxtValue = convert(RS!B_DIS_AMT)
    ElseIf Sub_Cmb.ListIndex = 3 Then
        RS.filter = "txn_typ='PCG'"
        If RS.RecordCount > 0 Then TxtValue = convert(RS!B_DIS_PCG & "")
    End If
    RS.filter = "txn_typ='FRE'"
        MSFree.Rows = RS.RecordCount + 1
        If RS.RecordCount > 0 Then
            If RS!F_CON_APP = "AND" Then
            CmbFree.ListIndex = 0
        ElseIf RS!F_CON_APP = "ORO" Then
            CmbFree.ListIndex = 1
        End If

        For A = 1 To RS.RecordCount
            MSFree.TextMatrix(A, 0) = GetValue("PRD_NME", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'")
            MSFree.TextMatrix(A, 1) = convert(GetValue("BRD.PRD_MRP", "PRD,BRD WHER PRD.PRD_IDY LIKE BRD.PRD_IDY AND BAT_IDY LIKE '" & Trim(RS!BAT_IDY) & "'"))
            MSFree.TextMatrix(A, 2) = Val(RS!B_PRD_QTY & "")
            RS.MoveNext
        Next
        End If
End If
End Sub
Private Sub smnid_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesman ID or Press <Enter> to get List of Salesman ID's"
End Sub

Private Sub smnid_KeyDown(K As Integer, S As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If RS.State = 1 Then RS.Close
        RS.Open "select * from smn where smn_idy like '" & Sincrement(smnid) & "'", CON, adOpenStatic
        If RS.RecordCount > 0 Then
            smnid = Sincrement(smnid)
            SMNNME = RS("smn_nme")
        Else
            smnlist
        End If
End If
End Sub

Private Sub SMNID_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub
Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.Open "select * from smn where smn_nme like '" & Trim(smnid) & "%'", CON, adOpenStatic
lv1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = lv1.ListItems.Add(, , Trim(RS("smn_nme")))
        ls.lv1.Add , , Trim(RS("smn_idy"))
        ls.lv1.Add , , Trim(RS("smn_ad4") & "")
        RS.MoveNext
    Next
    SSTab1.TabEnabled(2) = True
    SSTab1.Tab = 2
    Label20 = "Select Salesman"
    lv1.ColumnHeaders(1).Text = "Salesman Name"
    lv1.ColumnHeaders(2).Text = "Salesman ID"
    lv1.ColumnHeaders(3).Text = "Place"
    lv1.Visible = True
    lv1.SetFocus
Else
    MsgBox "Salesmen Not Found, Define Salesmen !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    smnid.Enabled = True
    smnid.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub

