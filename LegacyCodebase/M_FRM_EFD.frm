VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form M_FRM_EFD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export - Import File Data"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_EFD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   465
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&File Data Import - Retailers, SKUs"
      TabPicture(0)   =   "M_FRM_EFD.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cd1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtFileName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command6"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command8"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_EFD.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "lv"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Command8 
         Caption         =   "&Validate Batch"
         Height          =   375
         Left            =   4560
         TabIndex        =   47
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Update Retailers"
         Height          =   375
         Left            =   3120
         TabIndex        =   46
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "E&xport Rates"
         Height          =   375
         Left            =   1680
         TabIndex        =   45
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Export Retailers"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Close"
         Height          =   375
         Left            =   7560
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2040
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1125
         Width           =   1845
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Left            =   3960
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1125
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select File Data Category"
         Height          =   600
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   7215
         Begin VB.OptionButton opt_tax 
            Caption         =   "&Taxes - Rates"
            Height          =   195
            Left            =   5520
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   $"M_FRM_EFD.frx":047A
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton opt_prd 
            Caption         =   "&Products - SKU"
            Height          =   195
            Left            =   3595
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   $"M_FRM_EFD.frx":0529
            Top             =   300
            Width           =   1470
         End
         Begin VB.OptionButton opt_pgr 
            Caption         =   "&SKU Categories"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1670
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   300
            Width           =   1470
         End
         Begin VB.OptionButton opt_act 
            Caption         =   "&Retailers"
            Height          =   195
            Left            =   225
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   $"M_FRM_EFD.frx":05DC
            Top             =   300
            Width           =   990
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4440
         Width           =   975
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   600
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "xls"
         DialogTitle     =   "Select Excel File"
         FileName        =   "*.xls"
         Filter          =   "*.xls"
      End
      Begin MSComctlLib.ListView lv 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   7815
         _ExtentX        =   13785
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Default Values"
         Height          =   2535
         Left            =   240
         TabIndex        =   14
         Top             =   1475
         Visible         =   0   'False
         Width           =   7215
         Begin MSMAPI.MAPIMessages MAPIMessages1 
            Left            =   6120
            Top             =   720
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin VB.TextBox GRP_IDY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Width           =   1185
         End
         Begin VB.TextBox AREAID 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   600
            Width           =   1185
         End
         Begin VB.TextBox marId 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1185
         End
         Begin VB.TextBox Text32 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1185
         End
         Begin VB.TextBox CATID 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   960
            Width           =   1185
         End
         Begin VB.ComboBox ROU_LST 
            Height          =   315
            ItemData        =   "M_FRM_EFD.frx":06AE
            Left            =   1800
            List            =   "M_FRM_EFD.frx":071B
            Style           =   2  'Dropdown List
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1920
         End
         Begin MSMAPI.MAPISession MAPISession1 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   0   'False
            LogonUI         =   -1  'True
            NewSession      =   -1  'True
         End
         Begin MSCommLib.MSComm MSComm1 
            Left            =   120
            Top             =   960
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Group ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   33
            Top             =   315
            Width           =   645
         End
         Begin VB.Label LABEL4 
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
            Left            =   3720
            TabIndex        =   32
            Top             =   240
            Width           =   2625
         End
         Begin VB.Label Label29 
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
            Left            =   3720
            TabIndex        =   28
            Top             =   600
            Width           =   2625
         End
         Begin VB.Label Label30 
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
            Left            =   3720
            TabIndex        =   27
            Top             =   1320
            Width           =   2625
         End
         Begin VB.Label Label19 
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
            Left            =   3720
            TabIndex        =   26
            Top             =   960
            Width           =   2625
         End
         Begin VB.Label Label31 
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
            Left            =   3720
            TabIndex        =   25
            Top             =   1680
            Width           =   2625
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Route - &Beat"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   24
            Top             =   2100
            Width           =   900
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prime &Location ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   23
            Top             =   675
            Width           =   1260
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Market ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   4
            Left            =   675
            TabIndex        =   22
            Top             =   1395
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S&alesmen ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   21
            Top             =   1755
            Width           =   900
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categ&ory ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   2
            Left            =   540
            TabIndex        =   20
            Top             =   1035
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Default Values"
         Height          =   2175
         Left            =   240
         TabIndex        =   2
         Top             =   1700
         Visible         =   0   'False
         Width           =   5895
         Begin VB.TextBox grpid 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label GRPSEG 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2595
            TabIndex        =   41
            Top             =   840
            Width           =   2955
            WordWrap        =   -1  'True
         End
         Begin VB.Label GRPCAT 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2595
            TabIndex        =   40
            Top             =   360
            Width           =   2955
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   435
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   300
            TabIndex        =   5
            Top             =   1395
            Width           =   780
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2595
            TabIndex        =   4
            Top             =   1320
            Width           =   2955
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5280
         TabIndex        =   37
         Top             =   4080
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3360
         TabIndex        =   36
         Top             =   4080
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Added-Updated"
         Height          =   195
         Left            =   3720
         TabIndex        =   35
         Top             =   4080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Records"
         Height          =   195
         Left            =   2040
         TabIndex        =   34
         Top             =   4080
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -74640
         TabIndex        =   30
         Top             =   720
         Width           =   7815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Worksheet"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1193
         Width           =   1275
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
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
            Picture         =   "M_FRM_EFD.frx":0788
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":1094
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":13E8
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":1A20
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":1D74
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":21D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":24C8
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":28D4
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":29E0
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":2D34
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EFD.frx":3148
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
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
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "pre"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_EFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsact As New ADODB.Recordset, RS_VAN As New ADODB.Recordset
Dim rsloc As New ADODB.Recordset
Dim rsmar As New ADODB.Recordset
Dim RSSMN As New ADODB.Recordset
Dim RSCAT As New ADODB.Recordset
Dim rspgr As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim RSBRD As New ADODB.Recordset
Dim rsgrp As New ADODB.Recordset
Dim rsrdy As New ADODB.Recordset
'Dim MAPISession1 As MAPISession
'Dim MAPIMessages1 As MAPIMessages
Dim appXL As Excel.Application
Dim appX As Excel.Workbook
'Dim appXL As Object
'Dim appX As Object
Dim r As Integer, C As Integer, A As Integer
Dim Retail_Update As Boolean
Private Sub AREAID_keydown(K As Integer, Shift As Integer)
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsloc As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If rsloc.State = 1 Then rsloc.Close
        rsloc.Open " select * from LOC where LOC_IDY like '" & Sincrement(AREAID) & "'", con, adOpenStatic
        If rsloc.RecordCount > 0 Then
            AREAID = Sincrement(AREAID)
            Label29 = rsloc("LOC_NME")
            If K = 38 Then
                
            Else
                CATID.SetFocus
            End If
        Else
            arealist
        End If
End If
End Sub
Private Sub Command1_Click()
On Error GoTo errh
cd1.ShowOpen
cd1.CancelError = False
If Not Trim(cd1.FileName) = "" Then
    If LCase(Right(Trim(cd1.FileName), 3)) = "xls" Then ' Edited
        txtFileName.Text = cd1.FileName
        Set appX = GetObject(Trim(txtFileName.Text))
        Frame1.Enabled = True
        'opt_act.SetFocus
    Else
        MsgBox "Select A Valid WorkSheet,XLS File for Data Import! Remove AMPERSAND and SINGLE QUOTE characters!", vbCritical, "DAS Version 4.0"
        txtFileName.Text = ""
        appX.Close 0
        appXL.Quit
        Set appXL = Nothing
        Set appX = Nothing
        Exit Sub
    End If
End If
If opt_act.Value = True Then
    GRP_IDY.SetFocus
ElseIf opt_prd.Value = True Then
    grpid.SetFocus
ElseIf opt_pgr.Value = True Then
    Command2.SetFocus
ElseIf opt_tax.Value = True Then
    Command2.SetFocus
End If
Exit Sub
errh:
MsgBox "Select a Valid File Database for Data Import!", vbInformation, "DAS Version 4.0"
txtFileName.SetFocus
SendKeys "{HOME}+{END}"
End Sub
Private Sub Command2_Click()
'If appXL.State = 1 Then appXL.Close
On Error GoTo ERRH1
    Set appX = GetObject(Trim(txtFileName.Text))
    If opt_act.Value = True Then
        act
    ElseIf opt_pgr.Value = True Then
        pgr
    ElseIf opt_prd.Value = True Then
        prd
    ElseIf opt_tax.Value = True Then
        SKU_RAT
    End If
Exit Sub
ERRH1:
MsgBox "Select A Valid WorkSheet,XLS File for Data Import! Remove AMPERSAND and SINGLE QUOTE characters!", vbCritical, "DAS Version 4.0"
txtFileName.SetFocus
SendKeys "{home}+{end}"
End Sub
Private Sub Command3_Click()
Set appXL = Nothing
Set appX = Nothing
Unload Me
End Sub
Private Sub Command4_Click()
' * Start of Distributor ID Identification
    Dim DIR_RS As New ADODB.Recordset
    Dim M_WDS_IDY As String
    Dim M_WDS_SRL As String
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        If DIR_RS.BOF = False Then
            DIR_RS.MoveFirst
            M_WDS_IDY = Trim(DIR_RS("DIR_ALS"))
        End If
        If Len(M_WDS_IDY) < 8 Then
            MsgBox ("Distributor ID Required for this Process! Obtain Distributor ID and Start Again"), vbCritical, "DAS Version 4.0"
            Exit Sub
        End If
        M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
        If Len(Trim(M_WDS_SRL)) < 3 Then
             While Len(Trim(M_WDS_SRL)) < 3
                  M_WDS_SRL = "0" & M_WDS_SRL
             Wend
        End If
    End If
    
' * Start of Version Information
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'SYS' and DIR_IDX = 'VER'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        If DIR_RS.BOF = False Then
            DIR_RS.MoveFirst
            DIR_RS("DIR_MSG") = Trim("01.01.01")
            DIR_RS.Update
        End If
    End If
' * End of Version Information
' * End of Distributor ID Identification
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker4, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker3, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveFirst
            ST_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
            Pending_RS.MoveLast
            EN_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
        End If

        
    ' * : Beginning of _RETAILERS.TXT Creation Code for Text Data Export
    Me.MousePointer = vbHourglass
    Dim F_OMS, F_LOC, F_MAR, F_SMN
    Dim FS_SOU, FS_LOC, FS_MAR, FS_SMN As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
'    If FS_SOU.FileExists(CURDIR & "\HHT\PARTY.TXT") = True Then
'        FS_SOU.DeleteFile (CURDIR & "\HHT\PARTY.TXT")
'    End If
    
    M_FIL_NME = "C:\DEX\" & Trim(M_WDS_IDY) & "_RETAILER.TXT"
    Set F_OMS = FS_SOU.OpenTextFile("C:\DEX\" & Trim(M_WDS_IDY) & "_RETAILER.TXT", 2, True)
    
    Dim J As Integer
    Dim CCnt As Integer
    J = 0
    CCnt = 1
    
    Dim Rrs As New ADODB.Recordset
    Dim Category_RS As New ADODB.Recordset
    Dim Location_RS As New ADODB.Recordset
    Dim Market_RS As New ADODB.Recordset
    Dim Salesman_RS As New ADODB.Recordset
    Dim RDY_RS As New ADODB.Recordset
    
    If Rrs.State = 1 Then Rrs.Close
    Rrs.Open "select * from ACT where ACT_TYP = 'SAL' or ACT_TYP = 'PUR'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
    
    If Rrs.RecordCount > 0 Then Rrs.MoveFirst
       
       Dim L_CAT_NME As String
       Dim L_LOC_NME As String
       Dim L_MAR_NME As String
       Dim L_SMN_NME As String
       Dim L_ROU_IDY As String
       Dim L_DST_NME As String
       Dim L_MAR_GRD As String
       Dim D_MON, D_TUE, D_WED, D_THU, D_FRI, D_SAT, D_SUN As String
       
       If Category_RS.State = 1 Then Category_RS.Close
       Category_RS.Open "SELECT * FROM CAT", con, adOpenKeyset, adLockPessimistic
       If Location_RS.State = 1 Then Location_RS.Close
       Location_RS.Open "SELECT * FROM LOC", con, adOpenKeyset, adLockPessimistic
       If Market_RS.State = 1 Then Market_RS.Close
       Market_RS.Open "SELECT * FROM MAR", con, adOpenKeyset, adLockPessimistic
       If Salesman_RS.State = 1 Then Salesman_RS.Close
       Salesman_RS.Open "SELECT * FROM SMN", con, adOpenKeyset, adLockPessimistic
       If RDY_RS.State = 1 Then RDY_RS.Close
       RDY_RS.Open "SELECT * FROM RDY ORDER BY DAY_IDY", con, adOpenKeyset, adLockPessimistic
       
       ' Title Line Record for Text Data Export File
       
       F_OMS.WriteLine "DAS_ACT_IDY" & "|" & "NAME" & "|" & "GST" & "|" & "CST" & "|" & "ADDRESS1" & "|" & "ADDRESS2" & "|" & "ADDRESS3" & "|" & "ADDRESS4" & "|" & "ZIPCODE" & "|" & "TELEPHONE" & "|" & "CONTACT" & "|" & "LOCATION" & "|" & "MARKET" & "|" & "SALESMAN" & "|" & "CATEGORY" & "|" & "ROUTE" & "|" & "GRADE" & "|" & "ACTIVE" & "|" & "CAPACITY" & "|" & "RED CODE" & "|" & "HVO CODE" & "|" & "TYPE" & "|" & "CASH-CREDIT" & "|" & "MONDAY" & "|" & "TUESDAY" & "|" & "WEDNESDAY" & "|" & "THURSDAY" & "|" & "FRIDAY" & "|" & "SATURDAY" & "|" & "SUNDAY" & "|" & "CUSTOMER GROUP" & "|" & "RSA CLUSTER" & "|" & "RED OUTLET CLASS" & "|" & "TOWN CLASS" & "|" & "TOWN" & "|" & "DISTRICT" & "|" & "STATE" & "|" & "COUNTRY" & "|" & "GCC ID"
       
       While Not Rrs.EOF
       
          If Location_RS.RecordCount > 0 Then Location_RS.MoveFirst
          Location_RS.Find "LOC_IDY='" & Trim(Rrs("LOC_IDY")) & "'"
          If Location_RS.EOF = False Then
             L_LOC_NME = Location_RS("LOC_NME")
          End If
          
          If Market_RS.RecordCount > 0 Then Market_RS.MoveFirst
          Market_RS.Find "MAR_IDY='" & Trim(Rrs("MAR_IDY")) & "'"
          If Market_RS.EOF = False Then
             L_MAR_NME = Market_RS("MAR_NME")
             L_DST_NME = Market_RS("MAR_PGR")
             L_MAR_GRD = Market_RS("MAR_GRD")
                ' Added for Margin Minder Data Extract by PNR
             L_MAR_CSG = Market_RS("MAR_CSG")
             L_MAR_CLU = Market_RS("MAR_CLU")
             L_MAR_ROC = Market_RS("MAR_ROC")
             L_MAR_TWC = Market_RS("MAR_TWC")
             L_MAR_TWN = Market_RS("MAR_TWN")
             L_MAR_DST = Market_RS("MAR_DST")
             L_MAR_STA = Market_RS("MAR_STA")
             L_MAR_COU = Market_RS("MAR_COU")
          End If
          
          If RDY_RS.RecordCount > 0 Then RDY_RS.MoveFirst
          RDY_RS.Find "ACT_IDY='" & Trim(Rrs("ACT_IDY")) & "'"
          If RDY_RS.EOF = False Then
             L_SMN_IDY = Trim(RDY_RS("SMN_IDY"))
             L_ROU_IDY = Trim(RDY_RS("ROU_IDY"))
          Else
             L_SMN_IDY = ""
             L_ROU_IDY = ""
          End If
          
          While Not RDY_RS.EOF
          If RDY_RS("ACT_IDY") = Trim(Rrs("ACT_IDY")) Then
                If RDY_RS("DAY_IDY") = 1 Then
                    D_MON = "MONDAY"
                ElseIf RDY_RS("DAY_IDY") = 2 Then
                    D_TUE = "TUESDAY"
                ElseIf RDY_RS("DAY_IDY") = 3 Then
                    D_WED = "WEDNESDAY"
                ElseIf RDY_RS("DAY_IDY") = 4 Then
                    D_THU = "THURSDAY"
                ElseIf RDY_RS("DAY_IDY") = 5 Then
                    D_FRI = "FRIDAY"
                ElseIf RDY_RS("DAY_IDY") = 6 Then
                    D_SAT = "SATURDAY"
                ElseIf RDY_RS("DAY_IDY") = 7 Then
                    D_SUN = "SUNDAY"
                End If
          End If
          RDY_RS.MoveNext
          Wend
          
          If Salesman_RS.RecordCount > 0 Then Salesman_RS.MoveFirst
          Salesman_RS.Find "SMN_IDY='" & Trim(L_SMN_IDY) & "'"
          If Salesman_RS.EOF = False Then
             L_SMN_NME = Salesman_RS("SMN_NME")
          End If
          
          If Category_RS.RecordCount > 0 Then Category_RS.MoveFirst
          Category_RS.Find "CAT_IDY='" & Trim(Rrs("CAT_IDY")) & "'"
          If Category_RS.EOF = False Then
             L_CAT_NME = Category_RS("CAT_NME")
          End If


'          F_OMS.WriteLine Trim(M_WDS_IDY) & "|" & Trim(Rrs("ACT_IDY")) & "|" & Trim(Rrs("ACT_NME")) & "|" & Trim(Rrs("ACT_AD1")) & "|" & Trim(Rrs("ACT_AD2")) & "|" & Trim(Rrs("ACT_AD3")) & "|" & Trim(Rrs("ACT_AD4")) & "|" & Trim(L_DST_NME) & "|" & Trim(Rrs("ACT_TFX")) & "|" & Trim(Rrs("ACT_MBL")) & "|" & Trim(Rrs("ACT_EML")) & "|" & Trim(Rrs("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(Rrs("BBN_ONE")) & "|" & Trim(Rrs("CST_TYP")) & "|" & Trim(Rrs("RET_TYP")) & "|" & Trim(Rrs("VAN_CAP")) & "|" & Trim(Rrs("DBT_TYP"))
          
'          If Trim(Rrs("ACT_TYP")) = "PUR" Then
'                F_OMS.WriteLine Trim(Rrs("ACT_IDY")) & "|" & "S" & "|" & Trim(Rrs("CST_TYP")) & "|" & Trim(Rrs("ACT_NME")) & "|" & Trim(Rrs("ACT_AD1")) & "|" & Trim(Rrs("ACT_AD2")) & "|" & Trim(Rrs("ACT_AD3")) & "|" & Trim(Rrs("ACT_AD4")) & "|" & Trim(Rrs("ACT_TFX")) & "|" & Trim(Rrs("ACT_MBL")) & "|" & Trim(Rrs("ACT_EML")) & "|" & Trim(Rrs("ACT_CPN")) & "|" & Trim(Rrs("LOC_IDY")) & "|" & Trim(Rrs("MAR_IDY")) & "|" & Trim(Rrs("CAT_IDY")) & "|" & Trim(Rrs("RET_TYP")) & "|" & Trim(Rrs("BBN_ONE")) & "|" & Trim(Rrs("DBT_TYP")) & "|" & Trim(Rrs("VAN_CAP")) & "|" & Trim(Rrs("ROU_IDY"))
'                J = J + 1
'          End If
          
          If Trim(Rrs("ACT_TYP")) = "SAL" Then
                F_OMS.WriteLine Trim(Rrs("ACT_IDY")) & "|" & Trim(Rrs("ACT_NME")) & "|" & Trim(Rrs("ACT_GST")) & "|" & Trim(Rrs("ACT_CST")) & "|" & Trim(Rrs("ACT_AD1")) & "|" & Trim(Rrs("ACT_AD2")) & "|" & Trim(Rrs("ACT_AD3")) & "|" & Trim(Rrs("ACT_AD4")) & "|" & Trim(Rrs("ACT_ZIP")) & "|" & Trim(Rrs("ACT_TFX")) & "-" & Trim(Rrs("ACT_MBL")) & "|" & Trim(Rrs("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(L_SMN_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_ROU_IDY) & "|" & Trim(Rrs("RET_TYP")) & "|" & Trim(Rrs("DBT_TYP")) & "|" & Trim(Rrs("VAN_CAP")) & "|" & Trim(Rrs("ACN_TWO")) & "|" & Trim(Rrs("BNM_TWO")) & "|" & Trim(Rrs("CST_TYP")) & "|" & Trim(Rrs("MCL_REQ")) & "|" & D_MON & "|" & D_TUE & "|" & D_WED & "|" & D_THU & "|" & D_FRI & "|" & D_SAT & "|" & D_SUN & "|" & L_MAR_CSG & "|" & L_MAR_CLU & "|" & L_MAR_ROC & "|" & L_MAR_TWC & "|" & L_MAR_TWN & "|" & L_MAR_DST & "|" & L_MAR_STA & "|" & L_MAR_COU & "|" & Trim(Rrs("GCC_ID"))
                J = J + 1
          End If
          
          L_SMN_IDY = ""
          L_ROU_IDY = ""
          L_SMN_NME = ""
          L_CAT_NME = ""
          L_LOC_NME = ""
          L_MAR_NME = ""
          L_MAR_CSG = ""
          L_MAR_CLU = ""
          L_MAR_ROC = ""
          L_MAR_TWC = ""
          L_MAR_TWN = ""
          L_MAR_DST = ""
          L_MAR_STA = ""
          L_MAR_COU = ""
          D_MON = ""
          D_TUE = ""
          D_WED = ""
          D_THU = ""
          D_FRI = ""
          D_SAT = ""
          D_SUN = ""
          
          Rrs.MoveNext
       Wend
       
        F_OMS.Close
        MsgBox ("Text File " & M_FIL_NME & " with " & J & " Records Created in the Folder C:\DEX"), vbInformation, "DAS Version 4.0"
        Me.MousePointer = Normal
        Command3.SetFocus
    
    ' * : End of _RETAILERS.TXT Creation Code for Text Data Export
End Sub

Private Sub Command5_Click()
' * Start of Distributor ID Identification
    Dim DIR_RS As New ADODB.Recordset
    Dim M_WDS_IDY As String
    Dim M_WDS_SRL As String
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        If DIR_RS.BOF = False Then
            DIR_RS.MoveFirst
            M_WDS_IDY = Trim(DIR_RS("DIR_ALS"))
        End If
        If Len(M_WDS_IDY) < 8 Then
            MsgBox ("Distributor ID Required for this Process! Obtain Distributor ID and Start Again"), vbCritical, "DAS Version 4.0"
            Exit Sub
        End If
        M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
        If Len(Trim(M_WDS_SRL)) < 3 Then
             While Len(Trim(M_WDS_SRL)) < 3
                  M_WDS_SRL = "0" & M_WDS_SRL
             Wend
        End If
    End If
    
' * Start of Version Information
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'SYS' and DIR_IDX = 'VER'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        If DIR_RS.BOF = False Then
            DIR_RS.MoveFirst
            DIR_RS("DIR_MSG") = Trim("01.01.01")
            DIR_RS.Update
        End If
    End If
' * End of Version Information
' * End of Distributor ID Identification
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker4, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker3, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveFirst
            ST_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
            Pending_RS.MoveLast
            EN_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
        End If

        
    ' * : Beginning of _SKU_RATE.TXT Creation Code for Text Data Export
    Me.MousePointer = vbHourglass
    Dim F_SKU
    Dim FS_SOU As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    M_FIL_NME = "C:\DEX\" & Trim(M_WDS_IDY) & "_SKU_RATE.TXT"
    Set F_SKU = FS_SOU.OpenTextFile("C:\DEX\" & Trim(M_WDS_IDY) & "_SKU_RATE.TXT", 2, True)
    
    Dim L_BAS_RAT, L_REO_LEV, L_LST_RAT, L_LST_DIS As Integer
    Dim J As Integer
    J = 0
    
    Dim SKU_RS As New ADODB.Recordset
    Dim BAT_RS As New ADODB.Recordset
    
    If SKU_RS.State = 1 Then SKU_RS.Close
    SKU_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    If SKU_RS.RecordCount > 0 Then SKU_RS.MoveFirst
    
    If BAT_RS.State = 1 Then BAT_RS.Close
    BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
       
       ' Title Line Record for Text Data Export File
       
       F_SKU.WriteLine "DAS_SKU_IDY" & "|" & "SKU NAME" & "|" & "DAS_BATCH_IDY" & "|" & "BATCH NAME" & "|" & "SUPPLIER ID" & "|" & "PUR RATE" & "|" & "PUR TAX" & "|" & "SAL RATE" & "|" & "SAL TAX" & "|" & "MRP RATE" & "|" & "CRATE RENTAL" & "|" & "CRATE TAX" & "|" & "LIFE" & "|" & "SKU STATUS" & "|" & "BASE RATE" & "|" & "REORDER" & "|" & "PRESENT QTY" & "|" & "PUR INC-TVE" & "|" & "SAL INC-TVE"
       
       While Not BAT_RS.EOF
       
            If SKU_RS.RecordCount > 0 Then SKU_RS.MoveFirst
                SKU_RS.Find "PRD_IDY='" & Trim(BAT_RS("PRD_IDY")) & "'"
                If SKU_RS.EOF = False Then
'                    If BAT_RS("PRD_QTY") > 0 Then 'Commented by PNR as per PRINCE CHOPRA UPDATE
                        If IsNull(SKU_RS("BAS_RAT")) Or IsNull(SKU_RS("REO_LEV")) Or IsNull(SKU_RS("LST_RAT")) Or IsNull(SKU_RS("LST_DIS")) Then
                            L_BAS_RAT = 0#
                            L_REO_LEV = 0#
                            L_LST_RAT = 0#
                            L_LST_DIS = 0#
                        Else
                            L_BAS_RAT = 0#   'SKU_RS("BAS_RAT") Verified and changed based on Prince Chopra Update, either field should be made available in BRD or otherwise 0.00 default value
                            L_REO_LEV = SKU_RS("REO_LEV")
                            L_LST_RAT = SKU_RS("LST_RAT")
                            L_LST_DIS = SKU_RS("LST_DIS")
                        End If
                        If IsNull(BAT_RS("P_INC_DIS")) Or IsNull(BAT_RS("S_INC_DIS")) Then
                            LP_INC_DIS = 0#
                            LS_INC_DIS = 0#
                        Else
                            LP_INC_DIS = Val(BAT_RS("P_INC_DIS"))
                            LS_INC_DIS = Val(BAT_RS("S_INC_DIS"))
                        End If
                        
                        F_SKU.WriteLine Trim(SKU_RS("PRD_IDY")) & "|" & Trim(SKU_RS("PRD_NME")) & "|" & Trim(BAT_RS("BAT_IDY")) & "|" & Trim(BAT_RS("BAT_NME")) & "|" & Trim(SKU_RS("PRD_PRN")) & "|" & Trim(convert(BAT_RS("PRD_PDR"))) & "|" & Trim(convert(SKU_RS("PUR_TAX"))) & "|" & Trim(convert(BAT_RS("PRD_SDR"))) & "|" & Trim(convert(SKU_RS("SAL_TAX"))) & "|" & Trim(convert(BAT_RS("PRD_MRP"))) & "|" & Trim(str(L_LST_RAT)) & "|" & Trim(str(L_LST_DIS)) & "|" & Trim(convert(SKU_RS("TOD_PCG"))) & "|" & Trim(SKU_RS("STK_TYP")) & "|" & Trim(str(L_BAS_RAT)) & "|" & Format(Trim(str(L_REO_LEV)), "0.000") & "|" & Format(Trim(str(BAT_RS("PRD_QTY"))), "0.000") & "|" & Trim(convert(Val(LP_INC_DIS))) & "|" & Trim(convert(Val(LS_INC_DIS)))
                        J = J + 1
'                    End If                         'Commented by PNR as per PRINCE CHOPRA UPDATE
                End If
            BAT_RS.MoveNext
       Wend
       
        F_SKU.Close
        MsgBox ("Text File " & M_FIL_NME & " with " & J & " Records Created in the Folder C:\DEX"), vbInformation, "DAS Version 4.0"
        Me.MousePointer = Normal
        Command3.SetFocus
    
    ' * : End of _SKU_RATE.TXT Creation Code for Text Data Export
End Sub

Private Sub Command6_Click()
Dim appX As New Workbook

On Error GoTo errh
cd1.ShowOpen
cd1.CancelError = False
If Not Trim(cd1.FileName) = "" Then
    If LCase(Right(Trim(cd1.FileName), 3)) = "xls" Then ' Edited
        txtFileName.Text = cd1.FileName
'        Set appX = GetObject(Trim(txtFileName.Text))
        Frame1.Enabled = True
        'opt_act.SetFocus
    Else
        appX.Close 0
        appXL.Quit
        Set appXL = Nothing
        Set appX = Nothing
        MsgBox "Select A Valid WorkSheet,XLS File for Data Import! Remove AMPERSAND and SINGLE QUOTE characters!", vbCritical, "DAS Version 4.0"
        txtFileName.Text = ""
        Exit Sub
    End If
End If
'Set appX = GetObject(Trim(txtFileName.Text))
ARET_UPD 'Module to update Retailers List Back into Effmcg from XLS
Exit Sub
errh:
MsgBox "Select a Valid XLS File for Retailers Update!", vbInformation, "DAS Version 4.0"
'appX.Close 0
'appXL.Quit
Set appXL = Nothing
Set appX = Nothing
Command6.SetFocus
SendKeys "{HOME}+{END}"
End Sub
Private Sub Command7_Click()
    MailSend "info@consoftsystems.com", "MAPI Test", "Test Message!", "C:\DEX\MIM.ZIP"
    MailSend "info@consoftsystems.com", "MAPI Test", "Test Message!", "C:\DEX\MIM.ZIP"
End Sub

Private Sub Command8_Click()
Dim BRD_VRF As New ADODB.Recordset
If BRD_VRF.State = 1 Then BRD_VRF.Close
BRD_VRF.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
If BRD_VRF.RecordCount > 0 Then BRD_VRF.MoveFirst
While Not BRD_VRF.EOF
    If IsNull(BRD_VRF!P_INC_DIS) = True Then
        BRD_VRF!P_INC_DIS = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_INC_DIS) = True Then
        BRD_VRF!S_INC_DIS = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_PUR_TAX) = True Then
        BRD_VRF!B_PUR_TAX = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_SAL_TAX) = True Then
        BRD_VRF!B_SAL_TAX = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_CRT_RAT) = True Then
        BRD_VRF!B_CRT_RAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_CRT_DIS) = True Then
        BRD_VRF!B_CRT_DIS = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!SKU_SBU) = True Then
        BRD_VRF!SKU_SBU = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_PUR_VAT) = True Then
        BRD_VRF!P_PUR_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_CRT_RNT) = True Then
        BRD_VRF!P_CRT_RNT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_CRT_VAT) = True Then
        BRD_VRF!P_CRT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_TOT_VAT) = True Then
        BRD_VRF!P_TOT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!U_PUR_LPP) = True Then
        BRD_VRF!U_PUR_LPP = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!PUR_LPP) = True Then
        BRD_VRF!PUR_LPP = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_SAL_VAT) = True Then
        BRD_VRF!S_SAL_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_CRT_RNT) = True Then
        BRD_VRF!S_CRT_RNT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_CRT_VAT) = True Then
        BRD_VRF!S_CRT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_TOT_VAT) = True Then
        BRD_VRF!S_TOT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!U_SAL_LPP) = True Then
        BRD_VRF!U_SAL_LPP = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!SAL_LPP) = True Then
        BRD_VRF!SAL_LPP = 0#
        BRD_VRF.Update
    End If
    BRD_VRF.MoveNext
Wend
'If BRD_VRF.State = 1 Then BRD_VRF.Close
MsgBox "Batch Validation Completed!", vbInformation, "DAS Version 4.0"
End Sub
Private Sub Form_Activate()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
opt_act.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K = 81 And s = 2 Then
    CANCELRECORD
ElseIf K = 88 And s = 2 Then
    Set appXL = Nothing
    Set appX = Nothing
    Unload Me
ElseIf K = 112 Then
    M_HELP
ElseIf K = 80 And s = 2 Then
    M_FRM_AQY.Tag = "ACT"
    M_FRM_AQY.Show 1
ElseIf K = 82 And s = 2 Then
End If
End Sub
Public Sub ARET_UPD() ' Retailers Update Module from Excel Sheet, Updates Existing Records Only
If CHECKTB("act.dbf,rdy.dbf,loc.dbf,smn.dbf,mar.dbf,cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If MsgBox("Update from " & txtFileName, vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then Exit Sub
Command2.Enabled = False
If rsgrp.State = 1 Then rsgrp.Close
rsgrp.Open "select * from grp where grp_IDY like '" & Trim(GRP_IDY) & "'", con, adOpenKeyset, adLockPessimistic
C = 0
Set appX = GetObject(txtFileName.Text)
   appX.Sheets(1).Cells(1, 40).Formula = "=counta(a1:a65536)"
    MaxRows = appX.Sheets(1).Cells(1, 40)
If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "DAS_ACT_IDY" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 3))) = "GST" And UCase(Trim(appX.Sheets(1).Cells(1, 4))) = "CST" And UCase(Trim(appX.Sheets(1).Cells(1, 5))) = "ADDRESS1" And UCase(Trim(appX.Sheets(1).Cells(1, 6))) = "ADDRESS2" And UCase(Trim(appX.Sheets(1).Cells(1, 7))) = "ADDRESS3" And UCase(Trim(appX.Sheets(1).Cells(1, 8))) = "ADDRESS4" And UCase(Trim(appX.Sheets(1).Cells(1, 9))) = "ZIPCODE" And UCase(Trim(appX.Sheets(1).Cells(1, 10))) = "TELEPHONE" And UCase(Trim(appX.Sheets(1).Cells(1, 11))) = "CONTACT" And UCase(Trim(appX.Sheets(1).Cells(1, 12))) = "LOCATION" And UCase(Trim(appX.Sheets(1).Cells(1, 13))) = "MARKET" And UCase(Trim(appX.Sheets(1).Cells(1, 14))) = "SALESMAN" And UCase(Trim(appX.Sheets(1).Cells(1, 15))) = "CATEGORY" And UCase(Trim(appX.Sheets(1).Cells(1, 16))) = "ROUTE" And UCase(Trim(appX.Sheets(1).Cells(1, 17))) = "GRADE" And UCase(Trim(appX.Sheets(1).Cells(1, 18))) = "ACTIVE" _
        And UCase(Trim(appX.Sheets(1).Cells(1, 20))) = "RED CODE" And UCase(Trim(appX.Sheets(1).Cells(1, 21))) = "HVO CODE" And UCase(Trim(appX.Sheets(1).Cells(1, 22))) = "TYPE" And UCase(Trim(appX.Sheets(1).Cells(1, 23))) = "CASH-CREDIT" And UCase(Trim(appX.Sheets(1).Cells(1, 24))) = "MONDAY" And UCase(Trim(appX.Sheets(1).Cells(1, 25))) = "TUESDAY" And UCase(Trim(appX.Sheets(1).Cells(1, 26))) = "WEDNESDAY" And UCase(Trim(appX.Sheets(1).Cells(1, 27))) = "THURSDAY" And UCase(Trim(appX.Sheets(1).Cells(1, 28))) = "FRIDAY" And UCase(Trim(appX.Sheets(1).Cells(1, 29))) = "SATURDAY" And UCase(Trim(appX.Sheets(1).Cells(1, 30))) = "SUNDAY" And UCase(Trim(appX.Sheets(1).Cells(1, 31))) = "CUSTOMER GROUP" And UCase(Trim(appX.Sheets(1).Cells(1, 32))) = "RSA CLUSTER" And UCase(Trim(appX.Sheets(1).Cells(1, 33))) = "RED OUTLET CLASS" And UCase(Trim(appX.Sheets(1).Cells(1, 34))) = "TOWN CLASS" And UCase(Trim(appX.Sheets(1).Cells(1, 35))) = "TOWN" And UCase(Trim(appX.Sheets(1).Cells(1, 36))) = "DISTRICT" _
        And UCase(Trim(appX.Sheets(1).Cells(1, 37))) = "STATE" And UCase(Trim(appX.Sheets(1).Cells(1, 38))) = "COUNTRY" And UCase(Trim(appX.Sheets(1).Cells(1, 39))) = "GCC ID" Then
''''''  And UCase(Trim(appX.sheets(1).cells(1, 18))) = "CAPACITY"
''''''    If Trim(GRP_IDY) = "" Then
''''''        MsgBox "Select Group Account by Pressing Enter Key, SUNDRY DEBTORS - RETAILERS or Press 7", vbInformation, "DAS Version 4.0"
''''''        GRP_IDY.SetFocus
''''''        Exit Sub
''''''    End If
    Me.MousePointer = vbHourglass
    If rsrdy.State = 1 Then rsrdy.Close
    rsrdy.Open "select * from rdy where act_idy=''", con, adOpenKeyset, adLockPessimistic
    
'       appX.Sheets(1).Cells(1, 29).Formula = "=counta(a1:a65536)"
'    MaxRows = appX.Sheets(1).Cells(1, 29)
'
'    appX.Sheets(1).Cells(1, 29).Formula = "=counta(a1:a65536)"
'    MaxRows = appX.Sheets(1).Cells(1, 29)
    Label7 = MaxRows - 1
'    If Len(Trim(GetValue("GCC_ID", "ACT WHERE ACT_IDY = '" & M_CST_IDY & "'"))) = 0 Then
'        GCC_Validate = True
'    Else
'        GCC_Validate = False
'    End If

    For r = 2 To MaxRows + 1
        SHOWBUT (True)
        If rsloc.State = 1 Then rsloc.Close
        If rsmar.State = 1 Then rsmar.Close
        If RSSMN.State = 1 Then RSSMN.Close
        If RSCAT.State = 1 Then RSCAT.Close
        If Trim(appX.Sheets(1).Cells(r, 1)) = "" Then
            appX.Close 0
            appXL.Quit
            Set appXL = Nothing
            Set appX = Nothing
            Exit For
        End If
        CheckGCCStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'GCC'")
        If Trim(CheckGCCStatus) = "T" Then
            If Len(Trim(appX.Sheets(1).Cells(r, 39))) = 0 Then
                'Retail_Update = False
                Retail_Update = True 'Changed to true to allow record updates without GCC as per Prince Chopra feedback
            Else
                Retail_Update = True
            End If
        Else
            Retail_Update = True
        End If
        If rsact.State = 1 Then rsact.Close
        rsact.Open "SELECT * FROM ACT WHERE ACT_IDY = '" & Trim(appX.Sheets(1).Cells(r, 1)) & "'", con, adOpenKeyset, adLockPessimistic

            If rsact.RecordCount > 0 Then rsact.MoveFirst
'            rsact.Find "ACT_IDY='" & Trim(appX.Sheets(1).Cells(r, 1)) & "'"
            
            If rsact.EOF = False Then
'                If Trim(CheckGCCStatus) = "T" Then
'                    CheckGCCDuplicate = GetValue("GCC_ID", "ACT WHERE GCC_ID = '" & Trim(appX.Sheets(1).Cells(r, 39)) & "'")
'                    If Trim(CheckGCCDuplicate) = Trim(appX.Sheets(1).Cells(r, 39)) Then
'                        Retail_Update = False
'                    Else
'                        Retail_Update = True
'                    End If
'                End If
            
'                If Retail_Update = True Then
'                CheckGCCDuplicate = GetValue("GCC_ID", "ACT WHERE GCC_ID = '" & Trim(appX.Sheets(1).Cells(r, 39)) & "'")
'                If Trim(CheckGCCDuplicate) <> Trim(appX.Sheets(1).Cells(r, 39)) Then
                    rsact("act_nme") = UCase(ADS(rsact, "ACT_NME", appX.Sheets(1).Cells(r, 2)))
                    rsact("act_gst") = UCase(ADS(rsact, "act_gst", appX.Sheets(1).Cells(r, 3)))
                    rsact("act_cst") = UCase(ADS(rsact, "act_cst", appX.Sheets(1).Cells(r, 4)))
                    rsact("act_ad1") = UCase(ADS(rsact, "act_ad1", appX.Sheets(1).Cells(r, 5)))
                    rsact("act_ad2") = UCase(ADS(rsact, "act_ad2", appX.Sheets(1).Cells(r, 6)))
                    rsact("act_ad3") = UCase(ADS(rsact, "act_ad3", appX.Sheets(1).Cells(r, 7)))
                    rsact("act_ad4") = UCase(ADS(rsact, "act_ad4", appX.Sheets(1).Cells(r, 8)))
                    rsact("act_zip") = ADS(rsact, "act_zip", appX.Sheets(1).Cells(r, 9))
                    rsact("act_tfx") = ADS(rsact, "act_tfx", appX.Sheets(1).Cells(r, 10))
                    rsact("act_cpn") = UCase(ADS(rsact, "act_cpn", appX.Sheets(1).Cells(r, 11)))
                    rsact("ret_typ") = UCase(ADS(rsact, "ret_typ", appX.Sheets(1).Cells(r, 17)))
                    If Len(Trim(appX.Sheets(1).Cells(r, 39))) > 0 Then
                        rsact("GCC_ID") = UCase(ADS(rsact, "GCC_ID", appX.Sheets(1).Cells(r, 39)))
                    End If
                    rsact("sku_rng") = 0#
                    If ADS(rsact, "dbt_typ", appX.Sheets(1).Cells(r, 18)) = "" Then
                        rsact("dbt_typ") = "Y"
                    Else
                        rsact("dbt_typ") = UCase(ADS(rsact, "dbt_typ", appX.Sheets(1).Cells(r, 18)))
                    End If
                    rsact("van_cap") = Val(UCase(ADS(rsact, "van_cap", appX.Sheets(1).Cells(r, 19))))
                    rsact("acn_two") = UCase(ADS(rsact, "acn_two", appX.Sheets(1).Cells(r, 20)))
                    rsact("bnm_two") = UCase(ADS(rsact, "bnm_two", appX.Sheets(1).Cells(r, 21)))
                    If ADS(rsact, "mcl_req", appX.Sheets(1).Cells(r, 23)) = "" Then
                        rsact("mcl_req") = "Y"
                    ElseIf ADS(rsact, "mcl_req", appX.Sheets(1).Cells(r, 23)) = "N" Then
                        rsact("mcl_req") = "N"
                    Else
                        rsact("mcl_req") = UCase(ADS(rsact, "mcl_req", appX.Sheets(1).Cells(r, 23)))
                    End If
                    'rsact("cst_typ") = UCase(ADS(rsact, "cst_typ", appX.Sheets(1).cells(r, 21)))
                    If (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "Retailer" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "AMC" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "SubDistributor" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "Key A/c" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "Local Key A/c" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "Regnl Key A/c" Or UCase(ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "HVO-P" Or UCase(ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22))) = "HVO-NP" Or ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22)) = "Van" Then
                        rsact("cst_typ") = (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 22)))
                    Else
                        rsact("cst_typ") = "Retailer"
                    End If
                    If Trim(appX.Sheets(1).Cells(r, 12)) <> "" Then
                        If rsloc.State = 1 Then rsloc.Close
                        rsloc.Open "SELECT * FROM LOC WHERE TRIM(LOC_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 12))) & "'", con, adOpenKeyset, adLockPessimistic
                        If rsloc.EOF = False Then
                            rsact("LOC_IDY") = rsloc("LOC_IDY")
                        Else
                            rsloc.AddNew
                            rsloc("LOC_IDY") = increment("IDY_LOC")
                            rsloc("LOC_NME") = UCase(ADS(rsloc, "loc_nme", appX.Sheets(1).Cells(r, 12)))
                            rsloc("LOC_TYP") = ""
                            rsact("LOC_IDY") = rsloc("LOC_IDY")
                            rsloc.Update
                        End If
                    Else
                        rsact("LOC_IDY") = Trim(AREAID)
                    End If
                    If Trim(appX.Sheets(1).Cells(r, 14)) <> "" Then
                        If RSSMN.State = 1 Then RSSMN.Close
                        RSSMN.Open "SELECT * FROM SMN WHERE TRIM(SMN_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 14))) & "'", con, adOpenKeyset, adLockPessimistic
                        If RSSMN.EOF = False Then
                            rsact("SMN_IDY") = RSSMN("SMN_IDY")
                        Else
                            RSSMN.AddNew
                            For i = 0 To RSSMN.Fields.count - 1
                                If RSSMN(i).Type = adVarChar Or RSSMN(i).Type = adChar Or RSSMN(i).Type = adBSTR Or RSSMN(i).Type = adLongVarChar Or RSSMN(i).Type = adLongVarWChar Then
                                    RSSMN(i) = ""
                                ElseIf RSSMN(i).Type = adInteger Or RSSMN(i).Type = adNumeric Or RSSMN(i).Type = adBigInt Or RSSMN(i).Type = adSingle Or RSSMN(i).Type = adDouble Or RSSMN(i).Type = adBigInt Or RSSMN(i).Type = adCurrency Or RSSMN(i).Type = adNumeric Or RSSMN(i).Type = adSmallInt Or RSSMN(i).Type = adDecimal Then
                                    RSSMN(i) = 0
                                ElseIf RSSMN(i).Type = adDate Or RSSMN(i).Type = adDBDate Or RSSMN(i).Type = adDBTimeStamp Then
                                    RSSMN(i) = Date
                                End If
                            Next
                            RSSMN("SMN_IDY") = increment("IDY_SMN")
                            RSSMN("SMN_NME") = UCase(ADS(RSSMN, "smn_nme", appX.Sheets(1).Cells(r, 14)))
                            rsact("SMN_IDY") = RSSMN("SMN_IDY")
                            RSSMN.Update
                        End If
                    Else
                        rsact("SMN_IDY") = Trim(Text32)
                    End If
                    If Trim(appX.Sheets(1).Cells(r, 13)) <> "" Then
                        If rsmar.State = 1 Then rsmar.Close
                        rsmar.Open "SELECT * FROM MAR WHERE TRIM(MAR_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 13))) & "'", con, adOpenKeyset, adLockPessimistic
                        If rsmar.EOF = False Then
                            rsact("MAR_IDY") = rsmar("MAR_IDY")
                            rsmar("MAR_NME") = UCase(ADS(rsmar, "mar_nme", appX.Sheets(1).Cells(r, 13)))
                            rsmar("MAR_CSG") = UCase(ADS(rsmar, "MAR_CSG", appX.Sheets(1).Cells(r, 31)))
                            rsmar("MAR_CLU") = UCase(ADS(rsmar, "MAR_CLU", appX.Sheets(1).Cells(r, 32)))
                            rsmar("MAR_ROC") = UCase(ADS(rsmar, "MAR_ROC", appX.Sheets(1).Cells(r, 33)))
                            rsmar("MAR_TWC") = UCase(ADS(rsmar, "MAR_TWC", appX.Sheets(1).Cells(r, 34)))
                            rsmar("MAR_TWN") = UCase(ADS(rsmar, "MAR_TWN", appX.Sheets(1).Cells(r, 35)))
                            rsmar("MAR_DST") = UCase(ADS(rsmar, "MAR_DST", appX.Sheets(1).Cells(r, 36)))
                            rsmar("MAR_STA") = UCase(ADS(rsmar, "MAR_STA", appX.Sheets(1).Cells(r, 37)))
                            rsmar("MAR_COU") = UCase(ADS(rsmar, "MAR_COU", appX.Sheets(1).Cells(r, 38)))
                            rsmar.Update
                        Else
                            rsmar.AddNew
                            For i = 0 To rsmar.Fields.count - 1
                                If rsmar(i).Type = adVarChar Or rsmar(i).Type = adChar Or rsmar(i).Type = adBSTR Or rsmar(i).Type = adLongVarChar Or rsmar(i).Type = adLongVarWChar Then
                                    rsmar(i) = ""
                                ElseIf rsmar(i).Type = adInteger Or rsmar(i).Type = adNumeric Or rsmar(i).Type = adBigInt Or rsmar(i).Type = adSingle Or rsmar(i).Type = adDouble Or rsmar(i).Type = adBigInt Or rsmar(i).Type = adCurrency Or rsmar(i).Type = adNumeric Or rsmar(i).Type = adSmallInt Or rsmar(i).Type = adDecimal Then
                                    rsmar(i) = 0
                                ElseIf rsmar(i).Type = adDate Or rsmar(i).Type = adDBDate Or rsmar(i).Type = adDBTimeStamp Then
                                    rsmar(i) = Date
                                End If
                            Next
                            rsmar("MAR_IDY") = increment("IDY_MAR")
                            rsmar("MAR_NME") = UCase(ADS(rsmar, "mar_nme", appX.Sheets(1).Cells(r, 13)))
                            rsmar("SMN_IDY") = rsact("SMN_IDY")
                            rsmar("ROU_IDY") = "B"
                            rsact("MAR_IDY") = rsmar("MAR_IDY")
                            ' Added for Margin Minder Data Extract by PNR
                            rsmar("MAR_CSG") = UCase(ADS(rsmar, "MAR_CSG", appX.Sheets(1).Cells(r, 31)))
                            rsmar("MAR_CLU") = UCase(ADS(rsmar, "MAR_CLU", appX.Sheets(1).Cells(r, 32)))
                            rsmar("MAR_ROC") = UCase(ADS(rsmar, "MAR_ROC", appX.Sheets(1).Cells(r, 33)))
                            rsmar("MAR_TWC") = UCase(ADS(rsmar, "MAR_TWC", appX.Sheets(1).Cells(r, 34)))
                            rsmar("MAR_TWN") = UCase(ADS(rsmar, "MAR_TWN", appX.Sheets(1).Cells(r, 35)))
                            rsmar("MAR_DST") = UCase(ADS(rsmar, "MAR_DST", appX.Sheets(1).Cells(r, 36)))
                            rsmar("MAR_STA") = UCase(ADS(rsmar, "MAR_STA", appX.Sheets(1).Cells(r, 37)))
                            rsmar("MAR_COU") = UCase(ADS(rsmar, "MAR_COU", appX.Sheets(1).Cells(r, 38)))
                            rsmar.Update
                        End If
                    Else
                        rsact("mar_idy") = Trim(marId)
                    End If
                    If Trim(appX.Sheets(1).Cells(r, 15)) <> "" Then
                        If RSCAT.State = 1 Then RSCAT.Close
                        RSCAT.Open "SELECT * FROM CAT WHERE TRIM(CAT_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 15))) & "'", con, adOpenKeyset, adLockPessimistic
                        If RSCAT.EOF = False Then
                            rsact("CAT_IDY") = RSCAT("CAT_IDY")
                        Else
                            RSCAT.AddNew
                            For i = 0 To RSCAT.Fields.count - 1
                                If RSCAT(i).Type = adVarChar Or RSCAT(i).Type = adChar Or RSCAT(i).Type = adBSTR Or RSCAT(i).Type = adLongVarChar Or RSCAT(i).Type = adLongVarWChar Then
                                    RSCAT(i) = " "
                                ElseIf RSCAT(i).Type = adInteger Or RSCAT(i).Type = adNumeric Or RSCAT(i).Type = adBigInt Or RSCAT(i).Type = adSingle Or RSCAT(i).Type = adDouble Or RSCAT(i).Type = adBigInt Or RSCAT(i).Type = adCurrency Or RSCAT(i).Type = adNumeric Or RSCAT(i).Type = adSmallInt Or RSCAT(i).Type = adDecimal Then
                                    RSCAT(i) = 0
                                ElseIf RSCAT(i).Type = adDate Or RSCAT(i).Type = adDBDate Or RSCAT(i).Type = adDBTimeStamp Then
                                    RSCAT(i) = Date
                                End If
                            Next
                            RSCAT("CAT_IDY") = increment("IDY_CAT")
                            RSCAT("CAT_NME") = UCase(ADS(RSCAT, "cat_nme", appX.Sheets(1).Cells(r, 15)))
                            rsact("CAT_IDY") = RSCAT("CAT_IDY")
                            RSCAT.Update
                        End If
                    Else
                        rsact("cat_idy") = Trim(CATID)
                    End If
                    If UCase(appX.Sheets(1).Cells(r, 16)) <> "" Then
                        rsact("rou_idy") = UCase(ADS(rsact, "rou_idy", UCase(appX.Sheets(1).Cells(r, 16))))
                    Else
                        rsact("rou_idy") = Trim(ROU_LST & "")
                    End If
            '        rsact("dbt_typ") = "Y" - Active - Inactive Status captured from Xcel only
                    rsact("key_act") = "N"
                    rsact("dbt_rap") = "B"
                    rsact("act_typ") = "SAL"
                    rsact.Update
'                    End If
                    'If Retail_Update = True Then
                    ' Retailers Update Module End Here and Beat Plan Module Starts From Here
                    con.Execute "DELETE FROM RDY WHERE ACT_IDY LIKE '" & Trim(appX.Sheets(1).Cells(r, 1)) & "'"
                    'CON.Execute "delete from act where act_idy like '" & Sincrement(Text1) & "'"
                    If UCase(Trim(appX.Sheets(1).Cells(r, 24))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 25))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 26))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 27))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 28))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 29))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 30))) = "" Then
                        For d = 1 To 7
                            rsrdy.AddNew
                            For K = 0 To rsrdy.Fields.count - 1
                                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                    rsrdy(K) = " "
                                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                    rsrdy(K) = 0
                                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                    rsrdy(K) = Date
                                End If
                            Next
                            rsrdy("act_idy") = rsact("act_idy")
                            rsrdy("day_idy") = d
                            rsrdy("smn_idy") = rsact("smn_idy")
                            rsrdy("rou_idy") = Trim(rsact("rou_idy") & "")
                        Next
                    End If
                    ' Beat Plan Logic incorporated by PNR from EXCEL sheet - not compulsory, if available, RDY table is updated
                    If UCase(Trim(appX.Sheets(1).Cells(r, 24))) = "MONDAY" Then
                        rsrdy.AddNew
                        For K = 0 To rsrdy.Fields.count - 1
                            If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                rsrdy(K) = " "
                            ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                rsrdy(K) = 0
                            ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                rsrdy(K) = Date
                            End If
                        Next
                        rsrdy("ACT_IDY") = rsact("ACT_IDY")
                        rsrdy("DAY_IDY") = 1
                        rsrdy("SMN_IDY") = rsact("SMN_IDY")
                        rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
                    End If
                    If UCase(Trim(appX.Sheets(1).Cells(r, 25))) = "TUESDAY" Then
                        rsrdy.AddNew
                        For K = 0 To rsrdy.Fields.count - 1
                            If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                rsrdy(K) = " "
                            ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                rsrdy(K) = 0
                            ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                rsrdy(K) = Date
                            End If
                        Next
                        rsrdy("ACT_IDY") = rsact("ACT_IDY")
                        rsrdy("DAY_IDY") = 2
                        rsrdy("SMN_IDY") = rsact("SMN_IDY")
                        rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
                    End If
                    If UCase(Trim(appX.Sheets(1).Cells(r, 26))) = "WEDNESDAY" Then
                        rsrdy.AddNew
                        For K = 0 To rsrdy.Fields.count - 1
                            If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                rsrdy(K) = " "
                            ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                rsrdy(K) = 0
                            ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                rsrdy(K) = Date
                            End If
                        Next
                        rsrdy("ACT_IDY") = rsact("ACT_IDY")
                        rsrdy("DAY_IDY") = 3
                        rsrdy("SMN_IDY") = rsact("SMN_IDY")
                        rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
                    End If
                    If UCase(Trim(appX.Sheets(1).Cells(r, 27))) = "THURSDAY" Then
                        rsrdy.AddNew
                        For K = 0 To rsrdy.Fields.count - 1
                            If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                rsrdy(K) = " "
                            ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                rsrdy(K) = 0
                            ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                rsrdy(K) = Date
                            End If
                        Next
                        rsrdy("ACT_IDY") = rsact("ACT_IDY")
                        rsrdy("DAY_IDY") = 4
                        rsrdy("SMN_IDY") = rsact("SMN_IDY")
                        rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
                    End If
                    If UCase(Trim(appX.Sheets(1).Cells(r, 28))) = "FRIDAY" Then
                        rsrdy.AddNew
                        For K = 0 To rsrdy.Fields.count - 1
                            If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                rsrdy(K) = " "
                            ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                rsrdy(K) = 0
                            ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                rsrdy(K) = Date
                            End If
                        Next
                        rsrdy("ACT_IDY") = rsact("ACT_IDY")
                        rsrdy("DAY_IDY") = 5
                        rsrdy("SMN_IDY") = rsact("SMN_IDY")
                        rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
                    End If
                    If UCase(Trim(appX.Sheets(1).Cells(r, 29))) = "SATURDAY" Then
                        rsrdy.AddNew
                        For K = 0 To rsrdy.Fields.count - 1
                            If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                rsrdy(K) = " "
                            ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                rsrdy(K) = 0
                            ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                rsrdy(K) = Date
                            End If
                        Next
                        rsrdy("ACT_IDY") = rsact("ACT_IDY")
                        rsrdy("DAY_IDY") = 6
                        rsrdy("SMN_IDY") = rsact("SMN_IDY")
                        rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
                    End If
                    If UCase(Trim(appX.Sheets(1).Cells(r, 30))) = "SUNDAY" Then
                        rsrdy.AddNew
                        For K = 0 To rsrdy.Fields.count - 1
                            If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                                rsrdy(K) = " "
                            ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                                rsrdy(K) = 0
                            ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                                rsrdy(K) = Date
                            End If
                        Next
                        rsrdy("ACT_IDY") = rsact("ACT_IDY")
                        rsrdy("DAY_IDY") = 7
                        rsrdy("SMN_IDY") = rsact("SMN_IDY")
                        rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
                    End If
                    C = C + 1
                    Label8 = r - 1
                    Label8.Refresh
                'End If
            End If
    Next
    If Val(Trim(Label8)) = Val(Trim(Label7)) Then
        Me.MousePointer = vbDefault
        MsgBox "Data Import Completed!", vbInformation, "DAS Version 4.0"
    Else
        Me.MousePointer = vbDefault
        MsgBox "GCC Related Update Control Effective! Update Not Fully Done, Updated only : " & Val(Trim(Label8)) & " rows", vbInformation, "DAS Version 4.0"
    End If
'    appX.Close 0
'    appXL.Quit
    Set appXL = Nothing
    Set appX = Nothing
    Unload Me
Else
    appX.Close 0
'    appXL.Quit
    Set appXL = Nothing
    Set appX = Nothing
    MsgBox "Invalid WorkSheet,XLS File Selected for Retailers Import!", vbExclamation, "DAS Version 4.0"
    clrall
    Command2.Enabled = True
End If
End Sub
Public Sub act()
If CHECKTB("act.dbf,rdy.dbf,loc.dbf,smn.dbf,mar.dbf,cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Command3.SetFocus
    Exit Sub
End If
CheckGCCStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'GCC'")
If Trim(CheckGCCStatus) = "T" Then
    MsgBox "GCC Enabled! Cannot Import New Retail Outlets! ", vbInformation, "DAS Version 4.0"
    Command3.SetFocus
    Exit Sub
End If
If MsgBox("Import from " & txtFileName, vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then Exit Sub
Command2.Enabled = False
If rsgrp.State = 1 Then rsgrp.Close
rsgrp.Open "select * from grp where grp_IDY like '" & Trim(GRP_IDY) & "'", con, adOpenKeyset, adLockPessimistic
C = 0
If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "GST" And UCase(Trim(appX.Sheets(1).Cells(1, 3))) = "CST" And UCase(Trim(appX.Sheets(1).Cells(1, 4))) = "ADDRESS1" And UCase(Trim(appX.Sheets(1).Cells(1, 5))) = "ADDRESS2" And UCase(Trim(appX.Sheets(1).Cells(1, 6))) = "ADDRESS3" And UCase(Trim(appX.Sheets(1).Cells(1, 7))) = "ADDRESS4" And UCase(Trim(appX.Sheets(1).Cells(1, 8))) = "ZIPCODE" And UCase(Trim(appX.Sheets(1).Cells(1, 9))) = "TELEPHONE" And UCase(Trim(appX.Sheets(1).Cells(1, 10))) = "CONTACT" And UCase(Trim(appX.Sheets(1).Cells(1, 11))) = "LOCATION" And UCase(Trim(appX.Sheets(1).Cells(1, 12))) = "MARKET" And UCase(Trim(appX.Sheets(1).Cells(1, 13))) = "SALESMAN" And UCase(Trim(appX.Sheets(1).Cells(1, 14))) = "CATEGORY" And UCase(Trim(appX.Sheets(1).Cells(1, 15))) = "ROUTE" And UCase(Trim(appX.Sheets(1).Cells(1, 16))) = "GRADE" And UCase(Trim(appX.Sheets(1).Cells(1, 17))) = "ACTIVE" And UCase(Trim(appX.Sheets(1).Cells(1, 18))) = "CAPACITY" Then
    If Trim(GRP_IDY) = "" Then
        MsgBox "Select Group Account by Pressing Enter Key, SUNDRY DEBTORS - RETAILERS or Press 7", vbInformation, "DAS Version 4.0"
        GRP_IDY.SetFocus
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    If rsact.State = 1 Then rsact.Close
    rsact.Open "SELECT * FROM ACT", con, adOpenKeyset, adLockPessimistic
    rsrdy.Open "select * from rdy where act_idy=''", con, adOpenKeyset, adLockPessimistic
    appX.Sheets(1).Cells(1, 29).Formula = "=counta(a1:a65536)"
    MaxRows = appX.Sheets(1).Cells(1, 29)
    Label7 = MaxRows - 1
    For r = 2 To MaxRows + 1
    SHOWBUT (True)
    If rsloc.State = 1 Then rsloc.Close
    If rsmar.State = 1 Then rsmar.Close
    If RSSMN.State = 1 Then RSSMN.Close
    If RSCAT.State = 1 Then RSCAT.Close
        If Trim(appX.Sheets(1).Cells(r, 1)) = "" Then
            Exit For
        End If
        rsact.AddNew
        For i = 0 To rsact.Fields.count - 1
            If rsact(i).Type = adVarChar Or rsact(i).Type = adChar Or rsact(i).Type = adBSTR Or rsact(i).Type = adLongVarChar Or rsact(i).Type = adLongVarWChar Then
                rsact(i) = " "
            ElseIf rsact(i).Type = adInteger Or rsact(i).Type = adNumeric Or rsact(i).Type = adBigInt Or rsact(i).Type = adSingle Or rsact(i).Type = adDouble Or rsact(i).Type = adBigInt Or rsact(i).Type = adCurrency Or rsact(i).Type = adNumeric Or rsact(i).Type = adSmallInt Or rsact(i).Type = adDecimal Then
                rsact(i) = 0
            ElseIf rsact(i).Type = adDate Or rsact(i).Type = adDBDate Or rsact(i).Type = adDBTimeStamp Then
                rsact(i) = Date
            End If
        Next
        rsact("act_idy") = increment("IDY_ACT")
        rsact("grp_clf") = rsgrp("GRP_CLF")
        rsact("grp_idy") = rsgrp("GRP_IDY")
        rsact("act_nme") = UCase(ADS(rsact, "ACT_NME", appX.Sheets(1).Cells(r, 1)))
        rsact("act_gst") = UCase(ADS(rsact, "act_gst", appX.Sheets(1).Cells(r, 2)))
        rsact("act_cst") = UCase(ADS(rsact, "act_cst", appX.Sheets(1).Cells(r, 3)))
        rsact("act_ad1") = UCase(ADS(rsact, "act_ad1", appX.Sheets(1).Cells(r, 4)))
        rsact("act_ad2") = UCase(ADS(rsact, "act_ad2", appX.Sheets(1).Cells(r, 5)))
        rsact("act_ad3") = UCase(ADS(rsact, "act_ad3", appX.Sheets(1).Cells(r, 6)))
        rsact("act_ad4") = UCase(ADS(rsact, "act_ad4", appX.Sheets(1).Cells(r, 7)))
        rsact("act_zip") = ADS(rsact, "act_zip", appX.Sheets(1).Cells(r, 8))
        rsact("act_tfx") = ADS(rsact, "act_tfx", appX.Sheets(1).Cells(r, 9))
        rsact("act_cpn") = UCase(ADS(rsact, "act_cpn", appX.Sheets(1).Cells(r, 10)))
        rsact("ret_typ") = UCase(ADS(rsact, "ret_typ", appX.Sheets(1).Cells(r, 16)))
        rsact("sku_rng") = 0#
        If ADS(rsact, "dbt_typ", appX.Sheets(1).Cells(r, 17)) = "" Then
            rsact("dbt_typ") = "Y"
        Else
            rsact("dbt_typ") = UCase(ADS(rsact, "dbt_typ", appX.Sheets(1).Cells(r, 17)))
        End If
        rsact("van_cap") = Val(UCase(ADS(rsact, "van_cap", appX.Sheets(1).Cells(r, 18))))
        rsact("acn_two") = UCase(ADS(rsact, "acn_two", appX.Sheets(1).Cells(r, 19)))
        rsact("bnm_two") = UCase(ADS(rsact, "bnm_two", appX.Sheets(1).Cells(r, 20)))
        If ADS(rsact, "mcl_req", appX.Sheets(1).Cells(r, 22)) = "" Then
            rsact("mcl_req") = "Y"
        ElseIf ADS(rsact, "mcl_req", appX.Sheets(1).Cells(r, 22)) = "N" Then
            rsact("mcl_req") = "N"
        Else
            rsact("mcl_req") = UCase(ADS(rsact, "mcl_req", appX.Sheets(1).Cells(r, 22)))
        End If
        'rsact("cst_typ") = UCase(ADS(rsact, "cst_typ", appX.Sheets(1).cells(r, 21)))
        If (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "Retailer" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "AMC" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "SubDistributor" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "Key A/c" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "Local Key A/c" Or (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "Regnl Key A/c" Or UCase(ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "HVO-P" Or UCase(ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21))) = "HVO-NP" Or ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21)) = "Van" Then
            rsact("cst_typ") = (ADS(rsact, "cst_typ", appX.Sheets(1).Cells(r, 21)))
        Else
            rsact("cst_typ") = "Retailer"
        End If
        If Trim(appX.Sheets(1).Cells(r, 11)) <> "" Then
            If rsloc.State = 1 Then rsloc.Close
            rsloc.Open "SELECT * FROM LOC WHERE TRIM(LOC_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 11))) & "'", con, adOpenKeyset, adLockPessimistic
            If rsloc.EOF = False Then
                rsact("LOC_IDY") = rsloc("LOC_IDY")
            Else
                rsloc.AddNew
                rsloc("LOC_IDY") = increment("IDY_LOC")
                rsloc("LOC_NME") = UCase(ADS(rsloc, "loc_nme", appX.Sheets(1).Cells(r, 11)))
                rsloc("LOC_TYP") = ""
                rsact("LOC_IDY") = rsloc("LOC_IDY")
                rsloc.Update
            End If
        Else
            rsact("LOC_IDY") = Trim(AREAID)
        End If
        If Trim(appX.Sheets(1).Cells(r, 13)) <> "" Then
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT * FROM SMN WHERE TRIM(SMN_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 13))) & "'", con, adOpenKeyset, adLockPessimistic
            If RSSMN.EOF = False Then
                rsact("SMN_IDY") = RSSMN("SMN_IDY")
            Else
                RSSMN.AddNew
                For i = 0 To RSSMN.Fields.count - 1
                If RSSMN(i).Type = adVarChar Or RSSMN(i).Type = adChar Or RSSMN(i).Type = adBSTR Or RSSMN(i).Type = adLongVarChar Or RSSMN(i).Type = adLongVarWChar Then
                    RSSMN(i) = ""
                ElseIf RSSMN(i).Type = adInteger Or RSSMN(i).Type = adNumeric Or RSSMN(i).Type = adBigInt Or RSSMN(i).Type = adSingle Or RSSMN(i).Type = adDouble Or RSSMN(i).Type = adBigInt Or RSSMN(i).Type = adCurrency Or RSSMN(i).Type = adNumeric Or RSSMN(i).Type = adSmallInt Or RSSMN(i).Type = adDecimal Then
                    RSSMN(i) = 0
                ElseIf RSSMN(i).Type = adDate Or RSSMN(i).Type = adDBDate Or RSSMN(i).Type = adDBTimeStamp Then
                    RSSMN(i) = Date
                End If
                Next
                RSSMN("SMN_IDY") = increment("IDY_SMN")
                RSSMN("SMN_NME") = UCase(ADS(RSSMN, "smn_nme", appX.Sheets(1).Cells(r, 13)))
                rsact("SMN_IDY") = RSSMN("SMN_IDY")
                RSSMN.Update
            End If
        Else
            rsact("SMN_IDY") = Trim(Text32)
        End If
        If Trim(appX.Sheets(1).Cells(r, 12)) <> "" Then
            If rsmar.State = 1 Then rsmar.Close
            rsmar.Open "SELECT * FROM MAR WHERE TRIM(MAR_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 12))) & "'", con, adOpenKeyset, adLockPessimistic
            If rsmar.EOF = False Then
                rsact("MAR_IDY") = rsmar("MAR_IDY")
            Else
                rsmar.AddNew
                For i = 0 To rsmar.Fields.count - 1
                If rsmar(i).Type = adVarChar Or rsmar(i).Type = adChar Or rsmar(i).Type = adBSTR Or rsmar(i).Type = adLongVarChar Or rsmar(i).Type = adLongVarWChar Then
                    rsmar(i) = ""
                ElseIf rsmar(i).Type = adInteger Or rsmar(i).Type = adNumeric Or rsmar(i).Type = adBigInt Or rsmar(i).Type = adSingle Or rsmar(i).Type = adDouble Or rsmar(i).Type = adBigInt Or rsmar(i).Type = adCurrency Or rsmar(i).Type = adNumeric Or rsmar(i).Type = adSmallInt Or rsmar(i).Type = adDecimal Then
                    rsmar(i) = 0
                ElseIf rsmar(i).Type = adDate Or rsmar(i).Type = adDBDate Or rsmar(i).Type = adDBTimeStamp Then
                    rsmar(i) = Date
                End If
                Next
                rsmar("MAR_IDY") = increment("IDY_MAR")
                rsmar("MAR_NME") = UCase(ADS(rsmar, "mar_nme", appX.Sheets(1).Cells(r, 12)))
                rsmar("SMN_IDY") = rsact("SMN_IDY")
                rsmar("ROU_IDY") = "B"
                rsact("MAR_IDY") = rsmar("MAR_IDY")
                ' Added for Margin Minder Data Extract by PNR
                rsmar("MAR_CSG") = UCase(ADS(rsmar, "MAR_CSG", appX.Sheets(1).Cells(r, 30)))
                rsmar("MAR_CLU") = UCase(ADS(rsmar, "MAR_CLU", appX.Sheets(1).Cells(r, 31)))
                rsmar("MAR_ROC") = UCase(ADS(rsmar, "MAR_ROC", appX.Sheets(1).Cells(r, 32)))
                rsmar("MAR_TWC") = UCase(ADS(rsmar, "MAR_TWC", appX.Sheets(1).Cells(r, 33)))
                rsmar("MAR_TWN") = UCase(ADS(rsmar, "MAR_TWN", appX.Sheets(1).Cells(r, 34)))
                rsmar("MAR_DST") = UCase(ADS(rsmar, "MAR_DST", appX.Sheets(1).Cells(r, 35)))
                rsmar("MAR_STA") = UCase(ADS(rsmar, "MAR_STA", appX.Sheets(1).Cells(r, 36)))
                rsmar("MAR_COU") = UCase(ADS(rsmar, "MAR_COU", appX.Sheets(1).Cells(r, 37)))
                rsmar.Update
            End If
        Else
            rsact("mar_idy") = Trim(marId)
        End If
        If Trim(appX.Sheets(1).Cells(r, 14)) <> "" Then
            If RSCAT.State = 1 Then RSCAT.Close
            RSCAT.Open "SELECT * FROM CAT WHERE TRIM(CAT_NME) LIKE '" & UCase(Trim(appX.Sheets(1).Cells(r, 14))) & "'", con, adOpenKeyset, adLockPessimistic
            If RSCAT.EOF = False Then
                rsact("CAT_IDY") = RSCAT("CAT_IDY")
            Else
                RSCAT.AddNew
                For i = 0 To RSCAT.Fields.count - 1
                If RSCAT(i).Type = adVarChar Or RSCAT(i).Type = adChar Or RSCAT(i).Type = adBSTR Or RSCAT(i).Type = adLongVarChar Or RSCAT(i).Type = adLongVarWChar Then
                    RSCAT(i) = " "
                ElseIf RSCAT(i).Type = adInteger Or RSCAT(i).Type = adNumeric Or RSCAT(i).Type = adBigInt Or RSCAT(i).Type = adSingle Or RSCAT(i).Type = adDouble Or RSCAT(i).Type = adBigInt Or RSCAT(i).Type = adCurrency Or RSCAT(i).Type = adNumeric Or RSCAT(i).Type = adSmallInt Or RSCAT(i).Type = adDecimal Then
                    RSCAT(i) = 0
                ElseIf RSCAT(i).Type = adDate Or RSCAT(i).Type = adDBDate Or RSCAT(i).Type = adDBTimeStamp Then
                    RSCAT(i) = Date
                End If
                Next
                RSCAT("CAT_IDY") = increment("IDY_CAT")
                RSCAT("CAT_NME") = UCase(ADS(RSCAT, "cat_nme", appX.Sheets(1).Cells(r, 14)))
                rsact("CAT_IDY") = RSCAT("CAT_IDY")
                RSCAT.Update
            End If
        Else
            rsact("cat_idy") = Trim(CATID)
        End If
        If UCase(appX.Sheets(1).Cells(r, 15)) <> "" Then
            rsact("rou_idy") = UCase(ADS(rsact, "rou_idy", UCase(appX.Sheets(1).Cells(r, 15))))
        Else
            rsact("rou_idy") = Trim(ROU_LST & "")
        End If
'        rsact("dbt_typ") = "Y" - Active - Inactive Status captured from Xcel only
        rsact("key_act") = "N"
        rsact("dbt_rap") = "B"
        rsact("act_typ") = "SAL"
        rsact.Update
        If UCase(Trim(appX.Sheets(1).Cells(r, 23))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 24))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 25))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 26))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 27))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 28))) = "" And UCase(Trim(appX.Sheets(1).Cells(r, 29))) = "" Then
            For d = 1 To 7
                rsrdy.AddNew
                For K = 0 To rsrdy.Fields.count - 1
                    If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                        rsrdy(K) = " "
                    ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                        rsrdy(K) = 0
                    ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                        rsrdy(K) = Date
                    End If
                Next
                rsrdy("act_idy") = rsact("act_idy")
                rsrdy("day_idy") = d
                rsrdy("smn_idy") = rsact("smn_idy")
                rsrdy("rou_idy") = Trim(rsact("rou_idy") & "")
            Next
        End If
        ' Beat Plan Logic incorporated by PNR from EXCEL sheet - not compulsory, if available, RDY table is updated
        If UCase(Trim(appX.Sheets(1).Cells(r, 23))) = "MONDAY" Then
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("ACT_IDY") = rsact("ACT_IDY")
            rsrdy("DAY_IDY") = 1
            rsrdy("SMN_IDY") = rsact("SMN_IDY")
            rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
        End If
        If UCase(Trim(appX.Sheets(1).Cells(r, 24))) = "TUESDAY" Then
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("ACT_IDY") = rsact("ACT_IDY")
            rsrdy("DAY_IDY") = 2
            rsrdy("SMN_IDY") = rsact("SMN_IDY")
            rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
        End If
        If UCase(Trim(appX.Sheets(1).Cells(r, 25))) = "WEDNESDAY" Then
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("ACT_IDY") = rsact("ACT_IDY")
            rsrdy("DAY_IDY") = 3
            rsrdy("SMN_IDY") = rsact("SMN_IDY")
            rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
        End If
        If UCase(Trim(appX.Sheets(1).Cells(r, 26))) = "THURSDAY" Then
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("ACT_IDY") = rsact("ACT_IDY")
            rsrdy("DAY_IDY") = 4
            rsrdy("SMN_IDY") = rsact("SMN_IDY")
            rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
        End If
        If UCase(Trim(appX.Sheets(1).Cells(r, 27))) = "FRIDAY" Then
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("ACT_IDY") = rsact("ACT_IDY")
            rsrdy("DAY_IDY") = 5
            rsrdy("SMN_IDY") = rsact("SMN_IDY")
            rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
        End If
        If UCase(Trim(appX.Sheets(1).Cells(r, 28))) = "SATURDAY" Then
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("ACT_IDY") = rsact("ACT_IDY")
            rsrdy("DAY_IDY") = 6
            rsrdy("SMN_IDY") = rsact("SMN_IDY")
            rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
        End If
        If UCase(Trim(appX.Sheets(1).Cells(r, 29))) = "SUNDAY" Then
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("ACT_IDY") = rsact("ACT_IDY")
            rsrdy("DAY_IDY") = 7
            rsrdy("SMN_IDY") = rsact("SMN_IDY")
            rsrdy("ROU_IDY") = Trim(rsact("ROU_IDY") & "")
        End If
        C = C + 1
        Label8 = r - 1
        Label8.Refresh
    Next
    Me.MousePointer = vbDefault
    MsgBox "Data Import Completed!", vbExclamation, "DAS Version 4.0"
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    Set appX = Nothing
    Unload Me
Else
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    MsgBox "Invalid WorkSheet,XLS File Selected for Retailers Import!", vbExclamation, "DAS Version 4.0"
    clrall
    Command2.Enabled = True
End If
End Sub
Public Sub pgr()
If MsgBox("Import from " & txtFileName, vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then Exit Sub
If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "CATEGORIES" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "SEGMENT" Then
    Me.MousePointer = vbHourglass
    If rspgr.State = 1 Then rspgr.Close
    rspgr.Open "select * from pgr", con, adOpenKeyset, adLockPessimistic
    appX.Sheets(1).Cells(1, 27).Formula = "=counta(a1:a65536)"
    MaxRows = appX.Sheets(1).Cells(1, 27)
    Label7 = MaxRows - 1
    SHOWBUT (True)
    For r = 2 To MaxRows + 1
        If Trim(appX.Sheets(1).Cells(r, 1)) = "" Then
            Exit For
        End If
        rspgr.AddNew
        rspgr("Grp_idy") = increment("idy_Pgr")
        rspgr("grp_cat") = UCase(ADS(rspgr, "grp_cat", Trim(appX.Sheets(1).Cells(r, 1))))
        rspgr("grp_cls") = UCase(ADS(rspgr, "grp_cls", Trim(appX.Sheets(1).Cells(r, 2))))
        rspgr.Update
        Label8 = r - 1
        Label8.Refresh
    Next
    Me.MousePointer = vbDefault
    MsgBox "Import Completed", vbExclamation
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    Set appX = Nothing
    Unload Me
Else
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    MsgBox "Invalid File Selected for Product,SKU Categories Import!", vbExclamation, "DAS Version 4.0"
    clrall
    Command2.Enabled = True
End If
End Sub
Public Sub prd()
DoEvents
' Check for TWP Version NO IMPORT
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    If Trim(DIR_RS("DIR_ALS")) = "Y" Then
        If Trim(DIR_RS("DIR_MSG")) <> "WRITETWP" Then
            MsgBox ("Cannot Import Product,SKUs Data for KO-TWP Version!"), vbCritical, "DAS Version 4.0"
            clrall
            Exit Sub
        End If
    End If
End If

If MsgBox("Import from " & txtFileName, vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then Exit Sub
If Trim(appX.Sheets(1).Cells(2, 6)) = "" And Trim(appX.Sheets(1).Cells(2, 7)) = "" Then
    If Trim(grpid) = "" Then
        MsgBox "Select Product,SKU Group - Category - Segment!", vbExclamation, "DAS Version 4.0"
        grpid.SetFocus
        Exit Sub
    End If
End If
If rspgr.State = 1 Then rspgr.Close
If rsact.State = 1 Then rsact.Close
If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "SKU NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "COMPANY ID" And UCase(Trim(appX.Sheets(1).Cells(1, 3))) = "SHORT NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 4))) = "BRAND NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 5))) = "GROUP ID" And UCase(Trim(appX.Sheets(1).Cells(1, 6))) = "CATEGORY" And UCase(Trim(appX.Sheets(1).Cells(1, 7))) = "SEGMENT" And UCase(Trim(appX.Sheets(1).Cells(1, 8))) = "COMPANY NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 9))) = "PUR RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 10))) = "SPL RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 11))) = "SAL RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 12))) = "MRP RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 13))) = "CON FACTOR" And UCase(Trim(appX.Sheets(1).Cells(1, 14))) = "SUB UNITS" Then
Me.MousePointer = vbHourglass
    If rsprd.State = 1 Then rsprd.Close
    If RSBRD.State = 1 Then RSBRD.Close
    rsprd.Open "select * from Prd", con, adOpenKeyset, adLockPessimistic
    RSBRD.Open "select * from Brd", con, adOpenKeyset, adLockPessimistic
    appX.Sheets(1).Cells(1, 27).Formula = "=counta(a1:a65536)"
    MaxRows = appX.Sheets(1).Cells(1, 27)
    Label7 = MaxRows - 1
    SHOWBUT (True)
    For r = 2 To MaxRows + 1
        If Trim(appX.Sheets(1).Cells(r, 1)) = "" Then
            Exit For
        End If
        rsprd.AddNew
        For i = 0 To rsprd.Fields.count - 1
            If rsprd(i).Type = adVarChar Or rsprd(i).Type = adChar Or rsprd(i).Type = adBSTR Or rsprd(i).Type = adLongVarChar Or rsprd(i).Type = adLongVarWChar Then
                rsprd(i) = ""
            ElseIf rsprd(i).Type = adInteger Or rsprd(i).Type = adNumeric Or rsprd(i).Type = adBigInt Or rsprd(i).Type = adSingle Or rsprd(i).Type = adDouble Or rsprd(i).Type = adBigInt Or rsprd(i).Type = adCurrency Or rsprd(i).Type = adNumeric Or rsprd(i).Type = adSmallInt Or rsprd(i).Type = adDecimal Then
                rsprd(i) = 0
            ElseIf rsprd(i).Type = adDate Or rsprd(i).Type = adDBDate Or rsprd(i).Type = adDBTimeStamp Then
                rsprd(i) = Date
            End If
        Next
        rsprd("prd_idy") = increment("idy_Prd")
        rsprd("prd_nme") = UCase(ADS(rsprd, "prd_nme", Trim(appX.Sheets(1).Cells(r, 1))))
        rsprd("alt_idy") = UCase(ADS(rsprd, "alt_idy", Trim(appX.Sheets(1).Cells(r, 2))))
        rsprd("stk_typ") = "Salable"
        rsprd("prd_prn") = UCase(ADS(rsprd, "prd_prn", Trim(appX.Sheets(1).Cells(r, 3))))
        rsprd("brd_nme") = UCase(ADS(rsprd, "brd_nme", Trim(appX.Sheets(1).Cells(r, 4))))
'       RSPRD("grp_idy") = rspgr("GRP_IDY") 'Trim(appX.Sheets(1).Cells(R, 6))
        rsprd("cmp_idy") = Trim(Text3) & "" 'RSACT("ACT_IDY") 'Trim(appX.Sheets(1).Cells(R, 7))
        rsprd("pur_rat") = ADS(rsprd, "pur_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 9)))))
        rsprd("bas_rat") = ADS(rsprd, "bas_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 10)))))
        rsprd("sal_rat") = ADS(rsprd, "sal_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 11)))))
        rsprd("mrp_rat") = ADS(rsprd, "mrp_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 12)))))
        rsprd("con_fac") = ADS(rsprd, "con_fac", Trim(appX.Sheets(1).Cells(r, 13)))
        'RSPRD("prd_uom") = Trim(appX.Sheets(1).Cells(R, 18))
        rsprd("sub_unt") = ADS(rsprd, "sub_unt", Trim(appX.Sheets(1).Cells(r, 14)))
        If Trim(appX.Sheets(1).Cells(r, 6)) <> "" And Trim(appX.Sheets(1).Cells(r, 7)) <> "" Then
            If rspgr.State = 1 Then rspgr.Close
            rspgr.Open "SELECT * FROM PGR WHERE grp_cat ='" & Trim(appX.Sheets(1).Cells(r, 6)) & "' and grp_cls ='" & Trim(appX.Sheets(1).Cells(r, 7)) & "'", con, adOpenKeyset, adLockPessimistic
            If rspgr.EOF = False Then
                rsprd("GRP_IDY") = rspgr("GRP_IDY")
            Else
                rspgr.AddNew
                For i = 0 To rspgr.Fields.count - 1
                If rspgr(i).Type = adVarChar Or rspgr(i).Type = adChar Or rspgr(i).Type = adBSTR Or rspgr(i).Type = adLongVarChar Or rspgr(i).Type = adLongVarWChar Then
                    rspgr(i) = ""
                ElseIf rspgr(i).Type = adInteger Or rspgr(i).Type = adNumeric Or rspgr(i).Type = adBigInt Or rspgr(i).Type = adSingle Or rspgr(i).Type = adDouble Or rspgr(i).Type = adBigInt Or rspgr(i).Type = adCurrency Or rspgr(i).Type = adNumeric Or rspgr(i).Type = adSmallInt Or rspgr(i).Type = adDecimal Then
                    rspgr(i) = 0
                ElseIf rspgr(i).Type = adDate Or rspgr(i).Type = adDBDate Or rspgr(i).Type = adDBTimeStamp Then
                    rspgr(i) = Date
                End If
                Next
                rspgr("GRP_IDY") = increment("IDY_PGR")
                rspgr("GRP_CAT") = UCase(ADS(rspgr, "GRP_CAT", Trim(appX.Sheets(1).Cells(r, 6))))
                rspgr("GRP_CLS") = UCase(ADS(rspgr, "GRP_CLS", Trim(appX.Sheets(1).Cells(r, 7))))
                rsprd("GRP_IDY") = rspgr("GRP_IDY")
                rspgr.Update
            End If
        Else
            rsprd("GRP_IDY") = Trim(grpid)
        End If
        rsprd.Update
        RSBRD.AddNew
        For i = 0 To RSBRD.Fields.count - 1
            If RSBRD(i).Type = adVarChar Or RSBRD(i).Type = adChar Or RSBRD(i).Type = adBSTR Or RSBRD(i).Type = adLongVarChar Or RSBRD(i).Type = adLongVarWChar Then
                RSBRD(i) = ""
            ElseIf RSBRD(i).Type = adInteger Or RSBRD(i).Type = adNumeric Or RSBRD(i).Type = adBigInt Or RSBRD(i).Type = adSingle Or RSBRD(i).Type = adDouble Or RSBRD(i).Type = adBigInt Or RSBRD(i).Type = adCurrency Or RSBRD(i).Type = adNumeric Or RSBRD(i).Type = adSmallInt Or RSBRD(i).Type = adDecimal Then
                RSBRD(i) = 0
            ElseIf RSBRD(i).Type = adDate Or RSBRD(i).Type = adDBDate Or RSBRD(i).Type = adDBTimeStamp Then
                RSBRD(i) = Date
            End If
        Next
        
        B_SUB_UNT = 0
        B_MRP_RAT = 0
        B_UNT_RAT = 0
        
        B_SUB_UNT = ADS(rsprd, "sub_unt", Trim(appX.Sheets(1).Cells(r, 14)))
        B_MRP_RAT = ADS(rsprd, "mrp_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 12)))))
        B_UNT_RAT = Val(B_MRP_RAT) / Val(B_SUB_UNT)
        
        RSBRD("bat_idy") = "BAT" & T7increment("idy_bat")
        RSBRD("prd_idy") = rsprd("PRD_IDY")
        RSBRD("PRD_PDR") = rsprd("PUR_RAT")
        RSBRD("PRD_SDR") = rsprd("SAL_RAT")
        RSBRD("PRD_MRP") = rsprd("MRP_RAT")
        RSBRD("BRD_DOM") = Date
        RSBRD("BRD_DOE") = Date
        RSBRD("bat_nme") = "MRP:" & Trim(str(B_UNT_RAT)) & " " & Format(Date, "YY") & Format(Date, "MM") & Format(Date, "dd")
        RSBRD.Update
        Label8 = r - 1
        Label8.Refresh
    Next
    Me.MousePointer = vbDefault
    MsgBox "Products, SKUs Import Complete!", vbExclamation, "DAS Version 4.0"
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    Set appX = Nothing
    Unload Me
Else
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    MsgBox "Invalid File Selected for Product,SKU Data Import", vbExclamation, "DAS Version 4.0"
    clrall
    Command2.Enabled = True
End If
End Sub
Public Sub SKU_RAT()
DoEvents
If RS_VAN.State = 1 Then RS_VAN.Close
RS_VAN.Open "SELECT * FROM VAN WHERE VAN_CLS = 'N' or VAN_CLS = 'S'", con, adOpenKeyset, adLockPessimistic
If RS_VAN.RecordCount > 0 Then
    MsgBox ("Van Challans are Open! Cannot Update Rates!"), vbInformation, "DAS Version 4.0"
    RS_VAN.Close
    Exit Sub
End If
If MsgBox("Update from " & txtFileName, vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then Exit Sub
''If Trim(appX.sheets(1).cells(2, 6)) = "" And Trim(appX.sheets(1).cells(2, 7)) = "" Then
''    If Trim(grpid) = "" Then
''        MsgBox "Select Product,SKU Group - Category - Segment!", vbExclamation, "DAS Version 4.0"
''        grpid.SetFocus
''        Exit Sub
''    End If
''End If
If rspgr.State = 1 Then rspgr.Close
If rsact.State = 1 Then rsact.Close
If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "SKU NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "SUPPLIER ID" And UCase(Trim(appX.Sheets(1).Cells(1, 3))) = "PUR RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 4))) = "PUR TAX" And UCase(Trim(appX.Sheets(1).Cells(1, 5))) = "SAL RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 6))) = "SAL TAX" And UCase(Trim(appX.Sheets(1).Cells(1, 7))) = "MRP RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 8))) = "CRATE RENTAL" And UCase(Trim(appX.Sheets(1).Cells(1, 9))) = "CRATE TAX" And UCase(Trim(appX.Sheets(1).Cells(1, 10))) = "LIFE" And UCase(Trim(appX.Sheets(1).Cells(1, 11))) = "SKU STATUS" And UCase(Trim(appX.Sheets(1).Cells(1, 12))) = "BASE RATE" Then
    Me.MousePointer = vbHourglass
    If rsprd.State = 1 Then rsprd.Close
    If RSBRD.State = 1 Then RSBRD.Close
    rsprd.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    RSBRD.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    appX.Sheets(1).Cells(1, 27).Formula = "=counta(a1:a65536)"
    MaxRows = appX.Sheets(1).Cells(1, 27)
    Label7 = MaxRows - 1
    SHOWBUT (True)
    For r = 2 To MaxRows + 1
        If Trim(appX.Sheets(1).Cells(r, 1)) = "" Then
            Exit For
        End If
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "SELECT * FROM PRD where ALT_IDY = '" & Trim(appX.Sheets(1).Cells(r, 2)) & "' or PRD_PRN = '" & Trim(appX.Sheets(1).Cells(r, 2)) & "'", con, adOpenKeyset, adLockPessimistic
        If rsprd.RecordCount > 0 Then
            While Not rsprd.EOF()
                rsprd("prd_nme") = ADS(rsprd, "prd_nme", UCase(Trim(appX.Sheets(1).Cells(r, 1))))
                rsprd("lst_rat") = ADS(rsprd, "lst_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 8))))) 'Crate Rental
                rsprd("lst_dis") = ADS(rsprd, "lst_dis", convert(amt(Trim(appX.Sheets(1).Cells(r, 9))))) 'Crate Rental VAT
                rsprd("pur_rat") = ADS(rsprd, "pur_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 3)))))
                rsprd("pur_tax") = ADS(rsprd, "pur_tax", convert(amt(Trim(appX.Sheets(1).Cells(r, 4)))))
                rsprd("sal_rat") = ADS(rsprd, "sal_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 5)))))
                rsprd("sal_tax") = ADS(rsprd, "sal_tax", convert(amt(Trim(appX.Sheets(1).Cells(r, 6)))))
                rsprd("mrp_rat") = ADS(rsprd, "mrp_rat", convert(amt(Trim(appX.Sheets(1).Cells(r, 7)))))
                rsprd("tod_pcg") = ADS(rsprd, "tod_pcg", Val(Trim(appX.Sheets(1).Cells(r, 10))))
                rsprd("bas_rat") = ADS(rsprd, "bas_rat", Val(Trim(appX.Sheets(1).Cells(r, 12))))            ' Base Rate
                rsprd("reo_lev") = ADS(rsprd, "reo_lev", Int(Val(Trim(appX.Sheets(1).Cells(r, 13)))))
                If Trim(appX.Sheets(1).Cells(r, 11)) = "Salable" Or Trim(appX.Sheets(1).Cells(r, 11)) = "Closed" Then
                    rsprd("STK_TYP") = ADS(rsprd, "stk_typ", Trim(appX.Sheets(1).Cells(r, 11)))          ' SKU Status ie, stock type - Salable or Closed
                Else
                    rsprd("STK_TYP") = "Salable"
                End If
                rsprd.Update
                If RSBRD.State = 1 Then RSBRD.Close
                    RSBRD.Open "SELECT * FROM BRD where PRD_IDY = '" & rsprd("PRD_IDY") & "'", con, adOpenKeyset, adLockPessimistic
                    If RSBRD.RecordCount > 0 Then
                    While Not RSBRD.EOF()
                        RSBRD("PRD_PDR") = ADS(RSBRD, "PRD_PDR", convert(amt(Trim(appX.Sheets(1).Cells(r, 3)))))
                        RSBRD("PRD_SDR") = ADS(RSBRD, "PRD_SDR", convert(amt(Trim(appX.Sheets(1).Cells(r, 5)))))
                        RSBRD("PRD_MRP") = ADS(RSBRD, "PRD_MRP", convert(amt(Trim(appX.Sheets(1).Cells(r, 7)))))
                        INC_CHK = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'INC'")
                        If Trim(INC_CHK) = "T" Then
                            RSBRD("P_INC_DIS") = ADS(RSBRD, "P_INC_DIS", convert(amt(Trim(appX.Sheets(1).Cells(r, 14))))) 'Purchase Incentive Update from worksheet for Maharashtra VAT
                            RSBRD("S_INC_DIS") = ADS(RSBRD, "S_INC_DIS", convert(amt(Trim(appX.Sheets(1).Cells(r, 15))))) 'Sale Incentive Update from worksheet for Maharashtra VAT
                        Else
                            RSBRD("P_INC_DIS") = 0#
                            RSBRD("S_INC_DIS") = 0#
                        End If
                        RSBRD.Update
                        RSBRD.MoveNext
                    Wend
                End If
                rsprd.MoveNext
            Wend
        End If
        Label8 = r - 1
        Label8.Refresh
    Next
    Me.MousePointer = vbDefault
    MsgBox "Products, SKU - Batches Rates, Taxes Update Complete!", vbExclamation, "DAS Version 4.0"
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    Set appX = Nothing
    Unload Me
Else
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    MsgBox "Invalid File Selected for Product,SKU - Batch Rates Data Import", vbExclamation, "DAS Version 4.0"
    clrall
    Command2.Enabled = True
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "opt_act" Then
        If MsgBox("Exit Export - Import File Data Module ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then Unload Me
        Set appXL = Nothing
        Set appX = Nothing
    Else
        Set appXL = Nothing
        Set appX = Nothing
        clrall
    End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo errh
Set appXL = CreateObject("Excel.Application")
Set appX = appXL.Workbooks.Add
Exit Sub
errh:
If err.Number = 429 Then
    MsgBox "Required External File Data Component Not Available!", vbCritical, "DAS Version 4.0"
    Unload M_FRM_EFD
    Exit Sub
ElseIf err.Number = 364 Then
    Exit Sub
End If
End Sub

Private Sub Form_Terminate()
rsact.Close
rsloc.Close
rsmar.Close
RSSMN.Close
RSCAT.Close
rspgr.Close
rsprd.Close
RSBRD.Close
rsgrp.Close
rsrdy.Close
Set rsact = Nothing
Set rsloc = Nothing
Set rsmar = Nothing
Set RSSMN = Nothing
Set RSCAT = Nothing
Set rspgr = Nothing
Set rsprd = Nothing
Set RSBRD = Nothing
Set rsgrp = Nothing
Set rsrdy = Nothing
End Sub

Private Sub GRP_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Group Account ID or Press <Enter> to get List of Group ID's"
End Sub
Private Sub GRP_IDY_KeyDown(K As Integer, s As Integer)
If CHECKTB("grp.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If K = 13 Or K = 40 Then
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
    rs.Open "select * from grp where grp_idy like '" & Sincrement(Trim(GRP_IDY)) & "'", con, adOpenKeyset
    If rs.RecordCount > 0 Then
        GRP_IDY = Sincrement(GRP_IDY)
        Label4 = Trim(rs("grp_nme"))
        AREAID.SetFocus
    Else
        grouplist
    End If
End If
End Sub
Private Sub grouplist()
If CHECKTB("grp.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from grp where grp_nme like '" & Trim(GRP_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    lv.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rs("grp_nme")))
        ls.ListSubItems.Add , , Trim(rs("grp_idy"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Group Account"
    lv.Visible = True
    lv.SetFocus
Else
   MsgBox "Group Account Not Found, Define Group Account !", vbExclamation, "DAS Version 4.0"
   SSTab1.Tab = 0
   GRP_IDY.SetFocus
End If
End Sub
Private Sub GRP_IDY_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub grpid_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
    If rs.State = 1 Then rs.Close
        rs.Open "select * from pgr where grp_idy like '" & Sincrement(grpid) & "'", con, adOpenStatic, adLockReadOnly
        If rs.RecordCount > 0 Then
        grpid = Sincrement(Trim(Text1))
        GRPCAT = Trim(rs("grp_cat"))
        GRPSEG = Trim(rs("grp_cls"))
        grpid.Enabled = False
        Text3.SetFocus
        MOD1 = True
        'Toolbar1.Buttons(12).Enabled = True
        Else
            prdgrplist
        End If
End If
End Sub

Private Sub opt_act_Click()
opt_act.Enabled = True
'opt_prd.Enabled = False
'opt_pgr.Enabled = False
'Frame1.Enabled = False
Frame3.Visible = False
Frame2.Visible = True
GRP_IDY.SetFocus
Command2.Enabled = True
opt_act.SetFocus
End Sub

Private Sub opt_act_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtFileName.SetFocus
End If
End Sub

Private Sub opt_pgr_Click()
opt_pgr.Enabled = True
Frame2.Visible = False
Frame3.Visible = False
Command2.Enabled = True
opt_pgr.SetFocus
End Sub

Private Sub opt_pgr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtFileName.SetFocus
End If
End Sub

Private Sub opt_prd_Click()
opt_prd.Enabled = True
Frame2.Visible = False
Frame3.Visible = True
grpid.SetFocus
Command2.Enabled = True
opt_prd.SetFocus
End Sub

Private Sub opt_prd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtFileName.SetFocus
End If
End Sub

Private Sub opt_tax_Click()
opt_tax.Enabled = True
Frame2.Visible = False
Frame3.Visible = False
Command2.Enabled = True
opt_tax.SetFocus
End Sub

Private Sub opt_tax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtFileName.SetFocus
End If
End Sub


Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsact As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If rs.State = 1 Then rs.Close
    If Len(Trim(Text3)) <> 0 Then
        rsact.Open "select * from ACT where ACT_idy like '" & Sincrement(Trim(Text3)) & "' AND (act_typ like 'PUR' OR act_typ like 'BOT')", con, adOpenStatic
        If rsact.RecordCount > 0 Then
            Text3 = Sincrement(GRP_IDY)
            Label12 = rsact("ACT_nme")
            If K = 38 Then
                Text3.SetFocus
                
                'Text2.SetFocus
            Else
                Text13.SetFocus
            End If
            Command2.SetFocus
        Else
            suplist
        End If
    Else
        suplist
    End If
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "c" Then
    CANCELRECORD
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
Else
    Set appXL = Nothing
    Set appX = Nothing
    Unload Me
End If
End Sub
Public Sub CANCELRECORD()
clrctr Me
Frame2.Visible = False
Frame3.Visible = False
Frame1.Enabled = True
opt_act.Enabled = True
opt_prd.Enabled = True
opt_pgr.Enabled = True
txtFileName.SetFocus
End Sub
Private Sub arealist()
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsloc As New ADODB.Recordset
If rsloc.State = 1 Then rsloc.Close
Dim ls As ListItem
rsloc.Open "select * from LOC where LOC_NME like '" & AREAID & "%'", con, adOpenKeyset, adLockPessimistic
If rsloc.RecordCount > 0 Then
    lv.ListItems.CLEAR
    For A = 1 To rsloc.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rsloc("LOC_nme")))
        ls.ListSubItems.Add , , Trim(rsloc("LOC_IDY"))
        rsloc.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    Label20 = "Select Location"
    lv.ColumnHeaders(1).Text = "Location Name"
    lv.ColumnHeaders(2).Text = "Location ID"
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Location Not Found, Define Location !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    AREAID.SetFocus
End If
End Sub
Private Sub LV_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If lv.SortOrder = lvwAscending Then
lv.SortOrder = lvwDescending
ElseIf lv.SortOrder = lvwDescending Then
lv.SortOrder = lvwAscending
End If
lv.SortKey = CH.Index - 1
End Sub

Private Sub LV_DblClick()
lv_KeyPress 13
End Sub

Private Sub lv_KeyPress(K As Integer)
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If K = 13 Then
    If lv.ListItems.count > 0 Then
        If Label20 = "Select Location" Then
            If lv.ListItems.count > 0 Then
                Dim lst As ListSubItem
                Set lst = lv.SelectedItem.ListSubItems.Item(1)
                If rs.State = 1 Then rs.Close
                rs.Open "select * from LOC where LOC_IDY like '" & Trim(lst.Text) & "'", con, adOpenStatic
                AREAID = rs("LOC_IDY")
                Label29 = rs("LOC_NME")
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                CATID.SetFocus
            End If
        ElseIf Label20 = "Select Market" Then
            If lv.ListItems.count > 0 Then
                Dim lst2 As ListSubItem
                Set lst2 = lv.SelectedItem.ListSubItems.Item(1)
                If rsmar.State = 1 Then rsmar.Close
                rsmar.Open "select * from mar where mar_idy like '" & Trim(lst2.Text) & "'", con, adOpenStatic
                If rsmar.RecordCount > 0 Then
                    marId = rsmar("mar_idy")
                    Label30 = rsmar("mar_nme")
                    If MOD1 = False Then
                        If RSSMN.State = 1 Then RSSMN.Close
                        RSSMN.Open "SELECT * FROM SMN WHERE SMN_IDY LIKE '" & Trim(rsmar("SMN_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                        If RSSMN.RecordCount > 0 Then Label31 = Trim(RSSMN("smn_nme"))
                        Text32 = Trim(rsmar!SMN_IDY & "")
                        rlst = Trim(rsmar!rou_idy & "")
                        For K = 0 To ROU_LST.ListCount - 1
                            If Trim(ROU_LST.List(K)) = Trim(rlst) Then
                                ROU_LST.ListIndex = K
                                Exit For
                            End If
                        Next
                   End If
                End If
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                Command2.SetFocus
            End If
        ElseIf Label20 = "Select Category" Then
            If lv.ListItems.count > 0 Then
                Dim lst3 As ListSubItem
                Set lst3 = lv.SelectedItem.ListSubItems.Item(1)
                                
                If RSCAT.State = 1 Then RSCAT.Close
                RSCAT.Open "select * from CAT where CAT_idy like '" & Trim(lst3.Text) & "'", con, adOpenStatic
                CATID = RSCAT("CAT_idy")
                Label19 = RSCAT("CAT_nme")
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                marId.SetFocus
            End If
        ElseIf Label20 = "Select Product Category" Then
            If lv.ListItems.count > 0 Then
                grpid = lv.SelectedItem.ListSubItems(1)
                GRPSEG = lv.SelectedItem.ListSubItems(2)
                GRPCAT = lv.SelectedItem
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                Text3.SetFocus
            End If
        ElseIf Label20 = "Select Supplier" Then
            If lv.ListItems.count > 0 Then
                Text3 = lv.SelectedItem.ListSubItems(1)
                Label12 = lv.SelectedItem
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                Command2.SetFocus
            End If
        ElseIf Label20 = "Select Group Account" Then
            If lv.ListItems.count > 0 Then
                GRP_IDY = lv.SelectedItem.ListSubItems(1)
                Label4 = lv.SelectedItem
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                AREAID.SetFocus
            End If
        End If
    End If
End If
End Sub

Private Sub list4_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub CATID_KeyDown(K As Integer, s As Integer)
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim RSCAT As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If RSCAT.State = 1 Then RSCAT.Close
        RSCAT.Open " select * from CAT where CAT_idy like '" & Sincrement(CATID) & "'", con, adOpenStatic
        If RSCAT.RecordCount > 0 Then
            CATID = Sincrement(CATID)
            Label19 = RSCAT("CAT_nme")
            If K = 38 Then
                AREAID.SetFocus
            Else
                marId.SetFocus
            End If
        Else
            chanlist
        End If
End If
End Sub

Private Sub catid_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub chanlist()
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim RSCAT As New ADODB.Recordset
If RSCAT.State = 1 Then rs.Close
RSCAT.Open "select * from CAT where CAT_nme like '" & CATID & "%'", con, adOpenKeyset, adLockPessimistic
If RSCAT.RecordCount > 0 Then
    lv.ListItems.CLEAR
    For A = 1 To RSCAT.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(RSCAT("CAT_nme")))
        ls.ListSubItems.Add , , Trim(RSCAT("CAT_idy"))
        RSCAT.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    Label20 = "Select Category"
    lv.ColumnHeaders(1).Text = "Category Name"
    lv.ColumnHeaders(2).Text = "Category ID"
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Category Not Found, Define Category !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    CATID.SetFocus
End If
End Sub

Private Sub MARID_KeyPress(K As Integer)
 If K > 96 And K < 123 Then K = K - 32
 If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub MARID_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf,mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsmar As New ADODB.Recordset
Dim RSSMN As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If rsmar.State = 1 Then rsmar.Close
        rsmar.Open " select * from mar where mar_idy like '" & Sincrement(marId) & "'", con, adOpenStatic
        If rsmar.RecordCount > 0 Then
            marId = Sincrement(marId)
            Label30 = rsmar("mar_nme")
            If MOD1 = False Then
                Text32 = Trim(rsmar!SMN_IDY)
                If RSSMN.State = 1 Then RSSMN.Close
                RSSMN.Open "SELECT SMN_NME FROM SMN WHERE SMN_IDY LIKE '" & Trim(rsmar("SMN_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If RSSMN.RecordCount > 0 Then Label31 = RSSMN("smn_nme")
                'ROU_LST = Trim(rsmar!ROU_IDY)
                rlst = Trim(rsmar!rou_idy & "")
                For K = 0 To ROU_LST.ListCount - 1
                    If Trim(ROU_LST.List(K)) = Trim(rlst) Then
                        ROU_LST.ListIndex = K
                        Exit For
                    End If
                Next
            Else
                
            End If
            If K = 38 Then
                CATID.SetFocus
            Else
                Text32.SetFocus
            End If
        Else
            marlist
        End If
End If
End Sub

Private Sub marlist()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsmar As New ADODB.Recordset
If rsmar.State = 1 Then rsmar.Close
rsmar.Open "select * from mar where mar_nme like '" & Trim(marId) & "%'", con, adOpenKeyset, adLockPessimistic
If rsmar.RecordCount > 0 Then
    lv.ListItems.CLEAR
    For A = 1 To rsmar.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rsmar("mar_nme")))
        ls.ListSubItems.Add , , Trim(rsmar("mar_idy"))
        rsmar.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    Label20 = "Select Market"
    lv.ColumnHeaders(1).Text = "Market Name"
    lv.ColumnHeaders(2).Text = "Market ID"
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    marId.SetFocus
End If
End Sub

Private Sub prdgrplist()
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rspgr As New ADODB.Recordset
If rspgr.State = 1 Then rspgr.Close
rspgr.Open "select * from pgr where grp_IDY like '" & Trim(grpid) & "%'", con, adOpenKeyset
If rspgr.RecordCount > 0 Then
    lv.ListItems.CLEAR
    lv.ColumnHeaders.CLEAR
    lv.ColumnHeaders.Add , , "Product Category"
    lv.ColumnHeaders.Add , , "ID"
    lv.ColumnHeaders.Add , , "Product Segment"
    For A = 1 To rspgr.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rspgr("grp_cat")))
        ls.ListSubItems.Add , , Trim(rspgr("grp_idy"))
        ls.ListSubItems.Add , , Trim(rspgr("GRP_CLS") & "")
        rspgr.MoveNext
    Next
    SSTab1.Tab = 1
    Label31 = "Select Product Category"
    lv.Visible = True
    lv.SetFocus
    Label20 = "Select Product Category"
Else
    MsgBox "Segment Not Found, Define Segment !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    grpid.Enabled = True
    grpid.SetFocus
End If
End Sub


Private Sub suplist()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsact As New ADODB.Recordset
If rsact.State = 1 Then rsact.Close
'If Check3.Value = 1 Then
'rs.Open "select * from ACT where ACT_nme like '" & Trim(Text3) & "%' and (act_typ like 'PUR' or act_typ like 'BOT')", CON, adOpenKeyset
'ElseIf Check4.Value = 1 Then
rsact.Open "select * from ACT where act_ad4 like '" & Trim(Text3) & "%' and (act_typ like 'PUR' OR act_typ like 'BOT')", con, adOpenKeyset
'End If
lv.ListItems.CLEAR
lv.ColumnHeaders.CLEAR
lv.ColumnHeaders.Add , , "Supplier Name"
lv.ColumnHeaders.Add , , "ID"
lv.ColumnHeaders.Add , , "Place"
Label20 = "Select Supplier"
If rsact.RecordCount > 0 Then
    For A = 1 To rsact.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rsact("ACT_nme")))
        ls.ListSubItems.Add , , Trim(rsact("ACT_idy"))
        ls.ListSubItems.Add , , Trim(rsact("act_ad4") & "")
        rsact.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Supplier"
    lv.Visible = True
    lv.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text3.Enabled = True
        Text3.SetFocus
    End If
End If
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.SetFocus
End If
End Sub
Private Sub SHOWBUT(b As Boolean)
Label5.Visible = b
Label6.Visible = b
Label7.Visible = b
Label8.Visible = b
End Sub
Private Function ADS(rs As ADODB.Recordset, FLD As String, DB As String)
If rs(FLD).DefinedSize < Len(Trim(DB)) Then
    ADS = Mid(DB, 1, rs(FLD).DefinedSize)
Else
    ADS = Trim(DB)
End If
End Function

Private Sub clrall()
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
opt_act.SetFocus
End Sub
Function MailSend(sSendTo As String, sSubject As String, sText As String, Optional sAttachFile As String) As Boolean
'Dim MAPISession1 As MAPISession
'Dim MAPIMessages1 As MAPIMessages
'    On Error GoTo ErrHandler
    With MAPISession1
        If .SessionID = 0 Then
            .DownLoadMail = False
            .LogonUI = True
            .SignOn
            .NewSession = True
            MAPIMessages1.SessionID = .SessionID
        End If
    End With
    With MAPIMessages1
        .Compose
        .RecipAddress = sSendTo
        .AddressResolveUI = True
        .ResolveName
        .MsgSubject = sSubject
        .MsgNoteText = sText
        If Len(sAttachFile) > 0 And Len(Dir$(sAttachFile)) > 0 Then
            .AttachmentPathName = sAttachFile
        Else
            .AttachmentCount = 0
        End If
        .Send False
    End With
    MailSend = True
    Exit Function
ErrHandler:
    Debug.Print err.Description
    MailSend = False
End Function
    '''''        End If
    '''''
    '''''
    '''''        rsact.AddNew
    '''''        For I = 0 To rsact.Fields.COUNT - 1
    '''''            If rsact(I).Type = adVarChar Or rsact(I).Type = adChar Or rsact(I).Type = adBSTR Or rsact(I).Type = adLongVarChar Or rsact(I).Type = adLongVarWChar Then
    '''''                rsact(I) = " "
    '''''            ElseIf rsact(I).Type = adInteger Or rsact(I).Type = adNumeric Or rsact(I).Type = adBigInt Or rsact(I).Type = adSingle Or rsact(I).Type = adDouble Or rsact(I).Type = adBigInt Or rsact(I).Type = adCurrency Or rsact(I).Type = adNumeric Or rsact(I).Type = adSmallInt Or rsact(I).Type = adDecimal Then
    '''''                rsact(I) = 0
    '''''            ElseIf rsact(I).Type = adDate Or rsact(I).Type = adDBDate Or rsact(I).Type = adDBTimeStamp Then
    '''''                rsact(I) = Date
    '''''            End If
    '''''        Next
    '        rsact("act_idy") = increment("IDY_ACT")
    '        rsact("grp_clf") = rsgrp("GRP_CLF")
    '        rsact("grp_idy") = rsgrp("GRP_IDY")

