VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_MAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Markets : Districts : Census Information"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "M_FRM_MAR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
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
            Picture         =   "M_FRM_MAR.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAR.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&Market : Districts : Census Information"
      TabPicture(0)   =   "M_FRM_MAR.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "M_SMN_NME"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "M_SMN_IDY"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ROU_LST"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "text4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TEXT7E"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "List3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "List2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TextDistrict"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_MAR.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).Control(2)=   "list4"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Geographic Information"
         Height          =   4455
         Left            =   3840
         TabIndex        =   32
         Top             =   480
         Width           =   4335
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   40
            Top             =   3840
            Width           =   2600
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   39
            Top             =   3360
            Width           =   2600
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   38
            Top             =   2880
            Width           =   2600
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   37
            Top             =   2400
            Width           =   2600
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   36
            Top             =   1920
            Width           =   2600
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   35
            Top             =   1440
            Width           =   2600
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   34
            Top             =   960
            Width           =   2600
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   33
            Top             =   480
            Width           =   2600
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Customer Group"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   3908
            Width           =   1140
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "RED Outlet Class"
            Height          =   195
            Left            =   150
            TabIndex        =   49
            Top             =   3428
            Width           =   1230
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "RSA Cluster"
            Height          =   195
            Left            =   525
            TabIndex        =   48
            Top             =   2948
            Width           =   855
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Town Class"
            Height          =   195
            Left            =   555
            TabIndex        =   47
            Top             =   1988
            Width           =   825
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Town"
            Height          =   195
            Left            =   975
            TabIndex        =   46
            Top             =   2468
            Width           =   405
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "District"
            Height          =   195
            Left            =   900
            TabIndex        =   45
            Top             =   1508
            Width           =   480
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "State"
            Height          =   195
            Left            =   1005
            TabIndex        =   44
            Top             =   1028
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            Height          =   195
            Left            =   840
            TabIndex        =   43
            Top             =   548
            Width           =   540
         End
      End
      Begin VB.TextBox TextDistrict 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6600
         MaxLength       =   15
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Population : Census Data"
         Height          =   2055
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   3465
         Begin VB.ComboBox Text6 
            Height          =   315
            Left            =   1455
            TabIndex        =   29
            Top             =   480
            Width           =   1560
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1455
            MaxLength       =   10
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
         End
         Begin VB.ComboBox TEXT7 
            Height          =   315
            Left            =   1455
            TabIndex        =   24
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Income Group"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   225
            TabIndex        =   28
            Top             =   1500
            Width           =   1005
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "District : State"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   540
            Width           =   990
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Population (Figs)"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   60
            TabIndex        =   26
            Top             =   1028
            Width           =   1170
         End
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "M_FRM_MAR.frx":3386
         Left            =   1560
         List            =   "M_FRM_MAR.frx":339F
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "M_FRM_MAR.frx":33E3
         Left            =   1560
         List            =   "M_FRM_MAR.frx":33FC
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6960
         MaxLength       =   5
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ComboBox TEXT7E 
         Height          =   315
         ItemData        =   "M_FRM_MAR.frx":3440
         Left            =   1800
         List            =   "M_FRM_MAR.frx":3442
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3780
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1850
         Width           =   1575
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3735
         Left            =   -74520
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1005
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   6588
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
            Text            =   "Salesman Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.ComboBox text4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6240
         TabIndex        =   11
         Text            =   "Van 1"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox ROU_LST 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "M_FRM_MAR.frx":3444
         Left            =   6240
         List            =   "M_FRM_MAR.frx":34B1
         TabIndex        =   9
         Top             =   4020
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox M_SMN_IDY 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   4560
         MaxLength       =   15
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3540
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   990
         Width           =   2025
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3735
         Left            =   -74520
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1005
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   6588
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
            Text            =   "Market Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Closing Day"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4680
         TabIndex        =   41
         Top             =   4560
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shandy Day"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   465
         TabIndex        =   31
         Top             =   2333
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Market Name"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   375
         TabIndex        =   22
         Top             =   1058
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Market ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   630
         TabIndex        =   21
         Top             =   653
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Day"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   495
         TabIndex        =   20
         Top             =   1493
         Width           =   840
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Market Distance"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   1918
         Width           =   1170
      End
      Begin VB.Label M_SMN_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5760
         TabIndex        =   12
         Top             =   3540
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route (Beat)"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   4995
         TabIndex        =   10
         Top             =   4080
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kms"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   1918
         Width           =   300
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Type"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4965
         TabIndex        =   6
         Top             =   3015
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Market"
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
         Left            =   -74520
         TabIndex        =   3
         Top             =   660
         Width           =   7905
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
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
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MAR_LST"
                  Text            =   "&Markets List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DST_LST"
                  Text            =   "&Districts List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "STA_LST"
                  Text            =   "&States - Districts List"
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
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Market ID "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Market Name"
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
      Height          =   240
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "M_FRM_MAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, idnew As Boolean
Private Sub CREATE()
MOD1 = False
Text1 = ""
save
End Sub

Public Sub delete()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Delete Market ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
    If check("mar", "mar_idy", Text1) = False Then
        MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        clrctr M_FRM_MAR
        Text1.Enabled = True
        Text1.SetFocus
        Exit Sub
    Else
        If deletecheck = False Then
            con.Execute "delete from mar where mar_idy like '" & Sincrement(Text1) & "'"
            SSTab1.Tab = 0
            Text1.Enabled = True
            Text1.SetFocus
        Else
            MsgBox "Cannot Delete, Transaction Exists !", vbExclamation, "DAS Version 4.0"
            Text2.SetFocus
            Exit Sub
        End If
    End If
    cancel
Else
    cancel
End If
End Sub

Public Function deletecheck() As Boolean
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Function
End If
If check("act", "mar_idy", Trim(Text1)) = False Then
    deletecheck = False
Else
    deletecheck = True
End If
End Function

Private Sub save()
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If CHECKTB("mar.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If Len(Trim(Text2)) = 0 Then
    MsgBox "Blank Market Name, Enter Name !", vbExclamation, "DAS Version 4.0"
    Text2.Enabled = True
    Text2.SetFocus
    Exit Sub
End If
If Len(Trim(Text6)) = 0 Then
    MsgBox "District - State Information Required!", vbExclamation, "DAS Version 4.0"
    Text6.Enabled = True
    Text6.SetFocus
    Exit Sub
End If
If Val(Text8) = 0 Then
    MsgBox "Population in Figs. Required!", vbExclamation, "DAS Version 4.0"
    Text8.Enabled = True
    Text8.SetFocus
    Exit Sub
End If
If Len(Trim(Text7)) = 0 Then
    MsgBox "Population Group Selection Required!", vbExclamation, "DAS Version 4.0"
    Text7.Enabled = True
    Text7.SetFocus
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If MOD1 = True Then
rs.Open "select * from mar where mar_idy like '" & Sincrement(Text1) & "'", con, adOpenKeyset, adLockPessimistic
Else
rs.Open "select * from mar", con, adOpenKeyset, adLockPessimistic
rs.AddNew
rs("mar_idy") = increment("idy_mar")
End If
rs("mar_nme") = Trim(Text2) & ""
rs("mar_vda") = Trim(List2) & ""
rs("mar_cld") = Trim(list3) & ""
rs("veh_typ") = Trim(Text4) & ""
rs("dis_wdp") = Val(Text5) & ""
'updated to add mar_dop,mar_pgr,Mar_grd,Mar_vda
rs("mar_dop") = Val(Text8) & ""
rs("mar_pgr") = Trim(Text6) & ""
rs("mar_grd") = Trim(Text7) & ""
rs!SMN_IDY = Trim(M_SMN_IDY)
rs!rou_idy = Trim(ROU_LST & "")
If RS1.State = 1 Then RS1.Close
RS1.Open "SELECT DIR_ALS FROM DIR WHERE DIR_IDY LIKE 'WDS'", con, adOpenKeyset, adLockPessimistic
If RS1.RecordCount > 0 Then rs("WDI_IDY") = Trim(RS1(0) & "")
If MOD1 = False Then MsgBox "New Market ID is " & rs("mar_idy"), vbExclamation, "DAS Version 4.0"
rs.Update
clrctr M_FRM_MAR
Text6.Text = ""
Text7.Text = ""
Text1.Enabled = True
idnew = False
SSTab1.Tab = 0
Text1.SetFocus
MOD1 = False
'Toolbar1.Buttons(12).Enabled = False
'Toolbar1.Buttons(13).Enabled = True
List2.Selected(0) = True
list3.Selected(0) = True
ROU_LST.Selected(0) = True
M_SMN_NME = ""
End Sub

Private Sub cancel()
clrctr M_FRM_MAR
Text6.Text = ""
Text7.Text = ""
Text1.Enabled = True
idnew = False
SSTab1.Tab = 0
Text1.SetFocus
MOD1 = False
'Toolbar1.Buttons(12).Enabled = False
'Toolbar1.Buttons(13).Enabled = True
M_SMN_NME = ""
List2.Selected(0) = True
list3.Selected(0) = True
ROU_LST.Selected(0) = True
End Sub

Private Sub newrec()
MOD1 = False
clrctr M_FRM_MAR
Text1.Enabled = False
SSTab1.Tab = 0
Text2.SetFocus
'Toolbar1.Buttons(12).Enabled = False
'Toolbar1.Buttons(13).Enabled = False
List2.Selected(0) = True
list3.Selected(0) = True
ROU_LST.Selected(0) = True
End Sub

Private Sub marlist()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from mar where mar_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("mar_nme")))
        ls.ListSubItems.Add , , Trim(rs("mar_idy"))
        'ls.ListSubItems.Add , , Trim(RS("mar_vda") & "")
        ls.ListSubItems.Add , , Trim(rs("mar_cld") & "")
        rs.MoveNext
    Next
    Label20 = "Select Market"
    Toolbar1.Enabled = True
    SSTab1.Tab = 1
    list1.Visible = True
    list1.ListItems(1).Selected = True 'Added by PNR
    list1.SetFocus
Else
    MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub exitfrm()
Unload Me
Set Form = Nothing
End Sub



Private Sub Form_Activate()
Text7.Text = ""
SSTab1.TabEnabled(1) = False
'SSTab1.TabEnabled(2) = False
SSTab1.Tab = 0
List2.Selected(0) = True
list3.Selected(0) = True
ROU_LST.Selected(0) = True
'Text1.Enabled = True
Text1.Locked = False
'Text1.Enabled = True
If Text1.Enabled = True Then
    Text1.SetFocus
Else
    If Text2.Enabled = True Then Text2.SetFocus
End If
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_MAR.ActiveControl Is TextBox Then M_FRM_MAR.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False Then
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        exitfrm
    ElseIf K = 67 And s = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And s = 2 Then
        delete
    ElseIf K = 66 And s = 2 Then
        'If Toolbar1.Buttons(13).Enabled = True Then BROWSE
    ElseIf K = 69 And s = 2 Then
        If M_FRM_MAR.ActiveControl.name = "Text1" Then marlist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        OpenReport CURDIR & "\lists.RP1", "Markets list", 0, "", 6
    ElseIf K = 82 And s = 2 Then
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        'SSTab1.TabEnabled(1) = False
        If CHECKTB("mar.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If
        If check("mar", "mar_idy", Trim(Text1)) = False Then
            Text1.Enabled = True
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Text2.SetFocus
        End If
    ElseIf list4.Visible = True Then
        list4.Visible = False
        SSTab1.Tab = 0
        M_SMN_IDY.SetFocus
    Else
        If M_FRM_MAR.ActiveControl.name = "Text1" Then
            Unload Me
            Set Form = Nothing
        ElseIf M_FRM_MAR.ActiveControl.name = "SSTab1" Then
        ElseIf M_FRM_MAR.ActiveControl.name = "MS" Then
            MS.HighLight = flexHighlightNever
            Dim rs As New ADODB.Recordset
            Text27.Visible = False
            List6.Visible = False
            MS.SetFocus
            M_FRM_MAR.MousePointer = vbHourglass
            For A = 1 To MS.Rows - 1
                If rs.State = 1 Then rs.Close
                rs.Open "select * from mar where mar_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'", con, adOpenKeyset, adLockPessimistic
                If rs.RecordCount > 0 Then
                    rs("wdi_idy") = Trim(MS.TextMatrix(A, 1))
                    rs("mar_nme") = Trim(MS.TextMatrix(A, 2))
                    rs("mar_vda") = Trim(MS.TextMatrix(A, 3))
                    rs("mar_cld") = Trim(MS.TextMatrix(A, 4))
                    rs("veh_typ") = Trim(MS.TextMatrix(A, 5))
                    rs("dis_wdp") = Val(MS.TextMatrix(A, 6) & "")
                 End If
                    rs.Update
                    rs.MoveNext
            Next
            M_FRM_MAR.MousePointer = vbDefault
            SSTab1.Tab = 0
            SSTab1.TabEnabled(2) = False
            Text1.SetFocus
        Else
                If Len(Trim(Text7)) > 15 Then
                    Text7 = ""
                    Text7.SetFocus
                Else
                    save
                End If
            
        End If
    End If
End If
End Sub

Private Sub Form_Load()

'checkdongle
'MOD1 = False
'idnew = False
''Toolbar1.Buttons(12).Enabled = False
'
'MS.ColWidth(0) = 1200 'mar id
'MS.ColWidth(1) = 1200 'wds id
'MS.ColWidth(2) = 3000 'mar name
'MS.ColWidth(3) = 1100 'visit day
'MS.ColWidth(4) = 1100 'closing day
'MS.ColWidth(5) = 1300 'veh type
'MS.ColWidth(6) = 1000 'distance
'
'MS.TextMatrix(0, 0) = "Market ID"
'MS.TextMatrix(0, 1) = "WD ID"
'MS.TextMatrix(0, 2) = "Market Name"
'MS.TextMatrix(0, 3) = "Visit Day"
'MS.TextMatrix(0, 4) = "Closing Day"
'MS.TextMatrix(0, 5) = "Vehicle"
'MS.ColAlignment(5) = 7
'MS.TextMatrix(0, 6) = "Kms"
'MS.ColAlignment(6) = 7
'SCl MS, 6
'Text1.Enabled = True
'Text1.SetFocus
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub List2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Visit Day"
End Sub

Private Sub List3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Closing Day"
End Sub

Private Sub list4_DblClick()
list4_KeyDown 13, 1
End Sub

Private Sub M_SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesman ID or Press <Enter> to get List of Salesman ID's"
SendKeys "{home}+{end}"
End Sub

Private Sub M_SMN_IDY_KeyPress(K As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
    Dim rs As New ADODB.Recordset
    rs.Open "select * from smn where smn_idy like '" & Sincrement(M_SMN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        M_SMN_IDY = Sincrement(M_SMN_IDY)
        M_SMN_NME = Trim(rs!SMN_NME)
        ROU_LST.SetFocus
    Else
        smnlist
    End If
ElseIf K = 38 Then
    Text5.SetFocus
End If
End Sub

Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from smn where smn_nme like '" & Trim(M_SMN_IDY) & "%'", con, adOpenStatic
list4.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
        ls.ListSubItems.Add , , Trim(rs("smn_ad4") & "")
        rs.MoveNext
    Next
    Label20 = "Select Salesman"
    SSTab1.Tab = 1
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K <> 27 Then
        If MS.Col = 6 Then
            If K > 47 And K < 58 Or K = 46 Or K = 8 Then
                Text27.MaxLength = 5
                Text27.Visible = True
                Text27.Top = MS.Top + MS.CellTop
                Text27.Left = MS.CellLeft + MS.Left
                Text27.Width = MS.CellWidth
                Text27.Height = MS.CellHeight
                Text27 = ""
                Text27.SetFocus
                If K <> 13 Then Text27 = Chr(K)
                Text27.SelStart = Len(Text27)
            End If
        ElseIf MS.Col = 3 Or MS.Col = 4 Then
            If K > 64 And K < 92 Then
                List6.Visible = True
                List6.Top = MS.Top + MS.CellTop
                List6.Left = MS.CellLeft + MS.Left
                List6.Width = MS.CellWidth
                List6.Height = MS.CellHeight
                If Trim(MS.Text) = "" Then
                    List6.Selected(0) = True
                Else
                    List6 = Trim(MS.Text)
                End If
                List6.SetFocus
            End If
        ElseIf MS.Col = 5 Then
            If K > 64 And K < 92 Or K = 8 Then
                Text27.MaxLength = 15
                Text27.Visible = True
                Text27.Top = MS.Top + MS.CellTop
                Text27.Left = MS.CellLeft + MS.Left
                Text27.Width = MS.CellWidth
                Text27.Height = MS.CellHeight
                Text27 = ""
                Text27.SetFocus
                If K <> 13 Then Text27 = Chr(K)
                Text27.SelStart = Len(Text27)
            End If
        End If
ElseIf K = 13 Then
If MS.Col <> 6 Then MS.Col = MS.Col + 1
Else
K = 0
End If
End Sub

Private Sub List6_KeyPress(K As Integer)
If K = 13 Then
    MS.Text = Trim(List6)
    List6.Visible = False
    MS.Col = MS.Col + 1
    MS.SetFocus
ElseIf K = 27 Then
    List6.Visible = False
    MS.Col = MS.Col + 1
    MS.SetFocus
End If
End Sub
Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub
Private Sub List4_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list4.SortOrder = lvwAscending Then
list4.SortOrder = lvwDescending
ElseIf list4.SortOrder = lvwDescending Then
list4.SortOrder = lvwAscending
End If
list4.SortKey = CH.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Market or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(K As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
    If list1.ListItems.count > 0 Then
        Dim lst As ListSubItem
        Set lst = list1.SelectedItem.ListSubItems.Item(1)
        If rs.State = 1 Then rs.Close
        rs.Open "select * from mar where mar_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
        Text1 = Trim(rs("mar_idy"))
        Text2 = Trim(rs("mar_nme"))
        List2 = Trim(rs("mar_vda") & "")
        list3 = Trim(rs("mar_cld") & "")
        Text4 = Trim(rs("veh_typ") & "")
        Text5 = Val(rs("dis_wdp") & "")
        Text9 = Trim(rs("MAR_COU") & "")
        Text10 = Trim(rs("MAR_STA") & "")
        Text11 = Trim(rs("MAR_DST") & "")
        Text12 = Trim(rs("MAR_TWC") & "")
        Text13 = Trim(rs("MAR_TWN") & "")
        Text14 = Trim(rs("MAR_CLU") & "")
        Text15 = Trim(rs("MAR_ROC") & "")
        Text16 = Trim(rs("MAR_CSG") & "")
        'updated to add mar_dop,mar_pgr,Mar_grd,Mar_vda
        Text8 = Trim(rs("mar_dop") & "")
        Text6 = Trim(rs("mar_pgr") & "")
        Text7 = Trim(rs("mar_grd") & "")
        ROU_LST = Trim(rs!rou_idy & "")
        M_SMN_IDY = Trim(rs!SMN_IDY & "")
        M_SMN_NME = GetValue("smn_nme", "smn where smn_idy like '" & Trim(rs!SMN_IDY) & "'")
        list1.Visible = False
        SSTab1.Tab = 0
        Text1.Enabled = False
        Text2.SetFocus
        MOD1 = True
        Toolbar1.Buttons(12).Enabled = True
        'Toolbar1.Buttons(13).Enabled = False
    End If
End If
End Sub
Private Sub list4_KeyDown(K As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
    If list4.ListItems.count > 0 Then
        M_SMN_IDY = list4.SelectedItem.ListSubItems(1)
        M_SMN_NME = list4.SelectedItem
        list4.Visible = False
        SSTab1.Tab = 0
        ROU_LST.SetFocus
    End If
End If
End Sub

Private Sub list1_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub List2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'save
    Text6.SetFocus
    
ElseIf KeyCode = 38 And Shift = 1 Then
    Text5.SetFocus
End If
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text5.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    Text2.SetFocus
End If
End Sub

Private Sub ROU_LST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Route"
End Sub

Private Sub ROU_LST_Keydown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
    save
ElseIf KeyAscii = 38 And s = 1 Then
    M_SMN_IDY.SetFocus
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'If MS.Visible = True Then MS.Visible = False
If SSTab1.Tab = 1 Then
'Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
list1.Visible = False
If SSTab1.Tab = 0 Then
    If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Market ID or Press <Enter> to get List of Market ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
Set Form = Nothing
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Market Name"
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If rs.State = 1 Then rs.Close
        rs.Open " select * from mar where mar_idy like '" & Sincrement(Text1) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text1 = Sincrement(Text1)
            Text2 = Trim(rs("mar_nme"))
            List2 = Trim(rs("mar_vda") & "")
            list3 = Trim(rs("mar_cld") & "")
            Text4 = Trim(rs("veh_typ") & "")
            Text5 = Val(rs("dis_wdp") & "")
            
            'Updated to add mar_dop,mar_pgr,Mar_grd,Mar_vda
            
            Text8 = Trim(rs("mar_dop") & "")
            Text6 = Trim(rs("mar_pgr") & "")
            Text7 = Trim(rs("mar_grd") & "")
            
            Text9 = Trim(rs("MAR_COU") & "")
            Text10 = Trim(rs("MAR_STA") & "")
            Text11 = Trim(rs("MAR_DST") & "")
            Text12 = Trim(rs("MAR_TWC") & "")
            Text13 = Trim(rs("MAR_TWN") & "")
            Text14 = Trim(rs("MAR_CLU") & "")
            Text15 = Trim(rs("MAR_ROC") & "")
            Text16 = Trim(rs("MAR_CSG") & "")
            
            ROU_LST = Trim(rs!rou_idy & "")
            M_SMN_IDY = Trim(rs!SMN_IDY & "")
            M_SMN_NME = GetValue("smn_nme", "smn where smn_idy like '" & Trim(rs!SMN_IDY) & "'")
            Text1.Enabled = False
            Text2.SetFocus
            MOD1 = True
            Toolbar1.Buttons(12).Enabled = True
            'Toolbar1.Buttons(13).Enabled = False
        Else
            marlist
        End If
End If
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text2)) > 0 Then
       list3.SetFocus
     Else
        MsgBox "Blank Market Name, Enter Name !", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
     End If
     ElseIf K = 38 Then
    Text1.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text27_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MS.Text = Trim(Text27) & ""
    Text27.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Text27_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text4_GotFocus()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
M_MDI.STS_BAR.Panels(1).Text = "Enter Vehicle Type"
b = Text4
Dim rs As New ADODB.Recordset
rs.Open "select veh_typ from mar group by veh_typ", con, adOpenKeyset, adLockPessimistic
Text4.CLEAR
For A = 1 To rs.RecordCount
    Text4.AddItem Trim(rs(0))
    rs.MoveNext
Next
Text4 = b
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Distance from the WD"
SendKeys "{home}+{end}"
End Sub

Private Sub TEXT5_KeyDown(K As Integer, Shift As Integer)
'If K = 13 Or K = 40 Then
'    M_SMN_IDY.SetFocus
'ElseIf K = 38 Then
'    Text4.SetFocus
'End If
If K = 13 Then
    List2.SetFocus
    'save
ElseIf K = 38 Then
    list3.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)

If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text6_GotFocus()
Dim rs As New ADODB.Recordset
T = ""
rs.Open "select alltrim(DIR_MSG) from dir where substr(alltrim(DIR_IDY),1,2) like 'ST%' .and. substr(alltrim(DIR_IDX),1,2) like 'DI%' ORDER BY DIR_MSG", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Text6)
    Text6.CLEAR
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Text6.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Text6 = Trim(T)
End If

End Sub

Private Sub TEXT6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text8.SetFocus
ElseIf KeyCode = 38 Then
    List2.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
Dim rs As New ADODB.Recordset
T = ""
rs.Open "select alltrim(DIR_ALS) from dir where DIR_MSG = '" & Trim(Text6.Text) & "'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Text7)
    Text8.Text = rs(0)
    ElseIf Val(Text8.Text) < 1 Then
    
    Text8.Text = 0
End If

End Sub

Private Sub Text7_Change()
If Len(Trim(Text7)) > 14 Then SLT
End Sub

Private Sub Text7_GotFocus()
'updated to add mar_dop,mar_pgr,Mar_grd,Mar_vda
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Population Grade from list"

Dim rs As New ADODB.Recordset
T = ""
rs.Open "select alltrim(DIR_ALS) from dir where substr(alltrim(DIR_IDY),1,2) like 'MA%' .and. substr(alltrim(DIR_IDX),1,2) like 'IN%'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Text7)
    Text7.CLEAR
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Text7.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Text7 = Trim(T)
End If

''Dim rs As New ADODB.Recordset
''t = ""
''
''rs.Open "select alltrim(MAR_GRD) from MAR where substr(MAR_IDY,1,2) like ''", CON, adOpenKeyset, adlockpessimistic 'Modified by Ramanesh
''
''If rs.RecordCount > 0 Then
''    If MOD1 = True Then t = Trim(TEXT7)
''    TEXT7.CLEAR
''    For A = 1 To rs.RecordCount
''        If Len(Trim(rs(0))) <> 0 Then TEXT7.AddItem Trim(rs(0))
''        rs.MoveNext
''    Next
''    If MOD1 = True Then TEXT7 = Trim(t)
''End If
''If rs.State = 1 Then rs.Close
''rs.Open "select mar_grd from mar group by mar_grd", CON, adOpenKeyset, adlockpessimistic
''If rs.RecordCount > 0 Then
''    If MOD1 = True Then t = Trim(TEXT7)
'''    text7.Clear
''    For A = 1 To rs.RecordCount
''        If Len(Trim(rs(0))) <> 0 Then TEXT7.AddItem Trim(rs(0))
''        rs.MoveNext
''    Next
''    If MOD1 = True Then TEXT7 = Trim(t)
''End If

End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then LIST2.SetFocus


If KeyCode = 13 Then
   save
ElseIf KeyCode = 38 And Shift = 1 Then
    Text8.SetFocus
End If


End Sub

Private Sub Text7_keypress(KeyAscii As Integer)

If Len(Trim(Text7)) > 15 Then
Text7 = ""
Text7.SetFocus
End If
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0

End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then Text6.SetFocus
If KeyCode = 13 Then
Text7.SetFocus
ElseIf KeyCode = 38 Then
Text6.SetFocus
End If
End Sub

Private Sub Text8_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    exitfrm
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If M_FRM_MAR.ActiveControl.name = "Text1" Then marlist
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
'    M_FRM_PRE.Caption = "Market ID Prefix"
'    M_FRM_PRE.Show 1
    SYS_PSW.Tag = "MAR_PRE"
    SYS_PSW.Show 1
ElseIf Button.KEY = "brw" Then
'    BROWSE
ElseIf Button.KEY = "p" Then
    OpenReport CURDIR & "\lists.RP1", "Markets list", 0, "", 6
    cancel
ElseIf Button.KEY = "pr" Then
    'MDIForm1.RRReport1.WindowTitle = "WD M_FRM_MAR List Report"
    'MDIForm1.RRReport1.ReportName = App.Path & "\M_FRM_MAR list.rrw"
    'MDIForm1.RRReport1.Destination = 1
    'MDIForm1.RRReport1.Printer = "?"
    'MDIForm1.RRReport1.RunReport 1
ElseIf Button.KEY = "crt" Then
    CREATE
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
With MDIForm1
    If ButtonMenu.KEY = "MAR_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "Markets List", 0, "", 6
        cancel
    ElseIf ButtonMenu.KEY = "DST_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "Districts List", 2, "TRIM(DIR->DIR_IDY) = 'STA' .AND. TRIM(DIR->DIR_IDX) = 'DIS'", 6
        cancel
    ElseIf ButtonMenu.KEY = "STA_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "States List", 2, "TRIM(DIR->DIR_IDY) = 'STA' .AND. TRIM(DIR->DIR_IDX) = 'DIS'", 6
        cancel
    End If
End With
End Sub
