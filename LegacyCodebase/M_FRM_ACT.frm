VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_ACT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business Accounts & Parties"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "M_FRM_ACT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   480
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
            Picture         =   "M_FRM_ACT.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ACT.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&Account Details"
      TabPicture(0)   =   "M_FRM_ACT.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label13(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label32"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label27"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label37"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label38"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label39"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "List5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "act_dat"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Ad&dress"
      TabPicture(1)   =   "M_FRM_ACT.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label10(0)"
      Tab(1).Control(2)=   "Label11(0)"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(5)=   "Label13(0)"
      Tab(1).Control(6)=   "Label14(0)"
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(8)=   "Label18"
      Tab(1).Control(9)=   "Label12(0)"
      Tab(1).Control(10)=   "Label6"
      Tab(1).Control(11)=   "ZIPCODE"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text8"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text7"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text6"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text5"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text11"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "WWW"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "EMAILID"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "MOBILE"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text9"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "FAX"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text12"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "&Retailer"
      TabPicture(2)   =   "M_FRM_ACT.frx":3386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&Supplier"
      TabPicture(3)   =   "M_FRM_ACT.frx":33A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame1"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Retailer &Visit Day"
      TabPicture(4)   =   "M_FRM_ACT.frx":33BE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label28"
      Tab(4).Control(1)=   "ms"
      Tab(4).Control(2)=   "Text32"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmbRoute"
      Tab(4).Control(4)=   "txtSmn"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "View"
      TabPicture(5)   =   "M_FRM_ACT.frx":33DA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label20"
      Tab(5).Control(1)=   "list1"
      Tab(5).Control(2)=   "list2"
      Tab(5).Control(3)=   "list4"
      Tab(5).Control(4)=   "Frame5"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Ass&et Details"
      TabPicture(6)   =   "M_FRM_ACT.frx":33F6
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "astidy"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "msa"
      Tab(6).ControlCount=   2
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6360
         MaxLength       =   25
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2170
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   25
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3705
         Width           =   2385
      End
      Begin MSComCtl2.DTPicker act_dat 
         Height          =   330
         Left            =   7315
         TabIndex        =   115
         Top             =   1387
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16580611
         CurrentDate     =   37971
      End
      Begin VB.TextBox txtSmn 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72120
         TabIndex        =   114
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmbRoute 
         Height          =   315
         Left            =   -72120
         TabIndex        =   113
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox astidy 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72360
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid msa 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   103
         Top             =   720
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   7435
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   16711680
         FocusRect       =   0
         GridLines       =   2
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid ms 
         Height          =   2700
         Left            =   -73080
         TabIndex        =   101
         Top             =   1200
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4763
         _Version        =   393216
         Rows            =   8
         Cols            =   3
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   16711680
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   0
      End
      Begin VB.Frame Frame8 
         Caption         =   "Tax Information"
         Height          =   1455
         Left            =   4440
         TabIndex        =   94
         Top             =   3000
         Width           =   4215
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1920
            MaxLength       =   25
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   990
            Width           =   2170
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   615
            Width           =   2170
         End
         Begin VB.TextBox GST 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   240
            Width           =   2170
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old GST.SST Number"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   210
            TabIndex        =   100
            Top             =   1058
            Width           =   1575
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST Number"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   870
            TabIndex        =   99
            Top             =   690
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT - TIN Number"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   465
            TabIndex        =   98
            Top             =   315
            Width           =   1320
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Party Additional Information - Bank and Supplier "
         Height          =   2175
         Left            =   120
         TabIndex        =   81
         Top             =   2280
         Width           =   4215
         Begin VB.TextBox Text26 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1875
            MaxLength       =   25
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   2550
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1875
            MaxLength       =   25
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   735
            Width           =   2145
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1875
            MaxLength       =   15
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   360
            Width           =   2145
         End
         Begin VB.TextBox Text28 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1875
            MaxLength       =   25
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   1695
            Width           =   2145
         End
         Begin VB.TextBox Text27 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1875
            MaxLength       =   15
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2145
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            X1              =   0
            X2              =   4180
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank - Branch"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   690
            TabIndex        =   93
            Top             =   810
            Width           =   1020
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank A/C Number"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   405
            TabIndex        =   92
            Top             =   435
            Width           =   1305
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1."
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   435
            Width           =   135
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2."
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   1395
            Width           =   135
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RED Outlet ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   690
            TabIndex        =   89
            Top             =   1395
            Width           =   1020
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HVO Outlet ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   690
            TabIndex        =   88
            Top             =   1770
            Width           =   1020
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   735
            TabIndex        =   87
            Top             =   2625
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sales && Distribution Tracking Information"
         Enabled         =   0   'False
         Height          =   2415
         Left            =   -74640
         TabIndex        =   69
         Top             =   2640
         Width           =   8295
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "M_FRM_ACT.frx":3412
            Left            =   3360
            List            =   "M_FRM_ACT.frx":3414
            TabIndex        =   121
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox Text14 
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
            Enabled         =   0   'False
            Height          =   330
            Left            =   6500
            MaxLength       =   5
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   1900
            Width           =   825
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "M_FRM_ACT.frx":3416
            Left            =   3960
            List            =   "M_FRM_ACT.frx":3426
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   1900
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_FRM_ACT.frx":3449
            Left            =   1905
            List            =   "M_FRM_ACT.frx":346E
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   1900
            Width           =   1455
         End
         Begin VB.ListBox List8 
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
            ItemData        =   "M_FRM_ACT.frx":34E1
            Left            =   1935
            List            =   "M_FRM_ACT.frx":34EB
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox AREAID 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1935
            MaxLength       =   10
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   240
            Width           =   1185
         End
         Begin VB.TextBox Text31 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1935
            MaxLength       =   10
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   960
            Width           =   1185
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1935
            MaxLength       =   10
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Cases"
            Height          =   195
            Left            =   7440
            TabIndex        =   122
            Top             =   1965
            Width           =   435
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capacity"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   0
            Left            =   5760
            TabIndex        =   110
            Top             =   1965
            Width           =   615
         End
         Begin VB.Label lblgrd 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grade"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   3375
            TabIndex        =   107
            Top             =   1960
            Width           =   435
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   495
            TabIndex        =   106
            Top             =   1965
            Width           =   1005
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "National Key A/c"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   80
            Top             =   1485
            Width           =   1215
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
            Left            =   3360
            TabIndex        =   78
            Top             =   240
            Width           =   3945
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
            Left            =   3360
            TabIndex        =   77
            Top             =   960
            Width           =   3945
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
            Left            =   3360
            TabIndex        =   76
            Top             =   600
            Width           =   3945
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prime Location ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   75
            Top             =   315
            Width           =   1260
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Market ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   4
            Left            =   795
            TabIndex        =   74
            Top             =   1035
            Width           =   705
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Channel ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   2
            Left            =   705
            TabIndex        =   73
            Top             =   675
            Width           =   795
         End
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         ItemData        =   "M_FRM_ACT.frx":34F5
         Left            =   7315
         List            =   "M_FRM_ACT.frx":3508
         TabIndex        =   68
         Top             =   960
         Width           =   495
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74760
         TabIndex        =   60
         Top             =   4560
         Width           =   8055
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            Caption         =   "Sub Distributor"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5760
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            Caption         =   "AMC"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4800
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            Caption         =   "Van"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   7320
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            Caption         =   "Supplier"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3720
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            Caption         =   "Retailer"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2520
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   120
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1440
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -74640
         TabIndex        =   53
         Top             =   540
         Width           =   1815
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Supplier"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Purchase Invoice Defaults"
         Enabled         =   0   'False
         Height          =   1695
         Left            =   -74640
         TabIndex        =   44
         Top             =   900
         Width           =   8295
         Begin VB.ListBox List7 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            ItemData        =   "M_FRM_ACT.frx":351B
            Left            =   1995
            List            =   "M_FRM_ACT.frx":3525
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   825
            Width           =   495
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1995
            MaxLength       =   5
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text23 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4995
            MaxLength       =   9
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4995
            MaxLength       =   3
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MDL Required"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   4
            Left            =   675
            TabIndex        =   52
            Top             =   855
            Width           =   1035
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Discount %"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   51
            Top             =   435
            Width           =   1515
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Credit Days"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   2
            Left            =   3285
            TabIndex        =   50
            Top             =   818
            Width           =   1425
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Credit Limit"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   2
            Left            =   3285
            TabIndex        =   49
            Top             =   375
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -74520
         TabIndex        =   42
         Top             =   540
         Width           =   2175
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Retailer"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   30
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sale Invoice Defaults"
         Enabled         =   0   'False
         Height          =   1695
         Left            =   -74640
         TabIndex        =   32
         Top             =   900
         Width           =   8295
         Begin VB.ListBox List6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            ItemData        =   "M_FRM_ACT.frx":352F
            Left            =   1920
            List            =   "M_FRM_ACT.frx":353F
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text25 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1935
            MaxLength       =   9
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1230
            Width           =   975
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            ItemData        =   "M_FRM_ACT.frx":354F
            Left            =   4920
            List            =   "M_FRM_ACT.frx":355C
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1935
            MaxLength       =   5
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text20 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4920
            MaxLength       =   9
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   367
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Party"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   705
            TabIndex        =   41
            Top             =   870
            Width           =   810
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sale Rate (Products)"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   3090
            TabIndex        =   40
            Top             =   870
            Width           =   1470
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Due Days"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   3255
            TabIndex        =   39
            Top             =   435
            Width           =   1275
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Outstanding Limit"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   38
            Top             =   1305
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Discount %"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   37
            Top             =   435
            Width           =   1365
         End
      End
      Begin VB.TextBox FAX 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   25
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1095
         Width           =   2385
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   25
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   660
         Width           =   2385
      End
      Begin VB.TextBox MOBILE 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   25
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1500
         Width           =   2385
      End
      Begin VB.TextBox EMAILID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   30
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2355
         Width           =   2385
      End
      Begin VB.TextBox WWW 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   30
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2805
         Width           =   2385
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   25
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3255
         Width           =   2385
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68925
         MaxLength       =   25
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2385
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73485
         MaxLength       =   25
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   660
         Width           =   2625
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73485
         MaxLength       =   25
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2625
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73485
         MaxLength       =   25
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1380
         Width           =   2625
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73485
         MaxLength       =   25
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1860
         Width           =   2625
      End
      Begin VB.TextBox ZIPCODE 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73485
         MaxLength       =   6
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2295
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1995
         MaxLength       =   40
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1380
         Width           =   4305
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   1185
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   960
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
      Begin MSComctlLib.ListView list2 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   945
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
            Text            =   "Group Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group ID"
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   945
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Market"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Active"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GCC Name - Code"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   390
         TabIndex        =   127
         Top             =   1875
         Width           =   1305
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6360
         TabIndex        =   126
         Top             =   1800
         Width           =   2170
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1995
         TabIndex        =   125
         Top             =   1800
         Width           =   4305
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier XRef ID"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   5115
         TabIndex        =   120
         Top             =   2595
         Width           =   1185
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Created On"
         Height          =   195
         Left            =   6360
         TabIndex        =   112
         Top             =   1455
         Width           =   810
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Active A/c"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   6405
         TabIndex        =   67
         Top             =   990
         Width           =   765
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   -71070
         TabIndex        =   63
         Top             =   480
         Width           =   4665
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
         Left            =   -74520
         TabIndex        =   56
         Top             =   600
         Width           =   7815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   -70035
         TabIndex        =   31
         Top             =   728
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   -69525
         TabIndex        =   30
         Top             =   1163
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   -69735
         TabIndex        =   29
         Top             =   1568
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   -70365
         TabIndex        =   28
         Top             =   3323
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E Mail ID"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   -69915
         TabIndex        =   27
         Top             =   2423
         Width           =   645
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   -69930
         TabIndex        =   26
         Top             =   2873
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   -70110
         TabIndex        =   25
         Top             =   3773
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pager"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   -69690
         TabIndex        =   24
         Top             =   1988
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   -74595
         TabIndex        =   15
         Top             =   2363
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "City/Place"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   14
         Top             =   1928
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   -74520
         TabIndex        =   13
         Top             =   728
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3240
         TabIndex        =   7
         Top             =   960
         Width           =   3045
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   630
         TabIndex        =   6
         Top             =   1455
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group ID"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   1050
         TabIndex        =   5
         Top             =   1035
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account ID"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   885
         TabIndex        =   4
         Top             =   615
         Width           =   810
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   64
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
               NumButtonMenus  =   19
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACT_LST"
                  Text            =   "&a All Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_CAT"
                  Text            =   "&b Retailer Category"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_LOC"
                  Text            =   "&c Retailer Location"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_MAR"
                  Text            =   "&d Retailer Market"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ROU"
                  Text            =   "&e Retailer Route"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_SMN"
                  Text            =   "&f Retailer Salesman"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SUP_LST"
                  Text            =   "&g Supplier List"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ID_LST"
                  Text            =   "&h Retailer ID List"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "wk"
                  Text            =   "&i Retailer Week Day"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RDY_LST"
                  Text            =   "&j Routes List"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RRC"
                  Text            =   "&k Route Retailers Count"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MRC"
                  Text            =   "&l Market Retailers Count"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SRC"
                  Text            =   "&m Salesman Retailers Count"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RSR"
                  Text            =   "&n Retailer SKU Range"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NRL"
                  Text            =   "&o New Retailers List"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CLS_RET_ID"
                  Text            =   "&p Closed Outlets List"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VAN_LST"
                  Text            =   "&q Van List"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "GRD_KEY_LST"
                  Text            =   "&r Grade - Key - Count"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "GCC"
                  Text            =   "&s GCC Report by Page"
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
            Object.ToolTipText     =   "Prefix Account ID "
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
      Begin VB.CommandButton GCC_Xtract 
         Caption         =   "&2 GCC Xtract"
         Height          =   375
         Left            =   5670
         TabIndex        =   124
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton GCC_Import 
         Caption         =   "&1 GCC Import"
         Height          =   375
         Left            =   4440
         TabIndex        =   123
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label31 
      Caption         =   "Label31"
      Height          =   495
      Left            =   3840
      TabIndex        =   111
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "M_FRM_ACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, ls As ListItem, RV As Integer
Dim rsrdy As New ADODB.Recordset
Public Sub CREATE()
MOD1 = False
Text1 = ""
save
End Sub

Public Sub DISPDETAILS()
On Error GoTo lab
If CHECKTB("act.dbf,grp.dbf,loc.dbf,cat.dbf,smn.dbf,mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset, GCC_RS As New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from act where  act_idy like '" & Sincrement(Text1) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        Text1 = Sincrement(Text1)
        Text2 = Trim(rs("act_nme"))
        Text3 = Trim(rs("grp_idy"))
        If GCC_RS.State = 1 Then GCC_RS.Close
        GCC_RS.Open "SELECT * FROM GCC where GCC_ID like '" & Trim(rs("GCC_ID")) & "'", con, adOpenKeyset, adLockPessimistic
        If GCC_RS.RecordCount > 0 Then
            Label37 = Trim(GCC_RS("CST_NME"))
            Label38 = Trim(GCC_RS("GCC_ID"))
        End If
        GCC_RS.Close
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select * from grp where grp_idy like'" & Text3 & "'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then Label4 = RS1("grp_nme")
        Text5 = Trim(rs("act_ad1") & "")
        Text6 = Trim(rs("act_ad2") & "")
        If Trim(rs!CST_TYP) = "" Then
            Combo1.ListIndex = 0
        Else
            Combo1 = Trim(rs!CST_TYP)
        End If
        If Trim(rs!RET_TYP) = "" Then
            Combo2.ListIndex = 0
        Else
            If Trim(rs!RET_TYP) = "" Then
                Combo2.ListIndex = 0
            Else
                If IsNull(Trim(rs!RET_TYP)) = False Then
                    If Trim(rs!RET_TYP) = "A" Then
                        Combo2.ListIndex = 0
                    End If
                    If Trim(rs!RET_TYP) = "B" Then
                        Combo2.ListIndex = 1
                    End If
                    If Trim(rs!RET_TYP) = "C" Then
                        Combo2.ListIndex = 2
                    End If
                    If Trim(rs!RET_TYP) = "D" Then
                        Combo2.ListIndex = 3
                    Else
                        Combo2.ListIndex = 3
                    End If
                Else
                    Combo2.ListIndex = 0
                End If
            End If
        End If
        
        
        If Trim(rs!CST_TYP) = "" Then
            Combo1.ListIndex = 0
        Else
            Combo1 = Trim(rs!CST_TYP)
        End If
        
        Text7 = Trim(rs("act_ad3") & "")
        Text8 = Trim(rs("act_ad4") & "")
        ZIPCODE = Trim(rs("act_zip") & "")
        Text9 = Trim(rs("act_tfx") & "")
        FAX = Trim(rs("act_fax") & "")
        MOBILE = Trim(rs("act_mbl") & "")
        Text4 = Trim(rs("act_pgr") & "")
        EMAILID = Trim(rs("act_eml") & "")
        WWW = Trim(rs("ACT_WWW") & "")
        Text11 = Trim(rs("act_cpn") & "")
        Text12 = Trim(rs("act_dsg") & "")
        Text10 = Trim(rs("act_cst") & "")
        GST = Trim(rs("act_gst") & "")
        Text16 = Trim(rs("act_rgd") & "")
        Text17 = Trim(rs("acn_one") & "")
        Text19 = Trim(rs("bnm_one") & "")
'        Text26 = Trim(rs("bbn_one") & "")
        If Trim(rs!bbn_one) <> "" Then
            Combo3.Text = Trim(rs!bbn_one)
        Else
            Combo3.Text = Trim(rs!bbn_one)
        End If
        Text27 = Trim(rs("acn_two") & "")
        Text28 = Trim(rs("bnm_two") & "")
        Text29 = Trim(rs("bbn_two") & "")
        LIST8 = Trim(rs!key_act)
        If Trim(Combo1.Text) = "Van" Or Trim(Combo1.Text) = "Retailer" Or Trim(Combo1.Text) = "SubDistributor" Or Trim(Combo1.Text) = "AMC" Or Trim(Combo1.Text) = "Key A/c" Or Trim(Combo1.Text) = "Local Key A/c" Or Trim(Combo1.Text) = "Regnl Key A/c" Or Trim(Combo1.Text) = "HVO-P" Or Trim(Combo1.Text) = "HVO-NP" Or Trim(Combo1.Text) = "HVO-P-WO" Or Trim(Combo1.Text) = "HVO-NP-WO" Then
            Text14.Enabled = True
            Text14 = Val(Trim(rs("van_cap")) & "")
        End If
        'Else
'''''        ' Code commented by PNR - necessity not felt
'''''        If Combo2.Enabled = True Then
'''''            For K = 0 To Combo2.ListCount - 1
'''''                If Trim(Combo2.List(K)) = Trim(rs("RET_TYP")) Then
'''''                    Combo2.ListIndex = K
'''''                    Exit For
'''''                End If
'''''            Next
'''''        End If
        
        list5.Text = Trim(rs("dbt_typ")) & ""
        act_dat = rs("ACT_DOC")
        AREAID = Trim(rs("LOC_IDY") & "")
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM ASM,ACA WHERE ACA.ASM_IDY LIKE ASM.AsM_IDY AND ACA.Act_IDY LIKE '" & Trim(Text1) & "'", con, adOpenKeyset, adLockPessimistic
        MSA.Rows = RS1.RecordCount + 1
        With MSA
        If RS1.RecordCount = 0 Then MSA.Rows = 2
        For A = 1 To RS1.RecordCount
            .TextMatrix(A, 0) = Trim(RS1!ASM_IDY)
            .TextMatrix(A, 1) = Trim(RS1!ASM_NME)
            .TextMatrix(A, 2) = Trim(RS1!ASM_NOS)
            .TextMatrix(A, 3) = Trim(RS1!asm_cod)
            .TextMatrix(A, 4) = Trim(RS1!ASM_RMK)
            .TextMatrix(A, 5) = GetValue("IND_IDY", "ASM WHERE ASM_IDY LIKE '" & Trim(RS1!ASM_IDY) & "'")
            RS1.MoveNext
        Next
        If RS1.RecordCount > 0 Then
        MSA.Rows = MSA.Rows + 1
        MSA.SetFocus
        End If
        End With
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM RDY,SMN WHERE SMN.SMN_IDY LIKE RDY.SMN_IDY AND RDY.ACT_IDY LIKE '" & Trim(Text1) & "'", con, adOpenKeyset, adLockPessimistic
        For A = 1 To RS1.RecordCount
            MS.TextMatrix(Val(RS1!DAY_IDY & ""), 1) = Trim(RS1!SMN_NME)
            MS.TextMatrix(Val(RS1!DAY_IDY & ""), 2) = Trim(RS1!rou_idy & "")
            MS.TextMatrix(Val(RS1!DAY_IDY & ""), 3) = Trim(RS1!SMN_IDY)
            RS1.MoveNext
        Next
        If Trim(AREAID) <> "" Then
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from LOC where LOC_IDY like '" & Trim(AREAID) & "'", con, adOpenKeyset, adLockPessimistic
            If RS1.RecordCount > 0 Then Label29 = RS1("LOC_NME")
        End If
        
        Text13 = Trim(rs("CAT_idy") & "")
        If Trim(Text13) <> "" Then
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from CAT where cAT_idy like '" & Trim(Text13) & "'", con, adOpenKeyset, adLockPessimistic
            If RS1.RecordCount > 0 Then Label19 = RS1("CAT_nme")
        End If
        
        Text31 = Trim(rs("mar_idy") & "")
        If Trim(Text31) <> "" Then
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from mar where mar_idy like '" & Trim(Text31) & "'", con, adOpenKeyset, adLockPessimistic
            If RS1.RecordCount > 0 Then Label30 = RS1("mar_nme")
        End If

        Text32 = Trim(rs("smn_idy") & "")
        If Trim(Text32) <> "" Then
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from smn where smn_idy like '" & Trim(Text32) & "'", con, adOpenKeyset, adLockPessimistic
            If RS1.RecordCount > 0 Then Label31 = RS1("smn_nme")
        End If
        
        ROU_LST = Trim(rs("rou_idy") & "")
        
        If Trim(rs("act_typ")) = "SAL" Then
            Check1.Value = 1
            Frame2.Enabled = True
            Frame3.Enabled = True
            SSTab1.TabEnabled(4) = True
        ElseIf Trim(rs("act_typ")) = "PUR" Then
            Check2.Value = 1
            Frame1.Enabled = True
        ElseIf Trim(rs("act_typ")) = "BOT" Then
            Check1.Value = 1
            Check2.Value = 1
            Frame1.Enabled = True
            Frame3.Enabled = True
            Frame2.Enabled = True
'            SSTab1.TabEnabled(4) = True
        ElseIf Trim(rs("act_typ")) = "" Then
            Check1.Value = 0
            Check2.Value = 0
            
        End If
        If Check1.Value = 1 Then
        SSTab1.TabEnabled(6) = True
        End If
        Text18 = discount(rs("dbt_dis") & "")
        List6 = Trim(rs("mcl_req"))
        If Trim(rs("mcl_req")) = "Y" Then
            Text25.Enabled = True
            Text20.Enabled = True
        Else
            Text25.Enabled = False
            Text20.Enabled = False
        End If
        Text25 = amt(rs("dbt_mcl") & "")
        Text20 = Trim(rs("dbt_mcp") & "")
        
        list3 = Trim(rs("dbt_rap") & "")
        Text15 = discount(rs("crt_dis") & "")
        List7 = Trim(rs("mdl_req"))
        If Trim(rs("mdl_req")) = "Y" Then
            Text23.Enabled = True
            Text24.Enabled = True
        Else
            Text23.Enabled = False
            Text24.Enabled = False
        End If
        Text23 = amt(rs("crt_mdl") & "")
        Text24 = Trim(rs("crt_ppd") & "")
        RS1.Close
        MOD1 = True
        Text1.Enabled = False
        Toolbar1.Buttons(12).Enabled = True
        Text3.SetFocus
    Else
        accountslist
    End If
Exit Sub
lab:
If err.Number = -2147217887 Then
    MsgBox "Invalid Table! Validate Fields in the Table!", vbExclamation, "DAS Version 4.0"
Else
    MsgBox err.Description
End If
End Sub

Private Sub delete()
On Error GoTo lab
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & "  Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Confirm Account Deletion ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
    If check("act", "act_idy", Text1) = False Then
        MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        clrctr M_FRM_ACT
        Text1.Enabled = True
        Text1.SetFocus
        setval
        Exit Sub
    Else
        If deletecheck = False Then
            con.Execute "delete from act where act_idy like '" & Sincrement(Text1) & "'"
            SSTab1.Tab = 0
            Text1.Enabled = True
            Text1.SetFocus
            setval
            clrctr M_FRM_ACT
            cancel
        Else
            MsgBox "Cannot Delete, Transaction Entries Already Done!", vbCritical, "DAS Version 4.0"
            Text2.SetFocus
            Exit Sub
        End If
    End If
    
Else
    cancel
End If
Exit Sub
lab:
MsgBox err.Description
End Sub

Private Sub setval()
Text25.Enabled = False
Text24.Enabled = False
Text23.Enabled = False
Text20.Enabled = False
Text18 = "0.00"
Text25 = "0.00"
Text20 = "0"
Text15 = "0.00"
Text23 = "0.00"
Text24 = "0"
Label4 = ""
Label28 = ""
Label29 = ""
Label30 = ""
Label31 = ""
Label19 = ""
list3.Selected(0) = True
list5.Selected(0) = True
act_dat.Value = Date
List6.Selected(0) = True
List7.Selected(0) = True
LIST8.Selected(1) = True
Check1.Value = 0
Check2.Value = 0
Frame2.Enabled = False
Frame1.Enabled = False
SSTab1.TabEnabled(5) = False
Combo1.ListIndex = 0
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
MS.CLEAR
MS.TextMatrix(0, 0) = "Day"
MS.TextMatrix(0, 1) = "SalesMan"
MS.ColWidth(1) = 1800
MS.ColAlignment(1) = 1
SCl MS, 1
MS.TextMatrix(0, 2) = "Route"
MS.ColWidth(2) = 2400
MS.ColAlignment(2) = 1
SCl MS, 2
MS.Cols = 4
MS.ColWidth(0) = 1300
'MS.ColWidth(1) = MS.Width - 2200
'MS.ColWidth(2) = 3500
MS.ColWidth(3) = 0
MS.TextMatrix(1, 0) = "Monday"
MS.TextMatrix(2, 0) = "Tuesday"
MS.TextMatrix(3, 0) = "Wednesday"
MS.TextMatrix(4, 0) = "Thursday"
MS.TextMatrix(5, 0) = "Friday"
MS.TextMatrix(6, 0) = "Saturday"
MS.TextMatrix(7, 0) = "Sunday"
With MSA
.CLEAR
.Cols = 6
.Rows = 2
.TextMatrix(0, 0) = "Asset ID"
SCl MSA, 0
.TextMatrix(0, 1) = "Asset Name"
.ColAlignment(2) = 7
.TextMatrix(0, 2) = "Nos"
SCl MSA, 2
.ColAlignment(2) = 7
.TextMatrix(0, 3) = "Asset Code"
SCl MSA, 3
.TextMatrix(0, 4) = "Remarks"
SCl MSA, 4
.ColAlignment(0) = 1
.ColAlignment(1) = 1
.ColAlignment(2) = 7
.ColAlignment(3) = 1
.ColAlignment(4) = 1
.ColWidth(0) = 1200
.ColWidth(1) = 2500
.ColWidth(2) = 800
.ColWidth(3) = 1200
.ColWidth(5) = 0
.ColWidth(4) = .Width - (1200 + 2550 + 800 + 1200)
End With
End Sub

Private Sub ACT_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text17.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    list5.SetFocus
End If
End Sub

Private Sub AREAID_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Location ID or Press <Enter> to get List of Location ID's"
End Sub

Private Sub AREAID_keydown(K As Integer, s As Integer)
On Error GoTo lab
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If rs.State = 1 Then rs.Close
        rs.Open " select * from LOC where LOC_IDY like '" & Sincrement(AREAID) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            AREAID = Sincrement(AREAID)
            Label29 = rs("LOC_NME")
            If K = 38 Then
                
            Else
                Text13.SetFocus
            End If
        Else
            arealist
        End If
End If
Exit Sub
lab:
MsgBox err.Description
End Sub

Private Sub AREAID_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub ASTIDY_KeyDown(K As Integer, Shift As Integer)
On Error GoTo lab
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("asm.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
If MSA.Col = 0 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from ASM where ASM_idy like '" & Sincrement(Trim(ASTIDY)) & "'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        With MSA
            .TextMatrix(MSA.Row, 0) = Sincrement(Trim(ASTIDY))
            .TextMatrix(MSA.Row, 1) = Trim(REC_SET("ASM_nme") & "")
            .TextMatrix(MSA.Row, 5) = Trim(REC_SET("ind_idy") & "")
            If .TextMatrix(MSA.Row, 5) = "Y" Then .TextMatrix(MSA.Row, 2) = 1
            If .TextMatrix(.Rows - 1, 0) <> "" Then .AddItem ""
            ASTIDY.Visible = False
            .SetFocus
            .Col = .Col + 1
        End With
    Else
        V_AST_LST
        MSA.Col = MSA.Col + 2
    End If
ElseIf MSA.Col = 2 Then
    MSA.Text = ASTIDY
    ASTIDY = ""
    ASTIDY.Visible = False
    MSA.SetFocus
    MSA.Col = MSA.Col + 1
ElseIf MSA.Col = 3 Then
    For A = 1 To MSA.Rows - 1
        If MSA.TextMatrix(MSA.Row, 0) = MSA.TextMatrix(A, 0) And MSA.TextMatrix(A, 3) = ASTIDY Then
            MsgBox "Asset Code Already Exists", vbExclamation
            ASTIDY.Visible = False
            MSA.SetFocus
            Exit Sub
        End If
    Next
    A = Check_Asset_Code(MSA.TextMatrix(MS.Row, 0), ASTIDY)
    If A <> 0 Then
        MsgBox "Asset Code Already Exists", vbExclamation
        ASTIDY.Visible = False
        MSA.SetFocus
        Exit Sub
    End If
    MSA.Text = ASTIDY
    ASTIDY = ""
    ASTIDY.Visible = False
    MSA.SetFocus
    MSA.Col = MSA.Col + 1
ElseIf MSA.Col = 4 Then
    MSA.Text = ASTIDY
    ASTIDY = ""
    ASTIDY.Visible = False
    MSA.SetFocus
    If MSA.Row <> MSA.Rows - 1 Then
        MSA.Row = MSA.Row + 1
        MSA.Col = 0
    End If
End If
End If
Exit Sub
lab:
MsgBox err.Description, vbExclamation, "DAS Version 4.0"
End Sub
Private Sub V_AST_LST()
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from asm,att WHERE asm.att_idy like att.att_idy and aSM_nme like '" & Trim(ASTIDY) & "%'", con, adOpenKeyset, adLockPessimistic
list4.ListItems.CLEAR
list4.ColumnHeaders.CLEAR
list4.ColumnHeaders.Add , , "Name"
list4.ColumnHeaders.Add , , "ID"
list4.ColumnHeaders.Add , , , 0
list4.ColumnHeaders.Add , , "Asset Type", 2000
list4.ColumnHeaders.Add , , "Company ID", 2000
list4.ColumnHeaders.Item(1).Width = 2500
list4.ColumnHeaders.Item(2).Width = 1500
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("ASM_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ASM_idy"))
        'ls.ListSubItems.Add , , Val(REC_SET("ASM_qty") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("IND_IDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET!att_typ & "")
        ls.ListSubItems.Add , , Trim(REC_SET!ASM_CID & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 5
    Label20 = "Select Asset"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 6 Then
        MsgBox "Asset Not Found, Define Asset !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 6
        If ASTIDY.Visible = True Then
            ASTIDY.SetFocus
        Else
            MSA.SetFocus
        End If
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub

Private Sub ASTIDY_KEYPRESS(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If MSA.Col = 2 Then
    If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
    Else
        K = 0
    End If
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
    Check4.Value = 0
    Check3.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check8.Value = 0
    Check9.Value = 0
    Label20.Caption = "Select Account"
    accountslist
ElseIf Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 And Check7.Value = 0 And Check8.Value = 0 And Check9.Value = 0 Then
    accountslist
End If
SSTab1.TabEnabled(5) = False
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then
    Check4.Value = 0
    Check3.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check9.Value = 0
    Label20.Caption = "Select Account"
    accountslist
ElseIf Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 And Check7.Value = 0 And Check8.Value = 0 And Check9.Value = 0 Then
    accountslist
End If
SSTab1.TabEnabled(5) = False
End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then
    Check4.Value = 0
    Check3.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check8.Value = 0
    Label20.Caption = "Select Account"
    accountslist
ElseIf Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 And Check7.Value = 0 And Check8.Value = 0 And Check9.Value = 0 Then
    accountslist
End If
SSTab1.TabEnabled(5) = False
End Sub

Private Sub cmbRoute_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MS.Text = Trim(cmbRoute.Text)
    If FINDITEM(Trim(MS.Text)) = 0 Then
        cmbRoute.AddItem Trim(MS.Text)
    End If
    cmbRoute.Visible = False
    If MS.Row <> MS.Rows - 1 Then
        MS.Row = MS.Row + 1
        MS.Col = 1
    End If
    MS.SetFocus
End If
End Sub

Private Sub cmbRoute_keypress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Combo2.Enabled = True
    Combo2.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
If Combo3.Enabled = True Then
    Combo3.SetFocus
Else
    LIST8.SetFocus
End If
End If
End Sub

Private Sub Combo1_keypress(K As Integer)
If K = 13 Then
'    If Trim(Combo1.Text) = "Van" Then
    If Trim(Combo1.Text) = "Van" Or Trim(Combo1.Text) = "Retailer" Or Trim(Combo1.Text) = "SubDistributor" Or Trim(Combo1.Text) = "AMC" Or Trim(Combo1.Text) = "Key A/c" Or Trim(Combo1.Text) = "Local Key A/c" Or Trim(Combo1.Text) = "Regnl Key A/c" Or Trim(Combo1.Text) = "HVO-P" Or Trim(Combo1.Text) = "HVO-NP" Or Trim(Combo1.Text) = "HVO-P-WO" Or Trim(Combo1.Text) = "HVO-NP-WO" Then
        Text14.Enabled = True
'        Text14.SetFocus
        Combo2.Enabled = True
        Combo2.SetFocus
    Else
        Combo2.Enabled = True
        Combo2.SetFocus
        Text14.Enabled = True
    End If
ElseIf K = 38 And s = 1 Then
    list5.SetFocus
End If
End Sub

Private Sub COMBO2_GotFocus()
'''''Dim RS_DIR As New ADODB.Recordset ' Code necessity not felt, commented by PNR
'''''t = ""
'''''If RS_DIR.State = 1 Then RS_DIR.Close
'''''RS_DIR.Open "select alltrim(DIR_ALS) from dir where substr(alltrim(DIR_IDY),1,3) like 'RET' .and. substr(alltrim(DIR_IDX),1,2) like 'TY%'", CON, adOpenKeyset, adlockpessimistic 'Modified by Ramanesh
'''''If RS_DIR.RecordCount > 0 Then
'''''    If MOD1 = True Then t = Trim(Combo2)
'''''    Combo2.CLEAR
'''''    For A = 1 To RS_DIR.RecordCount
'''''        If Len(Trim(RS_DIR(0))) <> 0 Then Combo2.AddItem Trim(RS_DIR(0))
'''''        RS_DIR.MoveNext
'''''    Next
''''''    If MOD1 = True Then Combo2 = Trim(t)
'''''End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text14.Enabled = True
    Text14.SetFocus
'    sstab1.Tab = 4
'    ms.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Combo2_keypress(K As Integer)
If K = 13 Then
    If Trim(Combo2.Text) = "DIAMOND" Or Trim(Combo2.Text) = "GOLD" Or Trim(Combo2.Text) = "SILVER" Or Trim(Combo2.Text) = "BRONZE" Then
        Text14.Enabled = True
        Text14.SetFocus
    Else
        Text14.Enabled = True
        Text14.SetFocus
    End If
ElseIf K = 38 And s = 1 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Combo3_GotFocus()
Dim RS_DIR As New ADODB.Recordset
T = ""
If RS_DIR.State = 1 Then RS_DIR.Close
RS_DIR.Open "select alltrim(DIR_ALS) from dir where substr(alltrim(DIR_IDY),1,2) like 'KA%' .and. substr(alltrim(DIR_IDX),1,2) like 'NK%' ORDER BY DIR_ALS", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If RS_DIR.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Combo3)
    Combo3.CLEAR
    For A = 1 To RS_DIR.RecordCount
        If Len(Trim(RS_DIR(0))) <> 0 Then Combo3.AddItem Trim(RS_DIR(0))
        RS_DIR.MoveNext
    Next
'    If MOD1 = True Then Combo3 = Trim(t)
End If

End Sub
Private Sub Combo3_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
ElseIf KeyAscii = 38 And s = 1 Then
    LIST8.SetFocus
End If
End Sub
Private Sub GCC_Import_Click()
GCC_FileImport
cancel
Text1.SetFocus
End Sub
Private Sub GCC_Xtract_Click()
GCC_FileXtract
cancel
Text1.SetFocus
End Sub
Private Sub LIST8_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Or Select Key Account"
End Sub
Private Sub List8_Keydown(KeyAscii As Integer, s As Integer)
If Trim(LIST8.Text) = "Y" Then
    If Combo3.Enabled = False Then
       Combo3.Enabled = True
    End If
    If KeyAscii = 13 Then
        Combo3.SetFocus
    ElseIf KeyAscii = 38 And s = 1 Then
        Text31.SetFocus
    End If
Else
    Combo3.CLEAR
    Combo3.Enabled = False
    If KeyAscii = 13 Then
        Combo1.SetFocus
    ElseIf KeyAscii = 38 And s = 1 Then
        Text31.SetFocus
    End If
End If
End Sub
Private Sub MS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    MS.TextMatrix(MS.Row, 1) = ""
    MS.TextMatrix(MS.Row, 2) = ""
    MS.TextMatrix(MS.Row, 3) = ""
End If
End Sub

Private Sub MS_KeyPress(K As Integer)
If MS.Row <> 0 Then
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
    If MS.Col = 1 Then
        If K = 13 Then
            smnlist
        Else
            txtSmn.Visible = True
            txtSmn = ""
            txtSmn.Top = MS.Top + MS.CellTop
            txtSmn.Left = MS.CellLeft + MS.Left
            txtSmn.Width = MS.CellWidth
            txtSmn.Height = MS.CellHeight
            txtSmn = Chr(K)
            txtSmn.SetFocus
            txtSmn.SelStart = Len(Trim(txtSmn))
        End If
'        If K = 13 And ms.Col = 1 Then
'            If Trim(ms.TextMatrix(ms.Row, 1)) <> "" Then
'                ms.Col = 2
'                ms.SetFocus
'            Else
'                Text32 = ""
'                smnlist
'            End If
'        'ElseIf K <> 27 Then
'        '    If K = vbKeySpace Then
'        '    Else
'        '    Text32 = ""
'        '    Text32.Top = ms.Top + ms.CellTop
'        '    Text32.Left = ms.CellLeft + ms.Left
'        '    Text32.Width = ms.CellWidth
'        '    Text32.Height = ms.CellHeight
'        '    If ms.Col = 1 Then
'        '        Text32.MaxLength = 15
'         '   Else
'        '        Text32.MaxLength = 15
'        '    End If
'         '   Text32.Visible = True
'         '   Text32.SetFocus
'
'          '  If K <> 13 Then Text32 = Chr(K)
'          '  Text32.SelStart = Len(Text32)
'          '  End If
'        End If
    ElseIf MS.Col = 2 Then
        If Trim(MS.TextMatrix(MS.Row, 1)) <> "" Then
            If K <> 27 Then
                If K = vbKeySpace Then
                Else
                    cmbRoute.Top = MS.Top + MS.CellTop
                    cmbRoute.Left = MS.CellLeft + MS.Left
                    cmbRoute.Width = MS.CellWidth - 620
'                    cmbRoute.Height = ms.CellHeight
                    cmbRoute.Enabled = True
                    cmbRoute.Visible = True
                    cmbRoute.SetFocus
                    SendKeys "{HOME}+{END}"
                    'If K <> 13 Then cmbRoute = Chr(K)
                    'TEXT32.SelStart = Len(TEXT32)
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub MS_LeaveCell()
Text32.Visible = False
cmbRoute.Visible = False
End Sub

Private Sub MS_Scroll()
If Text32.Visible = True Then
    Text32 = ""
    Text32.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub MSA_KeyPress(K As Integer)
If MSA.Row <> 0 Then
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
    
    If MSA.Col = 0 Or MSA.Col = 3 Or MSA.Col = 4 Then
        If K = 13 And MSA.Col = 0 Then
            If Trim(MSA.TextMatrix(MSA.Row, 0)) <> "" Then
                MSA.Col = 2
                MSA.SetFocus
        
            Else
                ASTIDY = ""
                V_AST_LST
            End If
        ElseIf K <> 27 Then
            If K = vbKeySpace Then
            Else
                ASTIDY.Visible = True
                ASTIDY.Top = MSA.Top + MSA.CellTop
                ASTIDY.Left = MSA.CellLeft + MSA.Left
                ASTIDY.Alignment = 0
                ASTIDY.Width = MSA.CellWidth
                ASTIDY.Height = MSA.CellHeight
                If MSA.Col = 0 Then
                    ASTIDY.MaxLength = 10
                ElseIf MSA.Col = 3 Then
                    ASTIDY.MaxLength = 15
                Else
                    ASTIDY.MaxLength = 75
                End If
                ASTIDY.Visible = True
                ASTIDY.SetFocus
                If K <> 13 Then ASTIDY = Chr(K)
                ASTIDY.SelStart = Len(ASTIDY)
            End If
        End If
    ElseIf MSA.Col = 2 And MSA.TextMatrix(MSA.Row, 5) = "N" Then
        If K = 39 Or K = 41 Or K = 37 Then K = 0
        If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
            If K = vbKeySpace Then
            Else
                ASTIDY.Visible = True
                ASTIDY.Alignment = 1
                ASTIDY.Top = MSA.Top + MSA.CellTop
                ASTIDY.Left = MSA.CellLeft + MSA.Left
                ASTIDY.Width = MSA.CellWidth
                ASTIDY.Height = MSA.CellHeight
                ASTIDY.MaxLength = 5
                ASTIDY.Visible = True
                ASTIDY.SetFocus
                If K <> 13 Then ASTIDY = Chr(K)
                ASTIDY.SelStart = Len(ASTIDY)
            End If
        End If
    End If
End If

End Sub

Private Sub MSA_LeaveCell()
ASTIDY.Visible = False
End Sub


Private Sub msa_Scroll()
If ASTIDY.Visible = True Then
    ASTIDY = ""
    ASTIDY.Visible = False
    MSA.SetFocus
End If
End Sub

Private Sub Text13_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Category ID or Press <Enter> to get List of Category ID's"
End Sub

Private Sub Text13_KeyDown(K As Integer, s As Integer)
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If rs.State = 1 Then rs.Close
        rs.Open " select * from CAT where CAT_idy like '" & Sincrement(Text13) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            Text13 = Sincrement(Text13)
            Label19 = rs("CAT_nme")
            If K = 38 Then
                AREAID.SetFocus
            Else
                Text31.SetFocus
            End If
        Else
            chanlist
        End If
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub Check1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Retailer to Enter Retailer Details "
End Sub

Private Sub Check2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Supplier to Enter Supplier Details "
End Sub

Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then
    If Check1.Value = 1 Then
        Frame2.Enabled = True
        Frame3.Enabled = True
        SSTab1.TabEnabled(4) = True
        SSTab1.TabEnabled(6) = True
        Text18.SetFocus
    Else
        Frame2.Enabled = False
        Frame3.Enabled = False
        SSTab1.Tab = 3
        Check2.SetFocus
    End If
End If
End Sub

Private Sub CHECK2_KeyPress(K As Integer)
If K = 13 Then
    If Check2.Value = 1 Then
        Frame1.Enabled = True
        Text15.SetFocus
    Else
        Frame1.Enabled = False
        If Check1.Value = 1 Then
'            SSTab1.TabEnabled(4) = True
'            SSTab1.Tab = 4
'            AREAID.SetFocus
             save
        Else
            save
        End If
    End If
End If
End Sub
Private Sub save()
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If CHECKTB("act.dbf,grp.dbf,loc.dbf,cat.dbf,smn.dbf,mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If check("grp", "grp_idy", Trim(Text3)) = False Then
MsgBox "No Group Accounts Found, Define Group Account !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Text3.Enabled = True
Text3.SetFocus
Exit Sub
End If

If Len(Trim(Text2)) = 0 Then
MsgBox "Account Name Should Not Be Blank, Give Name !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Text2.Enabled = True
Text2.SetFocus
Exit Sub
End If
If Check1.Value = 1 Then
    If Len(Trim(AREAID)) > 0 Then
        If check("LOC", "LOC_IDY", Trim(AREAID)) = False Then
            MsgBox "Location Not Found, Define Location !", vbExclamation, "DAS Version 4.0"
            SSTab1.Tab = 2
            AREAID.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Location Information Missing! Location Information Link is Required for Retailer!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 2
        AREAID.SetFocus
        Exit Sub
    End If
End If
If Check1.Value = 1 Then
    If Len(Trim(Text13)) > 0 Then
        If check("CAT", "CAT_idy", Trim(Text13)) = False Then
            MsgBox "Category - Channel Not Found, Define Category - Channel!", vbExclamation, "DAS Version 4.0"
            SSTab1.Tab = 2
            Text13.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Category - Channel Information Missing! Category - Channel Information Link is Required for Retailer!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 2
        Text13.SetFocus
        Exit Sub
    End If
End If
If Check1.Value = 1 Then
    If Len(Trim(Text31)) > 0 Then
        If check("mar", "mar_idy", Trim(Text31)) = False Then
            MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
            SSTab1.Tab = 2
            Text31.SetFocus
            Exit Sub
        End If
    Else
        mar = 1
        MsgBox "Market Information Missing! Market Information Link is Required While Invoicing", vbExclamation, "DAS Version 4.0"
        Text31.SetFocus 'to be check with testing
        Exit Sub
    End If
End If
If Check1.Value = 1 And Trim(Combo1.Text) <> "Van" Then
    If Len(Trim(Combo2)) = 0 Then
        MsgBox "Retailer Grade Information Required!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 2
        Combo2.SetFocus
        Exit Sub
    End If
End If
If Check1.Value = 1 Then
    SMN = 0
    For A = 1 To MS.Rows - 1
        'If Trim(ms.TextMatrix(A, 1)) <> "" Then
        If Len(Trim(MS.TextMatrix(A, 1))) > 0 Then
            If Len(Trim(MS.TextMatrix(A, 2))) > 0 Then
                SMN = 1
                Exit For
            End If
        End If
    Next
    If A = MS.Rows Then
        MsgBox "Salesman Information Missing! Salesman Information Link is Required With Routes While Invoicing", vbCritical, "DAS Version 4.0"
        SSTab1.Tab = 4
        MS.SetFocus
        MS.Col = 1
        MS.Row = 1
        Exit Sub
    End If
End If
'If Val(mar) = 1 Then MsgBox "Market Information Missing! Market Information Link is Required While Invoicing", vbExclamation
'If Val(SMN) = 1 Then MsgBox "Salesman Information Missing! Salesman Information Link is Required While Invoicing", vbExclamation

For A = 1 To MSA.Rows - 1
    If Trim(MSA.TextMatrix(A, 5)) = "Y" And Trim(MSA.TextMatrix(A, 3)) = "" Then
        MsgBox "Asset Code Required!", vbInformation + vbOKOnly, "DAS Version 4.0"
        MSA.Row = A
        MSA.Col = 3
        MSA.SetFocus
        Exit Sub
    End If
    b = Check_Asset_Code(MSA.TextMatrix(A, 0), MSA.TextMatrix(A, 3))
    If b <> 0 Then
        MsgBox "Asset already Exists", vbExclamation
        SSTab1.Tab = 6
        MSA.Row = b
        MSA.SetFocus
        Exit Sub
    End If
Next

CheckGCCStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'GCC'")
If Trim(CheckGCCStatus) = "T" Then
    If Check1.Value = 1 And Check2.Value = 0 Then
        If MOD1 = False Then
            MsgBox "Unable to Create New Retail Outlet and Code as GCC Details Required! ", vbInformation, "DAS Version 4.0"
            cancel
            Exit Sub
        End If
    End If
End If

If MOD1 = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from act where act_idy like '" & Sincrement(Text1) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from act where act_idy like ''", con, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs("act_idy") = increment("idy_act")
    MsgBox "New Account ID is " & rs("act_idy"), vbExclamation, "DAS Version 4.0"
End If
    rs("grp_idy") = Trim(Text3) & ""
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from grp where grp_idy like '" & Trim(Text3) & "'", con, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount > 0 Then rs("grp_clf") = RS1("grp_clf")
        rs("act_nme") = Text2 & ""
        rs("act_ad1") = Text5 & ""
        rs("act_ad2") = Text6 & ""
        rs("act_ad3") = Text7 & ""
        rs("act_ad4") = Text8 & ""
        rs("act_zip") = ZIPCODE & ""
        rs!CST_TYP = Combo1
        rs("act_tfx") = Text9 & ""
        rs("act_fax") = FAX & ""
        rs("act_mbl") = MOBILE & ""
        rs("act_pgr") = Text4 & ""
        rs("act_eml") = EMAILID & ""
        rs("ACT_WWW") = WWW & ""
        rs("act_cpn") = Text11 & ""
        rs("act_dsg") = Text12
        rs("DBT_TYP") = Trim(list5)
        rs("ACT_DOC") = act_dat.Value
        rs("act_gst") = GST & ""
        rs("act_cst") = Trim(Text10 & "")
        rs("act_rgd") = Trim(Text16 & "")
        rs("acn_one") = Trim(Text17 & "")
        rs("bnm_one") = Trim(Text19 & "")
        'rs("bbn_one") = Trim(Text26 & "")
        rs("bbn_one") = Trim(Combo3.Text & "")
        rs("acn_two") = Trim(Text27 & "")
        rs("bnm_two") = Trim(Text28 & "")
        rs("bbn_two") = Trim(Text29 & "")
        rs("sku_rng") = 0#
        If Combo2.Enabled = True Then
            rs("ret_typ") = Trim(Combo2.Text & "")
            rs("van_cap") = Val(Text14)
        Else
            rs("van_cap") = Val(Text14)
        End If
        rs!key_act = LIST8
        
        If (Check1.Value = 1 And Check2.Value = 1) Or (Check1.Value = 1 And Check2.Value = 0) Then
            rs("LOC_IDY") = Trim(AREAID & "")
            rs("CAT_idy") = Trim(Text13 & "")
            rs("mar_idy") = Trim(Text31 & "")
'            rs("smn_idy") = ""
'            rs("rou_idy") = ""
        Else
            rs("zon_idy") = ""
            rs("LOC_IDY") = ""
            rs("CAT_idy") = ""
            rs("mar_idy") = ""
'            rs("smn_idy") = ""
'            rs("rou_idy") = ""
        End If
        'rs("smn_idy") = Trim(Text32 & "")
If Check1.Value = 1 And Check2.Value = 0 Then
    rs("act_typ") = "SAL"
    rs("dbt_DIs") = Val(Trim(Text18))
    rs("mcl_req") = Trim(List6)

    If List6 = "Y" Then
        rs("dbt_mcl") = Val(Text25)
        rs("dbt_mcp") = Val(Text20)
    Else
        rs("dbt_mcl") = "0.00"
        rs("dbt_mcp") = 0
    End If
    rs("dbt_rap") = list3
    
    rs("crt_dis") = "0.00"
    rs("mdl_req") = "N"
    rs("crt_mdl") = "0.00"
    rs("crt_ppd") = 0
    If MOD1 = True Then
        con.Execute "delete from rdy where act_idy like '" & Trim(Text1) & "'"
    Else
        con.Execute "delete from rdy where act_idy like '" & Trim(rs!ACT_IDY) & "'"
    End If
    For A = 1 To MS.Rows - 1
        If Trim(MS.TextMatrix(A, 3)) <> "" Then
            If MOD1 = False Then
'                If Trim(list5) = "Y" Then   'Check and Update Beat Plan based on the Outlet Status
                    con.Execute "insert into rdy values('" & Trim(rs!ACT_IDY) & "'," & A & ",'" & Trim(MS.TextMatrix(A, 3)) & "','" & Trim(MS.TextMatrix(A, 2)) & "',0,0,'0','0')"
'                End If
            Else
'                If Trim(list5) = "Y" Then   'Check and Update Beat Plan based on the Outlet Status
                    con.Execute "insert into rdy values('" & Trim(Text1) & "'," & A & ",'" & Trim(MS.TextMatrix(A, 3)) & "','" & Trim(MS.TextMatrix(A, 2)) & "',0,0,'0','0')"
'                End If
            End If
        End If
    Next
    rs("crt_rap") = "A"
ElseIf Check2.Value = 1 And Check1.Value = 0 Then
    rs("act_typ") = "PUR"
    rs("crt_dis") = Val(Text15)
    rs("mdl_req") = Trim(List7)
    If List7 = "Y" Then
        rs("crt_mdl") = Val(Text23)
        rs("crt_ppd") = Val(Text24)
    Else
        rs("crt_mdl") = "0.00"
        rs("crt_ppd") = 0
    End If
    
    rs("dbt_DIs") = "0.00"
    rs("mcl_req") = "N"
    rs("dbt_mcl") = "0.00"
    rs("dbt_mcp") = 0
    rs("dbt_rap") = "A"
        
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
    rs("act_typ") = "BOT"
    
    rs("dbt_DIs") = Val(Trim(Text18))
    rs("mcl_req") = List6
    If List6 = "Y" Then
        rs("dbt_mcl") = Val(Text25)
        rs("dbt_mcp") = Val(Text20)
    Else
        rs("dbt_mcl") = "0.00"
        rs("dbt_mcp") = 0
    End If
    rs("dbt_rap") = list3
    
    rs("crt_dis") = Val(Text15)
    rs("mdl_req") = List7
    If List7 = "Y" Then
        rs("crt_mdl") = Val(Text23)
        rs("crt_ppd") = Val(Text24)
    Else
        rs("crt_mdl") = "0.00"
        rs("crt_ppd") = 0
    End If
    
    'modified on 01/07/2004
    
    If MOD1 = True Then
        con.Execute "delete from rdy where act_idy like '" & Trim(Text1) & "'"
    Else
        con.Execute "delete from rdy where act_idy like '" & Trim(rs!ACT_IDY) & "'"
    End If
    For A = 1 To MS.Rows - 1
        If Trim(MS.TextMatrix(A, 3)) <> "" Then
            If MOD1 = False Then
'                If Trim(list5) = "Y" Then   'Check and Update for Active Outlet or not in Beat Plan
                    con.Execute "insert into rdy values('" & Trim(rs!ACT_IDY) & "'," & A & ",'" & Trim(MS.TextMatrix(A, 3)) & "','" & Trim(MS.TextMatrix(A, 2)) & "',0,0,'0','0')"
'                End If
            Else
'                If Trim(list5) = "Y" Then   'Check and Update for Active Outlet or not in Beat Plan
                    con.Execute "insert into rdy values('" & Trim(Text1) & "'," & A & ",'" & Trim(MS.TextMatrix(A, 3)) & "','" & Trim(MS.TextMatrix(A, 2)) & "',0,0,'0','0')"
'                End If
            End If
        End If
    Next
Else
    rs("act_typ") = ""
    rs("crt_dis") = "0.00"
    rs("mdl_req") = "N"
    rs("crt_mdl") = "0.00"
    rs("crt_ppd") = 0
    
    rs("crt_rap") = "A"
    rs("dbt_DIs") = "0.00"
    rs("mcl_req") = "N"
    rs("dbt_mcl") = "0.00"
    rs("dbt_mcp") = 0
    rs("dbt_rap") = "A"
End If
rs.Update
If MOD1 = True Then
    con.Execute "DELETE FROM ACA WHERE ACT_IDY LIKE '" & Trim(Text1) & "'"
Else
    con.Execute "DELETE FROM ACA WHERE ACT_IDY LIKE '" & Trim(rs!ACT_IDY) & "'"
End If
For A = 1 To MSA.Rows - 1
    If MSA.TextMatrix(A, 0) <> "" Then
        If MOD1 = True Then
            If MSA.TextMatrix(A, 5) = "Y" And MSA.TextMatrix(A, 3) <> "" Then
                con.Execute "INSERT INTO ACA VALUES('" & Trim(Text1) & "','" & Trim(MSA.TextMatrix(A, 0)) & "'," & Val(MSA.TextMatrix(A, 2)) & ",'" & Trim(MSA.TextMatrix(A, 3)) & "','" & Trim(MSA.TextMatrix(A, 4)) & "')"
            ElseIf MSA.TextMatrix(A, 5) = "N" Then
                con.Execute "INSERT INTO ACA VALUES('" & Trim(Text1) & "','" & Trim(MSA.TextMatrix(A, 0)) & "'," & Val(MSA.TextMatrix(A, 2)) & ",'','" & Trim(MSA.TextMatrix(A, 4)) & "')"
            End If
        Else
               'CON.Execute "INSERT INTO ACA VALUES('" & Trim(rs!ACT_IDY) & "','" & Trim(msa.TextMatrix(A, 0)) & "'," & Val(msa.TextMatrix(A, 2)) & ",'" & Trim(msa.TextMatrix(A, 3)) & "','" & Trim(msa.TextMatrix(A, 4)) & "')"
            If MSA.TextMatrix(A, 5) = "Y" And MSA.TextMatrix(A, 3) <> "" Then
                con.Execute "INSERT INTO ACA VALUES('" & Trim(rs!ACT_IDY) & "','" & Trim(MSA.TextMatrix(A, 0)) & "'," & Val(MSA.TextMatrix(A, 2)) & ",'" & Trim(MSA.TextMatrix(A, 3)) & "','" & Trim(MSA.TextMatrix(A, 4)) & "')"
            ElseIf MSA.TextMatrix(A, 5) = "N" Then
                con.Execute "INSERT INTO ACA VALUES('" & Trim(rs!ACT_IDY) & "','" & Trim(MSA.TextMatrix(A, 0)) & "'," & Val(MSA.TextMatrix(A, 2)) & ",'','" & Trim(MSA.TextMatrix(A, 4)) & "')"
            End If
        End If
    End If
Next
MOD1 = False
clrctr M_FRM_ACT
setval
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
cancel
End Sub

Private Sub cancel()
clrctr M_FRM_ACT
Label37 = ""
Label38 = ""
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
setval
SSTab1.TabEnabled(6) = False
CNT = 0
End Sub

Private Sub newrec()
clrctr M_FRM_ACT
Text3.SetFocus
SSTab1.Tab = 0
Text1.Enabled = False
setval
End Sub

Private Sub accountslist()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check3.Value = 1 Then
    'rs.Open "select * from act where act_nme like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy where act_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check4.Value = 1 Then
    'rs.Open "select * from act where act_ad4 like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy where mar_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check5.Value = 1 Then
    'rs.Open "select * from act where act_typ like 'SAL' and act_nme like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy where act_typ like 'SAL' and cst_typ in('Retailer','Key A/c','Local Key A/c','Regnl Key A/c', 'HVO-P', 'HVO-NP', 'HVO-P-WO', 'HVO-NP-WO') and act_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check6.Value = 1 Then
    'rs.Open "select * from act where act_typ like 'PUR' and act_nme like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy where act_typ like 'PUR' and act_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check7.Value = 1 Then
    'rs.Open "select * from act where act_typ like 'PUR' and act_nme like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy where act_typ like 'SAL' and cst_typ= 'Van' and act_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check8.Value = 1 Then
    'rs.Open "select * from act where act_typ like 'PUR' and act_nme like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy where act_typ like 'SAL' and cst_typ= 'AMC' and act_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check9.Value = 1 Then
    'rs.Open "select * from act where act_typ like 'PUR' and act_nme like '" & Trim(Text1) & "%'", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy where act_typ like 'SAL' and cst_typ= 'SubDistributor ' and act_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic

Else
    rs.Open "select * from act left join mar on act.mar_idy like mar.mar_idy", con, adOpenKeyset, adLockPessimistic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("ACT_NME")))
        ls.ListSubItems.Add , , Trim(rs("ACT_IDY"))
        ls.ListSubItems.Add , , Trim(rs("MAR_NME") & "")
        ls.ListSubItems.Add , , Trim(rs("DBT_TYP"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(5) = True
    SSTab1.Tab = 5
    Label20 = "Select Account"
    list1.ColumnHeaders(1).Text = "Account Name"
    list1.ColumnHeaders(2).Text = "Account ID"
    list1.ColumnHeaders(3).Text = "Market"
    list1.ColumnHeaders(4).Text = "Active"
    Check3.Visible = True
    Check4.Visible = True
    Check5.Visible = True
    Check6.Visible = True
    Check7.Visible = True
    Check8.Visible = True
    Check9.Visible = True
    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check3.Value = 1 Then
    rs.Open "select * from smn where smn_nme like '" & Trim(txtSmn) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check4.Value = 1 Then
    rs.Open "select * from smn where smn_ad4 like '" & Trim(txtSmn) & "%'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from smn where smn_nme like '" & Trim(txtSmn) & "%'", con, adOpenKeyset, adLockPessimistic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy"))
        ls.ListSubItems.Add , , Trim(rs("smn_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.TabEnabled(5) = True
    SSTab1.Tab = 5
    Label20 = "Select Salesmen"
    list1.ColumnHeaders(1).Text = "Salesmen Name"
    list1.ColumnHeaders(2).Text = "Salesmen ID"
    list1.ColumnHeaders(3).Text = "Place"
    
    Check3.Visible = True
    Check4.Visible = True
    Check5.Visible = False
    Check6.Visible = False
    Check7.Visible = False
    Check8.Visible = False
    Check9.Visible = False

    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    MsgBox "Salesmen Not Found, Define Salesmen !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 4
    txtSmn.Visible = False
    MS.SetFocus
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub grouplist()
If CHECKTB("grp.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from grp where grp_nme like '" & Text3 & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    List2.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(rs("grp_nme")))
        ls.ListSubItems.Add , , Trim(rs("grp_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(5) = True
    SSTab1.Tab = 5
    Label20 = "Select Group Account"
    List2.Visible = True
    List2.ListItems(1).Selected = True
    List2.SetFocus
Else
   MsgBox "Group Account Not Found, Define Group Account !", vbExclamation, "DAS Version 4.0"
   SSTab1.Tab = 0
   Text3.SetFocus
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub exitfrm()
Unload Me
Set Form = Nothing
End Sub

Private Sub arealist()
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from LOC where LOC_NME like '" & AREAID & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list4.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rs("LOC_nme")))
        ls.ListSubItems.Add , , Trim(rs("LOC_IDY"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(5) = True
    SSTab1.Tab = 5
    Label20 = "Select Location"
    list4.ColumnHeaders(1).Text = "Location Name"
    list4.ColumnHeaders(2).Text = "Location ID"
    list4.Visible = True
    list4.ListItems(1).Selected = True 'Added by PNR
    list4.SetFocus
Else
    MsgBox "Location Not Found, Define Location !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 2
    AREAID.SetFocus
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub chanlist()
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from CAT where CAT_nme like '" & Text13 & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list4.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rs("CAT_nme")))
        ls.ListSubItems.Add , , Trim(rs("CAT_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(5) = True
    SSTab1.Tab = 5
    Label20 = "Select Category"
    list4.ColumnHeaders(1).Text = "Category Name"
    list4.ColumnHeaders(2).Text = "Category ID"
    list4.Visible = True
    list4.ListItems(1).Selected = True
    list4.SetFocus
Else
    MsgBox "Category Not Found, Define Category !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 2
    Text13.SetFocus
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub marlist()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from mar where mar_nme like '" & Trim(Text31) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list4.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rs("mar_nme")))
        ls.ListSubItems.Add , , Trim(rs("mar_idy"))
        rs.MoveNext
    Next
    SSTab1.TabEnabled(5) = True
    SSTab1.Tab = 5
    Label20 = "Select Market"
    list4.ColumnHeaders(1).Text = "Market Name"
    list4.ColumnHeaders(2).Text = "Market ID"
    list4.Visible = True
    list4.ListItems(1).Selected = True
    list4.SetFocus
Else
    MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 2
    Text31.SetFocus
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub Check3_Click()
Dim rs As New ADODB.Recordset
If Check3.Value = 1 Then
    Check4.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check8.Value = 0
    Check9.Value = 0
    If Label20 = "Select Account" Then
        accountslist
    ElseIf Label20 = "Select Salesmen" Then
        smnlist
    End If
ElseIf Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 And Check7.Value = 0 And Check8.Value = 0 And Check9.Value = 0 Then
    accountslist
End If
SSTab1.TabEnabled(5) = False
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
    Check4.Value = 0
    Check3.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check8.Value = 0
    Check9.Value = 0
    Label20.Caption = "Select Account"
    accountslist
ElseIf Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 And Check7.Value = 0 And Check8.Value = 0 And Check9.Value = 0 Then
    accountslist
End If
SSTab1.TabEnabled(5) = False
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
    Check4.Value = 0
    Check3.Value = 0
    Check5.Value = 0
    Check7.Value = 0
    Check8.Value = 0
    Check9.Value = 0
    Label20.Caption = "Select Account"
    accountslist
ElseIf Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 And Check7.Value = 0 And Check8.Value = 0 And Check9.Value = 0 Then
    accountslist
End If
SSTab1.TabEnabled(5) = False
End Sub

Private Sub Check3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Name to Get List Sorted on Name "
End Sub

Private Sub Check4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Place to Get List Sorted on Place "
End Sub

Private Sub Check4_Click()
Dim rs As New ADODB.Recordset
If Check4.Value = 1 Then
    Check3.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    If Label20 = "Select Account" Then
        accountslist
    ElseIf Label20 = "Select Salesmen" Then
        smnlist
    End If
ElseIf Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 Then
    accountslist
End If
SSTab1.TabEnabled(5) = False
End Sub

Private Sub EMAILID_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter EMail ID"
End Sub

Private Sub EMAILID_Keydown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Or KeyAscii = 40 Then
    WWW.SetFocus
ElseIf KeyAscii = 38 Then
    Text4.SetFocus
End If
End Sub

Private Sub FAX_Change()
If Len(FAX) = 25 Then SendKeys "{home}+{end}"
End Sub

Private Sub FAX_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Fax Number"
End Sub

Private Sub FAX_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    MOBILE.SetFocus
ElseIf K = 38 Then
    Text9.SetFocus
End If
End Sub

Private Sub Form_Activate()
'SETYEAR M_FRM_ACT 'Modified by Radhika & PNR
checkdongle
Dim DIR_RS As New ADODB.Recordset
CheckGCCStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'GCC'")
If Trim(CheckGCCStatus) = "T" Then
    Me.Caption = "Business Accounts & Parties - Global Customer Code (GCC) Enabled Module"
Else
    GCC_Import.Enabled = False
    GCC_Xtract.Enabled = False
End If
SSTab1.Tab = 0
SSTab1.TabEnabled(5) = False
SSTab1.TabEnabled(6) = False
list3.Selected(0) = True
list5.Selected(0) = True
act_dat.Value = Date
List6.Selected(0) = True
List7.Selected(0) = True
LIST8.Selected(1) = True
Combo1.ListIndex = 0
If Text1.Enabled = True Then
    Text1.SetFocus
Else
    If Text3.Enabled = True Then Text3.SetFocus
End If
If rsrdy.State = 1 Then rsrdy.Close
rsrdy.Open "select rou_idy from rdy GROUP BY ROU_IDY", con, adOpenKeyset, adLockPessimistic
While Not rsrdy.EOF
    cmbRoute.AddItem Trim(rsrdy(0))
    rsrdy.MoveNext
Wend
End Sub
Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_ACT.ActiveControl Is TextBox Then M_FRM_ACT.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False And List2.Visible = False And list4.Visible = False Then
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        exitfrm
    ElseIf K = 68 And s = 2 Then
        delete
    ElseIf K = 67 And s = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 69 And s = 2 Then
        If M_FRM_ACT.ActiveControl.name = "Text1" Then
            accountslist
            Check3.Visible = True
            Check4.Visible = True
            Check5.Visible = True
            Check6.Visible = True
            Check7.Visible = True
            Check8.Visible = True
            Check9.Visible = True

        ElseIf M_FRM_ACT.ActiveControl.name = "Text32" Then
            smnlist
            Check3.Visible = True
            Check4.Visible = True
            Check5.Visible = False
            Check6.Visible = False
            Check7.Visible = False
            Check8.Visible = False
            Check9.Visible = False
        End If
        If M_FRM_ACT.ActiveControl.name = "Text3" Then grouplist
        If M_FRM_ACT.ActiveControl.name = "AREAID" Then arealist
        If M_FRM_ACT.ActiveControl.name = "Text13" Then chanlist
        
        If M_FRM_ACT.ActiveControl.name = "Text31" Then marlist
    ElseIf K = 80 And s = 2 Then
        M_FRM_AQY.Tag = "ACT"
        M_FRM_AQY.Show 1
    ElseIf K = 82 And s = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\Accounts list.RRW"
        'MDIForm1.RRReport1.Destination = 1
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then
KeyAscii = 0
End If
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        Check3.Visible = False
        Check4.Visible = False
        Check5.Visible = False
        Check6.Visible = False
        Check7.Visible = False
        Check8.Visible = False
        Check9.Visible = False

        SSTab1.TabEnabled(5) = False
        
        Toolbar1.Buttons(12).Enabled = False
        If Label20 = "Select Account" Then
        If CHECKTB("act.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If
            SSTab1.Tab = 0
            If check("act", "act_idy", Trim(Text1)) = False Then
                Text1.Enabled = True
                Text1.SetFocus
            Else
                Text1.Enabled = False
                TEXT1_KeyDown 13, 0
                Text3.SetFocus
            End If
        ElseIf Label20 = "Select Salesmen" Then
            SSTab1.Tab = 4
            If Text32.Visible = True Then
                SLT
                Text32.SetFocus
            Else
                MS.SetFocus
            End If
        End If
    ElseIf List2.Visible = True Then
        List2.Visible = False
        SSTab1.TabEnabled(5) = False
        SSTab1.Tab = 0
        Text1.Enabled = False
        Text3.SetFocus
    ElseIf list4.Visible = True Then
        list4.Visible = False
        SSTab1.TabEnabled(5) = False
        If Label20 = "Select Location" Then
            SSTab1.Tab = 2
            AREAID.SetFocus
        ElseIf Label20 = "Select Market" Then
            SSTab1.Tab = 2
            Text31.SetFocus
        ElseIf Label20 = "Select Category" Then
            SSTab1.Tab = 2
            Text13.SetFocus
        ElseIf Label20 = "Select Asset" Then
            SSTab1.Tab = 6
            ASTIDY.Visible = False
            MSA.SetFocus
        End If
    Else
        If M_FRM_ACT.ActiveControl.name = "Text1" Then
            Unload Me
            Set Form = Nothing
            'CODE ADDED BY PKB
            If M_FRM_ACT.Tag = "PACT" Then
            M_FRM_PUR.M_SUP_IDY.SetFocus
            ElseIf M_FRM_ACT.Tag = "SACT" Then
            M_FRM_SAL.M_CST_IDY.SetFocus
            End If
        Else
            save
             'CODE ADDED BY PKB
            If M_FRM_ACT.Tag = "PACT" Then
            M_FRM_PUR.M_SUP_IDY.SetFocus
            ElseIf M_FRM_ACT.Tag = "SACT" Then
            M_FRM_SAL.M_CST_IDY.SetFocus
            End If
            KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
If CHECKTB("grp.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
List2.FullRowSelect = True
If rs.State = 1 Then rs.Close
rs.Open "select * from grp", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
    For A = 1 To rs.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(rs("grp_nme")))
        ls.ListSubItems.Add , , Trim(rs("GRP_idy"))
        rs.MoveNext
    Next
End If
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
setval
rs.Close
Set rs = Nothing
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub GST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter State Sales Tax Number"
End Sub

Private Sub GST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text10.SetFocus
ElseIf KeyCode = 38 Then
    Text29.SetFocus
End If
End Sub

Private Sub GST_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select ID or Press <Esc> to Quit"
End Sub

Private Sub List2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select ID or Press <Esc> to Quit"
End Sub

Private Sub list4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select ID or Press <Esc> to Quit"
End Sub

Private Sub list2_DblClick()
List2_Keydown 13, 0
End Sub

Private Sub list1_KeyDown(KeyAscii As Integer, s As Integer)
Dim rs As New ADODB.Recordset
Dim lst As ListSubItem, Lst1 As ListSubItem
If KeyAscii = 13 Then
    If Label20 = "Select Account" Then
        If list1.ListItems.count > 0 Then
            Set lst = list1.SelectedItem.ListSubItems.Item(1)
            rs.Open "select act_idy from act where act_idy like '" & Trim(lst.Text) & "'", con, adOpenKeyset, adLockPessimistic
            If rs.RecordCount > 0 Then
                SSTab1.Tab = 0
                SSTab1.TabEnabled(5) = False
                Text1 = rs(0)
                DISPDETAILS
                Text3.SetFocus
            End If
        End If
    ElseIf Label20 = "Select Salesmen" Then
        If list1.ListItems.count > 0 Then
            Set Lst1 = list1.SelectedItem.ListSubItems.Item(1)
            rs.Open "select * from smn where smn_idy like '" & Trim(Lst1.Text) & "'", con, adOpenKeyset, adLockPessimistic
            If rs.RecordCount > 0 Then
                SSTab1.TabEnabled(5) = False
                MS.TextMatrix(MS.Row, 3) = Trim(rs("smn_idy"))
                RV = MS.Row
                MS.TextMatrix(MS.Row, 1) = Trim(rs("Smn_nme"))
                MS.SetFocus
                SSTab1.Tab = 4
                Text32.Visible = False
                If MS.Col = 1 Then
                    MS.Col = MS.Col + 1
                    MS.Row = RV
                Else
                    If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                End If
            End If
        End If
    End If
    
End If
End Sub

Private Sub list2_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If List2.SortOrder = lvwAscending Then
List2.SortOrder = lvwDescending
ElseIf List2.SortOrder = lvwDescending Then
List2.SortOrder = lvwAscending
End If
List2.SortKey = CH.Index - 1
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub list2_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub List3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Sale Slab Using Arrow Keys "
End Sub

Private Sub List3_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    AREAID.SetFocus
ElseIf K = 38 And s = 1 Then
    If Text20.Enabled = True Then
        Text20.SetFocus
    Else
        List6.SetFocus
    End If
End If
End Sub

Private Sub List4_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list4.SortOrder = lvwAscending Then
list4.SortOrder = lvwDescending
ElseIf list4.SortOrder = lvwDescending Then
list4.SortOrder = lvwAscending
End If
list4.SortKey = CH.Index - 1
End Sub

Private Sub list4_DblClick()
list4_KeyPress 13
End Sub

Private Sub list4_KeyPress(K As Integer)
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If K = 13 Then
    If list4.ListItems.count > 0 Then
        If Label20 = "Select Location" Then
            If list4.ListItems.count > 0 Then
                Dim lst As ListSubItem
                Set lst = list4.SelectedItem.ListSubItems.Item(1)
                If rs.State = 1 Then rs.Close
                rs.Open "select * from LOC where LOC_IDY like '" & Trim(lst.Text) & "'", con, adOpenKeyset, adLockPessimistic
                AREAID = rs("LOC_IDY")
                Label29 = rs("LOC_NME")
                SSTab1.TabEnabled(5) = False
                list4.Visible = False
                SSTab1.Tab = 2
                Text13.SetFocus
            End If
        ElseIf Label20 = "Select Market" Then
            If list4.ListItems.count > 0 Then
                Dim lst2 As ListSubItem
                Set lst2 = list4.SelectedItem.ListSubItems.Item(1)
                If rs.State = 1 Then rs.Close
                rs.Open "select * from mar where mar_idy like '" & Trim(lst2.Text) & "'", con, adOpenKeyset, adLockPessimistic
                If rs.RecordCount > 0 Then
                    Text31 = rs("mar_idy")
                    Label30 = rs("mar_nme")
                    If MOD1 = False Then
                        If RS1.State = 1 Then RS1.Close
                        RS1.Open "SELECT * FROM SMN WHERE SMN_IDY LIKE '" & Trim(rs("SMN_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                        If RS1.RecordCount > 0 Then Label31 = Trim(RS1("smn_nme"))
                        Text32 = Trim(rs!SMN_IDY & "")
                        ROU_LST = Trim(rs!rou_idy & "")
                    End If
                End If
                SSTab1.TabEnabled(5) = False
                list4.Visible = False
                SSTab1.Tab = 2
                LIST8.SetFocus
            End If
        ElseIf Label20 = "Select Category" Then
            If list4.ListItems.count > 0 Then
                Dim lst3 As ListSubItem
                Set lst3 = list4.SelectedItem.ListSubItems.Item(1)
                If rs.State = 1 Then rs.Close
                rs.Open "select * from CAT where CAT_idy like '" & Trim(lst3.Text) & "'", con, adOpenKeyset, adLockPessimistic
                Text13 = rs("CAT_idy")
                Label19 = rs("CAT_nme")
                SSTab1.TabEnabled(5) = False
                list4.Visible = False
                SSTab1.Tab = 2
                Text31.SetFocus
            End If
        ElseIf Label20 = "Select Asset" Then
            If list4.ListItems.count > 0 Then
                MSA.TextMatrix(MSA.Row, 1) = list4.SelectedItem
                r = MSA.Row
                MSA.TextMatrix(MSA.Row, 0) = Trim(list4.SelectedItem.ListSubItems.Item(1))
                MSA.TextMatrix(MSA.Row, 5) = GetValue("IND_IDY", "ASM WHERE ASM_IDY LIKE '" & Trim(list4.SelectedItem.ListSubItems.Item(1)) & "'")
                If MSA.TextMatrix(MSA.Row, 5) = "Y" Then MSA.TextMatrix(MSA.Row, 2) = 1
                If MSA.TextMatrix(MSA.Rows - 1, 0) <> "" Then MSA.AddItem ""
                SSTab1.TabEnabled(5) = False
                list4.Visible = False
                SSTab1.Tab = 6
                MSA.SetFocus
                MSA.Row = r
                MSA.Col = 3
                If MSA.TextMatrix(MSA.Row, 5) = "N" Then
                MSA.SetFocus
                MSA.Row = r
                MSA.Col = 2
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub list4_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub List5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Active Y/N "
End Sub

Private Sub List5_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    act_dat.SetFocus
ElseIf K = 38 And s = 1 Then
    Text2.SetFocus
End If
End Sub

Private Sub List6_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Maximum Credit Limit Required or Not"
End Sub

Private Sub List6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 And Shift = 1 Then
    Text18.SetFocus
ElseIf KeyCode = 13 Then
    If List6.Selected(1) = True Then
        Text25.Enabled = True
        Text20.Enabled = True
        Text25.SetFocus
    Else
        Text25.Enabled = False
        Text20.Enabled = False
        list3.SetFocus
    End If
End If
End Sub

Private Sub List7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Maximum Debit Limit Required or Not"
End Sub

Private Sub List7_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 And Shift = 1 Then
    Text15.SetFocus
ElseIf KeyCode = 13 Then
    If List7.Selected(1) = True Then
        Text23.Enabled = True
        Text24.Enabled = True
        Text23.SetFocus
    Else
        Text23.Enabled = False
        Text24.Enabled = False
        save
    End If
End If
End Sub

Private Sub MOBILE_Change()
If Len(MOBILE) = 25 Then SendKeys "{home}+{end}"
End Sub

Private Sub MOBILE_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Mobile Number"
End Sub

Private Sub MOBILE_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text4.SetFocus
ElseIf K = 38 Then
    FAX.SetFocus
End If
End Sub

Private Sub MOBILE_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 32 Or K = 45 Then
Else
K = 0
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
list1.Visible = False
List2.Visible = False
list4.Visible = False
SSTab1.TabEnabled(5) = False
Check3.Visible = False
Check4.Visible = False
Check5.Visible = False
Check6.Visible = False
Check7.Visible = False
Check8.Visible = False
Check9.Visible = False

If SSTab1.Tab = 0 Then
    If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        Text3.SetFocus
    End If
ElseIf SSTab1.Tab = 1 Then
    Text5.SetFocus
ElseIf SSTab1.Tab = 2 Then
    Check1.SetFocus
ElseIf SSTab1.Tab = 3 Then
    Check2.SetFocus
ElseIf SSTab1.Tab = 4 Then
    txtSmn.Visible = False
    If MS.Rows > 1 Then
        MS.SetFocus
        MS.Row = 1
        MS.Col = 1
    End If
ElseIf SSTab1.Tab = 6 Then
    If MSA.Rows > 1 Then
    MSA.Col = 0
    MSA.Row = 1
    MSA.SetFocus
    End If
End If
If SSTab1.Tab = 5 Then
    'Toolbar1.Enabled = False
Else
    Toolbar1.Enabled = True
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Account ID or Press <Enter> to get List of Account ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SSTab1.Tab = 4
    MS.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    Combo2.SetFocus
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 0
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text15_Change()
If Len(Trim(Text15)) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text15_Click()
SendKeys "{home}+{end}"
End Sub

Private Sub Text15_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Purchase Discount in Percentage"
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text15 = discount(Trim(Text15))
    List7.SetFocus
End If
End Sub

Private Sub Text15_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Or K = 45 Then
Else
K = 0
End If
End Sub

Private Sub Text16_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Others"
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    SSTab1.Tab = 1
    Text5.SetFocus
ElseIf KeyCode = 38 Then
    Text10.SetFocus
End If
End Sub

Private Sub Text17_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Bank Account Number"
End Sub

Private Sub Text17_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text19.SetFocus
ElseIf KeyCode = 38 Then
    act_dat.SetFocus
End If
End Sub

Private Sub Text16_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text17_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text19_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Bank Name"
End Sub

Private Sub Text19_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text20_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text24_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

'Private Sub Text26_GotFocus()
'M_MDI.STS_BAR.Panels(1).Text = "Enter Branch Name"
'End Sub

Private Sub Text27_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Bank Account Number"
End Sub

Private Sub Text28_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Bank Name"
End Sub

Private Sub Text29_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Branch Name"
End Sub

'Private Sub Text26_KeyPress(K As Integer)
'If K > 96 And K < 123 Then K = K - 32
'End Sub

Private Sub Text27_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text28_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text29_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text18_Change()
If Len(Text18) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text18_Click()
SendKeys "{home}+{end}"
End Sub

Private Sub Text18_GotFocus()
SendKeys "{home}+{End}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Sales Discount Percentage"
End Sub

Private Sub Text18_KeyDown(KeyCode As Integer, s As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text18 = discount(Trim(Text18))
    List6.SetFocus
End If
End Sub

Private Sub Text18_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Or K = 45 Then
Else
K = 0
End If
End Sub

Private Sub Text19_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text27.SetFocus
ElseIf KeyCode = 38 Then
    Text17.SetFocus
End If
End Sub

Private Sub Text27_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text28.SetFocus
ElseIf KeyCode = 38 Then
    Text19.SetFocus
End If
End Sub

Private Sub Text28_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text29.SetFocus
ElseIf KeyCode = 38 Then
    Text27.SetFocus
End If
End Sub

Private Sub Text29_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    GST.SetFocus
ElseIf KeyCode = 38 Then
    Text28.SetFocus
End If
End Sub

Private Sub Text31_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Market ID or Press <Enter> to get List of Market ID's"
End Sub

Private Sub Text32_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesmen ID or Press <Enter> to get List of Salesmen ID's"
End Sub

Private Sub ROU_LST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Or Select Route <A-Z> or <1-9>"
End Sub

Private Sub Text31_KeyPress(K As Integer)
 If K > 96 And K < 123 Then K = K - 32
 If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub Text31_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf,mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If rs.State = 1 Then rs.Close
        rs.Open " select * from mar where mar_idy like '" & Sincrement(Text31) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            Text31 = Sincrement(Text31)
            Label30 = rs("mar_nme")
            If MOD1 = False Then
                Text32 = Trim(rs!SMN_IDY)
                If RS1.State = 1 Then RS1.Close
                RS1.Open "SELECT SMN_NME FROM SMN WHERE SMN_IDY LIKE '" & Trim(rs("SMN_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If RS1.RecordCount > 0 Then Label31 = RS1("smn_nme")
                ROU_LST = Trim(rs!rou_idy & "")
            Else
                
            End If
            If K = 38 Then
                Text13.SetFocus
            Else
                LIST8.SetFocus
            End If
        Else
            marlist
        End If
End If
End Sub

Private Sub Text32_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Then
    If MS.Col = 1 Then
        If rs.State = 1 Then rs.Close
        If Asc(Mid(Trim(Text32), 1, 1)) = 13 Or Asc(Mid(Trim(Text32), Len(Trim(Text32)), 1)) Then Text32 = ""
        rs.Open " select * from smn where smn_idy like '" & Sincrement(Trim(Text32)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            MS.TextMatrix(MS.Row, 3) = Sincrement(Text32)
            MS.Text = Trim(rs("smn_nme") & "")
            Text32.Visible = False
            MS.Col = MS.Col + 1
            If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
        Else
            smnlist
        End If
    Else
        MS.Text = Trim(Text32)
        Text32 = ""
        Text32.Visible = False
        MS.SetFocus
        If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
        MS.Col = 1
    End If
End If
End Sub

Private Sub Text32_KeyPress(K As Integer)
 If K > 96 And K < 123 Then K = K - 32
 If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub ROU_LST_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    LIST8.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    Text32.SetFocus
End If
End Sub

Private Sub ROU_LST_kEYPRESS(K As Integer)
 If K > 96 And K < 123 Then K = K - 32
 If (K > 64 And K < 92) Or (K > 96 And K < 123) Or (K > 48 And K < 58) Or K = 8 Then
 Else
 K = 0
 End If
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Account Name"
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 38 Or KeyCode = 40 Then
    If Len(Trim(Text2)) = 0 Then
        MsgBox "Name Should Not Be Blank, Give Name !", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
    Else
        If KeyCode = 38 Then
            Text3.SetFocus
        Else
            list5.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text20_Change()
If Len(Text20) = 3 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text20_GotFocus()
SendKeys "{home}+{End}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Bill Due Days"
End Sub

Private Sub Text20_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    If Val(Text20) > 365 Then
        MsgBox "Credit Period Cannot be Greater Than 365 Days !", vbExclamation, "DAS Version 4.0"
        Text20.SetFocus
    Else
        If K = 38 Then
            Text25.SetFocus
        Else
            list3.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text23_Change()
If Len(Text23) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text23_GotFocus()
SendKeys "{home}+{End}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Maximum Credit Limit"
End Sub

Private Sub Text23_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text23 = convert(amt(Val(Text23)))
Text24.SetFocus
ElseIf K = 38 Then
Text23 = convert(amt(Val(Text23)))
List7.SetFocus
End If
End Sub

Private Sub Text23_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text24_Change()
If Len(Text24) = 3 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text24_GotFocus()
SendKeys "{home}+{End}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Payable Period in Days"
End Sub

Private Sub Text24_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    If Val(Text24) > 365 Then
        MsgBox "Payable Period Cannot be Greater Than 365 Days !", vbExclamation, "DAS Version 4.0"
        Text24.SetFocus
    Else
        If K = 38 Then
            Text23.SetFocus
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Text25_Change()
If Len((Text25)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text25_GotFocus()
SendKeys "{home}+{End}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Maximum Credit Limit"
End Sub

Private Sub Text25_KeyDown(KeyCode As Integer, s As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text25 = convert(amt(Val(Text25)))
    Text20.SetFocus
ElseIf KeyCode = 38 Then
    Text25 = convert(amt(Val(Text25)))
    List6.SetFocus
End If
End Sub

Private Sub Text25_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Group Account ID or Press <Enter> to get List of Group ID's"
End Sub

Private Sub text3_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub Text4_Change()
If Len(Text4) = 25 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text4_GotFocus()
SendKeys "{home}+{End}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Pager Number"
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, s As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    EMAILID.SetFocus
ElseIf KeyCode = 38 Then
    MOBILE.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 32 Or K = 45 Then
Else
K = 0
End If
End Sub

Private Sub Text5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Address"
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text6_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Address"
End Sub

Private Sub Text7_keypress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text8_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter City or Place"
End Sub

Private Sub Text8_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text10_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text11_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text12_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text9_Change()
If Len(Text9) = 25 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text9_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Telephone Number"
SendKeys "{home}+{end}"
End Sub

Private Sub Text11_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Contact Person"
End Sub

Private Sub Text10_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Central Sales Tax Number"
End Sub

Private Sub Text12_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Designation"
End Sub

Private Sub TEXT1_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then DISPDETAILS
End Sub

Private Sub Text3_KeyDown(K As Integer, s As Integer)
If CHECKTB("grp.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If K = 13 Or K = 40 Then
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
    rs.Open "select * from grp where grp_idy like '" & Sincrement(Trim(Text3)) & "'", con, adOpenKeyset
    If rs.RecordCount > 0 Then
        Text3 = Sincrement(Text3)
        Label4 = Trim(rs("grp_nme"))
        Text2.SetFocus
    Else
        grouplist
    End If
End If
End Sub

Private Sub List2_Keydown(KeyAscii As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii = 13 Then
    If List2.ListItems.count > 0 Then
      If rs.State = 1 Then rs.Close
        Dim lst As ListSubItem
        Set lst = List2.SelectedItem.ListSubItems.Item(1)
        Label4 = List2.SelectedItem
        Text3 = List2.SelectedItem.ListSubItems.Item(1)
        List2.SelectedItem.Ghosted = True
        SSTab1.TabEnabled(5) = False
        List2.Visible = False
        SSTab1.Tab = 0
        Text2.SetFocus
      End If
    End If
End Sub

Private Sub TEXT5_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text6.SetFocus
ElseIf K = 38 Then
SSTab1.Tab = 0
Text16.SetFocus
End If
End Sub

Private Sub TEXT6_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text7.SetFocus
ElseIf K = 38 Then
Text5.SetFocus
End If
End Sub

Private Sub Text7_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text8.SetFocus
ElseIf K = 38 Then
Text6.SetFocus
End If
End Sub

Private Sub Text8_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
ZIPCODE.SetFocus
ElseIf K = 38 Then
Text7.SetFocus
End If
End Sub

Private Sub Text9_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
FAX.SetFocus
ElseIf K = 38 Then
ZIPCODE.SetFocus
End If
End Sub

Private Sub Text10_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text16.SetFocus
ElseIf K = 38 Then
GST.SetFocus
End If
End Sub

Private Sub Text11_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text12.SetFocus
ElseIf K = 38 Then
WWW.SetFocus
End If
End Sub

Private Sub Text12_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    SSTab1.Tab = 2
    Check1.SetFocus
ElseIf K = 38 Then
    Text11.SetFocus
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
    If M_FRM_ACT.ActiveControl.name = "Text1" Then
        accountslist
        Check3.Visible = True
        Check4.Visible = True
        Check5.Visible = True
        Check6.Visible = True
        Check7.Visible = True
        Check8.Visible = True
        Check9.Visible = True

    ElseIf M_FRM_ACT.ActiveControl.name = "Text32" Then
        smnlist
        Check3.Visible = True
        Check4.Visible = True
        Check5.Visible = False
        Check6.Visible = False
        Check7.Visible = False
        Check8.Visible = False
        Check9.Visible = False

    End If
    If M_FRM_ACT.ActiveControl.name = "Text3" Then grouplist
    If M_FRM_ACT.ActiveControl.name = "AREAID" Then arealist
    If M_FRM_ACT.ActiveControl.name = "Text31" Then marlist
    If M_FRM_ACT.ActiveControl.name = "Text13" Then chanlist
ElseIf Button.KEY = "p" Then
    M_FRM_AQY.Tag = "ACT"
    M_FRM_AQY.Show 1
ElseIf Button.KEY = "pr" Then
    'MDIForm1.RRReport1.ReportName = App.Path & "\AccountS List.RRW"
    'MDIForm1.RRReport1.Destination = 1
    'MDIForm1.RRReport1.Printer = "?"
    'M'DIForm1.RRReport1.RunReport 1
    Exit Sub
er1:
    MsgBox "Accounts Not Found, Define Accounts...!", vbExclamation, "DAS Version 4.0"
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
    SYS_PSW.Tag = "ACT_PRE"
    SYS_PSW.Show 1
ElseIf Button.KEY = "crt" Then
    CREATE
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
With MDIForm1
    If ButtonMenu.KEY = "ACT_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "Accounts List", 0, "", 6
'        OpenWinReport "", "Accounts List.RRW", 0, "", 6 'CURDIR &
    ElseIf ButtonMenu.KEY = "RET_CAT" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers List (Category)", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "RET_LOC" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers List (Location)", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "RET_MAR" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers List (Market)", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "RET_ROU" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers List (Route)", 0, "", 6
    ElseIf ButtonMenu.KEY = "RET_SMN" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers List (Salesman)", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "SUP_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "Suppliers List", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "RET_ID_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers List ID", 2, "ACT->ACT_TYP = 'SAL' .OR. ACT->ACT_TYP = 'BOT' .AND. ACT->CST_TYP <> 'Van'", 6
    ElseIf ButtonMenu.KEY = "CLS_RET_ID" Then
        OpenReport CURDIR & "\LISTS.RP1", "Closed Outlets List", 2, "ACT->DBT_TYP = 'N' .AND. (ACT->ACT_TYP = 'SAL' .OR. ACT->ACT_TYP = 'BOT') .AND. ACT->CST_TYP <> 'Van'", 6
    ElseIf ButtonMenu.KEY = "VAN_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "Van List", 2, "ACT->CST_TYP = 'Van'", 6
    ElseIf ButtonMenu.KEY = "GRD_KEY_LST" Then
        Grade_Key_Count
        OpenReport CURDIR & "\LISTS.RP1", "Retail Count List", 0, "", 6
    ElseIf ButtonMenu.KEY = "NRL" Then
        M_FRM_DSN.Tag = "NRQ"
        M_FRM_DSN.Show 1
    ElseIf ButtonMenu.KEY = "wk" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers List (WeekDay)", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "RDY_LST" Then
        OpenReport CURDIR & "\LISTS.RP1", "Scheduled Routes", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "RRC" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers Route Count", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "MRC" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers Market Count", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "SRC" Then
        OpenReport CURDIR & "\LISTS.RP1", "Retailers Salesman Count", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "RSR" Then
        Dim REC_SET As New ADODB.Recordset
        Dim REC_SET1 As New ADODB.Recordset
        Dim COUNTPRD As Integer
        con.Execute "DELETE FROM TOP"
        If REC_SET1.State = 1 Then REC_SET1.Close
           REC_SET1.Open "select * from PRD where STK_TYP = 'Salable'", con, adOpenKeyset, adLockPessimistic
           COUNTPRD = REC_SET1.RecordCount
        If REC_SET.State = 1 Then REC_SET.Close
'           REC_SET.Open "select CST_IDY, COUNT(DISTINCT PRD_IDY) FROM INL, INV where INV.INV_IDY like INL.INV_IDY and INL.TXN_TYP = 'SAL' group by INL.CST_IDY", CON, adOpenKeyset, adlockpessimistic
           'REC_SET.Open "select CST_IDY, COUNT(DISTINCT PRD_IDY) FROM INL, INV where INV.INV_IDY like INL.INV_IDY and INL.TXN_TYP = 'SAL' group by INL.CST_IDY", CON, adOpenKeyset, adlockpessimistic
           REC_SET.Open "SELECT COUNT(DISTINCT PRD_IDY), CST_IDY FROM INV INNER JOIN INL ON INV.INV_IDY = INL.INV_IDY WHERE INV.TXN_TYP = 'SAL' GROUP BY INL.CST_IDY", con, adOpenKeyset, adLockPessimistic
        While REC_SET.EOF = False
           con.Execute "INSERT INTO TOP VALUES('" & Trim(REC_SET(1) & "") & "'," & Val(REC_SET(0)) & ", " & COUNTPRD & ")"
           REC_SET.MoveNext
        Wend
        OpenReport CURDIR & "\MIR.RP1", "Retailer SKU Range", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "GCC" Then
        OpenReport CURDIR & "\KOR.RP1", "GCC Record by Page", 0, SAL_QRY, 6
    End If
End With
cancel
End Sub
Private Sub Grade_Key_Count()
Dim CLD_RS As New ADODB.Recordset
Dim ACT_RS As New ADODB.Recordset
con.Execute "DELETE FROM RP1"
If CLD_RS.State = 1 Then CLD_RS.Close
CLD_RS.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic
If ACT_RS.State = 1 Then ACT_RS.Close
ACT_RS.Open "SELECT CST_TYP, count(CST_TYP) FROM ACT group by CST_TYP", con, adOpenKeyset, adLockPessimistic
If ACT_RS.RecordCount > 0 Then ACT_RS.MoveFirst
While Not ACT_RS.EOF()
    CLD_RS.AddNew
    CLD_RS("CHR_TEN") = Trim(ACT_RS(0))
    CLD_RS("NUM_ONE") = Trim(ACT_RS(1))
    CLD_RS.Update
    ACT_RS.MoveNext
Wend
If ACT_RS.State = 1 Then ACT_RS.Close
ACT_RS.Open "SELECT RET_TYP, count(RET_TYP) FROM ACT group by RET_TYP", con, adOpenKeyset, adLockPessimistic
If ACT_RS.RecordCount > 0 Then ACT_RS.MoveFirst
While Not ACT_RS.EOF()
    CLD_RS.AddNew
    CLD_RS("CHR_TEN") = Trim(ACT_RS(0))
    CLD_RS("NUM_ONE") = Trim(ACT_RS(1))
    CLD_RS.Update
    ACT_RS.MoveNext
Wend
If ACT_RS.State = 1 Then ACT_RS.Close
ACT_RS.Open "SELECT BBN_ONE, count(BBN_ONE) FROM ACT group by BBN_ONE", con, adOpenKeyset, adLockPessimistic
If ACT_RS.RecordCount > 0 Then ACT_RS.MoveFirst
While Not ACT_RS.EOF()
    CLD_RS.AddNew
    CLD_RS("CHR_TEN") = Trim(ACT_RS(0))
    CLD_RS("NUM_ONE") = Trim(ACT_RS(1))
    CLD_RS.Update
    ACT_RS.MoveNext
Wend
End Sub
Private Sub WWW_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Web Site"
End Sub

Private Sub WWW_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text11.SetFocus
ElseIf K = 38 Then
EMAILID.SetFocus
End If
End Sub

Private Sub ZIPCODE_Change()
If Len(ZIPCODE) = 6 Then SendKeys "{home}+{end}"
End Sub

Private Sub ZIPCODE_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Zip Code"
End Sub

Private Sub zipcode_keydown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
Text9.SetFocus
ElseIf K = 38 Then
Text8.SetFocus
End If
End Sub

Private Sub ZIPCODE_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text9_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 32 Or K = 45 Then
Else
K = 0
End If
End Sub

Private Sub fax_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 32 Or K = 45 Then
Else
K = 0
End If
End Sub

Public Function deletecheck() As Boolean
If CHECKTB("inl.dbf,vch.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Function
End If
If check("inl", "cst_idy", Trim(Text1)) = False Then
    If check("inl", "act_idy", Trim(Text1)) = False Then
        If check("vch", "crt_idy", Trim(Text1)) = False Then
            If check("vch", "dbt_idy", Trim(Text1)) = False Then
                deletecheck = False
            Else
                deletecheck = True
            End If
        Else
            deletecheck = True
        End If
    Else
        deletecheck = True
    End If
Else
    deletecheck = True
End If
End Function

Public Function Check_Asset_Code(VASTIDY, code) As Integer
Dim RSINW As New ADODB.Recordset
Dim RSOUT As New ADODB.Recordset
RSINW.Open "select count(*),asm_cod,txn_typ from ast where asm_idy like '" & Trim(VASTIDY) & "' GROUP by ASM_COD,txn_typ", con, adOpenKeyset, adLockPessimistic
Set RSOUT = RSINW.Clone
RSINW.filter = "txn_typ='AIN'"
RSOUT.filter = "txn_typ='AOU'"
For A = 1 To MSA.Rows - 1
    If RSOUT.RecordCount > 0 Then RSOUT.MoveFirst
    RSOUT.Find "ASM_COD='" & Trim(code) & "'"
    If RSOUT.EOF = False Then
        AOUCOUNT = Val(RSOUT(0) & "")
        AINCOUNT = Val(RSINW(0) & "")
        If Val(AINCOUNT) <> Val(AOUCOUNT) Then
            Check_Asset_Code = A
            MSA.Row = A
            MSA.SetFocus
            Exit Function
        End If
    Else
        If RSINW.RecordCount > 0 Then RSINW.MoveFirst
        RSINW.Find "ASM_COD='" & Trim(code) & "'"
        If RSINW.EOF = False Then
            Check_Asset_Code = A
            MSA.Row = A
            MSA.SetFocus
            Exit Function
        End If
    End If
Next
Check_Asset_Code = 0
End Function
Private Sub txtSmn_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Then
        If rs.State = 1 Then rs.Close
        'If Asc(MID(Trim(Text32), 1, 1)) = 13 Or Asc(MID(Trim(Text32), Len(Trim(Text32)), 1)) Then Text32 = ""
        rs.Open " select * from smn where smn_idy like '" & Sincrement(Trim(txtSmn)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            MS.TextMatrix(MS.Row, 3) = Sincrement(txtSmn)
            MS.Text = Trim(rs("smn_nme") & "")
            txtSmn.Visible = False
            MS.Col = MS.Col + 1
            If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row ' + 1 ' CHANGED BY PKB/PNR
        Else
            smnlist
        End If
'    Else
'        ms.Text = Trim(txtSmn)
'        txtSmn = ""
'        txtSmn.Visible = False
'        ms.SetFocus
'        If ms.Row <> ms.Rows - 1 Then ms.Row = ms.Row + 1
'        ms.Col = 1
'    End If
    txtSmn = ""
End If
End Sub

Private Sub Txtsmn_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub
Private Function FINDITEM(ISTR As String)
Dim K As Integer, f As Integer
f = 0
For K = 0 To cmbRoute.ListCount - 1
    If Trim(ISTR) = cmbRoute.List(K) Then
        f = 1
        Exit For
    End If
Next
FINDITEM = f
End Function
Private Sub GCC_FileXtract()
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

        
    ' * : Beginning of _GCC.TXT Creation Code for Text Data Export
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
    
    M_FIL_NME = "C:\DEX\" & Trim(M_WDS_IDY) & "_GCC.TXT"
    Set F_OMS = FS_SOU.OpenTextFile("C:\DEX\" & Trim(M_WDS_IDY) & "_GCC.TXT", 2, True)
    
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
    Rrs.Open "select * from GCC", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
    
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
       
       'F_OMS.WriteLine "DAS_ACT_IDY" & "|" & "NAME" & "|" & "GST" & "|" & "CST" & "|" & "ADDRESS1" & "|" & "ADDRESS2" & "|" & "ADDRESS3" & "|" & "ADDRESS4" & "|" & "ZIPCODE" & "|" & "TELEPHONE" & "|" & "CONTACT" & "|" & "LOCATION" & "|" & "MARKET" & "|" & "SALESMAN" & "|" & "CATEGORY" & "|" & "ROUTE" & "|" & "GRADE" & "|" & "ACTIVE" & "|" & "CAPACITY" & "|" & "RED CODE" & "|" & "HVO CODE" & "|" & "TYPE" & "|" & "CASH-CREDIT" & "|" & "MONDAY" & "|" & "TUESDAY" & "|" & "WEDNESDAY" & "|" & "THURSDAY" & "|" & "FRIDAY" & "|" & "SATURDAY" & "|" & "SUNDAY" & "|" & "CUSTOMER GROUP" & "|" & "RSA CLUSTER" & "|" & "RED OUTLET CLASS" & "|" & "TOWN CLASS" & "|" & "TOWN" & "|" & "DISTRICT" & "|" & "STATE" & "|" & "COUNTRY"
       
       F_OMS.WriteLine "GC_ID" & "|" & "Entity_ID" & "|" & "Site_ID" & "|" & "SE_ID" & "|" & "Distributor_Route_ID" & "|" & "Customer_Type" & "|" & "Customer_Name" & "|" & "Customer_Sht_Name" & "|" & "Address_1" & "|" & "Address_2" & "|" & "Locality ID" & "|" & "Locality Name" & "|" & "PIN_Code" & "|" & "CSS_Cluster_ID" & "|" & "RSA_Cluster_ID" & "|" & "Channel_ID" & "|" & "State_ID" & "|" & "District_ID" & "|" & "Town_ID" & "|" & "Contact_Person" & "|" & "TelePhone" & "|" & "Mobile_No" & "|" & "NKA_Customer_ID" & "|" & "VPO_Class_ID" & "|" & "Market_Area_ID" & "|" & "Market_Area_Name" & "|" & "Territory_ID" & "|" & "Territory_Name" & "|" & "MD_ID" & "|" & "MD_Name" & "|" & "DAS_Customer_ID" & "|" & "Email_ID" & "|" & "Latitude" & "|" & "Longitude"
       
       
       While Not Rrs.EOF
       
          If Location_RS.RecordCount > 0 Then Location_RS.MoveFirst
          Location_RS.Find "LOC_IDY='" & Trim(Rrs("LOC_ID")) & "'"
          If Location_RS.EOF = False Then
             L_LOC_NME = Location_RS("LOC_NME")
          End If
          
          If Market_RS.RecordCount > 0 Then Market_RS.MoveFirst
          Market_RS.Find "MAR_IDY='" & Trim(Rrs("MAR_ID")) & "'"
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
          RDY_RS.Find "ACT_IDY='" & Trim(Rrs("DAS_ID")) & "'"
          If RDY_RS.EOF = False Then
             L_SMN_IDY = Trim(RDY_RS("SMN_IDY"))
             L_ROU_IDY = Trim(RDY_RS("ROU_IDY"))
          Else
             L_SMN_IDY = ""
             L_ROU_IDY = ""
          End If
          
          While Not RDY_RS.EOF
          If RDY_RS("ACT_IDY") = Trim(Rrs("DAS_ID")) Then
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
          Category_RS.Find "CAT_IDY='" & Trim(Rrs("CHN_ID")) & "'"
          If Category_RS.EOF = False Then
             L_CAT_NME = Category_RS("CAT_NME")
          End If


'          F_OMS.WriteLine Trim(M_WDS_IDY) & "|" & Trim(Rrs("ACT_IDY")) & "|" & Trim(Rrs("ACT_NME")) & "|" & Trim(Rrs("ACT_AD1")) & "|" & Trim(Rrs("ACT_AD2")) & "|" & Trim(Rrs("ACT_AD3")) & "|" & Trim(Rrs("ACT_AD4")) & "|" & Trim(L_DST_NME) & "|" & Trim(Rrs("ACT_TFX")) & "|" & Trim(Rrs("ACT_MBL")) & "|" & Trim(Rrs("ACT_EML")) & "|" & Trim(Rrs("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(Rrs("BBN_ONE")) & "|" & Trim(Rrs("CST_TYP")) & "|" & Trim(Rrs("RET_TYP")) & "|" & Trim(Rrs("VAN_CAP")) & "|" & Trim(Rrs("DBT_TYP"))
          
'          If Trim(Rrs("ACT_TYP")) = "PUR" Then
'                F_OMS.WriteLine Trim(Rrs("ACT_IDY")) & "|" & "S" & "|" & Trim(Rrs("CST_TYP")) & "|" & Trim(Rrs("ACT_NME")) & "|" & Trim(Rrs("ACT_AD1")) & "|" & Trim(Rrs("ACT_AD2")) & "|" & Trim(Rrs("ACT_AD3")) & "|" & Trim(Rrs("ACT_AD4")) & "|" & Trim(Rrs("ACT_TFX")) & "|" & Trim(Rrs("ACT_MBL")) & "|" & Trim(Rrs("ACT_EML")) & "|" & Trim(Rrs("ACT_CPN")) & "|" & Trim(Rrs("LOC_IDY")) & "|" & Trim(Rrs("MAR_IDY")) & "|" & Trim(Rrs("CAT_IDY")) & "|" & Trim(Rrs("RET_TYP")) & "|" & Trim(Rrs("BBN_ONE")) & "|" & Trim(Rrs("DBT_TYP")) & "|" & Trim(Rrs("VAN_CAP")) & "|" & Trim(Rrs("ROU_IDY"))
'                J = J + 1
'          End If
          
          'If Trim(Rrs("ACT_TYP")) = "SAL" Then
                'F_OMS.WriteLine Trim(Rrs("ACT_IDY")) & "|" & Trim(Rrs("ACT_NME")) & "|" & Trim(Rrs("ACT_GST")) & "|" & Trim(Rrs("ACT_CST")) & "|" & Trim(Rrs("ACT_AD1")) & "|" & Trim(Rrs("ACT_AD2")) & "|" & Trim(Rrs("ACT_AD3")) & "|" & Trim(Rrs("ACT_AD4")) & "|" & Trim(Rrs("ACT_ZIP")) & "|" & Trim(Rrs("ACT_TFX")) & "-" & Trim(Rrs("ACT_MBL")) & "|" & Trim(Rrs("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(L_SMN_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_ROU_IDY) & "|" & Trim(Rrs("RET_TYP")) & "|" & Trim(Rrs("DBT_TYP")) & "|" & Trim(Rrs("VAN_CAP")) & "|" & Trim(Rrs("ACN_TWO")) & "|" & Trim(Rrs("BNM_TWO")) & "|" & Trim(Rrs("CST_TYP")) & "|" & Trim(Rrs("MCL_REQ")) & "|" & D_MON & "|" & D_TUE & "|" & D_WED & "|" & D_THU & "|" & D_FRI & "|" & D_SAT & "|" & D_SUN & "|" & L_MAR_CSG & "|" & L_MAR_CLU & "|" & L_MAR_ROC & "|" & L_MAR_TWC & "|" & L_MAR_TWN & "|" & L_MAR_DST & "|" & L_MAR_STA & "|" & L_MAR_COU
                
                F_OMS.WriteLine Trim(Rrs("GCC_ID")) & "|" & Trim(Rrs("ENT_ID")) & "|" & Trim(Rrs("SIT_ID")) & "|" & Trim(Rrs("SE_ID")) & "|" & Trim(Rrs("DRI_ID")) & "|" & Trim(Rrs("CST_TYP")) & "|" & Trim(Rrs("CST_NME")) & "|" & Trim(Rrs("CST_SME")) & "|" & Trim(Rrs("ADR_ONE")) & "|" & Trim(Rrs("ADR_TWO")) & "|" & Trim(Rrs("LOC_ID")) & "|" & Trim(Rrs("LOC_NME")) & "|" & Trim(Rrs("PIN_COD")) & "|" & Trim(Rrs("CSS_CST")) & "|" & Trim(Rrs("RSA_CST")) & "|" & Trim(Rrs("CHN_ID")) & "|" & Trim(Rrs("STA_ID")) & "|" & Trim(Rrs("DST_ID")) & "|" & Trim(Rrs("TWN_ID")) & "|" & Trim(Rrs("CNT_CPN")) & "|" & Trim(Rrs("TEL_FAX")) & "|" & Trim(Rrs("MBL_NUM")) & "|" & Trim(Rrs("NKA_ID")) & "|" & Trim(Rrs("VPO_ID")) & "|" & Trim(Rrs("MAR_ID")) & "|" & Trim(Rrs("MAR_NME")) & "|" & Trim(Rrs("TER_ID")) & "|" & Trim(Rrs("TER_NME")) & "|" & Trim(Rrs("MD_ID")) & "|" & Trim(Rrs("MD_NME")) & "|" & Trim(Rrs("DAS_ID")) & "|" & Trim(Rrs("EML_ID")) & "|" & Trim(Rrs("LAT_ID")) & "|" & Trim(Rrs("LON_ID"))
                J = J + 1
          'End If
          
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
        cancel
    
    ' * : End of _GCC.TXT Creation Code for Text Data Export
End Sub
Private Sub GCC_FileImport()

Dim rs As New ADODB.Recordset
Dim Pgr_RS As New ADODB.Recordset
Dim BAT_RS As New ADODB.Recordset
Dim ACT_RS As New ADODB.Recordset

Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer
M_WDS_IDY = GetValue("DIR_ALS", "DIR WHERE DIR_IDY = 'WDS' AND DIR_IDX = 'IDY'")
If Len(Trim(M_WDS_IDY)) < 8 Then
    MsgBox ("Distributor ID Required for this Process! Obtain Distributor ID and Start Again"), vbCritical, "DAS Version 4.0"
    Exit Sub
End If
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\" & Trim(M_WDS_IDY) & "_GCC.TXT") = True Then
    FILE_NAME = "C:\DEX\" & Trim(M_WDS_IDY) & "_GCC.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("GCC Text File for Import Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
    Do
        Line Input #1, input_data
        TOT_REC = TOT_REC + 1
    Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("GCC Text File for Import Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    
    Dim M_GRP_IDY As String
    Dim M_PRD_IDY As String
    Dim M_ACT_IDY As String
    
    Dim ACT_NME As String
    Dim GRP_CAT As String
    Dim GRP_CLS As String
    Dim PRD_NME As String
    Dim ALT_IDY As String
    Dim BRD_NME As String
    Dim PRD_PCK As String
    Dim PRD_PKT As String
    Dim PUR_RAT As String
    Dim PUR_TAX As String
    Dim PUR_DIS As String
    Dim SAL_RAT As String
    Dim SAL_TAX As String
    Dim SAL_DIS As String
    Dim MRP_RAT As String
    Dim VOL_WGT As String
    Dim PRD_UOM As String
    Dim SUB_UNT As Integer
    Dim TOD_PCG As String
    Dim LST_RAT As String
    Dim LST_DIS As String
    Dim SAF_LEV As String
    Dim SAF_DAY As Integer
    Dim stk_typ As String
    
    Dim STR_CMP As String
    
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\" & Trim(M_WDS_IDY) & "_GCC.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        STR_CMP = Mid(Trim(INL_STR), STR_VAL, STR_LEN)
        If STR_CMP = "GCC" Then
            For M_CTR = 1 To 34
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> "|" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            GC_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            Entity_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 3 Then
                            Site_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 4 Then
                            SE_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 5 Then
                            Distributor_Route_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 6 Then
                            Customer_Type = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 7 Then
                            Customer_Name = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 8 Then
                            Customer_Sht_Name = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 9 Then
                            Address_1 = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 10 Then
                            Address_2 = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 11 Then
                            Locality_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 12 Then
                            Locality_Name = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 13 Then
                            PIN_Code = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 14 Then
                            CSS_Cluster_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 15 Then
                            RSA_Cluster_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 16 Then
                            Channel_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 17 Then
                            State_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 18 Then
                            District_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 19 Then
                            Town_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 20 Then
                            Contact_Person = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 21 Then
                            TelePhone = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 22 Then
                            Mobile_No = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 23 Then
                            NKA_Customer_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 24 Then
                            VPO_Class_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 25 Then
                            Market_Area_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 26 Then
                            Market_Area_Name = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 27 Then
                            Territory_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 28 Then
                            Territory_Name = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 29 Then
                            MD_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 30 Then
                            MD_Name = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 31 Then
                            DAS_Customer_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 32 Then
                            Email_ID = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 33 Then
                            Latitude = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 34 Then
                            Longitude = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid GCC Record Encountered!, Contact Support !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Entries in GCC Updated "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from GCC where GCC_ID = '" & Trim(GC_ID) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.EOF = True Then
            If UCase(Trim(Distributor_Route_ID)) = Trim(M_WDS_IDY) Then
                rs.AddNew
                rs("GCC_ID") = UCase(GC_ID)
                rs("ENT_ID") = UCase(Entity_ID)
                rs("SIT_ID") = UCase(Site_ID)
                rs("SE_ID") = UCase(SE_ID)
                rs("DRI_ID") = UCase(Distributor_Route_ID)
                rs("CST_TYP") = UCase(Customer_Type)
                rs("CST_NME") = UCase(Customer_Name)
                rs("CST_SME") = UCase(Customer_Sht_Name)
                rs("ADR_ONE") = UCase(Address_1)
                rs("ADR_TWO") = UCase(Address_2)
                rs("LOC_ID") = Mid(UCase(Locality_ID), 1, 10)
                rs("LOC_NME") = UCase(Locality_Name)
                rs("PIN_COD") = UCase(PIN_Code)
                rs("CSS_CST") = UCase(CSS_Cluster_ID)
                rs("RSA_CST") = UCase(RSA_Cluster_ID)
                rs("CHN_ID") = Mid(UCase(Channel_ID), 1, 10)
                rs("STA_ID") = UCase(Trim(State_ID))
                rs("DST_ID") = UCase(Trim(District_ID))
                rs("TWN_ID") = UCase(Trim(Town_ID))
                rs("CNT_CPN") = UCase(Trim(Contact_Person))
                rs("TEL_FAX") = UCase(Trim(TelePhone))
                rs("MBL_NUM") = UCase(Trim(Mobile_No))
                rs("NKA_ID") = UCase(Trim(NKA_Customer_ID))
                rs("VPO_ID") = UCase(Trim(VPO_Class_ID))
                rs("MAR_ID") = UCase(Trim(Market_Area_ID))
                rs("MAR_NME") = UCase(Trim(Market_Area_Name))
                rs("TER_ID") = UCase(Trim(Territory_ID))
                rs("TER_NME") = UCase(Trim(Territory_Name))
                rs("MD_ID") = UCase(Trim(MD_ID))
                rs("MD_NME") = UCase(Trim(MD_Name))
                rs("DAS_ID") = UCase(Mid(Trim(DAS_Customer_ID), 1, 10))
                rs("EML_ID") = Trim(Email_ID)
                rs("LAT_ID") = Trim(Latitude)
                rs("LON_ID") = Trim(Longitude)
                rs.Update
                TOT_UPD = TOT_UPD + 1
                If rs.State = 1 Then rs.Close
                rs.Open "SELECT * FROM ACT where ACT.GCC_ID LIKE '" & Trim(GC_ID) & "'", con, adOpenKeyset, adLockPessimistic
                If rs.EOF = True Then
                    rs.AddNew
                    rs("ACT_NME") = UCase(Customer_Name)
                    rs("GRP_IDY") = "0000000007"
                    rs("GRP_CLF") = "Debtor"
                    rs("ACT_AD1") = UCase(Address_1)
                    rs("ACT_AD2") = UCase(Address_2)
                    rs("ACT_CPN") = UCase(Trim(Contact_Person))
                    rs("ACT_TFX") = UCase(Trim(TelePhone))
                    rs("ACT_ZIP") = Val(Trim(UCase(PIN_Code)))
                    rs("ACT_EML") = Trim(Email_ID)
                    rs("ACT_MBL") = UCase(Trim(Mobile_No))
                    rs("ACT_TYP") = "SAL"
                    rs("DBT_DIS") = 0#
                    rs("DBT_MCP") = 0#
                    rs("DBT_TYP") = "Y"
                    rs("DBT_RAP") = "B"
                    rs("MDL_REQ") = ""
                    rs("MCL_REQ") = "Y"
                    rs("CRT_MDL") = 0#
                    rs("CRT_PPD") = 0#
                    rs("DBT_MCL") = 0#
                    rs("LOC_IDY") = ""
                    rs("CRT_TYP") = ""
                    rs("CRT_DIS") = 0#
                    rs("CRT_RAP") = ""
                    rs("ZON_IDY") = ""
                    rs("MAR_IDY") = ""
                    rs("SMN_IDY") = ""
                    rs("ACT_RGD") = ""
                    rs("ACN_ONE") = ""
                    rs("BNM_ONE") = ""
                    rs("BBN_ONE") = ""
                    rs("ACN_TWO") = ""
                    rs("BNM_TWO") = ""
                    rs("BBN_TWO") = ""
                    rs("CAT_IDY") = ""
                    rs("KEY_ACT") = "N"
                    rs("CST_TYP") = "Retailer"
                    rs("RET_TYP") = "GOLD"
                    rs("VAN_CAP") = 0#
                    rs("ACT_DOC") = Date$
                    rs("SKU_RNG") = 0#
                    rs("GCC_ID") = UCase(GC_ID)
                    rs("ACT_IDY") = increment("idy_act")
                    rs.Update
                End If
            End If
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Entries in GCC Added - Updated "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\" & Trim(M_WDS_IDY) & "_GCC.TXT") = True Then
'        If MsgBox("Delete GCC Text File ?", vbQuestion + vbOKCancel, "DAS Version 4.0") = vbOK Then
'            FS_SOU.DeleteFile ("C:\DEX\" & Trim(M_WDS_IDY) & "_GCC.TXT")
'        End If
    End If
End Sub

'            If ACT_RS.State = 1 Then ACT_RS.Close
'            ACT_RS.Open "select * from ACT where alltrim(ACT_NME) = '" & UCase(Trim(ACT_NME)) & "'", CON, adOpenKeyset, adLockPessimistic
'            If ACT_RS.RecordCount > 0 Then
'                M_ACT_IDY = ACT_RS("ACT_IDY")
'            Else
'                ACT_RS.AddNew
'                M_ACT_IDY = increment("IDY_ACT")
'                ACT_RS("ACT_IDY") = M_ACT_IDY
'                ACT_RS("GRP_IDY") = "0000000008"
'                ACT_RS("act_nme") = UCase(ACT_NME)
'                ACT_RS("act_ad1") = ""
'                ACT_RS("act_ad2") = ""
'                ACT_RS("act_ad3") = ""
'                ACT_RS("act_ad4") = ""
'                ACT_RS("act_zip") = ""
'                ACT_RS("cst_typ") = "Retailer"
'                ACT_RS("KEY_ACT") = "N"
'                ACT_RS("ACT_TYP") = "PUR"
'                ACT_RS("act_tfx") = ""
'                ACT_RS("act_fax") = ""
'                ACT_RS("act_mbl") = ""
'                ACT_RS("act_pgr") = ""
'                ACT_RS("act_eml") = ""
'                ACT_RS("ACT_WWW") = ""
'                ACT_RS("act_cpn") = ""
'                ACT_RS("act_dsg") = ""
'                ACT_RS("DBT_TYP") = "Y"
'                ACT_RS("ACT_DOC") = Date
'                ACT_RS("act_gst") = ""
'                ACT_RS("act_cst") = ""
'                ACT_RS("act_rgd") = ""
'                ACT_RS("acn_one") = ""
'                ACT_RS("bnm_one") = ""
'                ACT_RS("bbn_one") = ""
'                ACT_RS("acn_two") = ""
'                ACT_RS("bnm_two") = ""
'                ACT_RS("bbn_two") = ""
'                ACT_RS("ret_typ") = ""
'                ACT_RS("van_cap") = "0"
'                ACT_RS("GRP_CLF") = "Creditor"
'                ACT_RS("LOC_IDY") = ""
'                ACT_RS("CAT_idy") = ""
'                ACT_RS("mar_idy") = ""
'                ACT_RS("smn_idy") = ""
'                ACT_RS("rou_idy") = ""
'                ACT_RS("zon_idy") = ""
'                ACT_RS("dbt_DIs") = "0"
'                ACT_RS("mcl_req") = "N"
'                ACT_RS("dbt_mcl") = "0.00"
'                ACT_RS("dbt_mcp") = "0.00"
'                ACT_RS("dbt_rap") = ""
'                ACT_RS("crt_dis") = "0.00"
'                ACT_RS("mdl_req") = "N"
'                ACT_RS("crt_mdl") = "0.00"
'                ACT_RS("crt_ppd") = 0
'                ACT_RS.Update

