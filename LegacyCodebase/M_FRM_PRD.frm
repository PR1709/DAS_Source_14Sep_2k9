VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_PRD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products - SKU"
   ClientHeight    =   5595
   ClientLeft      =   5610
   ClientTop       =   3735
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_PRD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Tag             =   "D"
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Product - SKU Details"
      TabPicture(0)   =   "M_FRM_PRD.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label17"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label36"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label37"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label35"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label34"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label38"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label30"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label20"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label21"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label22"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label23"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label24"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label26"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "M_DAT_DOM"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "M_DAT_DOE"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text5"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text4"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text3"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text10"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text11"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text12"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text26"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "List5"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text15"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "RAT_FRM"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text6"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Combo1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Combo2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Frame2"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text14"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "excinf"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "List7"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "&Sequence"
      TabPicture(1)   =   "M_FRM_PRD.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text16"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MS_SEQ"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "M_FRM_PRD.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label31"
      Tab(2).Control(1)=   "list1"
      Tab(2).Control(2)=   "list2"
      Tab(2).Control(3)=   "list3"
      Tab(2).Control(4)=   "Frame1"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "&Batches Browse"
      TabPicture(3)   =   "M_FRM_PRD.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "bname"
      Tab(3).Control(1)=   "date1"
      Tab(3).Control(2)=   "List6"
      Tab(3).Control(3)=   "Text28"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Text27"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "MS"
      Tab(3).ControlCount=   6
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         ItemData        =   "M_FRM_PRD.frx":04B2
         Left            =   7560
         List            =   "M_FRM_PRD.frx":04D4
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   780
         Width           =   1215
      End
      Begin VB.ListBox excinf 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         ItemData        =   "M_FRM_PRD.frx":0539
         Left            =   8130
         List            =   "M_FRM_PRD.frx":0543
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   4800
         Width           =   615
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5280
         TabIndex        =   83
         Top             =   1250
         Width           =   465
      End
      Begin VB.Frame Frame2 
         Caption         =   "Taxes and Discounts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   6600
         TabIndex        =   71
         Top             =   1440
         Width           =   2310
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "M_FRM_PRD.frx":054D
            Left            =   960
            List            =   "M_FRM_PRD.frx":054F
            TabIndex        =   81
            Top             =   1710
            Width           =   1215
         End
         Begin VB.TextBox Text25 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   930
            Width           =   1215
         End
         Begin VB.TextBox PUR_DIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   200
            Width           =   1215
         End
         Begin VB.TextBox PUR_TAX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   560
            Width           =   1215
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tax Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   80
            Top             =   1740
            Width           =   675
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sale Tax %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   45
            TabIndex        =   79
            Top             =   1395
            Width           =   795
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sale CD%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   78
            Top             =   1005
            Width           =   705
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pur. CD%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   165
            TabIndex        =   77
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pur.Tax %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   615
            Width           =   720
         End
      End
      Begin VB.TextBox bname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74400
         MaxLength       =   20
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   330
         Left            =   -72840
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
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
         Format          =   80281603
         CurrentDate     =   38808
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2055
         TabIndex        =   64
         Top             =   2005
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "M_FRM_PRD.frx":0551
         Left            =   2055
         List            =   "M_FRM_PRD.frx":0553
         TabIndex        =   62
         Top             =   1650
         Width           =   1815
      End
      Begin VB.ComboBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "M_FRM_PRD.frx":0555
         Left            =   7560
         List            =   "M_FRM_PRD.frx":0557
         TabIndex        =   61
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Frame RAT_FRM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   50
         Top             =   3120
         Width           =   3135
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1940
            MaxLength       =   9
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   495
            Width           =   945
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1940
            MaxLength       =   9
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   855
            Width           =   945
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1940
            MaxLength       =   9
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1940
            MaxLength       =   9
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   150
            Width           =   945
         End
         Begin VB.TextBox BAT_NME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1940
            MaxLength       =   20
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1560
            Width           =   945
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MRP - C - CASE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   315
            TabIndex        =   60
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Rate - B - CASE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   59
            Top             =   900
            Width           =   1380
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rate - A - CASE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   58
            Top             =   540
            Width           =   1140
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pur Rate - CASE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   285
            TabIndex        =   57
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   585
            TabIndex        =   56
            Top             =   1620
            Width           =   885
         End
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   -74040
         MaxLength       =   5
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         ItemData        =   "M_FRM_PRD.frx":0559
         Left            =   -72840
         List            =   "M_FRM_PRD.frx":0563
         TabIndex        =   39
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   -73440
         MaxLength       =   3
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox Text27 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   -74760
         MaxLength       =   9
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         MaxLength       =   9
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1215
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         ItemData        =   "M_FRM_PRD.frx":056D
         Left            =   7560
         List            =   "M_FRM_PRD.frx":0583
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3990
         Width           =   1215
      End
      Begin VB.TextBox Text26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         MaxLength       =   3
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4340
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   440
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   -74880
         TabIndex        =   23
         Top             =   4560
         Width           =   8295
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            Caption         =   "Promotional Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5760
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            Caption         =   "POP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4440
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Product Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   200
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Short Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2640
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2640
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1065
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2060
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2060
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2375
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2060
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1185
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2060
         MaxLength       =   20
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1250
         Width           =   1785
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2060
         MaxLength       =   40
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   860
         Width           =   3705
      End
      Begin MSComctlLib.ListView list3 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6165
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
            Text            =   "Supplier Name"
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
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6165
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
            Text            =   "Product Category"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Product Segment"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6165
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product (SKU) Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Short Name"
            Object.Width           =   1413
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "UPC"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Vol.Wgt"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "UOM"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Segment"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Brand Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Supplier Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Suplr.Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5040
         Top             =   -120
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
               Picture         =   "M_FRM_PRD.frx":05A8
               Key             =   "p"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":0AEC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":0EB4
               Key             =   "c"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":1208
               Key             =   "e"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":1840
               Key             =   "s"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":1B94
               Key             =   "h"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":1FF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":22E8
               Key             =   "pr"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":26F4
               Key             =   "d"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":2800
               Key             =   "pre"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":2B54
               Key             =   "brw"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_PRD.frx":2F68
               Key             =   "cpy"
            EndProperty
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4755
         Left            =   -74950
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   8387
         _Version        =   393216
         Cols            =   36
         FixedCols       =   2
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   12582912
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   1
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
      End
      Begin MSFlexGridLib.MSFlexGrid MS_SEQ 
         Height          =   4755
         Left            =   -74935
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   360
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   8387
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   4210752
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         GridLines       =   2
         AllowUserResizing=   1
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
      End
      Begin MSComCtl2.DTPicker M_DAT_DOE 
         Height          =   330
         Left            =   4440
         TabIndex        =   46
         Top             =   3600
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   80281603
         CurrentDate     =   38808
      End
      Begin MSComCtl2.DTPicker M_DAT_DOM 
         Height          =   330
         Left            =   4440
         TabIndex        =   48
         Top             =   3240
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   80281603
         CurrentDate     =   38808
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   " Shelf Life - Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   82
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excise - TCom, Spl SKU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   68
         Top             =   4815
         Width           =   1710
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pack Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   885
         TabIndex        =   65
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pack Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   825
         TabIndex        =   63
         Top             =   2065
         Width           =   780
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mfg. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3510
         TabIndex        =   49
         Top             =   3315
         Width           =   705
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BBD-Expiry "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3375
         TabIndex        =   47
         Top             =   3675
         Width           =   840
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4440
         TabIndex        =   45
         Top             =   4665
         Width           =   1065
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6510
         TabIndex        =   42
         Top             =   795
         Width           =   960
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier's Product Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5775
         TabIndex        =   41
         Top             =   470
         Width           =   1695
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volume - Weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6300
         TabIndex        =   40
         Top             =   3660
         Width           =   1170
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Units Per Case - UPC)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5865
         TabIndex        =   34
         Top             =   4475
         Width           =   1605
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6780
         TabIndex        =   31
         Top             =   4280
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UOM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   190
         Left            =   7095
         TabIndex        =   30
         Top             =   4000
         Width           =   375
      End
      Begin VB.Label Label31 
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   600
         Width           =   8775
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3480
         TabIndex        =   16
         Top             =   2375
         Width           =   2235
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3480
         TabIndex        =   15
         Top             =   2760
         Width           =   2235
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   518
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Segment ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   765
         TabIndex        =   13
         Top             =   2443
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   825
         TabIndex        =   12
         Top             =   2828
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Short Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   765
         TabIndex        =   11
         Top             =   1318
         Width           =   840
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product - SKU Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   930
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6585
         TabIndex        =   9
         Top             =   1155
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   8
         Top             =   4020
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-order Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3315
         TabIndex        =   7
         Top             =   4380
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3390
         TabIndex        =   6
         Top             =   4740
         Width           =   825
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
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
                  Key             =   "PRD_LST"
                  Text            =   "&a Products List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BAT_LST"
                  Text            =   "&b Batch List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRC_LST"
                  Text            =   "&c Price List"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "STK_LST"
                  Text            =   "&d Stock List"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RAT_LST"
                  Text            =   "&e Rates List"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LEV_REO"
                  Text            =   "&f Re-Order Level"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LEV_SAF"
                  Text            =   "&g Safety Level"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BAL_PRD"
                  Text            =   "&h Balance Product"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BAL_SEG"
                  Text            =   "&i Balance Segment"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BAL_VAL"
                  Text            =   "&j Balance and Value"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LST_IDY"
                  Text            =   "&k ID List"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DMG_STK"
                  Text            =   "&l Damage Stock - SKU"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BRD_DMG"
                  Text            =   "&m Damage Stock - Brand"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SEG_DMG"
                  Text            =   "&n Damage Stock - Segment"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CAT_DMG"
                  Text            =   "&o Damage Stock - Category"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SKU_BRK"
                  Text            =   "&p SKU Stock Breakup"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SNV_STT"
                  Text            =   "&q Stock and Value Statement"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRD_PKS"
                  Text            =   "&r Stock - Pack Size"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRD_PKT"
                  Text            =   "&s Stock - Pack Type"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Prefix Product ID "
            ImageIndex      =   10
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A"
                  Text            =   "Purchase Rate Check"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "B"
                  Text            =   "Prefix Product ID"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.CommandButton Command2 
         Caption         =   "&Update Batch"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6360
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Validate Batch"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7800
         TabIndex        =   86
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "M_FRM_PRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, idnum As Boolean, ls As ListItem

Private Sub BROWSE()
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from prd,brd where prd.prd_idy like brd.prd_idy order by prd.prd_idy", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    MS.Rows = rs.RecordCount + 1
    For A = 1 To rs.RecordCount
        MS.TextMatrix(A, 0) = Trim(rs("prd_idy"))
        MS.TextMatrix(A, 4) = Trim(rs("alt_idy") & "")
        MS.TextMatrix(A, 2) = Trim(rs("prd_prn"))
        MS.TextMatrix(A, 1) = Trim(rs("prd_nme"))
        MS.TextMatrix(A, 3) = Trim(rs("brd_nme") & "")
        MS.TextMatrix(A, 5) = Trim(rs("grp_idy"))
        MS.TextMatrix(A, 6) = Trim(rs("bat_nme") & "")
        MS.TextMatrix(A, 7) = convert(amt(rs("prd_pdr") & ""))
        MS.TextMatrix(A, 8) = convert(amt(rs("prd_sdr") & ""))
        MS.TextMatrix(A, 9) = convert(amt(rs("prd_mrp") & ""))
        MS.TextMatrix(A, 10) = FAC(rs("con_fac") & "")
        MS.TextMatrix(A, 11) = QTY(rs("saf_lev") & "")
        MS.TextMatrix(A, 12) = QTY(rs("reo_lev") & "")
        MS.TextMatrix(A, 13) = QTY(rs("PRD_QTY") & "")
        MS.TextMatrix(A, 14) = Trim(rs("prd_uom"))
        MS.TextMatrix(A, 15) = Trim(Val(rs("sub_unt") & ""))
        MS.TextMatrix(A, 16) = convert(amt(Val(rs("bas_rat") & "")))
        MS.TextMatrix(A, 17) = discount(rs("tod_pcg") & "")
        MS.TextMatrix(A, 18) = discount(rs("pur_dis") & "")
        MS.TextMatrix(A, 19) = discount(rs("pur_tax") & "")
        MS.TextMatrix(A, 20) = discount(rs("sal_dis") & "")
        MS.TextMatrix(A, 21) = discount(rs("sal_tax") & "")
        MS.TextMatrix(A, 22) = convert(amt(rs("lst_rat") & ""))
        MS.TextMatrix(A, 23) = discount(rs("lst_dis") & "")
        MS.TextMatrix(A, 24) = discount(rs("lst_cst") & "")
        MS.TextMatrix(A, 25) = discount(rs("add_Dis") & "")
        MS.TextMatrix(A, 26) = discount(rs("cas_dis") & "")
        MS.TextMatrix(A, 27) = discount(rs("exc_add") & "")
        MS.TextMatrix(A, 28) = Trim(rs("exc_inf") & "")
        MS.TextMatrix(A, 29) = Trim(rs!BAT_IDY)
        MS.TextMatrix(A, 30) = Format(rs("brd_dom"), "dd/MMM/yyyy")
        MS.TextMatrix(A, 31) = Format(rs("brd_doe"), "dd/MMM/yyyy")
        If IsNull(rs!P_INC_DIS) = False Then
            MS.TextMatrix(A, 32) = convert(rs("P_INC_DIS"))
        Else
            MS.TextMatrix(A, 32) = convert(0#)
        End If
        If IsNull(rs!S_INC_DIS) = False Then
            MS.TextMatrix(A, 33) = convert(rs("S_INC_DIS"))
        Else
            MS.TextMatrix(A, 33) = convert(0#)
        End If
        If IsNull(rs!LST_RAT) = False Then
            MS.TextMatrix(A, 34) = convert(rs!LST_RAT)
        Else
            MS.TextMatrix(A, 34) = convert(0#)
        End If
        If IsNull(rs!LST_DIS) = False Then
            MS.TextMatrix(A, 35) = convert(rs!LST_DIS)
        Else
            MS.TextMatrix(A, 35) = convert(0#)
        End If
        rs.MoveNext
    Next
    If SSTab1.Tab <> 3 Then SSTab1.Tab = 3
    MS.Visible = True
    MS.SetFocus
    MS.Col = 2
    MS.Row = 1
Else
    MsgBox "Product, SKU Not Found!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        Text5.SetFocus
    End If
    Exit Sub
End If
End Sub

Private Sub CREATE()
MOD1 = False
Text1 = ""
save
End Sub

Private Sub setval()
MOD1 = False
Text7 = "0.00"
Text8 = "0.00"
Text9 = "0.00"
Text10 = "0.000"
Text11 = "0.000"
Text14 = "000"
Text13 = "0.00"
Text15 = "0.000"
Text19 = "0.00"
Text20 = "0.00"
Text21 = "0.00"
Text22 = "0.00"
Text23 = "0.00"
Text24 = "0.00"
Text25 = "0.00"
PUR_TAX = "0.00"
PUR_DIS = "0.00"
Text26 = "0"
Label11 = ""
Label12 = ""
Label30 = ""

List5.Selected(0) = True
Toolbar1.Buttons(12).Enabled = False
Toolbar1.Buttons(13).Enabled = True
End Sub

Public Sub DEL_REC()
If CHECKTB("prd.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Are you sure you want to delete Product ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
    If check("prd", "prd_idy", Text1) = False Then
        MsgBox "Product Not Found, Define Product !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        clrctr M_FRM_PRD
        Text1.Enabled = True
        Text1.SetFocus
        Exit Sub
    Else
        ' Check for TWP Version NO DELETE
        Dim DIR_RS As New ADODB.Recordset
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            DIR_RS.MoveFirst
            If Trim(DIR_RS("DIR_ALS")) = "Y" Then
                If Trim(DIR_RS("DIR_MSG")) <> "WRITETWP" Then
                    MsgBox ("Cannot Delete Master Data for KO-TWP Version!"), vbCritical, "DAS Version 4.0"
                    cancel
                    Exit Sub
                Else
                    If check("inv", "prd_idy", Trim(Text1)) = False Then
                        con.Execute "delete from prd where prd_idy like '" & Trim(Text1) & "'"
                        con.Execute "delete from Brd where prd_idy like '" & Trim(Text1) & "'"
                        SSTab1.Tab = 0
                        Text1.Enabled = True
                        Text1.SetFocus
                        clrctr M_FRM_PRD
                    Else
                        MsgBox "Cannot Delete, Transaction Exists", vbExclamation, "DAS Version 4.0"
                        Text2.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        End If
    cancel
Else
    cancel
End If
End Sub
Private Sub suplist()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check3.Value = 1 Then
rs.Open "select * from ACT where ACT_nme like '" & Trim(Text3) & "%' and (act_typ like 'PUR' or act_typ like 'BOT')", con, adOpenKeyset
ElseIf Check4.Value = 1 Then
rs.Open "select * from ACT where act_ad4 like '" & Trim(Text3) & "%' and (act_typ like 'PUR' OR act_typ like 'BOT')", con, adOpenKeyset
End If
list3.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list3.ListItems.Add(, , Trim(rs("ACT_nme")))
        ls.ListSubItems.Add , , Trim(rs("ACT_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label31 = "Select Supplier"
    Check4.Visible = True
    Check3.Visible = True
    list3.Visible = True
    list3.SetFocus
    
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text3.Enabled = True
        Text3.SetFocus
    End If
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub prdgrplist()
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from pgr where grp_cLS like '" & Trim(Text2) & "%'", con, adOpenKeyset
If rs.RecordCount > 0 Then
    list2.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list2.ListItems.Add(, , Trim(rs("grp_cat")))
        ls.ListSubItems.Add , , Trim(rs("grp_idy"))
        ls.ListSubItems.Add , , Trim(rs("GRP_CLS") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label31 = "Select Product Category"
    list2.Visible = True
    list2.ListItems(1).Selected = True
    list2.SetFocus
Else
    MsgBox "Segment Not Found, Define Segment !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text2.Enabled = True
    Text2.SetFocus
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub save()
Dim rs As New ADODB.Recordset
Dim REC_BAT As New ADODB.Recordset

If MsgBox("Confirm Record Save ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
              
' Check for TWP Version NO SAVE
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    If Trim(DIR_RS("DIR_ALS")) = "Y" Then
        If Trim(DIR_RS("DIR_MSG")) <> "WRITETWP" Then
            If rs.State = 1 Then rs.Close
            If MOD1 = True Then
                rs.Open "select * from prd where prd_idy like '" & Trim(Text1) & "'", con, adOpenKeyset, adLockPessimistic
                If Trim(List7) = "" Then
                    List7.Selected(0) = True
                    rs("stk_typ") = Trim(List7)
                Else
                    rs("stk_typ") = Trim(List7)
                End If
                rs("TOD_PCG") = Val(Trim(Text14))
                rs.Update
               
                MsgBox "Change Made to SKU - Product Type, Shelf Life Days!", vbInformation, "DAS Version 4.0"
            End If
            MsgBox ("Cannot Save Complete Master Data for KO-TWP Version!"), vbCritical, "DAS Version 4.0"
            cancel
            Exit Sub
        End If
    End If
End If

'SEQUENCE
If CHECKTB("pgr.dbf,prd.dbf,act.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If SSTab1.TabVisible(1) = True And MS_SEQ.Rows > 2 Then
    For A = 1 To MS_SEQ.Rows - 1
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(MS_SEQ.TextMatrix(A, 0)) & "'", con, adOpenDynamic, adLockPessimistic
    rs("BRW_SEQ") = Val(MS_SEQ.TextMatrix(A, 2))
    rs.Update
    Next
    SSTab1.Tab = 0
'    If Trim(Text5) = "" Then
    Exit Sub
End If
'END OF SEQUENCE
If Len(Trim(Text5)) = 0 Then
MsgBox "Product(SKU) Name Required, Enter Name !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Text5.SetFocus
Exit Sub
End If

If Len(Trim(Text4)) = 0 Then
MsgBox "Product Short Name Required, Enter Short Name !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Text4.SetFocus
Exit Sub
End If

If Len(Trim(Combo1)) = 0 Then
MsgBox "Product Pack Size Required, Enter Pack Size !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Combo1.SetFocus
Exit Sub
End If

If Len(Trim(Combo2)) = 0 Then
MsgBox "Product Pack Type Required, Enter Pack Type !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Combo2.SetFocus
Exit Sub
End If

If Len(Trim(Text6)) = 0 Then
MsgBox "Brand Entry Required, Enter Product Brand Name !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Text6.SetFocus
Exit Sub
End If

If check("pgr", "grp_idy", Trim(Text2)) = False Then
MsgBox "Segment Not Found, Define Segment !", vbExclamation, "DAS Version 4.0"
SSTab1.Tab = 0
Text2.SetFocus
Exit Sub
End If

If rs.State = 1 Then rs.Close
rs.Open "select act_idy from ACT where ACT_idy like '" & Sincrement(Trim(Text3)) & "' and (act_typ like 'PUR' OR act_typ like 'BOT')", con, adOpenStatic
If rs.RecordCount > 0 Then
Else
    MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text3.Enabled = True
    Text3.SetFocus
    Exit Sub
End If
If DateDiff("d", M_DAT_DOM, M_DAT_DOE) < 0 Then
    MsgBox "Invalid Date, BBD-Exp < Mfg Date!", vbInformation, "DAS Version 4.0"
    M_DAT_DOE.SetFocus
    Exit Sub
End If
If Val(Text15) <= 0 Then
    MsgBox "Volume Weight Input Required, Enter Volume Weight Value !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text15.SetFocus
    Exit Sub
End If

If Val(Text26) = 0 Then
    MsgBox "Sub Units (Units Per Case) Required, Enter Units Per Case - Sub Units !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text26.SetFocus
    Exit Sub
End If

 
If rs.State = 1 Then rs.Close
If MOD1 = True Then
    rs.Open "select * from prd where prd_idy like '" & Trim(Text1) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from prd", con, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs("prd_idy") = increment("idy_prd")
    MsgBox "New Product ID is " & rs("prd_idy"), vbExclamation, "DAS Version 4.0"
End If
If Trim(List7) = "" Then
    List7.Selected(0) = True
    rs("stk_typ") = Trim(List7)
Else
    rs("stk_typ") = Trim(List7)
End If
PRDID = rs("prd_idy")
rs("grp_idy") = Trim(Text2)
rs("alt_idy") = Trim(Text12)
rs("cmp_idy") = Trim(Text3)
rs("prd_prn") = Trim(Text4)
rs("prd_nme") = Trim(Text5)
rs("brd_nme") = Trim(Text6)
rs("prd_pck") = Trim(Combo1)
rs("prd_pkt") = Trim(Combo2)
rs("bas_rat") = convert(amt(Text7))
rs("sal_rat") = convert(amt(Text8))
rs("mrp_rat") = convert(amt(Text9))
rs("saf_lev") = QTY(Text10 & "")
rs("reo_lev") = QTY(Text11 & "")
rs("TOD_PCG") = Val(Trim(Text14))
rs("pur_rat") = convert(amt(Text13)) 'Added by PNR - TOD_PCG
rs!PUR_TAX = Val(PUR_TAX)
rs!PUR_DIS = Val(PUR_DIS)
rs!SAL_TAX = Val(Text25)
rs!SAL_DIS = Val(Text24)
rs!TAX_TYP = Trim(Combo3)


rs("exc_inf") = Trim(excinf.Text)
If MOD1 = False Then rs!PRD_PCS = 0
If Val(Text15) & "" = 0 Then
    rs("con_fac") = 1
Else
    rs("con_fac") = FAC(Text15) & ""
End If
rs("prd_uom") = List5
If Val(Text26) & "" = 0 Then
    rs("sub_unt") = 1
Else
    rs("sub_unt") = Trim(Val(Text26) & "")
End If
If Label30 <> "" Then
    rs("CUR_LEV") = QTY(Val(Label30) & "")
Else
    rs("CUR_LEV") = "0.000"
End If
If MOD1 = False Then rs("CUR_LEV") = "0.000"
rs("AVG_RAT") = 0
rs.Update
'BATCH UPDATION
If MOD1 = False Then
    If REC_BAT.State = 1 Then REC_BAT.Close
    REC_BAT.Open "select * from BRD", con, adOpenKeyset, adLockPessimistic
    REC_BAT.AddNew
    REC_BAT("bat_idy") = "BAT" & T7increment("idy_bat")
    REC_BAT("prd_idy") = PRDID
    If MOD1 = False Then REC_BAT("BAT_NME") = Trim(Bat_nme)
    REC_BAT("prd_pdr") = convert(amt(Val(Text13)))
    REC_BAT("prd_sdr") = convert(amt(Val(Text8)))
    REC_BAT("prd_mrp") = convert(amt(Val(Text9)))
    REC_BAT("BRD_DOM") = M_DAT_DOM
    REC_BAT("BRD_DOE") = M_DAT_DOE
    REC_BAT("PRD_QTY") = 0
    REC_BAT.Update
End If
'END OF BATCH UPDATION
clrctr M_FRM_PRD
SSTab1.Tab = 0
Text1.Enabled = True
idnum = False
Text1.SetFocus
cancel
setval
End Sub

Private Sub cancel()
clrctr M_FRM_PRD
SSTab1.Tab = 0
Combo1.CLEAR
Combo1 = ""
Combo2.CLEAR
Combo2 = ""
Text6.CLEAR
Text6 = ""
Text1.Enabled = True
idnum = False
Text26.BackColor = Text1.BackColor
Text26.ForeColor = Text1.ForeColor
Bat_nme.BackColor = Text1.BackColor
Bat_nme.ForeColor = Text1.ForeColor
M_DAT_DOM.Enabled = True
M_DAT_DOE.Enabled = True
M_DAT_DOM = Date$
M_DAT_DOE = Date$
Bat_nme.Locked = False
Text1.SetFocus
setval
End Sub

Private Sub newrec()
MOD1 = False
clrctr M_FRM_PRD
cancel
SSTab1.Tab = 0
Text12.SetFocus
Text1.Enabled = False
setval
Toolbar1.Buttons(13).Enabled = False
End Sub

Private Sub exitfrm()
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    DIR_RS("DIR_MSG") = ""
    DIR_RS.Update
End If
Unload Me
End Sub

Private Sub BAT_NME_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Batch Name"
End Sub

Private Sub Bat_nme_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    If M_DAT_DOM.Enabled = True Then
        M_DAT_DOM.SetFocus
    Else
        Text10.SetFocus
    End If
ElseIf K = 38 Then
    Text9.SetFocus
End If
End Sub

Private Sub Bat_nme_Keypress(KeyAscii As Integer)
KeyAscii = CHGCASE(KeyAscii)
End Sub

Private Sub bname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MS.Col = 6 Then
        MS.Text = Trim(bname)
        bname.Visible = False
        MS.SetFocus
    End If
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    PRDLIST
Else
    PRDLIST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    PRDLIST
Else
    PRDLIST
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Check4.Value = 0
    suplist
Else
    Check4.Value = 1
    Check4_Click
End If
End Sub

Private Sub Check3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Name to Get Suppliers List Sorted on Name "
End Sub

Private Sub Check4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Place to Get Suppliers List Sorted on Place "
End Sub

Private Sub Check1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Name to Get Product's List Sorted on Name "
End Sub

Private Sub Check2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Short Name to Get Product's List Sorted on Short Name "
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    Check3.Value = 0
    suplist
Else
    Check3.Value = 1
    Check3_Click
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
    Check1.Value = 0
    Check2.Value = 0
    Check6.Value = 0
    PRDLIST
Else
    PRDLIST
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
    Check1.Value = 0
    Check2.Value = 0
    Check5.Value = 0
    PRDLIST
Else
    PRDLIST
End If
End Sub

Private Sub Combo1_keypress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Combo2_keypress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Combo3_Change()
If Len(Trim(Combo3)) > 15 Then SLT
End Sub

Private Sub Combo3_GotFocus()
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
T = ""
'M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Pack Name from list"
'Dim rs As New ADODB.Recordset
'rs.Open "select alltrim(DIR_ALS) from dir where substr(DIR_IDX,1,2) like 'PD%'", CON, adOpenKeyset, adlockpessimistic 'Modified by Ramanesh
'If rs.RecordCount > 0 Then
'    If MOD1 = True Then t = Trim(Combo2)
'    Combo2.CLEAR
'    For A = 1 To rs.RecordCount
'        If Len(Trim(rs(0))) <> 0 Then Combo2.AddItem Trim(rs(0))
'        rs.MoveNext
'    Next
'    If MOD1 = True Then Combo2 = Trim(t)
'End If

If rs.State = 1 Then rs.Close
rs.Open "select DISTINCT(TRIM(TAX_TYP)) from prd group by TAX_TYP", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Combo3)
    Combo3.CLEAR
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Combo3.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Combo3 = Trim(T)
End If
End Sub

Private Sub Combo3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
   Text15.SetFocus
'ElseIf K = 38 Then
'    Text25.SetFocus
End If
End Sub

Private Sub Combo3_keypress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub
Private Sub Command1_Click()
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

Private Sub Command2_Click()
frmNewRateTemplate.Show 1
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MS.Col = 30 Then
        MS.Text = Format(date1.Value, "dd/MMM/yyyy")
        date1.Visible = False
        MS.SetFocus
    End If
    If MS.Col = 31 Then
        If Format(MS.TextMatrix(MS.Row, 30), "dd/MM/yyyy") > Format(date1.Value, "dd/MM/yyyy") Then
            MsgBox "BBD - Expiry date Should not be Lesser than Mfg. Date", vbExclamation, "DAS Version 4.0"
            date1.Visible = True
            date1.SetFocus
        Else
            MS.Text = Format(date1.Value, "dd/MMM/yyyy")
            date1.Visible = False
            MS.SetFocus
        End If
    End If
End If
End Sub

Private Sub excinf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text26)) = 0 Then Text26 = "0"
     save
ElseIf KeyCode = 38 And Shift = 1 Then
    If Text26.Enabled = True Then
        Text26.SetFocus
        SendKeys "{home}+{end}"
    Else
        List5.SetFocus
    End If
End If
End Sub

Private Sub MS_Scroll()
If Text27.Visible = True Then
    Text27 = ""
    Text27.Visible = False
    MS.SetFocus
ElseIf Text28.Visible = True Then
    Text28 = ""
    Text28.Visible = False
    MS.SetFocus
ElseIf List6.Visible = True Then
    List6.Visible = False
    MS.SetFocus
ElseIf bname.Visible = True Then
    bname = ""
    bname.Visible = False
    MS.SetFocus
ElseIf date1.Visible = True Then
    date1.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub MS_SEQ_Scroll()
If Text16.Visible = True Then
    Text16 = ""
    Text16.Visible = False
    MS_SEQ.SetFocus
End If
End Sub

Private Sub Text14_Change()
If Len(Text14) = 4 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text14_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "SKU Shelf Life in Days to Calculate BBD-Exp Date"
SendKeys "{home}+{end}"
End Sub

Private Sub Text14_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text6.SetFocus
ElseIf K = 38 Then
    If Text4.Enabled = True Then
        Text4.SetFocus
    Else
        Text6.SetFocus
    End If
'     If MOD1 = False Then
'       Text4.SetFocus
'     Else
'       Text6.SetFocus
'     End If
End If
End Sub

Private Sub Text14_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
'    Text4.SetFocus
Else
    K = 0
End If
End Sub

Private Sub Text6_Change()
If Len(Trim(Text6)) > 15 Then SLT
End Sub
Private Sub Combo1_Change()
If Len(Trim(Combo1)) > 15 Then SLT
End Sub
Private Sub Combo2_Change()
If Len(Trim(Combo2)) > 15 Then SLT
End Sub

Private Sub Text6_GotFocus()
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Brand Name from list"
Dim rs As New ADODB.Recordset
T = ""
rs.Open "select alltrim(DIR_MSG) from dir where substr(DIR_IDX,1,2) like 'BR%'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Text6)
    Text6.CLEAR
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Text6.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Text6 = Trim(T)
End If
If rs.State = 1 Then rs.Close
rs.Open "select brd_nme from prd group by brd_nme", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Text6)
'    Text6.Clear
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Text6.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Text6 = Trim(T)
End If
End Sub
Private Sub COMBO1_GotFocus()
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
T = ""
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Pack Name from list"
Dim rs As New ADODB.Recordset
rs.Open "select alltrim(DIR_MSG) from dir where substr(DIR_IDX,1,2) like 'PK%'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Combo1)
    Combo1.CLEAR
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Combo1.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Combo1 = Trim(T)
End If
If rs.State = 1 Then rs.Close
rs.Open "select PRD_PCK from prd group by PRD_PCK", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Combo1)
'    Combo1.Clear
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Combo1.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Combo1 = Trim(T)
End If
End Sub

Private Sub COMBO2_GotFocus()
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
T = ""
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Pack Name from list"
Dim rs As New ADODB.Recordset
rs.Open "select alltrim(DIR_ALS) from dir where substr(DIR_IDX,1,2) like 'PD%'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Combo2)
    Combo2.CLEAR
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Combo2.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Combo2 = Trim(T)
End If

If rs.State = 1 Then rs.Close
rs.Open "select PRD_PKT from prd group by PRD_PKT", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If MOD1 = True Then T = Trim(Combo2)
'    Combo2.Clear
    For A = 1 To rs.RecordCount
        If Len(Trim(rs(0))) <> 0 Then Combo2.AddItem Trim(rs(0))
        rs.MoveNext
    Next
    If MOD1 = True Then Combo2 = Trim(T)
End If
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabVisible(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = True
If Text1.Enabled = True Then
    Text1.SetFocus
Else
'    Text12.SetFocus
End If
List5.Selected(0) = True
List7.Selected(0) = True
excinf.ListIndex = 1
M_DAT_DOM = Date$
M_DAT_DOE = Date$
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_PRD.ActiveControl Is TextBox Then M_FRM_PRD.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False And list2.Visible = False And list3.Visible = False And SSTab1.Tab <> 3 Then
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        exitfrm
    ElseIf K = 66 And s = 2 Then
        If Toolbar1.Buttons(13).Enabled = True Then BROWSE
        SSTab1.TabEnabled(3) = True
    ElseIf K = 67 And s = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And s = 2 Then
        DEL_REC
    ElseIf K = 69 And s = 2 Then
        If M_FRM_PRD.ActiveControl.name = "Text1" Then PRDLIST
        If M_FRM_PRD.ActiveControl.name = "Text2" Then prdgrplist
        If M_FRM_PRD.ActiveControl.name = "Text3" Then suplist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        M_FRM_PPQ.Tag = "PRD"
        M_FRM_PPQ.Show 1
    ElseIf K = 82 And s = 2 Then
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        Check1.Visible = False
        Check2.Visible = False
        Check5.Visible = False
        Check6.Visible = False
        SSTab1.Tab = 0
        If CHECKTB("prd.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If

        If check("prd", "prd_idy", Trim(Text1)) = False Then
            Text1.SetFocus
            Exit Sub
        Else
            Text1.Enabled = False
            Text2.SetFocus
        End If
    ElseIf list2.Visible = True Then
        list2.Visible = False
        SSTab1.Tab = 0
        Text2.SetFocus
    ElseIf list3.Visible = True Then
        list3.Visible = False
        Check4.Visible = False
        Check3.Visible = False
        SSTab1.Tab = 0
        Text3.SetFocus
    ElseIf MS.Visible = True Then
        MS.HighLight = flexHighlightNever
    Dim rs As New ADODB.Recordset
    Text27.Visible = False
    Text28.Visible = False
    MS.SetFocus
'    If MS.Col <> 0 Or MS.Col <> 1 Or MS.Col <> 2 Or MS.Col <> 3 Or MS.Col <> 4 Or MS.Col <> 5 Or MS.Col <> 6 Or MS.Col <> 14 Or MS.Col <> 28 Then
        M_FRM_PRD.MousePointer = vbHourglass
        For A = 1 To MS.Rows - 1
            
            If rs.State = 1 Then rs.Close
            rs.Open "select * from brd where bat_idy like '" & Trim(MS.TextMatrix(A, 29)) & "'", con, adOpenKeyset, adLockPessimistic
            If rs.RecordCount > 0 Then
                rs("prd_pdr") = convert(amt(MS.TextMatrix(A, 7)))
                rs("prd_sdr") = convert(amt(MS.TextMatrix(A, 8)))
                rs("prd_mrp") = convert(amt(MS.TextMatrix(A, 9)))
                rs("brd_dom") = Format(MS.TextMatrix(A, 30), "mm/DD/YYYY")
                rs("brd_doe") = Format(MS.TextMatrix(A, 31), "mm/DD/YYYY")
                rs.Update
            End If
        Next
        M_FRM_PRD.MousePointer = vbDefault
        'End If
        MS.Visible = False
        Text27.Visible = False
        Text28.Visible = False
        SSTab1.Tab = 0
        If Text1.Enabled = True Then
            Text1.SetFocus
        Else
            Text12.SetFocus
        End If
        KeyAscii = 0
    ElseIf SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        If Text1.Enabled = True Then
            Text1.SetFocus
        Else
            Text12.SetFocus
        End If
    Else
        If M_FRM_PRD.ActiveControl.name = "Text1" And SSTab1.Tab = 0 Then
            Unload Me
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
MOD1 = False
idnum = False
Toolbar1.Buttons(12).Enabled = False
setval
MS.MergeCells = flexMergeFree
MS.MergeCol(0) = True
MS.MergeCol(1) = True
MS.MergeCol(2) = True
MS.MergeCol(3) = True
MS.MergeCol(4) = True
MS.ColWidth(0) = 1150 'prd id
MS.ColWidth(4) = 1100 'sup cod
MS.ColWidth(1) = 2500 'prd name
MS.ColWidth(2) = 1800 'prt num
MS.ColWidth(3) = 1500 'brd num
MS.ColWidth(5) = 0  'seg id
MS.ColWidth(6) = 1100 'sup id
MS.ColWidth(7) = 950 'bill rate A
MS.ColWidth(8) = 950 'Bill Rate B
MS.ColWidth(9) = 950 'Bill Rate C
MS.ColWidth(10) = 920 'con fac
MS.ColWidth(11) = 1100 'saf qty
MS.ColWidth(12) = 1100 'reo qty
MS.ColWidth(13) = 1100 'cur qty
MS.ColWidth(14) = 550 'uom
MS.ColWidth(15) = 500 'sub unit
MS.ColWidth(16) = 950 'pur rate
MS.ColWidth(17) = 0 'turn dis
MS.ColWidth(18) = 650 'pur dis
MS.ColWidth(19) = 650 'pur tax
MS.ColWidth(20) = 650 'sal dis
MS.ColWidth(21) = 650 'sal tax
MS.ColWidth(22) = 0 'list price
MS.ColWidth(23) = 0 'crd dis
MS.ColWidth(24) = 0 'deb dis
MS.ColWidth(25) = 0 'ad
MS.ColWidth(26) = 0 'cd
MS.ColWidth(27) = 0 'exd
MS.ColWidth(28) = 0 'exduty applicable
MS.ColWidth(29) = 0
MS.TextMatrix(0, 0) = "Product ID"
MS.ColAlignment(0) = 1
MS.TextMatrix(0, 4) = "Splr.ID  "
MS.TextMatrix(0, 1) = "Product Name"
MS.TextMatrix(0, 2) = "Short Name"
MS.TextMatrix(0, 3) = "Brand Name"
MS.TextMatrix(0, 5) = "Segment ID"
MS.ColAlignment(5) = 1
MS.TextMatrix(0, 6) = "Batch Name"
MS.ColAlignment(6) = 1
SCl MS, 6
MS.TextMatrix(0, 7) = "Pur  Price"
SCl MS, 7
MS.TextMatrix(0, 8) = " Sale Rate"
SCl MS, 8
MS.TextMatrix(0, 9) = "   MRP Rs."
SCl MS, 9
MS.TextMatrix(0, 10) = " Cnv.Factor"
'SCl MS, 10
MS.TextMatrix(0, 11) = "   Safety Qty"
MS.ColWidth(11) = 0
'SCl MS, 11
MS.TextMatrix(0, 12) = " Reorder Qty"
MS.ColWidth(12) = 0
MS.TextMatrix(0, 13) = " Present Qty"
'SCl MS, 13
MS.TextMatrix(0, 14) = "UOM"
MS.TextMatrix(0, 15) = " UPC"
MS.TextMatrix(0, 16) = "Special Rate"
MS.ColWidth(16) = 0
'SCl MS, 16
MS.TextMatrix(0, 17) = "   TD %"
MS.ColWidth(17) = 0
'SCl MS, 17
MS.TextMatrix(0, 18) = "   PD %"
MS.ColWidth(18) = 0
'SCl MS, 18
MS.TextMatrix(0, 19) = "   PT %"
MS.ColWidth(19) = 0
'SCl MS, 19
MS.TextMatrix(0, 20) = "   SD %"
MS.ColWidth(20) = 0
'SCl MS, 20
MS.TextMatrix(0, 21) = "   ST %"
MS.ColWidth(21) = 0
'SCl MS, 21
MS.TextMatrix(0, 22) = "  List Price"
MS.ColWidth(22) = 0
'SCl MS, 22
MS.TextMatrix(0, 23) = "CrtD %"
MS.ColWidth(23) = 0
'SCl MS, 23
MS.TextMatrix(0, 24) = "DbtD %"
MS.ColWidth(24) = 0
'SCl MS, 24
MS.TextMatrix(0, 25) = "   AD %"
MS.ColWidth(25) = 0
'SCl MS, 25
MS.TextMatrix(0, 26) = "   CD %"
MS.ColWidth(26) = 0
'SCl MS, 26
MS.TextMatrix(0, 27) = "   ED %"
MS.ColWidth(27) = 0
'SCl MS, 27
MS.TextMatrix(0, 28) = "ED App"
MS.ColWidth(28) = 0
MS.TextMatrix(0, 30) = "Mfg.Date"
MS.ColWidth(30) = 1500
SCl MS, 30
MS.ColAlignment(30) = 1
MS.TextMatrix(0, 31) = "BBD-Exp.Date"
MS.ColWidth(31) = 1500
SCl MS, 31
MS.ColAlignment(31) = 1

MS.TextMatrix(0, 32) = "Pur Inc'tve"
MS.ColWidth(32) = 1200
MS.ColAlignment(32) = 7
MS.TextMatrix(0, 33) = "Sal Inc'tve"
MS.ColWidth(33) = 1200
MS.ColAlignment(33) = 7
'SCl MS, 28

MS.TextMatrix(0, 34) = "List Rate"
MS.ColWidth(34) = 1200
MS.ColAlignment(34) = 7
MS.TextMatrix(0, 35) = "List %+"
MS.ColWidth(35) = 1200
MS.ColAlignment(35) = 7

MS_SEQ.ColWidth(0) = 1150
MS_SEQ.ColWidth(4) = 1100
MS_SEQ.ColWidth(1) = 2500
MS_SEQ.ColWidth(2) = 900
MS_SEQ.ColWidth(3) = 1500
MS_SEQ.ColWidth(5) = 1100
MS_SEQ.ColWidth(6) = 1000
MS_SEQ.TextMatrix(0, 0) = "Product ID"
MS_SEQ.ColAlignment(0) = 1
MS_SEQ.TextMatrix(0, 3) = "Brand    "
MS_SEQ.TextMatrix(0, 1) = "Product Name"
MS_SEQ.TextMatrix(0, 2) = "Sequence"
'SCl MS_SEQ, 2
MS_SEQ.TextMatrix(0, 4) = "Segment ID"
MS_SEQ.ColAlignment(4) = 1
MS_SEQ.TextMatrix(0, 5) = "Supplier ID"
MS_SEQ.ColAlignment(5) = 1
MS_SEQ.TextMatrix(0, 6) = "Short Name"
Text14.Enabled = False 'changed by PNR as per Prince Chopra request
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Product or Press <Esc> to Quit"
End Sub

Private Sub list2_DblClick()
List2_Keydown 13, 0
End Sub

Private Sub List2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Segment or Press <Esc> to Quit"
End Sub

Private Sub list3_DblClick()
List3_KeyDown 13, 0
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub list3_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list3.SortOrder = lvwAscending Then
list3.SortOrder = lvwDescending
ElseIf list3.SortOrder = lvwDescending Then
list3.SortOrder = lvwAscending
End If
list3.SortKey = CH.Index - 1
End Sub

Private Sub list2_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list2.SortOrder = lvwAscending Then
list2.SortOrder = lvwDescending
ElseIf list2.SortOrder = lvwDescending Then
list2.SortOrder = lvwAscending
End If
list2.SortKey = CH.Index - 1
End Sub

Private Sub list1_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
    If list1.ListItems.count > 0 Then
        Text1 = Trim(list1.SelectedItem.ListSubItems.Item(10))
        DISPDETAILS
        SSTab1.Tab = 0
        Check1.Visible = False
        Check2.Visible = False
        Check5.Visible = False
        Check6.Visible = False
        list1.Visible = False
        Text12.SetFocus
    End If
End If
End Sub

Private Sub list2_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub List3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Supplier or Press <Esc> to Quit"
End Sub

Private Sub List2_Keydown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
If list2.ListItems.count > 0 Then
    Text2 = list2.SelectedItem.ListSubItems.Item(1)
    Dim ls As ListItem
    Set ls = list2.SelectedItem
    Label11 = list2.SelectedItem.SubItems(2)
    SSTab1.Tab = 0
    list2.Visible = False
    Text3.SetFocus
End If
End If
End Sub

Private Sub List3_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
If list3.ListItems.count > 0 Then
    Text3 = list3.SelectedItem.ListSubItems.Item(1)
    Label12 = list3.SelectedItem
    SSTab1.Tab = 0
    Check3.Visible = False
    Check4.Visible = False
    list3.Visible = False
    Text13.SetFocus
End If
End If
End Sub
Private Sub List5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Unit Of Measurement"
End Sub

Private Sub List5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MOD1 = False Then
        Text26.Enabled = True
        Text26.SetFocus
    Else
        excinf.SetFocus
    End If
ElseIf KeyCode = 38 And Shift = 1 Then
    Text15.SetFocus
End If
End Sub

Private Sub List6_KeyPress(K As Integer)
If K = 13 Then
    MS.Text = Trim(List6)
    List6.Visible = False
    MS.Col = 2
    MS.SetFocus
End If
End Sub

Private Sub List7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Product Type"
End Sub

Private Sub List7_Keydown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
ElseIf KeyAscii = 38 And s = 1 Then
    Text5.SetFocus
End If
End Sub

Private Sub M_DAT_DOE_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select BBD - Expiry Date"
End Sub

Private Sub M_DAT_DOM_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Manufacturing Date"
End Sub

Private Sub M_DAT_DOM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then M_DAT_DOE.SetFocus
If KeyCode = 38 And Shift = 1 Then Bat_nme.SetFocus
End Sub

Private Sub M_DAT_DOE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text10.SetFocus
    If DateDiff("d", M_DAT_DOM, M_DAT_DOE) < 0 Then
        MsgBox "Invalid Date, BBD-Exp < Mfg Date!", vbExclamation, "DAS Version 4.0"
        M_DAT_DOE.SetFocus
    End If
ElseIf KeyCode = 38 And Shift = 1 Then
    M_DAT_DOM.SetFocus
End If

End Sub

Private Sub MS_EnterCell()
If MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Brand Name"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Supplier's Code"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Segment ID"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Supplier ID"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Rate A"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Rate B"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Rate C"
ElseIf MS.Col = 11 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Safety Quantity"
ElseIf MS.Col = 12 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Reorder Quantity"
ElseIf MS.Col = 13 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Present Quantity"
ElseIf MS.Col = 14 Then
    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measurement"
ElseIf MS.Col = 15 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Sub Units"
ElseIf MS.Col = 16 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Purchase Rate"
ElseIf MS.Col = 17 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Turnover Discount Percentage"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Conversion Factor"
ElseIf MS.Col = 18 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Purchase Discount Percentage"
ElseIf MS.Col = 19 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Purchase Tax Percentage"
ElseIf MS.Col = 20 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Sale Discount Percentage"
ElseIf MS.Col = 21 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Sale Tax Percentage"
ElseIf MS.Col = 22 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter List Price"
ElseIf MS.Col = 23 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Creditor Discount Percentage"
ElseIf MS.Col = 24 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Debtor Discount Percentage"
ElseIf MS.Col = 25 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Additional Discount Percentage"
ElseIf MS.Col = 26 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Cash Discount Percentage"
ElseIf MS.Col = 27 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Excise Duty Discount Percentage"
ElseIf MS.Col = 28 Then
    M_MDI.STS_BAR.Panels(1).Text = "Excise Duty Applicable"
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_KeyPress(K As Integer)
If K <> 27 Then
    If K > 47 And K < 58 Or K = 46 Or K = 8 Then
        If MS.Col = 7 Or MS.Col = 8 Or MS.Col = 9 Then
            Text27 = ""
            Text27.MaxLength = 10
            Text27.Visible = True
            Text27.Top = MS.Top + MS.CellTop
            Text27.Left = MS.CellLeft + MS.Left
            Text27.Width = MS.CellWidth
            Text27.Height = MS.CellHeight
            Text27 = ""
            Text27.Visible = True
            Text27.Enabled = True
            If K <> 13 Then Text27 = Chr(K)
            Text27.Visible = True
            Text27.SelStart = Len(Text27)
            Text27.Visible = True
            Text27.SetFocus
        ElseIf MS.Col = 6 Then
            bname = ""
            bname.Top = MS.Top + MS.CellTop
            bname.Left = MS.Left + MS.CellLeft
            bname.Height = MS.CellHeight
            bname.Width = MS.CellWidth
            bname.Visible = True
            bname.Enabled = True
            If K <> 13 Then bname = Chr(K)
            bname.Visible = True
            bname.SelStart = Len(Trim(bname))
            bname.Visible = True
            bname.SetFocus
        ElseIf MS.Col = 30 Or MS.Col = 31 Then
            date1.Top = MS.Top + MS.CellTop
            date1.Left = MS.Left + MS.CellLeft
            date1.Height = MS.CellHeight
            date1.Width = MS.CellWidth
            date1.Visible = True
            date1.SetFocus
        End If
    ElseIf K = 13 Then
        If MS.Col <> 31 Then
            MS.Col = MS.Col + 1
        End If
        'K = 0
    End If
End If
End Sub

Private Sub MS_LeaveCell()
List6.Visible = False
Text27.Visible = False
Text28.Visible = False
End Sub

Private Sub MS_SEQ_EnterCell()
If MS_SEQ.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product ID"
ElseIf MS_SEQ.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Name"
ElseIf MS_SEQ.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Sequence"
ElseIf MS_SEQ.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Brand"
ElseIf MS_SEQ.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Segment ID"
ElseIf MS_SEQ.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Supplier ID"
ElseIf MS_SEQ.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
End If
End Sub

Private Sub MS_SEQ_KeyPress(K As Integer)
If K <> 27 Then
    If K > 47 And K < 58 Or K = 46 Or K = 8 Then
        If MS_SEQ.Col = 2 Then
            Text16.MaxLength = 5
            Text16.Visible = True
            Text16.Top = MS_SEQ.Top + MS_SEQ.CellTop
            Text16.Left = MS_SEQ.CellLeft + MS_SEQ.Left
            Text16.Width = MS_SEQ.CellWidth
            Text16.Height = MS_SEQ.CellHeight
            Text16 = ""
            Text16.Visible = True
            Text16.SetFocus
            If K <> 13 Then Text16 = Chr(K)
            Text16.SelStart = Len(Text16)
        End If
    End If
End If
End Sub

Private Sub PUR_DIS_Change()
If Len(PUR_DIS) = 6 Then SendKeys "{home}+{end}"
End Sub

Private Sub PUR_DIS_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Purchase Cash Discount Percentage"
SendKeys "{HOME}+{END}"
End Sub

Private Sub PUR_TAX_Change()
If Len(PUR_TAX) = 6 Then SendKeys "{home}+{end}"
End Sub

Private Sub PUR_tax_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Purchase Tax Percentage"
SendKeys "{HOME}+{END}"
End Sub

Private Sub PUR_DIS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    PUR_DIS = discount(Val(PUR_DIS))
    PUR_TAX.SetFocus
ElseIf KeyCode = 38 Then
    PUR_DIS = discount(Val(PUR_DIS))
    Text11.SetFocus
End If
End Sub

Private Sub PUR_TAX_KeyPress(KeyAscii As Integer)
KeyAscii = CHKNUM(KeyAscii)
End Sub

Private Sub PUR_TAX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    PUR_TAX = discount(Val(PUR_TAX))
    Text24.SetFocus
ElseIf KeyCode = 38 Then
    PUR_TAX = discount(Val(PUR_TAX))
    PUR_DIS.SetFocus
End If
End Sub

Private Sub PUR_DIS_KeyPress(KeyAscii As Integer)
KeyAscii = CHKNUM(KeyAscii)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
list1.Visible = False
list2.Visible = False
list3.Visible = False
Check1.Visible = False
Check2.Visible = False
Check5.Visible = False
Check6.Visible = False

Check3.Visible = False
Check4.Visible = False

If SSTab1.Tab = 0 Then
    If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        Text12.SetFocus
    End If
ElseIf SSTab1.Tab = 1 Then
    SEQUENCE
    If MS_SEQ.Visible = True Then
        MS_SEQ.SetFocus
        MS.Col = 2
        MS.Row = 1
    End If
ElseIf SSTab1.Tab = 3 Then
BROWSE
End If
'If ms.Visible = True Then ms.Visible = False
If SSTab1.Tab = 3 Or SSTab1.Tab = 2 Then
'Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Product ID or Press <Enter> to get List of Product ID's"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Text12_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Supplier's Product Code"
SendKeys "{home}+{end}"
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then Text5.SetFocus
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Then KeyAscii = 0
End Sub

Private Sub Text10_Change()
If Len(Text10) = 10 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text11_Change()
If Len(Text11) = 10 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text12_Change()
If Len(Text12) = 10 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text13_Change()
If Len(Text13) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text15_Click()
SendKeys "{home}+{end}"
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text15 = FAC(Text15)
    List5.SetFocus
ElseIf KeyCode = 38 Then
    Text15 = FAC(Text15)
    Combo3.SetFocus
End If
End Sub

Private Sub Text15_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text15_Change()
If Len(Text15) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text16_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Or K = 45 Then
ElseIf K = 13 Then
    If MS_SEQ.Col = 2 Then
        MS_SEQ.Text = Trim(Text16)
    End If
    Text16.Visible = False
    MS_SEQ.SetFocus
    If MS_SEQ.Row = MS_SEQ.Rows - 1 Then
    Else
        MS_SEQ.Row = MS_SEQ.Row + 1
    End If
Else
    K = 0
End If
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Segment ID or Press <Enter> to get List of Segment ID's"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub Text24_Change()
If Len(Text24) = 6 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text24_Click()
SendKeys "{home}+{end}"
End Sub

Private Sub Text25_Change()
If Len(Text25) = 6 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text25_Click()
SendKeys "{home}+{end}"
End Sub

Private Sub Text26_Change()
If Len(Trim(Text26)) = 3 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text26_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Sub Units"
End Sub

Private Sub Text26_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
'    If Len(Trim(Text26)) = 0 Then Text26 = "0"
'        save
     excinf.SetFocus
ElseIf KeyCode = 38 Then
    If Len(Trim(Text26)) = 0 Then Text26 = "0"
    List5.SetFocus
End If
End Sub

Private Sub Text26_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub


Private Sub Text27_KeyPress(K As Integer)
If K <> 13 Then K = CHKNUM(K)
If K = 39 Or K = 41 Then K = 0
If K > 47 And K < 58 Or K = 46 Or K = 8 Or K = 45 Then
ElseIf K = 13 Then
    If MS.Col = 7 Or MS.Col = 8 Or MS.Col = 9 Then
        MS.Text = convert(amt(Val(Text27)))
        Text27.Visible = False
        MS.SetFocus
        If MS.Col <> 27 Then
            MS.Col = MS.Col + 1
        Else
            MS.Col = 2
        End If
    Else
        K = 0
    End If
End If
End Sub

Private Sub Text28_Change()
 If Len(Trim(Text28)) = 3 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text28_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
ElseIf K = 13 Then
    MS.Text = Val(Text28) & ""
    Text28.Visible = False
    MS.Col = MS.Col + 1
    MS.SetFocus
Else
    K = 0
End If
End Sub

Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Supplier ID or Press <Enter> to get List of Supplier ID's"
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub Text4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Short Name"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub Text5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Description of the Product"
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub Text7_Change()
If Len(Text7) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Rate A Per Case"
SendKeys "{home}+{end}"
End Sub

Private Sub Text8_Change()
If Len(Text8) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text8_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Billing Rate B Per Case"
SendKeys "{home}+{end}"
End Sub

Private Sub Text9_Change()
If Len(Text9) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Text9_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter MRP Rate C Per Case"
SendKeys "{home}+{end}"
End Sub

Private Sub Text10_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Safe Stock Keeping Level"
SendKeys "{home}+{end}"
End Sub

Private Sub Text11_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Reorder Level"
SendKeys "{home}+{end}"
End Sub

Private Sub Text13_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Purchase Rate Per Case"
SendKeys "{home}+{end}"
End Sub

Private Sub Text15_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Conversion "
SendKeys "{home}+{end}"
End Sub

Private Sub Text24_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Sale Cash Discount Percentage"
SendKeys "{home}+{end}"
End Sub

Private Sub Text25_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Sale Tax Percentage"
SendKeys "{home}+{end}"
End Sub

Private Sub TEXT1_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then DISPDETAILS
If K = 27 And M_FRM_PRD.ActiveControl.name = "Text1" Then Unload Me
End Sub

Private Sub TEXT2_KeyDown(K As Integer, s As Integer)
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If rs.State = 1 Then rs.Close
    If Len(Trim(Text2)) <> 0 Then
        rs.Open "select * from pgr where grp_idy like '" & Sincrement(Trim(Text2)) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text2 = Sincrement(Text2)
            Label11 = rs("grp_cLS")
            If K = 38 Then
                Text6.SetFocus
            Else
                Text3.SetFocus
            End If
        Else
            prdgrplist
        End If
    Else
        prdgrplist
   End If
End If
End Sub

Private Sub Text3_KeyDown(K As Integer, s As Integer)
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If rs.State = 1 Then rs.Close
    If Len(Trim(Text3)) <> 0 Then
        rs.Open "select * from ACT where ACT_idy like '" & Sincrement(Trim(Text3)) & "' AND (act_typ like 'PUR' OR act_typ like 'BOT')", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text3 = Sincrement(Text3)
            Label12 = rs("ACT_nme")
            If K = 38 Then
                Text2.SetFocus
            Else
                Text13.SetFocus
            End If
        Else
            suplist
        End If
    Else
        suplist
    End If
End If
End Sub

Private Sub Text4_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    If Len(Trim(Text4)) = 0 Then
        MsgBox "Short Name Should not be Blank, Give Short Name !", vbExclamation, "DAS Version 4.0"
        If Text14.Enabled = True Then
            Text14.SetFocus
        Else
            Text6.SetFocus
        End If
    Else
        If Text14.Enabled = True Then
            Text14.SetFocus
        Else
            Text6.SetFocus
        End If
    End If
ElseIf K = 38 Then
    If Len(Trim(Text4)) = 0 Then
        MsgBox "Short Name Should not be Blank, Give Short Name !", vbExclamation, "DAS Version 4.0"
        Text4.SetFocus
    Else
        List7.SetFocus
    End If
End If
End Sub

Private Sub TEXT5_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    If Len(Trim(Text5)) = 0 Then
        MsgBox "Name Should not be Blank, Give Name", vbExclamation, "DAS Version 4.0"
        Text5.SetFocus
    Else
        List7.SetFocus
    End If
ElseIf K = 38 Then
    If Len(Trim(Text5)) = 0 Then
        MsgBox "Name Should not be Blank, Give Name", vbExclamation, "DAS Version 4.0"
        Text5.SetFocus
    Else
        Text12.SetFocus
    End If
End If
End Sub

Private Sub TEXT6_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    Combo1.SetFocus
ElseIf K = 36 And s = 1 Then
    Text4.SetFocus
ElseIf K = 38 And s = 1 Then
    List7.SetFocus
End If
End Sub
Private Sub Combo1_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    Combo2.SetFocus
ElseIf K = 36 And s = 1 Then
    Text6.SetFocus
ElseIf K = 38 And s = 1 Then
    Text4.SetFocus
End If
End Sub
Private Sub Combo2_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    Text2.SetFocus
ElseIf K = 36 And s = 1 Then
    Combo1.SetFocus
ElseIf K = 38 And s = 1 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Text7_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text7 = convert(amt(Val(Text7)))
    Text8.SetFocus
ElseIf K = 38 Then
    Text7 = convert(amt(Val(Text7)))
    Text13.SetFocus
End If
End Sub

Private Sub Text7_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text8_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text9_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text10_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text11_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text13_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Text24_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Or K = 45 Then
Else
K = 0
End If
End Sub

Private Sub Text25_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Or K = 45 Then
Else
K = 0
End If
End Sub

Private Sub Text8_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text8 = convert(amt(Val(Text8)))
    Text9.SetFocus
ElseIf K = 38 Then
    Text8 = convert(amt(Val(Text8)))
    Text7.SetFocus
End If
End Sub

Private Sub Text9_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text9 = convert(amt(Val(Text9)))
    If MOD1 = False Then
       Bat_nme.SetFocus
    Else
       Text10.SetFocus
    End If
ElseIf K = 38 Then
    Text9 = convert(amt(Val(Text9)))
    Text8.SetFocus
End If
End Sub

Private Sub Text10_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text10 = QTY(Trim(Text10))
    Text11.SetFocus
ElseIf K = 38 Then
    Text10 = QTY(Trim(Text10))
    If MOD1 = False Then
       M_DAT_DOE.SetFocus
    Else
       Text9.SetFocus
    End If
End If
End Sub

Private Sub Text11_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text11 = QTY(Trim(Text11))
    PUR_DIS.SetFocus
ElseIf K = 38 Then
    Text11 = QTY(Trim(Text11))
    Text10.SetFocus
End If
End Sub

Private Sub Text13_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text13 = convert(amt(Val(Text13)))
    Text7.SetFocus
ElseIf K = 38 Then
    Text13 = convert(amt(Val(Text13)))
    Text3.SetFocus
End If
End Sub

Private Sub Text24_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text24 = discount(Trim(Text24))
    Text25.SetFocus
ElseIf K = 38 Then
    Text24 = discount(Trim(Text24))
    PUR_TAX.SetFocus
End If
End Sub

Private Sub Text25_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    Text25 = discount(Trim(Text25))
   ' Text15.SetFocus
   Combo3.SetFocus
ElseIf K = 38 Then
    Text25 = discount(Trim(Text25))
    Text24.SetFocus
End If
End Sub

Public Sub DISPDETAILS()
On Error Resume Next
If CHECKTB("prd.dbf,brd.dbf,pgr.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim REC_BAT As New ADODB.Recordset
If REC_BAT.State = 1 Then REC_BAT.Close
REC_BAT.Open "select * from brd where prd_idy like '" & Sincrement(Trim(Text1)) & "'", con, adOpenKeyset, adLockPessimistic
If rs.State = 1 Then rs.Close
rs.Open "select * from prd where prd_idy like '" & Sincrement(Trim(Text1)) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    Text1 = Sincrement(Text1)
    Text12 = Trim(rs("alt_idy") & "")
    Text2 = Trim(rs("grp_idy"))
    Text3 = Trim(rs("cmp_idy") & "")
    Text4 = Trim(rs("prd_prn"))
    Text5 = Trim(rs("prd_nme"))
    List7 = Trim(rs("stk_typ") & "")
    Text6 = Trim(rs("brd_nme") & "")
    Combo1 = Trim(rs("prd_pck") & "")
    Combo2 = Trim(rs("prd_pkt") & "")
    Text7 = convert(amt(rs("bas_rat") & ""))
    Text8 = convert(amt(rs("sal_rat") & ""))
    Text9 = convert(amt(rs("mrp_rat") & ""))
    Text10 = QTY(rs("saf_lev") & "")
    Text11 = QTY(rs("reo_lev") & "")
    'Text14 = convert(AMT(rs("TOD_PCG")) & "") ' Added by PNR - Safety Norm Days into field TOD_PCG
    Text14 = Format(Trim(str(rs("TOD_PCG"))), "000")
    Label30 = QTY(rs("cur_lev") & "")
    Text13 = convert(amt(Val(rs("pur_rat") & "")))
    Text15 = Format(FAC(rs("con_fac") & ""), "0.000")
    Text26 = Trim(Val(rs("sub_unt") & ""))
    Text26.BackColor = Label30.BackColor
    Text26.ForeColor = vbWhite
    Combo3 = Trim(rs("TAX_TYP") & "") ' ADDED BY PKB
    M_DAT_DOM = REC_BAT("brd_dom")
    M_DAT_DOE = REC_BAT("brd_doe")
    M_DAT_DOM.Enabled = False
    M_DAT_DOE.Enabled = False
    Bat_nme.Locked = True
    Bat_nme.BackColor = Label30.BackColor
    Bat_nme.ForeColor = vbWhite
    
    Text26.Enabled = False
    List5 = Trim(rs("prd_uom"))
    PUR_DIS = discount(rs("PUR_dis") & "")
    PUR_TAX = discount(rs("PUR_TAX") & "")
    
    Text24 = discount(rs("SAL_dis") & "")
    Text25 = discount(rs("SAL_TAX") & "")
    
    If Trim(rs("exc_inf")) = "Y" Then
        excinf.Selected(0) = True
    Else
        excinf.Selected(1) = True
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "select * from pgr where grp_idy like '" & Trim(Text2) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then Label11.Caption = rs("grp_cls")
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ACT where ACT_idy like '" & Trim(Text3) & "' and (act_typ like 'PUR' or act_typ like 'BOT')", con, adOpenStatic
    If rs.RecordCount > 0 Then Label12.Caption = rs("ACT_nme") & ""
    Text1.Enabled = False
    Text12.SetFocus
    MOD1 = True
    Toolbar1.Buttons(12).Enabled = True
    Toolbar1.Buttons(13).Enabled = False
Else
    PRDLIST
End If
End Sub

Private Sub PRDLIST()
If CHECKTB("act.dbf,pgr.dbf,prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check1.Value = 1 Then
rs.Open "select prd.prd_idy,prd.prd_nme,prd.prd_prn,prd.brd_nme,prd.alt_idy,prd.cur_lev,act.act_nme,pgr.grp_cat,prd.prd_uom,prd.sub_unt,prd.con_fac,prd.stk_typ from prd,act,pgr where prd.prd_nme like '" & Trim(Text1) & "%' and act.act_idy like prd.cmp_idy  and pgr.grp_idy like prd.grp_idy group by prd.prd_idy", con, adOpenKeyset, adLockPessimistic
ElseIf Check2.Value = 1 Then
rs.Open "select prd.prd_idy,prd.prd_nme,prd.prd_prn,prd.brd_nme,prd.alt_idy,prd.cur_lev,act.act_nme,pgr.grp_cat,prd.prd_uom,prd.sub_unt,prd.con_fac,prd.stk_typ from prd,act,pgr where prd.prd_prn like '" & Trim(Text1) & "%' and act.act_idy like prd.cmp_idy  and pgr.grp_idy like prd.grp_idy group by prd.prd_idy", con, adOpenKeyset, adLockPessimistic
ElseIf Check5.Value = 1 Then
rs.Open "select prd.prd_idy,prd.prd_nme,prd.prd_prn,prd.brd_nme,prd.alt_idy,prd.cur_lev,act.act_nme,pgr.grp_cat,prd.prd_uom,prd.sub_unt,prd.con_fac,prd.stk_typ from prd,act,pgr where prd.prd_nme like '" & Trim(Text1) & "%' and act.act_idy like prd.cmp_idy and pgr.grp_idy like prd.grp_idy and stk_typ ='POP' group by prd.prd_idy", con, adOpenKeyset, adLockPessimistic
ElseIf Check6.Value = 1 Then
rs.Open "select prd.prd_idy,prd.prd_nme,prd.prd_prn,prd.brd_nme,prd.alt_idy,prd.cur_lev,act.act_nme,pgr.grp_cat,prd.prd_uom,prd.sub_unt,prd.con_fac,prd.stk_typ from prd,act,pgr where prd.prd_nme like '" & Trim(Text1) & "%' and act.act_idy like prd.cmp_idy  and pgr.grp_idy like prd.grp_idy and stk_typ ='Promotional Item' group by prd.prd_idy", con, adOpenKeyset, adLockPessimistic
Else
rs.Open "select prd.prd_idy,prd.prd_nme,prd.prd_prn,prd.brd_nme,prd.alt_idy,prd.cur_lev,act.act_nme,pgr.grp_cat,prd.prd_uom,prd.sub_unt,prd.con_fac,prd.stk_typ from prd,act,pgr where prd.prd_nme like '" & Trim(Text1) & "%' and act.act_idy like prd.cmp_idy  and pgr.grp_idy like prd.grp_idy group by prd.prd_idy", con, adOpenKeyset, adLockPessimistic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("PRD_NME")))
        ls.ListSubItems.Add , , Trim(rs("PRD_PRN"))
        ls.ListSubItems.Add , , Trim(QTY(rs("CUR_LEV") & ""))
        ls.ListSubItems.Add , , Trim(rs("SUB_UNT") & "")
        ls.ListSubItems.Add , , QTY(rs("CON_FAC") & "")
        ls.ListSubItems.Add , , Trim(rs("PRD_UOM") & "")
        ls.ListSubItems.Add , , Trim(rs("GRP_CAT") & "")
        ls.ListSubItems.Add , , Trim(rs("BRD_NME") & "")
        ls.ListSubItems.Add , , Trim(rs("ACT_NME") & "")
        ls.ListSubItems.Add , , Trim(rs("ALT_IDY") & "")
        ls.ListSubItems.Add , , Trim(rs("PRD_IDY") & "")
        ls.ListSubItems.Add , , Trim(rs("STK_TYP"))
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label31 = "Select Product - SKU"
    Check1.Visible = True
    Check2.Visible = True
    Check5.Visible = True
    Check6.Visible = True
    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Product - SKU, Not Found, Define Product - SKU !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
rs.Close
Set rs = Nothing
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
    If M_FRM_PRD.ActiveControl.name = "Text1" Then PRDLIST
    If M_FRM_PRD.ActiveControl.name = "Text2" Then prdgrplist
    If M_FRM_PRD.ActiveControl.name = "Text3" Then suplist
ElseIf Button.KEY = "d" Then
    DEL_REC
ElseIf Button.KEY = "pre" Then
    If SSTab1.Tab <> 0 Then SSTab1.Tab = 0 ' CHANGED BY PKB
    SYS_PSW.Tag = "PRD_PRE"
    cancel
    SYS_PSW.Show 1
    M_MDI.STS_BAR.Panels(1).Text = "Set Prefix for Product ID"
ElseIf Button.KEY = "p" Then
'    M_FRM_PPQ.Tag = "PRD"
'    M_FRM_PPQ.Show 1
ElseIf Button.KEY = "pr" Then
ElseIf Button.KEY = "crt" Then
    CREATE
ElseIf Button.KEY = "brw" Then
    SSTab1.TabEnabled(2) = True
    BROWSE
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim rs As New ADODB.Recordset
With MDIForm1
If ButtonMenu.KEY = "A" Then
    SYS_PSW.Tag = "Pur Check"
    SYS_PSW.Show 1
    cancel
ElseIf ButtonMenu.KEY = "B" Then
    M_FRM_PRE.Caption = "Product ID Prefix"
    M_FRM_PRE.Show 1
    cancel
ElseIf ButtonMenu.KEY = "PRD_LST" Then
    M_FRM_PPQ.Tag = "PRD"
    M_FRM_PPQ.Show 1
    cancel
ElseIf ButtonMenu.KEY = "BAT_LST" Then
     M_FRM_PPQ.Tag = "BAT"
     M_FRM_PPQ.Show 1
     cancel
ElseIf ButtonMenu.KEY = "PRC_LST" Then
     M_FRM_PPQ.Tag = "PRC"
     M_FRM_PPQ.Show 1
     cancel
ElseIf ButtonMenu.KEY = "STK_LST" Then
     M_FRM_PPQ.Tag = "STK"
     M_FRM_PPQ.Show 1
     cancel
ElseIf ButtonMenu.KEY = "RAT_LST" Then
     M_FRM_PPQ.Tag = "RAT"
     M_FRM_PPQ.Show 1
     cancel
ElseIf ButtonMenu.KEY = "BAL_VAL" Then
     M_FRM_PPQ.Tag = "BAL_VAL"
     M_FRM_PPQ.Show 1
     cancel
ElseIf ButtonMenu.KEY = "LST_IDY" Then
     M_FRM_PPQ.Tag = "LST_IDY"
     M_FRM_PPQ.Show 1
     cancel
ElseIf ButtonMenu.KEY = "LEV_REO" Then
    If CHECKTB("prd.dbf,tpd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
    UPD_STK
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    con.Execute "DELETE FROM TPD"
    For A = 1 To rs.RecordCount
        con.Execute "INSERT INTO TPD VALUES('" & Trim(rs("PRD_IDY")) & "'," & CONQTY(rs("REO_LEV"), rs("PRD_IDY")) & ",0,0)"
        rs.MoveNext
    Next
    OpenReport CURDIR & "\stock.RP1", "Reorder level", 0, "", 6
    cancel
ElseIf ButtonMenu.KEY = "LEV_SAF" Then
    If CHECKTB("prd.dbf,tpd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
    UPD_STK
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    con.Execute "DELETE FROM TPD"
    For A = 1 To rs.RecordCount
        con.Execute "INSERT INTO TPD VALUES('" & Trim(rs("PRD_IDY")) & "'," & CONQTY(rs("REO_LEV"), rs("PRD_IDY")) & ",0,0)"
        rs.MoveNext
    Next
    OpenReport CURDIR & "\stock.RP1", "Safety Level", 0, "", 6
    cancel
ElseIf ButtonMenu.KEY = "BAL_PRD" Then
    M_FRM_PPQ.Tag = "Balance Stock"
    M_FRM_PPQ.Show 1
ElseIf ButtonMenu.KEY = "BAL_SEG" Then
    UPD_STK
    OpenReport CURDIR & "\stock.RP1", "Group Balance", 0, "", 6
ElseIf ButtonMenu.Index >= 12 Then
    If CHECKTB("prd.dbf,brd.dbf,cld.dbf,inv.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
    If RS1.State = 1 Then RS1.Close
    If rs.State = 1 Then rs.Close
    con.Execute "delete from cld"
    rs.Open "select * from brd,prd where prd.prd_idy like brd.prd_idy", con, adOpenKeyset, adLockPessimistic
    RS1.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        RS1.AddNew
        RS1!prd_idy = rs!BAT_IDY
        RS1(1) = MTS(Val(rs!prd_qty & ""), Val(rs!SUB_UNT))
        RS1(2) = 0
        RS1(3) = 0
        RS1(4) = 0
        RS1(5) = 0
        RS1(6) = 0
        RS1(7) = 0
        RS1(8) = Val(rs!SUB_UNT)
        RS1.Update
        rs.MoveNext
    Next
    
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT SUM(TOT_PCS),bat_idy,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('PUR','SAL','CNV') group by bat_idy", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        con.Execute "update cld set pur_stk = " & Val(rs(0)) & " where prd_idy like '" & Trim(rs(1)) & "'", d
        rs.MoveNext
    Next

    If rs.State = 1 Then rs.Close
    rs.Open "SELECT SUM(TOT_PCS),bat_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('STO','ADJ') group by bat_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        con.Execute "update cld set pur_stk = pur_stk -" & Val(rs(0)) & " where prd_idy like '" & Trim(rs(1)) & "'", d
        rs.MoveNext
    Next

    If rs.State = 1 Then rs.Close
    rs.Open "SELECT SUM(TOT_PCS),bat_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Salable' and txn_typ in('CNV') group by bat_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To rs.RecordCount
        con.Execute "update cld set pur_stk = pur_stk -" & Val(rs(0)) & " where prd_idy like '" & Trim(rs(1)) & "'", d
        rs.MoveNext
    Next
    If ButtonMenu.Index = 12 Then
        OpenReport CURDIR & "\LISTS.RP1", "Stock on Hand", 0, "", 6
        cancel
    ElseIf ButtonMenu.Index = 13 Then
        OpenReport CURDIR & "\LISTS.RP1", "Stock on Hand - Brand", 0, "", 6
        cancel
    ElseIf ButtonMenu.Index = 14 Then
        OpenReport CURDIR & "\LISTS.RP1", "Stock on Hand - Segment", 0, "", 6
        cancel
    ElseIf ButtonMenu.Index = 15 Then
        OpenReport CURDIR & "\LISTS.RP1", "Stock on Hand - Category", 0, "", 6
        cancel
    ElseIf ButtonMenu.Index = 16 Then
        OpenReport CURDIR & "\STOCK.RP1", "Stock Breakup", 0, "", 6
        cancel
    ElseIf ButtonMenu.Index = 17 Then
        M_FRM_DSN1.Tag = "SNV_STT"
        M_FRM_DSN1.Show 1
    ElseIf ButtonMenu.Index = 18 Then
        OpenReport CURDIR & "\LISTS.RP1", "Stock on Hand - Pack Size", 0, "", 6
        cancel
    ElseIf ButtonMenu.Index = 19 Then
        OpenReport CURDIR & "\LISTS.RP1", "Stock on Hand - Pack Type", 0, "", 6
        cancel
    End If
End If
End With
End Sub

Public Sub SEQUENCE()
If CHECKTB("prd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from prd order by prd_idy", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    MS_SEQ.Rows = rs.RecordCount + 1
    For A = 1 To rs.RecordCount
        MS_SEQ.TextMatrix(A, 0) = Trim(rs("prd_idy"))
        MS_SEQ.TextMatrix(A, 1) = Trim(rs("prd_nme"))
        MS_SEQ.TextMatrix(A, 2) = rs("BRW_SEQ") & ""
        MS_SEQ.TextMatrix(A, 3) = Trim(rs("brd_nme") & "")
        MS_SEQ.TextMatrix(A, 4) = Trim(rs("grp_idy")) & ""
        MS_SEQ.TextMatrix(A, 5) = Trim(rs("cmp_idy") & "")
        MS_SEQ.TextMatrix(A, 6) = Trim(rs("prd_prn")) & ""
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    MS_SEQ.Visible = True
    MS_SEQ.SetFocus
    MS_SEQ.Col = 2
    MS_SEQ.Row = 1
End If
End Sub
