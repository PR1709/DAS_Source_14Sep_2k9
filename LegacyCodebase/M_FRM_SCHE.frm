VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SCHE 
   Caption         =   "Schemes"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "M_FRM_SCHE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox PRO_CHK 
      Alignment       =   1  'Right Justify
      Caption         =   "&Pro-Rata"
      Height          =   195
      Left            =   7440
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   975
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
            Picture         =   "M_FRM_SCHE.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCHE.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Scheme  Details"
      TabPicture(0)   =   "M_FRM_SCHE.frx":334E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SCH_LST"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "S&KU Level"
      TabPicture(1)   =   "M_FRM_SCHE.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "M_FLX_AMT"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FLX_AMT"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MSFlexGrid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "BAS_PRD_VIW"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "FRE_PRD_VIW"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LBL_QTY"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LBL_AMT"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "LBL_PCG"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "&Bill Level"
      TabPicture(2)   =   "M_FRM_SCHE.frx":3386
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "BILL_LBL"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "MS_BILL"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "AMT_TXT"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "&Multiple SKU Level"
      TabPicture(3)   =   "M_FRM_SCHE.frx":33A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "BAS_CHK"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "FRE_CHK"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Text2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "SKU_AMT"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame6"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame7"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Check1"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "MS_BAS"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "MS_FRE"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label3"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label7"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "&UOM Product Group"
      TabPicture(4)   =   "M_FRM_SCHE.frx":33BE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).Control(1)=   "Frame9"
      Tab(4).Control(2)=   "UOM_AMT"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Text5"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame10"
      Tab(4).Control(5)=   "ALL_FRE"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "ALL_BAS"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "BAS_MS"
      Tab(4).Control(8)=   "FRE_MS"
      Tab(4).Control(9)=   "Label8"
      Tab(4).Control(10)=   "Label9"
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "&View"
      TabPicture(5)   =   "M_FRM_SCHE.frx":33DA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "list1"
      Tab(5).Control(1)=   "Label20"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Scheme Particulars"
         Height          =   2415
         Left            =   -74400
         TabIndex        =   32
         Top             =   480
         Width           =   7335
         Begin VB.ComboBox SUB_CMB 
            Height          =   315
            ItemData        =   "M_FRM_SCHE.frx":33F6
            Left            =   960
            List            =   "M_FRM_SCHE.frx":3403
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "(Select Discount Entry)"
            Top             =   1845
            Width           =   2175
         End
         Begin VB.ComboBox SCH_CMB 
            Height          =   315
            ItemData        =   "M_FRM_SCHE.frx":3432
            Left            =   960
            List            =   "M_FRM_SCHE.frx":343F
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "(Select Base Entry)"
            Top             =   1365
            Width           =   2175
         End
         Begin VB.TextBox SCH_NME 
            Height          =   330
            Left            =   945
            MaxLength       =   40
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   855
            Width           =   2865
         End
         Begin VB.ListBox SCH_ACT 
            Height          =   255
            ItemData        =   "M_FRM_SCHE.frx":346E
            Left            =   5535
            List            =   "M_FRM_SCHE.frx":3478
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   900
            Width           =   495
         End
         Begin VB.TextBox SCH_IDY 
            Height          =   330
            Left            =   945
            MaxLength       =   10
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   360
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker Dtpicker1 
            Height          =   330
            Left            =   5535
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1350
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
            Format          =   24641539
            CurrentDate     =   37099
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   5535
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1830
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
            Format          =   24641539
            CurrentDate     =   37099
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   4530
            TabIndex        =   44
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   4575
            TabIndex        =   43
            Top             =   1905
            Width           =   675
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   42
            Top             =   930
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   930
            Width           =   420
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   495
            TabIndex        =   40
            Top             =   435
            Width           =   165
         End
      End
      Begin VB.TextBox M_FLX_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -72000
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox FLX_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -67800
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73080
         MaxLength       =   40
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox AMT_TXT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   2880
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Scheme Value"
         Height          =   855
         Left            =   4800
         TabIndex        =   24
         Top             =   420
         Width           =   3975
         Begin VB.TextBox GRS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Invoice Gross Amount"
            Height          =   195
            Left            =   480
            TabIndex        =   26
            Top             =   435
            Width           =   1560
         End
      End
      Begin VB.CheckBox BAS_CHK 
         Alignment       =   1  'Right Justify
         Caption         =   "S&elect All SKUs"
         Height          =   195
         Left            =   -71950
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox FRE_CHK 
         Alignment       =   1  'Right Justify
         Caption         =   "Select A&ll SKUs"
         Height          =   255
         Left            =   -67680
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox SKU_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -72480
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -74860
         TabIndex        =   17
         Top             =   480
         Width           =   1935
         Begin VB.OptionButton BAS_OR 
            Caption         =   "&OR"
            Height          =   255
            Left            =   1080
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton BAS_AND 
            Caption         =   "&AND"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -70420
         TabIndex        =   14
         Top             =   480
         Width           =   1935
         Begin VB.OptionButton FRE_AND 
            Caption         =   "A&ND"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton FRE_OR 
            Caption         =   "O&R"
            Height          =   255
            Left            =   960
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -70440
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
         Begin VB.OptionButton FRE_OR_OPT 
            Caption         =   "OR"
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton FRE_AND_OPT 
            Caption         =   "AND"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -74760
         TabIndex        =   8
         Top             =   2160
         Width           =   1815
         Begin VB.OptionButton BAS_AND_OPT 
            Caption         =   "AND"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton BAS_OR_OPT 
            Caption         =   "OR"
            Height          =   255
            Left            =   1080
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox UOM_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -72600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame10 
         Caption         =   "Product (SKU) Name"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   4
         Top             =   360
         Width           =   8415
         Begin MSComctlLib.ListView UOM_LST 
            Height          =   1380
            Left            =   120
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   195
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2434
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SKU Name"
               Object.Width           =   4058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "MRP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Stock"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "UOM"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Batch"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Conversion Factor"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Product ID"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.CheckBox ALL_FRE 
         Alignment       =   1  'Right Justify
         Caption         =   "Sele&ct All SKUs"
         Height          =   255
         Left            =   -67800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1455
      End
      Begin VB.CheckBox ALL_BAS 
         Alignment       =   1  'Right Justify
         Caption         =   "Selec&t All SKUs"
         Height          =   375
         Left            =   -72120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mi&xed"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -71250
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   135
         Left            =   -72840
         TabIndex        =   45
         Top             =   2760
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   238
         _Version        =   393216
      End
      Begin MSComctlLib.ListView SCH_LST 
         Height          =   1935
         Left            =   -74400
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3000
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3413
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Scheme Name"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Active"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Start Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "End Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid BAS_PRD_VIW 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   5
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483641
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid FRE_PRD_VIW 
         Height          =   4095
         Left            =   -70440
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483641
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid MS_BILL 
         Height          =   3700
         Left            =   60
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8800
         _ExtentX        =   15531
         _ExtentY        =   6535
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483641
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   3
         GridLinesFixed  =   3
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
      Begin MSFlexGridLib.MSFlexGrid MS_BAS 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4400
         _ExtentX        =   7752
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483641
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
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
      Begin MSFlexGridLib.MSFlexGrid MS_FRE 
         Height          =   3975
         Left            =   -70600
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483641
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
      End
      Begin MSFlexGridLib.MSFlexGrid BAS_MS 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2760
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483641
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
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
      Begin MSFlexGridLib.MSFlexGrid FRE_MS 
         Height          =   2295
         Left            =   -70440
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2760
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483641
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
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
            Text            =   "Product Category"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Segment"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base SKU && Quantity (Ctrl+B)"
         Height          =   195
         Left            =   -74760
         TabIndex        =   64
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label LBL_QTY 
         AutoSize        =   -1  'True
         Caption         =   "Free SKU && Quantity (Ctrl+F)"
         Height          =   195
         Left            =   -70440
         TabIndex        =   63
         Top             =   2280
         Width           =   1995
      End
      Begin VB.Label LBL_AMT 
         AutoSize        =   -1  'True
         Caption         =   "Discount Amount (Ctrl+F)"
         Height          =   195
         Left            =   -70440
         TabIndex        =   62
         Top             =   2280
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label LBL_PCG 
         AutoSize        =   -1  'True
         Caption         =   "Discount Percentage (Ctrl+F)"
         Height          =   195
         Left            =   -70440
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label BILL_LBL 
         AutoSize        =   -1  'True
         Caption         =   "Free SKU && Quantity"
         Height          =   195
         Left            =   200
         TabIndex        =   60
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+B)"
         Height          =   195
         Left            =   -72840
         TabIndex        =   59
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+F)"
         Height          =   195
         Left            =   -68400
         TabIndex        =   58
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+B)"
         Height          =   195
         Left            =   -72720
         TabIndex        =   57
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+F)"
         Height          =   195
         Left            =   -68400
         TabIndex        =   56
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Product (SKU) Name"
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
         Left            =   -74760
         TabIndex        =   55
         Top             =   720
         Width           =   7905
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_SCHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
