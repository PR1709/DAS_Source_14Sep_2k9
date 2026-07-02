VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SCH 
   Caption         =   "Schemes"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "M_FRM_SCH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox PRO_CHK 
      Alignment       =   1  'Right Justify
      Caption         =   "&Pro-Rata"
      Height          =   195
      Left            =   7680
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   2640
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
            Picture         =   "M_FRM_SCH.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SCH.frx":2E02
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
      Tabs            =   5
      TabsPerRow      =   5
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
      TabPicture(0)   =   "M_FRM_SCH.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SCH_LST"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "S&KU Level"
      TabPicture(1)   =   "M_FRM_SCH.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text3"
      Tab(1).Control(1)=   "FLX_AMT"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "M_FLX_AMT"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MSFlexGrid2"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(5)=   "BAS_PRD_VIW"
      Tab(1).Control(6)=   "FRE_PRD_VIW"
      Tab(1).Control(7)=   "LBL_PCG"
      Tab(1).Control(8)=   "LBL_AMT"
      Tab(1).Control(9)=   "LBL_QTY"
      Tab(1).Control(10)=   "Label1"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "&Bill Level"
      TabPicture(2)   =   "M_FRM_SCH.frx":3386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "BILL_LBL"
      Tab(2).Control(1)=   "MS_BILL"
      Tab(2).Control(2)=   "Text1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "AMT_TXT"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame4"
      Tab(2).Control(5)=   "Frame3"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&Multiple SKU Level"
      TabPicture(3)   =   "M_FRM_SCH.frx":33A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label3"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "MS_FRE"
      Tab(3).Control(3)=   "MS_BAS"
      Tab(3).Control(4)=   "BAS_CHK"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "FRE_CHK"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame5"
      Tab(3).Control(7)=   "Text2"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "SKU_AMT"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Frame6"
      Tab(3).Control(10)=   "Frame7"
      Tab(3).Control(11)=   "Check1"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "&UOM Product Group"
      TabPicture(4)   =   "M_FRM_SCH.frx":33BE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label8"
      Tab(4).Control(1)=   "Label9"
      Tab(4).Control(2)=   "FRE_MS"
      Tab(4).Control(3)=   "BAS_MS"
      Tab(4).Control(4)=   "Frame8"
      Tab(4).Control(5)=   "Frame9"
      Tab(4).Control(6)=   "UOM_AMT"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Text5"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Frame10"
      Tab(4).Control(9)=   "ALL_FRE"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "ALL_BAS"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).ControlCount=   11
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mi&xed"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -71400
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox ALL_BAS 
         Alignment       =   1  'Right Justify
         Caption         =   "Selec&t All SKUs"
         Height          =   375
         Left            =   -72120
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox ALL_FRE 
         Alignment       =   1  'Right Justify
         Caption         =   "Sele&ct All SKUs"
         Height          =   255
         Left            =   -67800
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Frame Frame10 
         Caption         =   "Product (SKU) Name"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   58
         Top             =   360
         Width           =   8415
         Begin MSComctlLib.ListView UOM_LST 
            Height          =   1380
            Left            =   120
            TabIndex        =   59
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
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox UOM_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -72600
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -74760
         TabIndex        =   53
         Top             =   2160
         Width           =   1815
         Begin VB.OptionButton BAS_OR_OPT 
            Caption         =   "OR"
            Height          =   255
            Left            =   1080
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton BAS_AND_OPT 
            Caption         =   "AND"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -70440
         TabIndex        =   50
         Top             =   2160
         Width           =   1815
         Begin VB.OptionButton FRE_AND_OPT 
            Caption         =   "AND"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton FRE_OR_OPT 
            Caption         =   "OR"
            Height          =   255
            Left            =   1080
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -70440
         TabIndex        =   47
         Top             =   2160
         Width           =   1935
         Begin VB.OptionButton FRE_OR 
            Caption         =   "O&R"
            Height          =   255
            Left            =   960
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton FRE_AND 
            Caption         =   "A&ND"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Condition Applicable"
         Height          =   615
         Left            =   -74760
         TabIndex        =   44
         Top             =   2160
         Width           =   1935
         Begin VB.OptionButton BAS_AND 
            Caption         =   "&AND"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   240
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton BAS_OR 
            Caption         =   "&OR"
            Height          =   255
            Left            =   1080
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox SKU_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -72360
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73560
         MaxLength       =   40
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Product (SKU) Name"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   40
         Top             =   360
         Width           =   8415
         Begin MSComctlLib.ListView SKU_LST 
            Height          =   1380
            Left            =   120
            TabIndex        =   41
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
      Begin VB.CheckBox FRE_CHK 
         Alignment       =   1  'Right Justify
         Caption         =   "Select A&ll SKUs"
         Height          =   255
         Left            =   -67800
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1455
      End
      Begin VB.CheckBox BAS_CHK 
         Alignment       =   1  'Right Justify
         Caption         =   "S&elect All SKUs"
         Height          =   195
         Left            =   -72120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Product (SKU) Name"
         Height          =   1785
         Left            =   -74760
         TabIndex        =   34
         Top             =   360
         Width           =   8415
         Begin MSComctlLib.ListView PRD_VIW 
            Height          =   1380
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2434
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
      Begin VB.Frame Frame4 
         Caption         =   "Scheme Value"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   30
         Top             =   2580
         Width           =   3975
         Begin VB.TextBox GRS_AMT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   31
            TabStop         =   0   'False
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
      Begin VB.TextBox AMT_TXT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -67800
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -68880
         MaxLength       =   40
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -73080
         MaxLength       =   40
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox FLX_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -67800
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox M_FLX_AMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   -72000
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scheme Particulars"
         Height          =   2415
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   7335
         Begin VB.TextBox SCH_IDY 
            Height          =   330
            Left            =   945
            MaxLength       =   10
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   360
            Width           =   1095
         End
         Begin VB.ListBox SCH_ACT 
            Height          =   255
            ItemData        =   "M_FRM_SCH.frx":33DA
            Left            =   5535
            List            =   "M_FRM_SCH.frx":33E4
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   900
            Width           =   495
         End
         Begin VB.TextBox SCH_NME 
            Height          =   330
            Left            =   945
            MaxLength       =   40
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   855
            Width           =   2865
         End
         Begin VB.ComboBox SCH_CMB 
            Height          =   315
            ItemData        =   "M_FRM_SCH.frx":33EE
            Left            =   960
            List            =   "M_FRM_SCH.frx":33FB
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "(Select Base Entry)"
            Top             =   1365
            Width           =   2175
         End
         Begin VB.ComboBox SUB_CMB 
            Height          =   315
            ItemData        =   "M_FRM_SCH.frx":342A
            Left            =   960
            List            =   "M_FRM_SCH.frx":3437
            TabIndex        =   3
            TabStop         =   0   'False
            Text            =   "(Select Discount Entry)"
            Top             =   1845
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker Dtpicker1 
            Height          =   330
            Left            =   5535
            TabIndex        =   7
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
            TabIndex        =   8
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   495
            TabIndex        =   24
            Top             =   435
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   930
            Width           =   420
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   11
            Top             =   930
            Width           =   450
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   4575
            TabIndex        =   10
            Top             =   1905
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   4530
            TabIndex        =   9
            Top             =   1425
            Width           =   720
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   135
         Left            =   -72840
         TabIndex        =   1
         Top             =   2760
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   238
         _Version        =   393216
      End
      Begin MSComctlLib.ListView SCH_LST 
         Height          =   1935
         Left            =   720
         TabIndex        =   12
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
      Begin VB.Frame Frame2 
         Caption         =   "Product (SKU) Name"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   13
         Top             =   360
         Width           =   8415
         Begin MSComctlLib.ListView PRD_LST 
            Height          =   1380
            Left            =   120
            TabIndex        =   14
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
               Text            =   "Measure"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Product ID"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid BAS_PRD_VIW 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4048
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
         Height          =   2295
         Left            =   -70440
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4048
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
         Height          =   2295
         Left            =   -70440
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4048
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
         Height          =   2295
         Left            =   -74760
         TabIndex        =   36
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
      Begin MSFlexGridLib.MSFlexGrid MS_FRE 
         Height          =   2295
         Left            =   -70440
         TabIndex        =   37
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
      Begin MSFlexGridLib.MSFlexGrid BAS_MS 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   62
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
         TabIndex        =   63
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+F)"
         Height          =   195
         Left            =   -68400
         TabIndex        =   67
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+B)"
         Height          =   195
         Left            =   -72720
         TabIndex        =   66
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+F)"
         Height          =   195
         Left            =   -68400
         TabIndex        =   65
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(Ctrl+B)"
         Height          =   195
         Left            =   -72720
         TabIndex        =   64
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label BILL_LBL 
         AutoSize        =   -1  'True
         Caption         =   "Free SKU && Quantity"
         Height          =   195
         Left            =   -70440
         TabIndex        =   33
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label LBL_PCG 
         AutoSize        =   -1  'True
         Caption         =   "Discount Percentage (Ctrl+F)"
         Height          =   195
         Left            =   -70440
         TabIndex        =   26
         Top             =   2280
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label LBL_AMT 
         AutoSize        =   -1  'True
         Caption         =   "Discount Amount (Ctrl+F)"
         Height          =   195
         Left            =   -70440
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label LBL_QTY 
         AutoSize        =   -1  'True
         Caption         =   "Free SKU && Quantity (Ctrl+F)"
         Height          =   195
         Left            =   -70440
         TabIndex        =   18
         Top             =   2280
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base SKU && Quantity (Ctrl+B)"
         Height          =   195
         Left            =   -74760
         TabIndex        =   17
         Top             =   2280
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
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
Attribute VB_Name = "M_FRM_SCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, BASPRD As String

Private Sub CREATE()
MOD_REC = False
save
End Sub

Private Sub save()
Dim RS As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "Effmcg") = vbNo Then
    cancel
    Schemelist
    Exit Sub
End If
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
If Len(Trim(SCH_NME)) = 0 Then
    MsgBox "Scheme Name Should Not Be Blank, Give Name !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    SCH_NME.Enabled = True
    SCH_NME.SetFocus
    Exit Sub
End If
If Dtpicker1 > DTPicker2 Then
    SSTab1.Tab = 0
    MsgBox "End Date Should Be Greater Than Start Date.", vbExclamation, "Effmcg"
    DTPicker2.SetFocus
    Exit Sub
End If
If SSTab1.TabEnabled(2) = True And Val(GRS_AMT) < 0 Then
    SSTab1.Tab = 2
    MsgBox "Enter Gross Amount", vbExclamation, "Effmcg"
    GRS_AMT.SetFocus
    Exit Sub
End If
If MOD_REC = False Then
    If SSTab1.TabEnabled(1) = True And BAS_PRD_VIW.Rows < 3 Then
        If Trim(BAS_PRD_VIW.TextMatrix(1, 0)) = "" Then
        MsgBox "Select Base Product", vbExclamation, "Effmcg"
        SSTab1.Tab = 1
        d = True
        BAS_PRD_VIW.Col = 1
        BAS_PRD_VIW.SetFocus
        Exit Sub
        End If
    End If
End If
If MOD_REC = True Then
    If RS.State = 1 Then RS.Close
    RS.Open "select * from sch where sch_IDY like '" & Trim(SCH_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
    For A = 1 To RS.RecordCount
        RS("SCH_EDT") = DTPicker2
        RS("SCH_STS") = Trim(SCH_ACT)
        RS.Update
        RS.MoveNext
    Next
Else
    RS.Open "select * from sch", CON, adOpenKeyset, adLockOptimistic
    CON.BeginTrans
    schid = "SCH" & T7increment("idy_sch")
    If SCH_CMB.ListIndex = 0 And SUB_CMB.ListIndex = 0 Then
    'base quantity to free quantity
        For A = 1 To BAS_PRD_VIW.Rows - 1
            If Trim(BAS_PRD_VIW.TextMatrix(A, 3)) > 0 And Val(FRE_PRD_VIW.TextMatrix(A, 3)) > 0 Then
                RS.AddNew
                RS("sch_idy") = schid
                RS("sch_nme") = Trim(SCH_NME)
                RS("sch_typ") = SCH_CMB
                RS("f_sch_typ") = SUB_CMB
                RS("SCH_STS") = SCH_ACT
                RS("SCH_BDT") = Dtpicker1
                RS("sch_edt") = DTPicker2
                RS("B_PRD_IDY") = Trim(BAS_PRD_VIW.TextMatrix(A, 0))
                RS("B_PRD_QTY") = Val(BAS_PRD_VIW.TextMatrix(A, 3) & "")
                RS("BAT_IDY") = Trim(BAS_PRD_VIW.TextMatrix(A, 4))
                RS("txn_typ") = "BAS"
                RS("USR_IDY") = USERNAME
                RS("SYS_DAT") = Date
                RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                RS.Update
                savecheck = True
            End If
        Next
        For A = 1 To FRE_PRD_VIW.Rows - 1
            If Trim(BAS_PRD_VIW.TextMatrix(A, 3)) > 0 And Val(FRE_PRD_VIW.TextMatrix(A, 3)) > 0 Then
                RS.AddNew
                RS("sch_idy") = schid
                RS("sch_nme") = Trim(SCH_NME)
                RS("sch_typ") = SCH_CMB
                RS("f_sch_typ") = SUB_CMB
                RS("SCH_STS") = SCH_ACT
                RS("SCH_BDT") = Dtpicker1
                RS("sch_edt") = DTPicker2
                RS("B_PRD_IDY") = Trim(FRE_PRD_VIW.TextMatrix(A, 0)) & ""
                RS("B_PRD_QTY") = Val(FRE_PRD_VIW.TextMatrix(A, 3) & "")
                RS("f_prd_idy") = Trim(BAS_PRD_VIW.TextMatrix(A, 4))
                RS("BAT_IDY") = Trim(FRE_PRD_VIW.TextMatrix(A, 4))
                RS("txn_typ") = "FRE"
                If PRO_CHK.Value = 1 Then
                    RS("pro_rat") = "Y"
                Else
                    RS("pro_rat") = "N"
                End If
                RS("USR_IDY") = USERNAME
                RS("SYS_DAT") = Date
                RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                RS.Update
                savecheck = True
            End If
        Next
    End If
        'free quanity to cashoff or cash discount begin
        If SCH_CMB.ListIndex = 0 And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2) Then
            For A = 1 To BAS_PRD_VIW.Rows - 1
                If Val(BAS_PRD_VIW.TextMatrix(A, 3)) > 0 And Val(FRE_PRD_VIW.TextMatrix(1, 5)) > 0 Then
                    RS.AddNew
                    RS("sch_idy") = schid
                    RS("sch_nme") = Trim(SCH_NME)
                    RS("sch_typ") = SCH_CMB
                    RS("f_sch_typ") = SUB_CMB
                    RS("SCH_STS") = SCH_ACT
                    RS("SCH_BDT") = Dtpicker1
                    RS("sch_edt") = DTPicker2
                    RS("B_PRD_IDY") = Trim(BAS_PRD_VIW.TextMatrix(A, 0))
                    RS("B_PRD_QTY") = Val(BAS_PRD_VIW.TextMatrix(A, 3) & "")
                    RS("BAT_IDY") = Trim(BAS_PRD_VIW.TextMatrix(A, 4))
                    RS("txn_typ") = "BAS"
                    RS("USR_IDY") = USERNAME
                    RS("SYS_DAT") = Date
                    RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                    RS.Update
                    b = "ok"
                    savecheck = True
                End If
            Next
            If (convert(AMT(Val(FRE_PRD_VIW.TextMatrix(1, 5)))) > 0 Or discount(FRE_PRD_VIW.TextMatrix(1, 6)) > 0) And b = "ok" Then
                RS.AddNew
                RS("sch_idy") = schid
                RS("sch_nme") = Trim(SCH_NME)
                RS("sch_typ") = SCH_CMB
                RS("f_sch_typ") = SUB_CMB
                RS("SCH_STS") = SCH_ACT
                RS("SCH_BDT") = Dtpicker1
                RS("sch_edt") = DTPicker2
                If SUB_CMB.ListIndex = 1 Then RS("B_DIS_AMT") = convert(AMT(Val(FRE_PRD_VIW.TextMatrix(1, 5)) & ""))
                If SUB_CMB.ListIndex = 2 Then RS("B_DIS_PCG") = discount(FRE_PRD_VIW.TextMatrix(1, 6) & "")
                RS("f_prd_idy") = ""
                RS("BAT_IDY") = ""
                If SUB_CMB.ListIndex = 1 Then RS("txn_typ") = "CAS"
                If SUB_CMB.ListIndex = 2 Then RS("txn_typ") = "PCG"
                If PRO_CHK.Value = 1 Then
                    RS("pro_rat") = "Y"
                Else
                    RS("pro_rat") = "N"
                End If
                RS("USR_IDY") = USERNAME
                RS("SYS_DAT") = Date
                RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                RS.Update
                b = ""
                savecheck = True
            End If
        End If
        'free quanity to cashoff or cash discount end
        If SCH_CMB.ListIndex = 1 Then
            For A = 1 To MS_BILL.Rows - 1
                If (Val(GRS_AMT) > 0 And Val(MS_BILL.TextMatrix(A, 3)) > 0) Or ((convert(AMT(Val(MS_BILL.TextMatrix(A, 5)))) > 0 Or discount(MS_BILL.TextMatrix(A, 6)) > "0.00") And Val(GRS_AMT) > 0) Then
                    RS.AddNew
                    RS("sch_idy") = schid
                    RS("sch_nme") = Trim(SCH_NME)
                    RS("sch_typ") = SCH_CMB
                    RS("f_sch_typ") = SUB_CMB
                    RS("SCH_STS") = SCH_ACT
                    RS("SCH_BDT") = Dtpicker1
                    RS("sch_edt") = DTPicker2
                    If SUB_CMB.ListIndex = 0 Then
                        RS("B_PRD_IDY") = Trim(MS_BILL.TextMatrix(A, 0)) & ""
                        RS("B_PRD_QTY") = Val(MS_BILL.TextMatrix(A, 3)) & ""
                    End If
                    RS("B_BIL_AMT") = convert(AMT(Val(GRS_AMT)))
                    If SUB_CMB.ListIndex = 1 Then RS("B_DIS_AMT") = convert(AMT(Val(MS_BILL.TextMatrix(A, 5)) & ""))
                    If SUB_CMB.ListIndex = 2 Then RS("B_DIS_PCG") = discount(MS_BILL.TextMatrix(A, 6) & "")
                    RS("BAT_IDY") = Trim(MS_BILL.TextMatrix(A, 4))
                    If SUB_CMB.ListIndex = 0 Then RS("txn_typ") = "FRE"
                    If SUB_CMB.ListIndex = 1 Then RS("txn_typ") = "CAS"
                    If SUB_CMB.ListIndex = 2 Then RS("txn_typ") = "PCG"
                    If PRO_CHK.Value = 1 Then
                        RS("pro_rat") = "Y"
                    Else
                        RS("pro_rat") = "N"
                    End If
                    RS("USR_IDY") = USERNAME
                    RS("SYS_DAT") = Date
                    RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                    RS.Update
                    savecheck = True
                End If
            Next
        End If
    If SCH_CMB.ListIndex = 2 Then
        For A = 1 To MS_BAS.Rows - 1
        If Trim(MS_BAS.TextMatrix(A, 0)) <> "" Then
            RS.AddNew
            RS("sch_idy") = schid
            RS("sch_nme") = Trim(SCH_NME)
            RS("sch_typ") = SCH_CMB
            RS("f_sch_typ") = SUB_CMB
            RS("SCH_STS") = SCH_ACT
            RS("SCH_BDT") = Dtpicker1
            RS("sch_edt") = DTPicker2
            RS("B_PRD_IDY") = Trim(MS_BAS.TextMatrix(A, 0))
            If BAS_AND.Value = True Then
                RS("B_PRD_QTY") = Val(MS_BAS.TextMatrix(A, 3) & "")
            Else
                RS("B_PRD_QTY") = Val(MS_BAS.TextMatrix(1, 3) & "")
            End If
            RS("B_PRD_QTY") = Val(MS_BAS.TextMatrix(A, 3) & "")
            RS("BAT_IDY") = Trim(MS_BAS.TextMatrix(A, 4))
            RS("txn_typ") = "BAS"
            RS("USR_IDY") = USERNAME
            RS("SYS_DAT") = Date
            RS("SYS_TIM") = Format(Time, "hh:mm:ss")
            If BAS_AND.Value = True Then
                RS("b_con_app") = "AND"
            Else
                If Check1.Value = 0 Then
                    RS("B_CON_APP") = "ORO"
                Else
                    RS("B_CON_APP") = "ORM"
                End If
            End If
            RS.Update
            savecheck = True
            End If
        Next
        For A = 1 To MS_FRE.Rows - 1
            If Trim(MS_FRE.TextMatrix(A, 0)) <> "" Then
                RS.AddNew
                RS("sch_idy") = schid
                RS("sch_nme") = Trim(SCH_NME)
                RS("sch_typ") = SCH_CMB
                RS("f_sch_typ") = SUB_CMB
                RS("SCH_STS") = SCH_ACT
                RS("SCH_BDT") = Dtpicker1
                RS("sch_edt") = DTPicker2
                RS("B_PRD_IDY") = Trim(MS_FRE.TextMatrix(A, 0)) & ""
                If FRE_AND.Value = True Then
                    RS("B_PRD_QTY") = Val(MS_FRE.TextMatrix(A, 3) & "")
                Else
                    RS("B_PRD_QTY") = Val(MS_FRE.TextMatrix(1, 7) & "")
                End If
                RS("BAT_IDY") = Trim(MS_FRE.TextMatrix(A, 4))
                RS("txn_typ") = "FRE"
                If PRO_CHK.Value = 1 Then
                    RS("pro_rat") = "Y"
                Else
                    RS("pro_rat") = "N"
                End If
                RS("USR_IDY") = USERNAME
                RS("SYS_DAT") = Date
                RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                If FRE_AND.Value = True Then
                    RS("f_con_app") = "AND"
                Else
                    RS("f_CON_APP") = "ORO"
                End If
                RS.Update
                savecheck = True
            ElseIf convert(AMT(Val(MS_FRE.TextMatrix(A, 5)))) > 0 Or discount(MS_FRE.TextMatrix(A, 6)) > 0 Then
                RS.AddNew
                RS("sch_idy") = schid
                RS("sch_nme") = Trim(SCH_NME)
                RS("sch_typ") = SCH_CMB
                RS("f_sch_typ") = SUB_CMB
                RS("SCH_STS") = SCH_ACT
                RS("SCH_BDT") = Dtpicker1
                RS("sch_edt") = DTPicker2
                If SUB_CMB.ListIndex = 1 Then RS("B_DIS_AMT") = convert(AMT(Val(MS_FRE.TextMatrix(A, 5)) & ""))
                If SUB_CMB.ListIndex = 2 Then RS("B_DIS_PCG") = discount(MS_FRE.TextMatrix(A, 6) & "")
                RS("BAT_IDY") = Trim(MS_FRE.TextMatrix(A, 4))
                If SUB_CMB.ListIndex = 1 Then RS("txn_typ") = "CAS"
                If SUB_CMB.ListIndex = 2 Then RS("txn_typ") = "PCG"
                If PRO_CHK.Value = 1 Then
                    RS("pro_rat") = "Y"
                Else
                    RS("pro_rat") = "N"
                End If
                RS("USR_IDY") = USERNAME
                RS("SYS_DAT") = Date
                RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                RS.Update
                savecheck = True
            End If
        Next
    End If
    If SCH_CMB.ListIndex = 3 Then
        For A = 1 To BAS_MS.Rows - 1
            If Trim(BAS_MS.TextMatrix(A, 0)) <> "" Then
                RS.AddNew
                RS("sch_idy") = schid
                RS("sch_nme") = Trim(SCH_NME)
                RS("sch_typ") = SCH_CMB
                RS("f_sch_typ") = SUB_CMB
                RS("SCH_STS") = SCH_ACT
                RS("SCH_BDT") = Dtpicker1
                RS("sch_edt") = DTPicker2
                RS("B_PRD_IDY") = Trim(BAS_MS.TextMatrix(A, 0))
                RS("B_PRD_QTY") = Val(BAS_MS.TextMatrix(A, 3) & "")
                RS("BAT_IDY") = Trim(BAS_MS.TextMatrix(A, 4))
                RS("txn_typ") = "BAS"
                RS("USR_IDY") = USERNAME
                RS("SYS_DAT") = Date
                RS("SYS_TIM") = Format(Time, "hh:mm:ss")
                If BAS_AND_OPT.Value = True Then
                    RS("b_con_app") = "AND"
                Else
                    If Check1.Value = 1 Then
                        RS("B_CON_APP") = "ORM"
                    Else
                        RS("B_CON_APP") = "ORO"
                    End If
                End If
                RS.Update
                savecheck = True
            End If
        Next
        
        For A = 1 To FRE_MS.Rows - 1
        If Trim(FRE_MS.TextMatrix(A, 0)) <> "" Then
            RS.AddNew
            RS("sch_idy") = schid
            RS("sch_nme") = Trim(SCH_NME)
            RS("sch_typ") = SCH_CMB
            RS("f_sch_typ") = SUB_CMB
            RS("SCH_STS") = SCH_ACT
            RS("SCH_BDT") = Dtpicker1
            RS("sch_edt") = DTPicker2
            RS("B_PRD_IDY") = Trim(FRE_MS.TextMatrix(A, 0)) & ""
            If FRE_AND_OPT.Value = True Then
                RS("B_PRD_QTY") = Val(FRE_MS.TextMatrix(A, 3) & "")
            Else
                RS("B_PRD_QTY") = Val(FRE_MS.TextMatrix(1, 7) & "")
            End If
            RS("BAT_IDY") = Trim(FRE_MS.TextMatrix(A, 4))
            RS("txn_typ") = "FRE"
            If PRO_CHK.Value = 1 Then
                RS("pro_rat") = "Y"
            Else
                RS("pro_rat") = "N"
            End If
            RS("USR_IDY") = USERNAME
            RS("SYS_DAT") = Date
            RS("SYS_TIM") = Format(Time, "hh:mm:ss")
            If FRE_AND_OPT.Value = True Then
                RS("f_con_app") = "AND"
            Else
                RS("f_CON_APP") = "ORO"
            End If
            RS.Update
            savecheck = True
        ElseIf convert(AMT(Val(FRE_MS.TextMatrix(A, 5)))) > 0 Or discount(FRE_MS.TextMatrix(A, 6)) > 0 Then
            RS.AddNew
            RS("sch_idy") = schid
            RS("sch_nme") = Trim(SCH_NME)
            RS("sch_typ") = SCH_CMB
            RS("f_sch_typ") = SUB_CMB
            RS("SCH_STS") = SCH_ACT
            RS("SCH_BDT") = Dtpicker1
            RS("sch_edt") = DTPicker2
            If SUB_CMB.ListIndex = 1 Then RS("B_DIS_AMT") = convert(AMT(Val(FRE_MS.TextMatrix(A, 5)) & ""))
            If SUB_CMB.ListIndex = 2 Then RS("B_DIS_PCG") = discount(FRE_MS.TextMatrix(A, 6) & "")
            RS("BAT_IDY") = Trim(FRE_MS.TextMatrix(A, 4))
            If SUB_CMB.ListIndex = 1 Then RS("txn_typ") = "CAS"
            If SUB_CMB.ListIndex = 2 Then RS("txn_typ") = "PCG"
            If PRO_CHK.Value = 1 Then
                RS("pro_rat") = "Y"
            Else
                RS("pro_rat") = "N"
            End If
            RS("USR_IDY") = USERNAME
            RS("SYS_DAT") = Date
            RS("SYS_TIM") = Format(Time, "hh:mm:ss")
            RS.Update
            savecheck = True
        End If
        Next
    End If
End If
If MOD_REC = False Then
    If savecheck = True Then
        CON.CommitTrans
        MsgBox "New Scheme ID is " & schid, vbExclamation, "Effmcg"
    Else
        CON.RollbackTrans
    End If
End If
cancel
Schemelist
clrctr M_FRM_SCH
idnew = False
SSTab1.Tab = 0
SCH_IDY.Enabled = True
SCH_IDY.SetFocus
MOD_REC = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub cancel()
clrctr M_FRM_SCH
setval
Schemelist
idnew = False
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SCH_IDY.Enabled = True
SCH_IDY.SetFocus
MOD_REC = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub newrec()
MOD_REC = False
setval
clrctr M_FRM_SCH
PRO_CHK.Enabled = True
PRO_CHK.Visible = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.Tab = 0
SCH_NME.SetFocus
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub Schemelist()
If CHECKTB("sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim RS As New ADODB.Recordset
Dim ls As ListItem
If RS.State = 1 Then RS.Close
If MOD_REC = True Then
    RS.Open "select * from SCH where SCH_STS LIKE 'Y' GROUP BY SCH_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    RS.Open "select * from SCH where SCH_NME like '" & Trim(SCH_IDY) & "%' AND SCH_STS LIKE 'Y' GROUP BY SCH_IDY", CON, adOpenKeyset, adLockOptimistic
End If
If RS.RecordCount > 0 Then
    SCH_LST.ListItems.Clear
    For A = 1 To RS.RecordCount
        Set ls = SCH_LST.ListItems.Add(, , Trim(RS("SCH_nme")))
        ls.ListSubItems.Add , , Trim(RS("SCH_STS"))
        ls.ListSubItems.Add , , Format(RS("SCH_BDT"), "DD/MM/YYYY")
        ls.ListSubItems.Add , , Format(RS("SCH_EDT"), "DD/MM/YYYY")
        ls.ListSubItems.Add , , Trim(RS!SCH_IDY)
        RS.MoveNext
    Next
    SSTab1.Tab = 0
    SCH_LST.Visible = True
    If SCH_IDY <> "" Then SCH_LST.SetFocus
ElseIf RS.RecordCount = 0 Then
    MsgBox "No Schemes Found!", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    SCH_IDY.Enabled = True
    SCH_IDY.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub exitfrm()
Unload Me
End Sub

Private Sub ALL_BAS_Click()
If ALL_BAS.Value = 1 Then
    ALL_BAS.Enabled = False
    DISPLAYALL
ElseIf ALL_BAS.Value = 0 Then
    BAS_MS.Rows = 1
    BAS_MS.Enabled = True
    BAS_MS.AddItem ""
    BAS_MS.Col = 1
    BAS_MS.SetFocus
End If
End Sub

Private Sub ALL_FRE_Click()
If ALL_FRE.Value = 1 Then
    DISPLAYALL
ElseIf ALL_FRE.Value = 0 Then
    FRE_MS.Rows = 1
    FRE_MS.Enabled = True
    ALL_FRE.Enabled = True
    FRE_MS.AddItem ""
    FRE_MS.Col = 1
    FRE_MS.SetFocus
End If
End Sub

Private Sub AMT_TXT_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Then
ElseIf K = 13 And Trim(MS_BILL.TextMatrix(MS_BILL.Row, 0)) <> "" Then
    If MS_BILL.Col = 3 Then
        MS_BILL.TextMatrix(MS_BILL.Row, 3) = AMT_TXT
        AMT_TXT.Visible = False
        
'        If Val(MS_BILL.TextMatrix(MS_BILL.Row, 3)) > 0 Then
            MS_BILL.AddItem ""
            MS_BILL.Col = 1
            If MS_BILL.Row = MS_BILL.Rows - 1 Then
            Else
                MS_BILL.Row = MS_BILL.Row + 1
            End If
        MS_BILL.Col = 1
        MS_BILL.SetFocus
 '       Else
    '        MsgBox "Enter Free Quantity", vbExclamation, "Effmcg"
    '        MS_BILL.Col = 3
    '        MS_BILL.SetFocus
    '        Exit Sub
    
'        MS_BILL.Col = 1
 '       MS_BILL.SetFocus
 '       End If
    End If
ElseIf K = 13 And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2) Then
    If AMT_TXT > 0 Then
        MS_BILL.TextMatrix(MS_BILL.Row, 5) = convert(AMT(Val(AMT_TXT)))
        MS_BILL.TextMatrix(MS_BILL.Row, 6) = discount(AMT_TXT)
        AMT_TXT.Visible = False
        save
    Else
        MS_BILL.TextMatrix(MS_BILL.Row, 5) = convert(AMT(Val(AMT_TXT)))
        MS_BILL.TextMatrix(MS_BILL.Row, 6) = discount(AMT_TXT)
        AMT_TXT.Visible = False
    End If
ElseIf K = 46 Then
Else
    MS_BILL.Col = 1
    MS_BILL.SetFocus
End If
End Sub

Private Sub BAS_AND_Click()
If BAS_AND.Value = True Then
    If MS_BAS.Enabled = True Then MS_BAS.SetFocus
    Check1.Value = 0
    Check1.Enabled = False
Else
End If
End Sub

Private Sub BAS_AND_OPT_Click()
If BAS_AND_OPT.Value = True Then
    If BAS_MS.Enabled = True Then
        BAS_MS.SetFocus
    End If
Else
End If
End Sub

Private Sub BAS_CHK_Click()
If BAS_CHK.Value = 1 Then
    BAS_CHK.Enabled = False
    DISPLAYALL
ElseIf BAS_CHK.Value = 0 Then
    MS_BAS.Rows = 1
    MS_BAS.Enabled = True
    MS_BAS.AddItem ""
    MS_BAS.Col = 1
    MS_BAS.SetFocus
End If
End Sub

Private Sub BAS_MS_EnterCell()
If BAS_MS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product View"
ElseIf BAS_MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf BAS_MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter UOM"
End If
End Sub

Private Sub BAS_MS_GotFocus()
BAS_MS.HighLight = flexHighlightAlways
End Sub

Private Sub BAS_MS_KeyDown(K As Integer, S As Integer)
If SSTab1.TabEnabled(4) = True And K = 70 And S = 2 Then
 FRE_MS.Enabled = True
 If FRE_MS.Rows < 2 Then
    ALL_FRE.Enabled = True
    FRE_MS.AddItem ""
 Else
    FRE_MS.Enabled = True
    FRE_MS.SetFocus
 End If
 If Trim(FRE_MS.TextMatrix(FRE_MS.Row, 0)) = "" Then
    Text5.Visible = False
'    BAS_MS.Enabled = False
    FRE_MS.Enabled = True
    If SUB_CMB.ListIndex = 0 Then FRE_MS.Col = 1
    If SUB_CMB.ListIndex = 1 Then FRE_MS.Col = 5
    If SUB_CMB.ListIndex = 2 Then FRE_MS.Col = 6
    FRE_MS.SetFocus
 Else
    BAS_MS.SetFocus
 End If
End If
End Sub

Private Sub BAS_MS_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0
If BAS_MS.Col = 3 Then
    If K > 47 And K < 58 Or K = 46 Then
        UOM_AMT = ""
        UOM_AMT.MaxLength = 9
        UOM_AMT.Visible = True
        UOM_AMT.Top = BAS_MS.Top + BAS_MS.CellTop
        UOM_AMT.Left = BAS_MS.CellLeft + BAS_MS.Left
        UOM_AMT.Width = BAS_MS.CellWidth
        UOM_AMT.Height = BAS_MS.CellHeight
        UOM_AMT.SetFocus
        BASPRD = "BASE"
        UOM_AMT = Chr(K)
        UOM_AMT.SelStart = Len(UOM_AMT)
    End If
ElseIf BAS_MS.Col = 1 Then
If K = 39 Or K = 41 Then K = 0
If BAS_MS.Row <> 0 Then
    If BAS_MS.Col = 1 Then
        If K = 13 Then
            If Trim(BAS_MS.TextMatrix(BAS_MS.Row, 0)) <> "" Then
                BAS_MS.Col = 3
                BAS_MS.SetFocus
            Else
                BASPRD = "BASE"
                UOM_LST.Enabled = True
                UOM_LST.SetFocus
            End If
        ElseIf K <> 27 Then
            If K > 96 And K < 123 Then
            K = K - 32
            Text5.Visible = True
            Text5.Top = BAS_MS.Top + BAS_MS.CellTop
            Text5.Left = BAS_MS.CellLeft + BAS_MS.Left
            Text5.Width = BAS_MS.CellWidth
            Text5.Height = BAS_MS.CellHeight
            Text5.SetFocus
            If K <> 13 Then Text5 = Chr(K)
            Text5.SelStart = Len(Text5)
            Else
            K = 0
            End If
        End If
    End If
End If
End If
End Sub

Private Sub BAS_MS_LeaveCell()
Text5.Visible = False
UOM_AMT.Visible = False
BAS_MS.Enabled = True
BAS_MS.SetFocus
End Sub

Private Sub BAS_MS_LostFocus()
BAS_MS.HighLight = flexHighlightNever
End Sub


Private Sub BAS_MS_Scroll()
If Text5.Visible = True Then
    Text5 = ""
    Text5.Visible = False
    BAS_MS.SetFocus
ElseIf UOM_AMT.Visible = True Then
    UOM_AMT = ""
    UOM_AMT.Visible = False
    BAS_MS.SetFocus
End If
End Sub

Private Sub BAS_OR_Click()
If BAS_OR.Value = True Then
    If MS_BAS.Enabled = True Then
        MS_BAS.SetFocus
    End If
    Check1.Enabled = True
    For A = 1 To MS_BAS.Rows - 1
        MS_BAS.TextMatrix(A, 3) = ""
    Next
Else
End If
End Sub

Private Sub BAS_prd_viw_GotFocus()
BAS_PRD_VIW.HighLight = flexHighlightAlways
End Sub
Private Sub BAS_prd_viw_lostFocus()
BAS_PRD_VIW.HighLight = flexHighlightNever
End Sub

Private Sub BAS_PRD_VIW_Scroll()
If M_FLX_AMT.Visible = True Then
    M_FLX_AMT = ""
    M_FLX_AMT.Visible = False
    BAS_PRD_VIW.SetFocus
ElseIf Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    BAS_PRD_VIW.SetFocus
End If
End Sub

Private Sub FRE_AND_Click()
If FRE_AND.Value = True Then
    If MS_FRE.Enabled = True Then
        MS_FRE.SetFocus
    Else
        MS_BAS.Enabled = True
        MS_BAS.SetFocus
    End If
Else
End If
End Sub

Private Sub FRE_MS_Scroll()
If Text5.Visible = True Then
    Text5 = ""
    Text5.Visible = False
    FRE_MS.SetFocus
ElseIf UOM_AMT.Visible = True Then
    UOM_AMT = ""
    UOM_AMT.Visible = False
    FRE_MS.SetFocus
End If
End Sub

Private Sub fre_prd_viw_GotFocus()
FRE_PRD_VIW.HighLight = flexHighlightAlways
End Sub
Private Sub fre_prd_viw_lostFocus()
FRE_PRD_VIW.HighLight = flexHighlightNever
End Sub

Private Sub fre_MS_LostFocus()
FRE_MS.HighLight = flexHighlightNever
End Sub

Private Sub BAS_OR_OPT_KeyDown(K As Integer, S As Integer)
If K = 13 And BAS_OR_OPT.Value = True Then
    If BAS_MS.Enabled = True Then BAS_MS.SetFocus
    
 For A = 1 To BAS_MS.Rows - 1
        BAS_MS.TextMatrix(A, 3) = ""
    Next
Else

End If
End Sub

Private Sub BAS_PRD_VIW_EnterCell()
If BAS_PRD_VIW.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product View"
ElseIf BAS_PRD_VIW.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf BAS_PRD_VIW.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Quantity in Pieces"
End If
End Sub

Private Sub BAS_PRD_VIW_KeyDown(K As Integer, S As Integer)
If SSTab1.TabEnabled(1) = True And K = 70 And S = 2 Then
 If Trim(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 0)) = "" Then
    Text3.Visible = False
    'BAS_PRD_VIW.Enabled = False
    'FRE_PRD_VIW.Enabled = True
    FRE_PRD_VIW.SetFocus
 Else
    'BAS_PRD_VIW.Enabled = True
    BAS_PRD_VIW.SetFocus
 End If
End If
End Sub

Private Sub FRE_MS_GotFocus()
FRE_MS.HighLight = flexHighlightAlways
End Sub

Private Sub FRE_PRD_VIW_Scroll()
If FLX_AMT.Visible = True Then
    FLX_AMT = ""
    FLX_AMT.Visible = False
    FRE_PRD_VIW.SetFocus
ElseIf Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    FRE_PRD_VIW.SetFocus
End If
End Sub

Private Sub MS_BAS_GotFocus()
MS_BAS.HighLight = flexHighlightAlways
End Sub
Private Sub MS_BAS_lostFocus()
MS_BAS.HighLight = flexHighlightNever
End Sub

Private Sub MS_BAS_Scroll()
If Text2.Visible = True Then
    Text2 = ""
    Text2.Visible = False
    MS_BAS.SetFocus
ElseIf SKU_AMT.Visible = True Then
    SKU_AMT = ""
    SKU_AMT.Visible = False
    MS_BAS.SetFocus
End If
End Sub

Private Sub MS_BILL_Scroll()
If Text1.Visible = True Then
    Text1 = ""
    Text1.Visible = False
    MS_BILL.SetFocus
ElseIf AMT_TXT.Visible = True Then
    AMT_TXT = ""
    AMT_TXT.Visible = False
    MS_BILL.SetFocus
End If
End Sub

Private Sub MS_fre_GotFocus()
MS_FRE.HighLight = flexHighlightAlways
End Sub
Private Sub MS_fre_lostFocus()
MS_FRE.HighLight = flexHighlightNever
End Sub

Private Sub MS_BILL_LeaveCell()
Text1.Visible = False
AMT_TXT.Visible = False
MS_BILL.Enabled = True
If MS_BILL.Visible = True Then MS_BILL.SetFocus
End Sub

Private Sub MS_FRE_LeaveCell()
Text2.Visible = False
SKU_AMT.Visible = False
MS_FRE.Enabled = True
If MS_FRE.Visible = True Then MS_FRE.SetFocus
End Sub

Private Sub MS_FRE_Scroll()
If Text2.Visible = True Then
    Text2 = ""
    Text2.Visible = False
    MS_FRE.SetFocus
ElseIf SKU_AMT.Visible = True Then
    SKU_AMT = ""
    SKU_AMT.Visible = False
    MS_FRE.SetFocus
End If
End Sub

Private Sub PRO_CHK_keypress(K As Integer)
If K = 13 Then
    If SSTab1.TabEnabled(4) = True Then
        BAS_MS.Enabled = True
        BAS_MS.SetFocus
    ElseIf SSTab1.TabEnabled(3) = True Then
        MS_BAS.Enabled = True
        MS_BAS.SetFocus
    ElseIf SSTab1.TabEnabled(2) = True Then
        If SUB_CMB.ListIndex = 1 Then MS_BILL.Col = 5
        If SUB_CMB.ListIndex = 2 Then MS_BILL.Col = 6
        MS_BILL.SetFocus
    ElseIf SSTab1.TabEnabled(1) = True Then
        If SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2 Then
            If SUB_CMB.ListIndex = 2 Then FRE_PRD_VIW.Col = 6
            If SUB_CMB.ListIndex = 1 Then FRE_PRD_VIW.Col = 5
            FRE_PRD_VIW.SetFocus
        ElseIf SUB_CMB.ListIndex = 0 Or SUB_CMB.ListIndex = 0 Then
            FRE_PRD_VIW.SetFocus
        End If
    End If
End If
End Sub

Private Sub FRE_AND_OPT_KeyDown(K As Integer, S As Integer)
If K = 13 And FRE_AND_OPT.Value = True Then
    If FRE_MS.Enabled = True Then
        FRE_MS.SetFocus
    Else
        BAS_MS.Enabled = True
        BAS_MS.SetFocus
    End If
Else
End If
End Sub

Private Sub FRE_CHK_Click()
If FRE_CHK.Value = 1 Then
    DISPLAYALL
ElseIf FRE_CHK.Value = 0 Then
    MS_FRE.Rows = 1
    MS_FRE.Enabled = True
    FRE_CHK.Enabled = True
    MS_FRE.AddItem ""
    MS_FRE.Col = 1
    MS_FRE.SetFocus
End If
End Sub

Private Sub FRE_MS_KeyDown(K As Integer, S As Integer)
If SSTab1.TabEnabled(4) = True And K = 66 And S = 2 Then
 If Trim(BAS_MS.TextMatrix(BAS_MS.Rows - 1, 0)) = "" Then
    Text5.Visible = False
    BAS_MS.Enabled = True
    BAS_MS.SetFocus
  Else
    FRE_MS.Enabled = True
    FRE_MS.SetFocus
  End If
End If
End Sub

Private Sub FRE_MS_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0
If FRE_MS.Col = 3 Or FRE_MS.Col = 5 Or FRE_MS.Col = 6 Or FRE_MS.Col = 7 Then
    If K > 47 And K < 58 Or K = 46 Then
        UOM_AMT = ""
        UOM_AMT.MaxLength = 9
        UOM_AMT.Visible = True
        UOM_AMT.Top = FRE_MS.Top + FRE_MS.CellTop
        UOM_AMT.Left = FRE_MS.CellLeft + FRE_MS.Left
        UOM_AMT.Height = FRE_MS.CellHeight
        UOM_AMT.SetFocus
        
        BASPRD = "FREE"
        UOM_AMT = Chr(K)
        UOM_AMT.SelStart = Len(UOM_AMT)
    End If
ElseIf FRE_MS.Col = 1 Then
If K = 39 Or K = 41 Then K = 0
If FRE_MS.Row <> 0 Then
    If FRE_MS.Col = 1 Then
        If K = 13 Then
            If FRE_MS.Rows > 1 Then
            If Trim(FRE_MS.TextMatrix(FRE_MS.Row, 0)) <> "" Then
                FRE_MS.Col = 3
                FRE_MS.SetFocus
            Else
                BASPRD = "FREE"
                UOM_LST.Enabled = True
                UOM_LST.SetFocus
            End If
            End If
        ElseIf K <> 27 Then
            If K > 96 And K < 123 Then
            K = K - 32
            Text5.Visible = True
            Text5.Top = FRE_MS.Top + FRE_MS.CellTop
            Text5.Left = FRE_MS.CellLeft + FRE_MS.Left
            Text5.Width = FRE_MS.CellWidth
            Text5.Height = FRE_MS.CellHeight
            Text5.SetFocus
            If K <> 13 Then Text5 = Chr(K)
            Text5.SelStart = Len(Text5)
            Else
            K = 0
            End If
        End If
    End If
End If
End If
End Sub

Private Sub FRE_MS_LeaveCell()
Text5.Visible = False
UOM_AMT.Visible = False
FRE_MS.Enabled = True
FRE_MS.SetFocus
End Sub

Private Sub FRE_OR_Click()
If FRE_OR.Value = True Then
    MS_FRE.ColWidth(3) = 0
    MS_FRE.ColWidth(7) = 900
Else
    MS_FRE.ColWidth(3) = 900
    MS_FRE.ColWidth(7) = 0
End If
MS_FRE.ColWidth(1) = 2000
MS_FRE.ColWidth(2) = 1000
MS_FRE.ColWidth(4) = 0
MS_FRE.ColWidth(5) = 0
MS_FRE.ColWidth(6) = 0
If MS_FRE.Enabled = True And FRE_OR.Value = True Then
'    MS_BAS.Enabled = False
    If MS_FRE.Rows = 1 Then
        MS_FRE.AddItem ""
        FRE_CHK.Enabled = True
    End If
    If Trim(MS_FRE.TextMatrix(MS_FRE.Row, 0)) <> "" Then
        MS_FRE.Col = 7
    Else
        MS_FRE.Col = 1
    End If
    MS_FRE.Enabled = True
    MS_FRE.SetFocus
'ElseIf MS_BAS.Enabled = True Then MS_BAS.SetFocus
End If
End Sub

Private Sub FRE_OR_KeyDown(K As Integer, S As Integer)
If K = 13 And FRE_OR.Value = True Then
    If MS_FRE.Enabled = True Then
        MS_FRE.SetFocus
    Else
        MS_BAS.Enabled = True
        MS_BAS.SetFocus
    End If
Else
End If
End Sub

Private Sub FRE_OR_OPT_Click()
If FRE_OR_OPT.Value = True Then
    FRE_MS.ColWidth(3) = 0
    FRE_MS.ColWidth(7) = 900
Else
    FRE_MS.ColWidth(3) = 900
    FRE_MS.ColWidth(7) = 0
End If
FRE_MS.ColWidth(1) = 2000
FRE_MS.ColWidth(2) = 1000
FRE_MS.ColWidth(4) = 0
FRE_MS.ColWidth(5) = 0
FRE_MS.ColWidth(6) = 0
If FRE_MS.Enabled = True And FRE_OR_OPT.Value = True Then
'    PRO_CHK.Enabled = True
    If FRE_MS.Rows = 1 Then
        FRE_MS.AddItem ""
        ALL_FRE.Enabled = True
    End If
'    BAS_MS.Enabled = False
    If Trim(FRE_MS.TextMatrix(FRE_MS.Row, 0)) <> "" Then
        FRE_MS.Col = 7
    Else
        FRE_MS.Col = 1
    End If
    MS_FRE.Enabled = True
    FRE_MS.SetFocus
End If
'If BAS_MS.Enabled = True Then BAS_MS.SetFocus
End Sub

Private Sub FRE_OR_OPT_KeyDown(K As Integer, S As Integer)
If K = 13 Then
    If FRE_MS.Enabled = True Then FRE_MS.SetFocus
End If
End Sub

Private Sub MS_BAS_EnterCell()
If MS_BAS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product View"
ElseIf MS_BAS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf MS_BAS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Quantity in Pieces"
End If
End Sub

Private Sub MS_BAS_KeyDown(K As Integer, S As Integer)
If SSTab1.TabEnabled(3) = True And K = 70 And S = 2 Then
 MS_FRE.Enabled = True
 If MS_FRE.Rows < 2 Then
    FRE_CHK.Enabled = True
    MS_FRE.AddItem ""
 Else
    MS_FRE.Enabled = True
    MS_FRE.SetFocus
 End If
 If Trim(MS_FRE.TextMatrix(MS_FRE.Row, 0)) = "" Then
    Text2.Visible = False
'    MS_BAS.Enabled = False
    MS_FRE.Enabled = True
    If SUB_CMB.ListIndex = 0 Then MS_FRE.Col = 1
    If SUB_CMB.ListIndex = 1 Then MS_FRE.Col = 5
    If SUB_CMB.ListIndex = 2 Then MS_FRE.Col = 6
    MS_FRE.SetFocus
 Else
    MS_BAS.Enabled = True
    MS_BAS.SetFocus
 End If
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
If KeyCode = 38 And Shift = 1 Then SUB_CMB.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If DateDiff("D", Dtpicker1, DTPicker2) < 0 Then
        SSTab1.Tab = 0
        MsgBox "End Date Should Be Greater Than Start Date", vbExclamation, "Effmcg"
        DTPicker2.SetFocus
        Exit Sub
    End If
End If
If MOD_REC = False Then
    If KeyCode = 13 Then
        If SCH_CMB.ListIndex = 0 Then
            SSTab1.Tab = 1
            If BAS_PRD_VIW.Rows = 1 Then
            BAS_PRD_VIW.AddItem ""
            FRE_PRD_VIW.AddItem ""
            End If
            BAS_PRD_VIW.Col = 1
            BAS_PRD_VIW.Enabled = True
            BAS_PRD_VIW.SetFocus
        ElseIf SCH_CMB.ListIndex = 1 Then
            SSTab1.Tab = 2
            PRD_VIW.Enabled = True
            GRS_AMT.Enabled = True
            GRS_AMT.SetFocus
        ElseIf SCH_CMB.ListIndex = 2 Then
            SSTab1.Tab = 3
            SKU_LST.Enabled = True
            MS_BAS.Enabled = True
            MS_BAS.AddItem ""
            MS_BAS.Col = 1
            MS_BAS.SetFocus
        ElseIf SCH_CMB.ListIndex = 3 Then
            SSTab1.Tab = 4
            UOM_LST.Enabled = True
            BAS_MS.Enabled = True
            If SUB_CMB.ListIndex = 0 Then
            End If
            BAS_MS.AddItem ""
            BAS_MS.Col = 1
            BAS_MS.SetFocus
        End If
    End If
End If
If KeyCode = 38 And Shift = 1 Then
    If MOD_REC = False Then
        Dtpicker1.SetFocus
    Else
        SCH_ACT.SetFocus
    End If
ElseIf KeyCode = 13 Then
    If MOD_REC = False Then
        If DTPicker2 < Dtpicker1 Then
        SSTab1.Tab = 0
        MsgBox "End Date Should Be Greater Than Start Date", vbExclamation, "Effmcg"
        DTPicker2.SetFocus
        Exit Sub
        End If
    Else
        save
    End If
End If
End Sub

Private Sub FLX_AMT_Change()
If Len(Trim(FLX_AMT)) = 7 Then SendKeys "{home}+{end}"
End Sub

Private Sub FLX_AMT_KeyPress(K As Integer)
If SUB_CMB.ListIndex = 0 And K = 46 Then K = 0
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Then
ElseIf K = 13 And Trim(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 0)) <> "" Then
    If FRE_PRD_VIW.Col = 3 Then
    FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 3) = FLX_AMT
    FLX_AMT.Visible = False
    'If Val(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 3)) > 0 Then
        'BAS_PRD_VIW.Enabled = True
        If BAS_PRD_VIW.Row = BAS_PRD_VIW.Rows - 1 And BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Rows - 1, 0) <> "" Then
            BAS_PRD_VIW.AddItem ""
            FRE_PRD_VIW.AddItem ""
            BAS_PRD_VIW.Row = BAS_PRD_VIW.Rows - 1
        Else
'            BAS_PRD_VIW.Row = BAS_PRD_VIW.Row + 1
        End If
        BAS_PRD_VIW.Col = 1
        BAS_PRD_VIW.SetFocus
'    Else
'        MsgBox "Enter Free Quantity", vbExclamation, "Effmcg"
'        FRE_PRD_VIW.Col = 3
'        FRE_PRD_VIW.Enabled = True
'        FRE_PRD_VIW.SetFocus
'        Exit Sub
    'End If
    End If
ElseIf K = 13 And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2) Then
        FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 5) = convert(AMT(Val(FLX_AMT)))
        FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 6) = discount(FLX_AMT)
        FLX_AMT.Visible = False
        save
'        'BAS_PRD_VIW.Enabled = True
'        If BAS_PRD_VIW.Row = BAS_PRD_VIW.Rows - 1 Then
'        Else
'            BAS_PRD_VIW.Row = BAS_PRD_VIW.Row + 1
'        End If
'        BAS_PRD_VIW.Col = 1
'        BAS_PRD_VIW.SetFocus
ElseIf K = 46 Then
Else
    FRE_PRD_VIW.Col = 1
'    FRE_PRD_VIW.Enabled = True
    FRE_PRD_VIW.SetFocus
End If
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabVisible(4) = False
If MOD_REC = True Then
    SCH_ACT.SetFocus
Else
    SCH_IDY.Enabled = True
    SCH_IDY.SetFocus
End If
If CHECKTB("prd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "selec * from prd", CON, adOpenKeyset
If RS.RecordCount = 0 Then
    MsgBox "No SKUs Found", vbExclamation
    Unload Me
End If
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf M_FRM_SCH.ActiveControl Is TextBox Then M_FRM_SCH.ActiveControl.Text = Trim(FUNKEY(K))
    If K = 78 And S = 2 Then
        newrec
    ElseIf K = 83 And S = 2 Then
        save
    ElseIf K = 81 And S = 2 Then
        cancel
    ElseIf K = 88 And S = 2 Then
        exitfrm
    ElseIf SSTab1.TabEnabled(1) = True And K = 66 And S = 2 Then
        BAS_PRD_VIW.Enabled = True
        BAS_PRD_VIW.SetFocus
    ElseIf SSTab1.TabEnabled(1) = True And K = 70 And S = 2 Then
        FRE_PRD_VIW.Enabled = True
        FRE_PRD_VIW.SetFocus
    ElseIf SSTab1.TabEnabled(3) = True And K = 66 And S = 2 Then
        MS_BAS.Enabled = True
        MS_BAS.SetFocus
    ElseIf SSTab1.TabEnabled(4) = True And K = 66 And S = 2 Then
        BAS_MS.Enabled = True
        BAS_MS.SetFocus
    ElseIf K = 69 And S = 2 Then
        If M_FRM_SCH.ActiveControl.name = "SCH_IDY" Then
            Schemelist
            SCH_LST.SetFocus
        End If
    ElseIf SSTab1.TabEnabled(3) = True And K = 70 And S = 2 Then
        MS_FRE.Enabled = True
        MS_FRE.SetFocus
    ElseIf SSTab1.TabEnabled(4) = True And K = 70 And S = 2 Then
        FRE_MS.Enabled = True
        FRE_MS.SetFocus
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And S = 2 Then
        M_FRM_DSN.Tag = "Schemes List"
        M_FRM_DSN.Show 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If M_FRM_SCH.ActiveControl.name = "SCH_IDY" Then
        Unload Me
    ElseIf M_FRM_SCH.ActiveControl.name = "SSTab1" Then
        SSTab1.Tab = 0
    Else
        save
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
MOD_REC = False
idnew = False
Toolbar1.Buttons(12).Enabled = False
BAS_PRD_VIW.TextMatrix(0, 0) = "Product ID"
BAS_PRD_VIW.TextMatrix(0, 1) = "Product(SKU) Name"
BAS_PRD_VIW.TextMatrix(0, 2) = "MRP Rs."
BAS_PRD_VIW.TextMatrix(0, 3) = "Qty(Pcs)"
SCl BAS_PRD_VIW, 3
BAS_PRD_VIW.TextMatrix(0, 4) = "Batch"
BAS_PRD_VIW.ColAlignment(2) = 6
BAS_PRD_VIW.ColAlignment(3) = 6
BAS_PRD_VIW.ColWidth(0) = 0
BAS_PRD_VIW.ColWidth(1) = 2000
BAS_PRD_VIW.ColWidth(2) = 1000
BAS_PRD_VIW.ColWidth(3) = 900
BAS_PRD_VIW.ColWidth(4) = 0
FRE_PRD_VIW.TextMatrix(0, 0) = "Product ID"
FRE_PRD_VIW.TextMatrix(0, 1) = "Product(SKU) Name"
FRE_PRD_VIW.TextMatrix(0, 2) = "MRP Rs."
FRE_PRD_VIW.TextMatrix(0, 3) = "Qty(Pcs)"
SCl FRE_PRD_VIW, 3
FRE_PRD_VIW.TextMatrix(0, 4) = "Batch"
FRE_PRD_VIW.TextMatrix(0, 5) = "Cash Off"
FRE_PRD_VIW.TextMatrix(0, 6) = "Dis. Percent"
FRE_PRD_VIW.ColAlignment(2) = 6
FRE_PRD_VIW.ColAlignment(3) = 6
FRE_PRD_VIW.ColAlignment(5) = 6
FRE_PRD_VIW.ColAlignment(6) = 6
FRE_PRD_VIW.ColWidth(0) = 0
FRE_PRD_VIW.ColWidth(1) = 2000
FRE_PRD_VIW.ColWidth(2) = 1000
FRE_PRD_VIW.ColWidth(3) = 900
FRE_PRD_VIW.ColWidth(4) = 0
FRE_PRD_VIW.ColWidth(5) = 0
FRE_PRD_VIW.ColWidth(6) = 0
MS_BILL.TextMatrix(0, 0) = "Product ID"
MS_BILL.TextMatrix(0, 1) = "Product(SKU) Name"
SCl MS_BILL, 1
MS_BILL.TextMatrix(0, 2) = "MRP Rs."
MS_BILL.TextMatrix(0, 3) = "Qty (Pcs)"
SCl MS_BILL, 3
MS_BILL.TextMatrix(0, 4) = "Batch"
MS_BILL.TextMatrix(0, 5) = "Cash Off"
MS_BILL.TextMatrix(0, 6) = "Dis. Percent"
MS_BILL.ColAlignment(2) = 6
MS_BILL.ColAlignment(3) = 6
MS_BILL.ColAlignment(5) = 6
MS_BILL.ColAlignment(6) = 6
MS_BILL.ColWidth(0) = 0
MS_BILL.ColWidth(1) = 2000
MS_BILL.ColWidth(2) = 1000
MS_BILL.ColWidth(3) = 900
MS_BILL.ColWidth(4) = 0
MS_BILL.ColWidth(5) = 0
MS_BILL.ColWidth(6) = 0

MS_BAS.TextMatrix(0, 0) = "Product ID"
MS_BAS.TextMatrix(0, 1) = "Product(SKU) Name"
MS_BAS.TextMatrix(0, 2) = "MRP Rs."
MS_BAS.TextMatrix(0, 3) = "Qty (Pcs)"
SCl MS_BAS, 3
MS_BAS.TextMatrix(0, 4) = "Batch"
MS_BAS.ColAlignment(2) = 6
MS_BAS.ColAlignment(3) = 6
MS_BAS.ColWidth(0) = 0
MS_BAS.ColWidth(1) = 2000
MS_BAS.ColWidth(2) = 1000
MS_BAS.ColWidth(3) = 900
MS_BAS.ColWidth(4) = 0

MS_FRE.TextMatrix(0, 0) = "Product ID"
MS_FRE.TextMatrix(0, 1) = "Product(SKU) Name"
MS_FRE.TextMatrix(0, 2) = "MRP Rs."
MS_FRE.TextMatrix(0, 3) = "Qty (Pcs)"
SCl MS_FRE, 3
MS_FRE.TextMatrix(0, 4) = "Batch"
MS_FRE.TextMatrix(0, 5) = "Cash Off"
MS_FRE.TextMatrix(0, 6) = "Dis. Percent"
MS_FRE.TextMatrix(0, 7) = "Qty (Pcs)"
SCl MS_FRE, 7
MS_FRE.ColAlignment(2) = 6
MS_FRE.ColAlignment(3) = 6
MS_FRE.ColAlignment(5) = 6
MS_FRE.ColAlignment(6) = 6
MS_FRE.ColAlignment(7) = 6
MS_FRE.ColWidth(0) = 0
MS_FRE.ColWidth(1) = 2000
MS_FRE.ColWidth(2) = 1000
MS_FRE.ColWidth(3) = 900
MS_FRE.ColWidth(4) = 0
MS_FRE.ColWidth(5) = 0
MS_FRE.ColWidth(6) = 0
MS_FRE.ColWidth(7) = 0

BAS_MS.TextMatrix(0, 0) = "Product ID"
BAS_MS.TextMatrix(0, 1) = "Product(SKU) Name"
BAS_MS.TextMatrix(0, 2) = "MRP Rs."
BAS_MS.TextMatrix(0, 3) = "UOM"
BAS_MS.ColAlignment(3) = 1
BAS_MS.TextMatrix(0, 4) = "Batch"
BAS_MS.ColAlignment(2) = 6
BAS_MS.ColAlignment(3) = 6
BAS_MS.ColWidth(0) = 0
BAS_MS.ColWidth(1) = 2000
BAS_MS.ColWidth(2) = 1000
BAS_MS.ColWidth(3) = 900
BAS_MS.ColWidth(4) = 0

FRE_MS.TextMatrix(0, 0) = "Product ID"
FRE_MS.TextMatrix(0, 1) = "Product(SKU) Name"
FRE_MS.TextMatrix(0, 2) = "MRP Rs."
FRE_MS.TextMatrix(0, 3) = "UOM"
FRE_MS.ColAlignment(3) = 7
FRE_MS.TextMatrix(0, 4) = "Batch"
FRE_MS.TextMatrix(0, 5) = "Cash Off"
FRE_MS.TextMatrix(0, 6) = "Dis. Percent"
FRE_MS.TextMatrix(0, 7) = "UOM"
FRE_MS.ColAlignment(7) = 7
FRE_MS.ColAlignment(2) = 6
FRE_MS.ColAlignment(3) = 6
FRE_MS.ColAlignment(5) = 6
FRE_MS.ColAlignment(6) = 6
FRE_MS.ColAlignment(7) = 6
FRE_MS.ColWidth(0) = 0
FRE_MS.ColWidth(1) = 2000
FRE_MS.ColWidth(2) = 1000
FRE_MS.ColWidth(3) = 900
FRE_MS.ColWidth(4) = 0
FRE_MS.ColWidth(5) = 0
FRE_MS.ColWidth(6) = 0
FRE_MS.ColWidth(7) = 0
SETYEAR M_FRM_SCH
setval
Dim RS As New ADODB.Recordset
Dim ls As ListItem
If RS.State = 1 Then RS.Close
If MOD_REC = True Then
    RS.Open "select * from SCH where SCH_STS LIKE 'Y' GROUP BY SCH_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    RS.Open "select * from SCH where SCH_NME like '" & Trim(SCH_IDY) & "%' AND SCH_STS LIKE 'Y' GROUP BY SCH_IDY", CON, adOpenKeyset, adLockOptimistic
End If
SCH_LST.ListItems.Clear
For A = 1 To RS.RecordCount
    Set ls = SCH_LST.ListItems.Add(, , Trim(RS("SCH_nme")))
    ls.ListSubItems.Add , , Trim(RS("SCH_STS"))
    ls.ListSubItems.Add , , Format(RS("SCH_BDT"), "DD/MM/YYYY")
    ls.ListSubItems.Add , , Format(RS("SCH_EDT"), "DD/MM/YYYY")
    ls.ListSubItems.Add , , Trim(RS!SCH_IDY)
    RS.MoveNext
Next
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub FRE_PRD_VIW_EnterCell()
If FRE_PRD_VIW.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product View"
ElseIf FRE_PRD_VIW.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf FRE_PRD_VIW.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Quantity in Pieces"
ElseIf FRE_PRD_VIW.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Discount Amount"
ElseIf FRE_PRD_VIW.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Discount In Percentage"
End If
End Sub

Private Sub FRE_PRD_VIW_KeyDown(K As Integer, S As Integer)
If SSTab1.TabEnabled(1) = True And K = 66 And S = 2 Then
 If Trim(BAS_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 0)) = "" Then
    Text3.Visible = False
    BAS_PRD_VIW.Enabled = True
    BAS_PRD_VIW.SetFocus
  Else
    FRE_PRD_VIW.Enabled = True
    FRE_PRD_VIW.SetFocus
  End If
End If
End Sub

Private Sub MS_FRE_KeyDown(K As Integer, S As Integer)
If SSTab1.TabEnabled(3) = True And K = 66 And S = 2 Then
 If Trim(MS_BAS.TextMatrix(MS_BAS.Rows - 1, 0)) = "" Then
    Text2.Visible = False
'    MS_FRE.Enabled = False
    MS_BAS.Enabled = True
    MS_BAS.SetFocus
  Else
    MS_FRE.Enabled = True
    MS_FRE.SetFocus
  End If
End If
End Sub

Private Sub GRS_AMT_KeyDown(K As Integer, S As Integer)
If K = 13 Or K = 40 Then
    If convert(AMT(Val(GRS_AMT))) > 0 Then
        GRS_AMT.Text = convert(AMT(Val(GRS_AMT))) & ""
        If MS_BILL.Rows = 1 Then MS_BILL.AddItem ""
        If SUB_CMB.ListIndex = 0 Then MS_BILL.Col = 1
        If SUB_CMB.ListIndex = 1 Then MS_BILL.Col = 5
        If SUB_CMB.ListIndex = 2 Then MS_BILL.Col = 6
        If PRO_CHK.Visible = True Then
        PRO_CHK.SetFocus
        Else
            MS_BILL.Enabled = True
            MS_BILL.SetFocus
        End If
    Else
        MsgBox "Enter Gross Amount!", vbExclamation, "Effmcg"
        SendKeys "{home}+{end}"
        GRS_AMT.SetFocus
    End If
End If
End Sub

Private Sub GRS_AMT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_FLX_AMT_Change()
If Len(Trim(M_FLX_AMT)) = 7 Then SendKeys "{home}+{end}"
End Sub

Private Sub MS_FRE_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0
If MS_FRE.Col = 3 Or MS_FRE.Col = 5 Or MS_FRE.Col = 6 Or MS_FRE.Col = 7 Then
    If K > 47 And K < 58 Or K = 46 Then
        SKU_AMT = ""
        SKU_AMT.MaxLength = 9
        SKU_AMT.Visible = True
        SKU_AMT.Top = MS_FRE.Top + MS_FRE.CellTop
        SKU_AMT.Left = MS_FRE.CellLeft + MS_FRE.Left
        SKU_AMT.Width = MS_FRE.CellWidth
        SKU_AMT.Height = MS_FRE.CellHeight
        SKU_AMT.SetFocus
        BASPRD = "FREE"
        SKU_AMT = Chr(K)
        SKU_AMT.SelStart = Len(SKU_AMT)
    End If
ElseIf MS_FRE.Col = 1 Then
If K = 39 Or K = 41 Then K = 0
If MS_FRE.Row <> 0 Then
    If MS_FRE.Col = 1 Then
        If K = 13 Then
            If MS_FRE.Rows > 1 Then
                If Trim(MS_FRE.TextMatrix(MS_FRE.Row, 0)) <> "" Then
                MS_FRE.Enabled = True
                MS_FRE.Col = 3
                MS_FRE.SetFocus
                Else
                BASPRD = "FREE"
                SKU_LST.Enabled = True
                SKU_LST.SetFocus
                End If
            End If
        ElseIf K <> 27 Then
            If K > 96 And K < 123 Then
            K = K - 32
            Text2.Visible = True
            Text2.Top = MS_FRE.Top + MS_FRE.CellTop
            Text2.Left = MS_FRE.CellLeft + MS_FRE.Left
            Text2.Width = MS_FRE.CellWidth
            Text2.Height = MS_FRE.CellHeight
            Text2.SetFocus
            If K <> 13 Then Text2 = Chr(K)
            Text2.SelStart = Len(Text2)
            Else
            K = 0
            End If
        End If
    End If
End If
End If
End Sub

Private Sub SKU_AMT_Change()
If Len(Trim(SKU_AMT)) = 7 Then SendKeys "{home}+{end}"
End Sub

Private Sub FRE_PRD_VIW_LeaveCell()
Text3.Visible = False
FLX_AMT.Visible = False
FRE_PRD_VIW.Enabled = True
'FRE_PRD_VIW.SetFocus
End Sub

Private Sub BAS_PRD_VIW_LeaveCell()
Text3.Visible = False
M_FLX_AMT.Visible = False
BAS_PRD_VIW.Enabled = True
'BAS_PRD_VIW.SetFocus
End Sub

Private Sub MS_BAS_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0
If MS_BAS.Col = 3 Then
    If K > 47 And K < 58 Or K = 46 Then
        SKU_AMT = ""
        SKU_AMT.MaxLength = 9
        SKU_AMT.Visible = True
        SKU_AMT.Top = MS_BAS.Top + MS_BAS.CellTop
        SKU_AMT.Left = MS_BAS.CellLeft + MS_BAS.Left
        SKU_AMT.Width = MS_BAS.CellWidth
        SKU_AMT.Height = MS_BAS.CellHeight
        SKU_AMT.SetFocus
        SKU_AMT = Chr(K)
        BASPRD = "BASE"
        SKU_AMT.SelStart = Len(SKU_AMT)
    End If
ElseIf MS_BAS.Col = 1 Then
If K = 39 Or K = 41 Then K = 0
If MS_BAS.Row <> 0 Then
    If MS_BAS.Col = 1 Then
        If K = 13 Then
            If Trim(MS_BAS.TextMatrix(MS_BAS.Row, 0)) <> "" Then
                MS_BAS.Col = 3
                MS_BAS.Enabled = True
                MS_BAS.SetFocus
            Else
                BASPRD = "BASE"
                SKU_LST.Enabled = True
                SKU_LST.SetFocus
            End If
        ElseIf K <> 27 Then
            If K > 96 And K < 123 Then
            K = K - 32
            Text2.Visible = True
            Text2.Top = MS_BAS.Top + MS_BAS.CellTop
            Text2.Left = MS_BAS.CellLeft + MS_BAS.Left
            Text2.Width = MS_BAS.CellWidth
            Text2.Height = MS_BAS.CellHeight
            Text2.SetFocus
            If K <> 13 Then Text2 = Chr(K)
            Text2.SelStart = Len(Text2)
            Else
            K = 0
            End If
        End If
    End If
End If
End If
End Sub

Private Sub MS_BAS_LeaveCell()
Text2.Visible = False
SKU_AMT.Visible = False
MS_BAS.Enabled = True
If MS_BAS.Visible = True Then MS_BAS.SetFocus
End Sub

Private Sub MS_BILL_EnterCell()
If MS_BILL.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product View"
ElseIf MS_BILL.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "MRP"
ElseIf MS_BILL.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Quantity in Pieces"
ElseIf MS_BILL.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Discount Amount"
ElseIf MS_BILL.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Discount In Percentage"
End If
End Sub

Private Sub MS_BILL_KeyDown(K As Integer, S As Integer)
If Trim(MS_BILL.TextMatrix(MS_BILL.Row, 0)) = "" Then
    Text1.Visible = False
'    FRE_PRD_VIW.Enabled = False
    BAS_PRD_VIW.Enabled = True
'    BAS_PRD_VIW.SetFocus
  Else
    FRE_PRD_VIW.Enabled = True
'    FRE_PRD_VIW.SetFocus
  End If
End Sub

Private Sub MS_BILL_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0
If MS_BILL.Col = 3 Or MS_BILL.Col = 5 Or MS_BILL.Col = 6 Then
    If K > 47 And K < 58 Or K = 46 Then
        AMT_TXT = ""
        AMT_TXT.MaxLength = 9
        AMT_TXT.Visible = True
        AMT_TXT.Top = MS_BILL.Top + MS_BILL.CellTop
        AMT_TXT.Left = MS_BILL.CellLeft + MS_BILL.Left
        AMT_TXT.Width = MS_BILL.CellWidth
        AMT_TXT.Height = MS_BILL.CellHeight
        AMT_TXT.SetFocus
        AMT_TXT = Chr(K)
        AMT_TXT.SelStart = Len(AMT_TXT)
    End If
ElseIf MS_BILL.Col = 1 Then
If K = 39 Or K = 41 Then K = 0
If MS_BILL.Row <> 0 Then
    If MS_BILL.Col = 1 Then
        If K = 13 Then
            If Trim(MS_BILL.TextMatrix(MS_BILL.Row, 0)) <> "" Then
                MS_BILL.Col = 3
                MS_BILL.SetFocus
            Else
                PRD_VIW.SetFocus
            End If
        ElseIf K <> 27 Then
            Text1.Visible = True
            Text1.Top = MS_BILL.Top + MS_BILL.CellTop
            Text1.Left = MS_BILL.CellLeft + MS_BILL.Left
            Text1.Width = MS_BILL.CellWidth
            Text1.Height = MS_BILL.CellHeight
            Text1.SetFocus
            If K <> 13 Then Text3 = Chr(K)
            Text1.SelStart = Len(Text1)
        End If
    End If
End If
End If
End Sub

Private Sub PRD_LST_KeyPress(KeyAscii As Integer)
Dim REC_SET As New ADODB.Recordset
If KeyAscii = 13 Then
    If PRD_LST.ListItems.Count = 0 Then
        Exit Sub
    End If
    If PRD_LST.ListItems.Count > 0 Then
     If BASPRD = "BASE" Then
        For A = 1 To BAS_PRD_VIW.Rows - 1
            If Trim(PRD_LST.SelectedItem.ListSubItems.Item(4)) = Trim(BAS_PRD_VIW.TextMatrix(A, 4)) Then
            SSTab1.Tab = 1
            MsgBox "Product Already Exists!", vbExclamation, "Effmcg"
            Text3.Visible = False
            If BAS_PRD_VIW.Row = BAS_PRD_VIW.Rows - 1 Then
            Else
                BAS_PRD_VIW.Row = BAS_PRD_VIW.Row + 1
            End If
            BAS_PRD_VIW.SetFocus
            BAS_PRD_VIW.Col = 1
            Exit Sub
            End If
        Next
        BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 1) = PRD_LST.SelectedItem
        BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 2) = PRD_LST.SelectedItem.ListSubItems.Item(1)
        BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 4) = PRD_LST.SelectedItem.ListSubItems.Item(4)
        BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 0) = PRD_LST.SelectedItem.ListSubItems.Item(6)
        BAS_PRD_VIW.Col = 3
        PRD_LST_V
        BAS_PRD_VIW.SetFocus
     Else
        If BASPRD = "" Then
            PRD_LST.SetFocus
            Exit Sub
        End If
        If SUB_CMB.ListIndex = 0 Then
            FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 1) = PRD_LST.SelectedItem
            FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 2) = PRD_LST.SelectedItem.ListSubItems.Item(1)
            FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 4) = PRD_LST.SelectedItem.ListSubItems.Item(4)
            FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 0) = PRD_LST.SelectedItem.ListSubItems.Item(6)
        End If
        FRE_PRD_VIW.Col = 3
        PRD_LST_V
        FRE_PRD_VIW.SetFocus
     End If
    End If
End If
End Sub

Private Sub PRD_VIW_KeyPress(KeyAscii As Integer)
Dim REC_SET As New ADODB.Recordset
If KeyAscii = 13 Then
    If PRD_VIW.ListItems.Count = 0 Then
        Exit Sub
    End If
    If PRD_VIW.ListItems.Count > 0 Then
        For A = 1 To MS_BILL.Rows - 1
            If PRD_VIW.SelectedItem.ListSubItems.Item(4) = MS_BILL.TextMatrix(A, 4) Then
            SSTab1.Tab = 2
            MsgBox "Product Already Exists!", vbExclamation, "Effmcg"
            MS_BILL.SetFocus
            MS_BILL.Row = A + 1
            Text1.Visible = False
            Exit Sub
            End If
        Next
        If SUB_CMB.ListIndex = 0 Then
            MS_BILL.TextMatrix(MS_BILL.Row, 1) = PRD_VIW.SelectedItem
            MS_BILL.TextMatrix(MS_BILL.Row, 2) = PRD_VIW.SelectedItem.ListSubItems.Item(1)
            MS_BILL.TextMatrix(MS_BILL.Row, 4) = PRD_VIW.SelectedItem.ListSubItems.Item(4)
            MS_BILL.TextMatrix(MS_BILL.Row, 0) = PRD_VIW.SelectedItem.ListSubItems.Item(6)
        End If
        MS_BILL.Col = 3
        Text1.Visible = False
        PRD_LST_V
        MS_BILL.SetFocus
    End If
End If
End Sub

Private Sub SCH_ACT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select 'Y' or 'N'"
End Sub

Private Sub SCH_ACT_KeyDown(K As Integer, S As Integer)
If K = 13 Then
    If MOD_REC = False Then
        If SCH_ACT.Selected(1) = True Then
            MsgBox "Scheme Cannot Be Closed Before Creation", vbExclamation, "Effmcg"
            SCH_ACT.Selected(0) = True
            SCH_ACT.SetFocus
            Exit Sub
        End If
        SCH_CMB.SetFocus
    Else
        DTPicker2.SetFocus
    End If
ElseIf K = 38 And S = 1 Then
    SCH_NME.SetFocus
End If
End Sub

Private Sub SCH_CMB_KeyDown(K As Integer, S As Integer)
If K = 13 Then
    If MOD_REC = False Then
        SUB_CMB.SetFocus
    Else
        DTPicker2.SetFocus
    End If
    If SCH_CMB.ListIndex = 0 Then
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        V_PRD_LST
    ElseIf SCH_CMB.ListIndex = 1 Then
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        PRODUCTLIST
    ElseIf SCH_CMB.ListIndex = 2 Then
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = True
        SSTab1.TabEnabled(4) = False
        skulist
    ElseIf SCH_CMB.ListIndex = 3 Then
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = True
        skulist
    Else
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
    End If
ElseIf K = 38 And S = 1 Then
    SCH_ACT.SetFocus
End If
End Sub

Private Sub SCH_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Scheme Identification"
End Sub

Private Sub SCH_IDY_Keydown(K As Integer, S As Integer)
If K = 13 Or K = 40 Then
If Len(Trim(SCH_IDY)) = 0 Then
    If SCH_LST.ListItems.Count > 0 Then
        SCH_LST.SetFocus
    Else
        MsgBox "No Schemes Found!", vbExclamation
        SCH_IDY.SetFocus
    End If
Else
    dispsch
End If
End If
End Sub

Private Sub SCH_IDY_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub SKU_AMT_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If SUB_CMB.ListIndex = 0 And K = 46 Then K = 0
If K > 47 And K < 58 Or K = 8 Then
ElseIf K = 13 And BASPRD = "BASE" And (MS_BAS.TextMatrix(MS_BAS.Row, 0)) <> "" Then
    BASPRD = ""
    If MS_BAS.Col = 3 Then
    If BAS_AND.Value = True Then MS_BAS.TextMatrix(MS_BAS.Row, 3) = Val(SKU_AMT)
    If BAS_AND.Value = False Then MS_BAS.TextMatrix(1, 3) = Val(SKU_AMT)
    SKU_AMT.Visible = False
    MS_BAS.Enabled = True
    If MS_BAS.Row = MS_BAS.Rows - 1 Then
    Else
        MS_BAS.Row = MS_BAS.Row + 1
    End If
    If Trim(MS_BAS.TextMatrix(MS_BAS.Row, 0)) <> "" Then
        MS_BAS.Col = 3
    Else
        MS_BAS.Col = 1
    End If
    MS_BAS.Enabled = True
    MS_BAS.SetFocus
    End If
'ElseIf SUB_CMB.ListIndex = 0 Then
'    If MS_FRE.Col = 3 Or MS_FRE.Col = 7 Then
'    If MS_FRE.Col = 3 Then MS_FRE.TextMatrix(MS_FRE.Row, 3) = Val(SKU_AMT)
'    If MS_FRE.Col = 7 Then MS_FRE.TextMatrix(1, 7) = Val(SKU_AMT)
'    SKU_AMT.Visible = False
'    If Val(MS_FRE.TextMatrix(MS_FRE.Row, 3)) > 0 Then
'        MS_FRE.Enabled = True
'        If MS_FRE.Row = MS_FRE.Rows - 1 Then
'            MS_FRE.AddItem ""
'            FRE_CHK.Enabled = True
'            MS_FRE.Row = MS_FRE.Row + 1
'        Else
'            MS_FRE.Row = MS_FRE.Row + 1
'        End If
'        If Trim(MS_FRE.TextMatrix(MS_FRE.Row, 0)) <> "" Then
'            MS_FRE.Col = 3
'        Else
'            MS_FRE.Col = 1
'        End If
'        MS_FRE.Enabled = True
'        MS_FRE.SetFocus
'    ElseIf Val(MS_FRE.TextMatrix(MS_FRE.Row, 7)) > 0 Then
'        MS_FRE.Enabled = True
'        MS_FRE.Col = 7
'        MS_FRE.SetFocus
'    Else
'        If MS_FRE.Col = 3 Then
'            MsgBox "Enter Product Quantity", vbExclamation, "Effmcg"
'            MS_FRE.Col = 3
'            MS_FRE.Enabled = True
'            MS_FRE.SetFocus
'            Exit Sub
'        Else
'            MS_BAS.Enabled = True
'            MS_BAS.SetFocus
'        End If
'    End If
'    End If
ElseIf K = 13 And BASPRD = "FREE" And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2 Or SUB_CMB.ListIndex = 0) Then
BASPRD = ""
    If MS_FRE.Col = 3 Or MS_FRE.Col = 5 Or MS_FRE.Col = 6 Or MS_FRE.Col = 7 Then
        If MS_FRE.Col = 3 And FRE_OR.Value = False Then MS_FRE.TextMatrix(MS_FRE.Row, 3) = Val(SKU_AMT)
        If MS_FRE.Col = 3 And FRE_OR.Value = True Then MS_FRE.TextMatrix(1, 3) = Val(SKU_AMT)
        If MS_FRE.Col = 5 Then MS_FRE.TextMatrix(MS_FRE.Row, 5) = convert(AMT(Val(SKU_AMT)))
        If MS_FRE.Col = 6 Then MS_FRE.TextMatrix(MS_FRE.Row, 6) = discount(SKU_AMT)
        If MS_FRE.Col = 7 And FRE_OR.Value = False Then MS_FRE.TextMatrix(MS_FRE.Row, 7) = Val(SKU_AMT)
        If MS_FRE.Col = 7 And FRE_OR.Value = True Then MS_FRE.TextMatrix(1, 7) = Val(SKU_AMT)
        SKU_AMT.Visible = False
        If SUB_CMB.ListIndex = 0 Then
            If Val(MS_FRE.TextMatrix(MS_FRE.Row, 3)) > 0 Then
                MS_FRE.Enabled = True
                FRE_CHK.Enabled = True
                If MS_FRE.Rows = 1 Then MS_FRE.AddItem ""
                If MS_FRE.Row = MS_FRE.Rows - 1 Then
                    MS_FRE.AddItem ""
                Else
                    MS_FRE.Row = MS_FRE.Row + 1
                End If
                MS_FRE.Col = 1
                MS_FRE.Enabled = True
                MS_FRE.SetFocus
            End If
        ElseIf convert(AMT(Val(MS_FRE.TextMatrix(MS_FRE.Row, 5)))) > 0 Or discount(MS_FRE.TextMatrix(MS_FRE.Row, 6)) Then
            If MS_FRE.Row = MS_FRE.Rows - 1 Then
            Else
                MS_FRE.Row = MS_FRE.Row + 1
            End If
            MS_FRE.Col = 1
            MS_FRE.Enabled = True
            MS_FRE.SetFocus
        Else
'            MsgBox "Enter Free Quantity", vbExclamation, "Effmcg"
            MS_FRE.Col = 1
            MS_FRE.Row = MS_FRE.Rows - 1
'            MS_FRE.Enabled = True
'            MS_FRE.SetFocus
'            Exit Sub
        End If
        End If
ElseIf K = 46 Then
Else
'    If MS_BAS.Enabled = True Then
'        MS_BAS.Col = 1
'        MS_BAS.SetFocus
'    Else
'        MS_FRE.Col = 1
'        MS_FRE.Enabled = True
'        MS_FRE.SetFocus
'    End If
End If
End Sub

Private Sub SKU_LST_KeyPress(KeyAscii As Integer)
Dim REC_SET As New ADODB.Recordset
If KeyAscii = 13 Then
    If SKU_LST.ListItems.Count = 0 Then
        Exit Sub
    End If
    If SKU_LST.ListItems.Count > 0 Then
        If BASPRD = "BASE" Then
           For A = 1 To MS_BAS.Rows - 1
               If SKU_LST.SelectedItem.ListSubItems.Item(4) = MS_BAS.TextMatrix(A, 4) Then
                    SSTab1.Tab = 3
                    MsgBox "Product Already Exists!", vbExclamation, "Effmcg"
                    Text2.Visible = False
                    If MS_BAS.Row = MS_BAS.Rows - 1 Then
                    Else
                        MS_BAS.Row = MS_BAS.Row + 1
                    End If
                    MS_BAS.Enabled = True
                    MS_BAS.SetFocus
                    MS_BAS.Col = 1
                    'MS_BAS.Row = A
                    Exit Sub
               End If
           Next
           MS_BAS.AddItem ""
           MS_BAS.TextMatrix(MS_BAS.Row, 1) = SKU_LST.SelectedItem
           MS_BAS.TextMatrix(MS_BAS.Row, 2) = SKU_LST.SelectedItem.ListSubItems.Item(1)
           MS_BAS.TextMatrix(MS_BAS.Row, 4) = SKU_LST.SelectedItem.ListSubItems.Item(4)
           MS_BAS.TextMatrix(MS_BAS.Row, 0) = SKU_LST.SelectedItem.ListSubItems.Item(6)
           BASPRD = ""
           MS_BAS.Col = 3
           Text2.Visible = False
           PRD_LST_V
           MS_BAS.SetFocus
        Else
            If BASPRD = "" Then
                SKU_LST.SetFocus
            Exit Sub
            End If
            If SUB_CMB.ListIndex = 0 Then
            For A = 1 To MS_FRE.Rows - 1
                If Trim(SKU_LST.SelectedItem.ListSubItems.Item(4)) = Trim(MS_FRE.TextMatrix(A, 4)) Then
                    SSTab1.Tab = 3
                    MsgBox "Product Already Exists!", vbExclamation, "Effmcg"
                    Text5.Visible = False
                    If MS_FRE.Row = MS_FRE.Rows - 1 Then
                    Else
                        MS_FRE.Row = MS_FRE.Row + 1
                    End If
                    MS_FRE.SetFocus
                    MS_FRE.Col = 1
                    Exit Sub
                End If
            Next
                MS_FRE.TextMatrix(MS_FRE.Row, 1) = SKU_LST.SelectedItem
                MS_FRE.TextMatrix(MS_FRE.Row, 2) = SKU_LST.SelectedItem.ListSubItems.Item(1)
                MS_FRE.TextMatrix(MS_FRE.Row, 4) = SKU_LST.SelectedItem.ListSubItems.Item(4)
                MS_FRE.TextMatrix(MS_FRE.Row, 0) = SKU_LST.SelectedItem.ListSubItems.Item(6)
            End If
            Text2.Visible = False
            PRD_LST_V
            If FRE_AND.Value = True Then
                MS_FRE.Col = 3
            Else
                FRE_CHK.Enabled = True
                MS_FRE.AddItem ""
                MS_FRE.Col = 1
                If MS_FRE.Row = MS_FRE.Rows - 1 Then
                Else
                    MS_FRE.Row = MS_FRE.Row + 1
                End If
                End If
            skulist
            MS_FRE.Enabled = True
            MS_FRE.SetFocus
            BASPRD = ""
        End If
    End If
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If MOD_REC = False Then
    If SSTab1.Tab = 1 And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2) Then
        PRO_CHK.Visible = True
    ElseIf SSTab1.Tab = 1 And (SUB_CMB.ListIndex = 0 Or SUB_CMB.ListIndex = 0) Then
        PRO_CHK.Visible = True
    ElseIf SSTab1.Tab = 2 And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2) Then
        PRO_CHK.Visible = True
    ElseIf SSTab1.Tab = 3 And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2) Then
        PRO_CHK.Visible = True
    ElseIf SSTab1.Tab = 4 And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2) Then
        PRO_CHK.Visible = True
    Else
        PRO_CHK.Visible = False
        If SCH_IDY.Enabled = True Then
            SCH_IDY.SetFocus
        Else
            SCH_NME.SetFocus
        End If
    End If
Else
    If SSTab1.Tab = 1 And (SUB_CMB = "Cash Off" Or SUB_CMB = "Discount Percent" Or SUB_CMB = "Free Quantity") Then
        PRO_CHK.Enabled = False
        PRO_CHK.Visible = True
    ElseIf SSTab1.Tab = 2 And (SUB_CMB = "Cash Off" Or SUB_CMB = "Discount Percent") Then
        PRO_CHK.Enabled = False
        PRO_CHK.Visible = True
    ElseIf SSTab1.Tab = 3 And (SUB_CMB = "Cash Off" Or SUB_CMB = "Discount Percent") Then
        PRO_CHK.Enabled = False
        PRO_CHK.Visible = True
    ElseIf SSTab1.Tab = 4 And (SUB_CMB = "Cash Off" Or SUB_CMB = "Discount Percent") Then
        PRO_CHK.Enabled = False
        PRO_CHK.Visible = True
    Else
        PRO_CHK.Visible = False
        If SCH_IDY.Enabled = True Then
            SCH_IDY.SetFocus
        Else
            SCH_NME.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text1_Keydown(K As Integer, S As Integer)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
    If MS_BILL.Col = 1 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM BRD,prd WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.prd_IDY LIKE '" & Sincrement(Text1) & "' AND BRD.PRD_QTY > 0", CON, adOpenKeyset
        If REC_SET.RecordCount > 0 Then
            P_DIS_PRD Sincrement(Text1)
        Else
            PRODUCTLIST
            PRD_VIW.SetFocus
       End If
    End If
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub TEXT2_KeyDown(K As Integer, S As Integer)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then
End If
If K = 13 Then
    If MS_BAS.Col = 1 Or MS_FRE.Col = 1 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM BRD,prd WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.prd_IDY LIKE '" & Sincrement(Text2) & "' AND BRD.PRD_QTY > 0", CON, adOpenKeyset
        If REC_SET.RecordCount > 0 Then
            P_DIS_PRD Sincrement(Text2)
        Else
            PRODUCTLIST
            SKU_LST.SetFocus
       End If
    End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text3_KeyDown(K As Integer, S As Integer)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 13 Then
    If BAS_PRD_VIW.Col = 1 Or FRE_PRD_VIW.Col = 1 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM BRD,prd WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.prd_IDY LIKE '" & Sincrement(Text3) & "' AND BRD.PRD_QTY > 0", CON, adOpenKeyset
        If REC_SET.RecordCount > 0 Then
            P_DIS_PRD Sincrement(Text3)
        Else
            V_PRD_LST
            PRD_LST.SetFocus
       End If
    End If
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub FRE_PRD_VIW_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0
If FRE_PRD_VIW.Col = 3 Or FRE_PRD_VIW.Col = 5 Or FRE_PRD_VIW.Col = 6 Then
    If K > 47 And K < 58 Or K = 46 Then
        FLX_AMT = ""
        FLX_AMT.MaxLength = 9
        FLX_AMT.Visible = True
        FLX_AMT.Top = FRE_PRD_VIW.Top + FRE_PRD_VIW.CellTop
        FLX_AMT.Left = FRE_PRD_VIW.CellLeft + FRE_PRD_VIW.Left
        FLX_AMT.Width = FRE_PRD_VIW.CellWidth
        FLX_AMT.Height = FRE_PRD_VIW.CellHeight
        FLX_AMT.SetFocus
        FLX_AMT = Chr(K)
        FLX_AMT.SelStart = Len(FLX_AMT)
    End If
ElseIf FRE_PRD_VIW.Col = 1 Then
If K = 39 Or K = 41 Then K = 0
If FRE_PRD_VIW.Row <> 0 Then
    If FRE_PRD_VIW.Col = 1 Then
        If K = 13 Then
            If Trim(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 0)) <> "" Then
                FRE_PRD_VIW.Col = 3
'                FRE_PRD_VIW.Enabled = True
                FRE_PRD_VIW.SetFocus
            Else
                PRD_LST.SetFocus
                BASPRD = "FREE"
            End If
        ElseIf K <> 27 Then
            If K > 96 And K < 123 Then
            K = K - 32
            Text3.Visible = True
            Text3.Top = FRE_PRD_VIW.Top + FRE_PRD_VIW.CellTop
            Text3.Left = FRE_PRD_VIW.CellLeft + FRE_PRD_VIW.Left
            Text3.Width = FRE_PRD_VIW.CellWidth
            Text3.Height = FRE_PRD_VIW.CellHeight
            Text3.SetFocus
            If K <> 13 Then Text3 = Chr(K)
            Text3.SelStart = Len(Text3)
            Else
            K = 0
            End If
        End If
    End If
End If
End If
End Sub

Private Sub BAS_PRD_VIW_KeyPress(K As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K > 96 And K < 123 Then
K = K - 32
ElseIf K = 46 Then

End If
If K = 39 Or K = 41 Then K = 0

If BAS_PRD_VIW.Col = 3 Then
    If K > 47 And K < 58 Or K = 46 Then
        M_FLX_AMT = ""
        M_FLX_AMT.MaxLength = 9
        M_FLX_AMT.Visible = True
        M_FLX_AMT.Top = BAS_PRD_VIW.Top + BAS_PRD_VIW.CellTop
        M_FLX_AMT.Left = BAS_PRD_VIW.CellLeft + BAS_PRD_VIW.Left
        M_FLX_AMT.Width = BAS_PRD_VIW.CellWidth
        M_FLX_AMT.Height = BAS_PRD_VIW.CellHeight
        M_FLX_AMT.SetFocus
        M_FLX_AMT = Chr(K)
        M_FLX_AMT.SelStart = Len(M_FLX_AMT)
    End If
ElseIf BAS_PRD_VIW.Col = 1 Then
    If K = 39 Or K = 41 Then K = 0
    If BAS_PRD_VIW.Row <> 0 Then
        If BAS_PRD_VIW.Col = 1 Then
            If K = 13 Then
                If Trim(BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 0)) <> "" Then
                    BAS_PRD_VIW.Col = 3
                    'BAS_PRD_VIW.Enabled = True
                    BAS_PRD_VIW.SetFocus
                Else
                    BASPRD = "BASE"
                    PRD_LST.Enabled = True
                    PRD_LST.SetFocus
                End If
            ElseIf K <> 27 Then
                If K > 96 And K < 123 Then
                K = K - 32
                Text3.Visible = True
                Text3.Top = BAS_PRD_VIW.Top + BAS_PRD_VIW.CellTop
                Text3.Left = BAS_PRD_VIW.CellLeft + BAS_PRD_VIW.Left
                Text3.Width = BAS_PRD_VIW.CellWidth
                Text3.Height = BAS_PRD_VIW.CellHeight
                Text3.SetFocus
                If K <> 13 Then Text3 = Chr(K)
                Text3.SelStart = Len(Text3)
                Else
                K = 0
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub SCH_LST_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If SCH_LST.SortOrder = lvwAscending Then
SCH_LST.SortOrder = lvwDescending
ElseIf SCH_LST.SortOrder = lvwDescending Then
SCH_LST.SortOrder = lvwAscending
End If
SCH_LST.SortKey = CH.Index - 1
End Sub

Private Sub SCH_LST_DblClick()
SCH_LST_KeyDown 13, 0
End Sub

Private Sub SCH_LST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Scheme or Press <Esc> to Quit"
End Sub

Private Sub SCH_LST_KeyDown(K As Integer, S As Integer)
If K = 13 Then
'    MOD_REC = False
    setval
    If SCH_LST.ListItems.Count > 0 Then
    SCH_IDY = SCH_LST.SelectedItem.ListSubItems.Item(4)
    dispsch
    End If
End If
End Sub

Private Sub prd_lst_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If PRD_LST.SortOrder = lvwAscending Then
PRD_LST.SortOrder = lvwDescending
ElseIf PRD_LST.SortOrder = lvwDescending Then
PRD_LST.SortOrder = lvwAscending
End If
PRD_LST.SortKey = CH.Index - 1
End Sub

Private Sub prd_lst_DblClick()
PRD_LST_KeyPress 13
End Sub

Private Sub prd_lst_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Product or Press <Esc> to Quit"
End Sub

Private Sub SKU_LST_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If SKU_LST.SortOrder = lvwAscending Then
SKU_LST.SortOrder = lvwDescending
ElseIf SKU_LST.SortOrder = lvwDescending Then
SKU_LST.SortOrder = lvwAscending
End If
SKU_LST.SortKey = CH.Index - 1
End Sub

Private Sub SKU_LST_DblClick()
SKU_LST_KeyPress 13
End Sub

Private Sub SKU_LST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Product or Press <Esc> to Quit"
End Sub

Private Sub UOM_LST_DblClick()
UOM_LST_KeyPress 13
End Sub

Private Sub UOM_LST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Product or Press <Esc> to Quit"
End Sub

Private Sub SSTab1_DblClick()
Toolbar1.Enabled = True
If SSTab1.Tab = 0 And SCH_IDY.Enabled = True Then
   SCH_IDY.SetFocus
Else
    SCH_NME.SetFocus
End If
End Sub

Private Sub SCH_NME_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Scheme, Discount Entry Description"
End Sub

Private Sub SCH_NME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(SCH_NME)) > 0 Then
        If MOD_REC = False Then
            SCH_CMB.SetFocus
        Else
            DTPicker2.SetFocus
        End If
    Else
        MsgBox "Scheme Name Should Not Be Blank, Give Name !", vbExclamation, "Effmcg"
        SCH_NME.SetFocus
     End If
ElseIf KeyCode = 38 Then
    If SCH_IDY.Enabled = True Then
        SCH_IDY.SetFocus
    Else
        SCH_NME.SetFocus
    End If
End If
End Sub

Private Sub SCH_NME_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub SUB_CMB_KeyDown(K As Integer, S As Integer)
If K = 13 Then
    If MOD_REC = False Then
        Dtpicker1.SetFocus
    Else
        DTPicker2.SetFocus
    End If
    If SCH_CMB.ListIndex = 0 And SUB_CMB.ListIndex = 0 Then
        LBL_QTY.Visible = True
        LBL_AMT.Visible = False
        LBL_PCG.Visible = False
        FRE_PRD_VIW.ColWidth(1) = 2000
        FRE_PRD_VIW.ColWidth(2) = 1000
        FRE_PRD_VIW.ColWidth(3) = 900
        FRE_PRD_VIW.ColWidth(5) = 0
        FRE_PRD_VIW.ColWidth(6) = 0
    ElseIf SCH_CMB.ListIndex = 0 And SUB_CMB.ListIndex = 1 Then
        LBL_QTY.Visible = False
        LBL_AMT.Visible = True
        LBL_PCG.Visible = False
'        PRO_CHK.Enabled = True
'        PRO_CHK.Visible = True
        FRE_PRD_VIW.ColWidth(1) = 0
        FRE_PRD_VIW.ColWidth(2) = 0
        FRE_PRD_VIW.ColWidth(3) = 0
        FRE_PRD_VIW.ColWidth(5) = 900
        FRE_PRD_VIW.ColWidth(6) = 0
    ElseIf SCH_CMB.ListIndex = 0 And SUB_CMB.ListIndex = 2 Then
        LBL_QTY.Visible = False
        LBL_AMT.Visible = False
        LBL_PCG.Visible = True
'        PRO_CHK.Enabled = True
'        PRO_CHK.Visible = True
        FRE_PRD_VIW.ColWidth(1) = 0
        FRE_PRD_VIW.ColWidth(2) = 0
        FRE_PRD_VIW.ColWidth(3) = 0
        FRE_PRD_VIW.ColWidth(5) = 0
        FRE_PRD_VIW.ColWidth(6) = 1000
    ElseIf SCH_CMB.ListIndex = 1 And SUB_CMB.ListIndex = 0 Then
        BILL_LBL.Caption = "Free SKU && Quantity"
        MS_BILL.ColWidth(1) = 2000
        MS_BILL.ColWidth(2) = 1000
        MS_BILL.ColWidth(3) = 900
        MS_BILL.ColWidth(5) = 0
        MS_BILL.ColWidth(6) = 0
    ElseIf SCH_CMB.ListIndex = 1 And SUB_CMB.ListIndex = 1 Then
        BILL_LBL.Caption = "Discount Amount"
'        PRO_CHK.Enabled = True
'        PRO_CHK.Visible = True
        MS_BILL.ColWidth(1) = 0
        MS_BILL.ColWidth(2) = 0
        MS_BILL.ColWidth(3) = 0
        MS_BILL.ColWidth(5) = 900
        MS_BILL.ColWidth(6) = 0
    ElseIf SCH_CMB.ListIndex = 1 And SUB_CMB.ListIndex = 2 Then
        BILL_LBL.Caption = "Discount Percentage"
'        PRO_CHK.Enabled = True
'        PRO_CHK.Visible = True
        MS_BILL.ColWidth(1) = 0
        MS_BILL.ColWidth(2) = 0
        MS_BILL.ColWidth(3) = 0
        MS_BILL.ColWidth(5) = 0
        MS_BILL.ColWidth(6) = 1000
    ElseIf SCH_CMB.ListIndex = 2 And SUB_CMB.ListIndex = 0 Then
        MS_FRE.ColWidth(1) = 2000
        MS_FRE.ColWidth(2) = 1000
        MS_FRE.ColWidth(3) = 900
        MS_FRE.ColWidth(5) = 0
        MS_FRE.ColWidth(6) = 0
        MS_FRE.ColWidth(7) = 0
        PRO_CHK.Visible = False
        Frame7.Visible = True
        FRE_CHK.Visible = True
    ElseIf SCH_CMB.ListIndex = 2 And SUB_CMB.ListIndex = 1 Then
        MS_FRE.ColWidth(1) = 0
        MS_FRE.ColWidth(2) = 0
        MS_FRE.ColWidth(3) = 0
        MS_FRE.ColWidth(5) = 900
        MS_FRE.ColWidth(6) = 0
        MS_FRE.ColWidth(7) = 0
'        PRO_CHK.Visible = True
        Frame7.Visible = False
        FRE_CHK.Visible = False
    ElseIf SCH_CMB.ListIndex = 2 And SUB_CMB.ListIndex = 2 Then
        MS_FRE.ColWidth(1) = 0
        MS_FRE.ColWidth(2) = 0
        MS_FRE.ColWidth(3) = 0
        MS_FRE.ColWidth(5) = 0
        MS_FRE.ColWidth(6) = 1000
        MS_FRE.ColWidth(7) = 0
'        PRO_CHK.Visible = True
'        PRO_CHK.Enabled = True
        Frame7.Visible = False
        FRE_CHK.Visible = False
    ElseIf SCH_CMB.ListIndex = 3 And SUB_CMB.ListIndex = 0 Then
        FRE_MS.ColWidth(1) = 2000
        FRE_MS.ColWidth(2) = 1000
        FRE_MS.ColWidth(3) = 900
        FRE_MS.ColWidth(5) = 0
        FRE_MS.ColWidth(6) = 0
        FRE_MS.ColWidth(7) = 0
        PRO_CHK.Visible = False
    ElseIf SCH_CMB.ListIndex = 3 And SUB_CMB.ListIndex = 1 Then
        FRE_MS.ColWidth(1) = 0
        FRE_MS.ColWidth(2) = 0
        FRE_MS.ColWidth(3) = 0
        FRE_MS.ColWidth(5) = 900
        FRE_MS.ColWidth(6) = 0
        FRE_MS.ColWidth(7) = 0
'        PRO_CHK.Visible = True
'        PRO_CHK.Enabled = True
        ALL_FRE.Visible = False
        Frame8.Visible = False
    ElseIf SCH_CMB.ListIndex = 3 And SUB_CMB.ListIndex = 2 Then
        FRE_MS.ColWidth(1) = 0
        FRE_MS.ColWidth(2) = 0
        FRE_MS.ColWidth(3) = 0
        FRE_MS.ColWidth(5) = 0
        FRE_MS.ColWidth(6) = 1000
        FRE_MS.ColWidth(7) = 0
'        PRO_CHK.Visible = True
'        PRO_CHK.Enabled = True
        ALL_FRE.Visible = False
        Frame8.Visible = False
    End If
End If
If K = 38 And S = 1 Then SCH_CMB.SetFocus
End Sub

Private Sub Text5_keydown(K As Integer, S As Integer)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then
End If
If K = 13 Then
    If BAS_MS.Col = 1 Or FRE_MS.Col = 1 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM BRD,prd WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.prd_IDY LIKE '" & Sincrement(Text5) & "' AND BRD.PRD_QTY > 0", CON, adOpenKeyset
        If REC_SET.RecordCount > 0 Then
            P_DIS_PRD Sincrement(Text5)
        Else
            PRODUCTLIST
            UOM_LST.SetFocus
       End If
    End If
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
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
    exitfrm
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If M_FRM_SCH.ActiveControl.name = "SCH_IDY" Then
        Schemelist
        SCH_LST.SetFocus
    End If
ElseIf Button.KEY = "p" Then
    M_FRM_DSN.Tag = "Schemes List"
    M_FRM_DSN.Show 1
'    Openreport CURDIR & "\lists.rp1", "Schemes List", 0, "", 6
End If
End Sub

Private Sub setval()
MOD_REC = False
clrctr M_FRM_SCH
SCH_NME.Locked = False
SCH_CMB.Locked = False
SUB_CMB.Locked = False
Dtpicker1.Enabled = True
SCH_ACT.Selected(0) = True
PRD_LST.Enabled = True
SCH_CMB.ListIndex = 0
SUB_CMB.ListIndex = 0
BAS_PRD_VIW.Rows = 1
FRE_PRD_VIW.Rows = 1
MS_BILL.Rows = 1
MS_BAS.Rows = 1
MS_FRE.Rows = 1
FRE_PRD_VIW.ColWidth(1) = 2000
FRE_PRD_VIW.ColWidth(2) = 1000
FRE_PRD_VIW.ColWidth(3) = 900
FRE_PRD_VIW.ColWidth(5) = 0
FRE_PRD_VIW.ColWidth(6) = 0
MS_BILL.ColWidth(1) = 2000
MS_BILL.ColWidth(2) = 1000
MS_BILL.ColWidth(3) = 900
MS_BILL.ColWidth(5) = 0
MS_BILL.ColWidth(6) = 0
MS_FRE.ColWidth(1) = 2000
MS_FRE.ColWidth(2) = 1000
MS_FRE.ColWidth(3) = 900
MS_FRE.ColWidth(5) = 0
MS_FRE.ColWidth(6) = 0
FRE_MS.ColWidth(1) = 2000
FRE_MS.ColWidth(2) = 1000
FRE_MS.ColWidth(3) = 900
FRE_MS.ColWidth(5) = 0
FRE_MS.ColWidth(6) = 0
BILL_LBL.Caption = "Free SKU && Quantity"
PRO_CHK.Value = 0
BAS_AND.Value = True
BAS_OR.Value = False
FRE_AND.Value = True
FRE_OR.Value = False
Frame6.Enabled = True
Frame7.Enabled = True
FRE_CHK.Visible = True
BAS_AND_OPT.Value = True
BAS_OR_OPT.Value = False
FRE_AND_OPT.Value = True
FRE_OR_OPT.Value = False
Frame9.Enabled = True
Frame8.Enabled = True
PRO_CHK.Value = 0
ALL_BAS.Enabled = True
ALL_FRE.Enabled = False
ALL_BAS.Value = 0
ALL_FRE.Value = 0
BAS_MS.Rows = 1
FRE_MS.Rows = 1
AMT_TXT.Visible = False
LBL_QTY.Visible = True
LBL_AMT.Visible = False
LBL_PCG.Visible = False
PRO_CHK.Visible = False
PRO_CHK.Value = 0
BAS_CHK.Value = 0
BAS_CHK.Enabled = True
FRE_CHK.Value = 0
FRE_CHK.Enabled = False
PRO_CHK.Value = 0
Frame6.Visible = True
Frame7.Visible = True
BAS_AND.Value = True
BAS_OR.Value = False
FRE_AND.Value = True
FRE_OR.Value = False
Dtpicker1 = Date
DTPicker2 = Date
Text3.Visible = False
Text1.Visible = False
Text2.Visible = False
Text5.Visible = False
M_FLX_AMT.Visible = False
FLX_AMT.Visible = False
SKU_AMT.Visible = False
Frame8.Visible = True
ALL_FRE.Visible = True

Toolbar1.Buttons(12).Enabled = False
Toolbar1.Buttons(13).Enabled = True
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
REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(Text3) & "%'", CON, adOpenKeyset
PRD_LST.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = PRD_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("PRD_MRP") & "")))
        ls.ListSubItems.Add , , QTY(Val(REC_SET("PRD_QTY") & ""))
        ls.ListSubItems.Add , , REC_SET("PRD_UOM") & ""
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , FAC(REC_SET("con_fac") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("PRD_idy") & "")
        REC_SET.MoveNext
    Next
Else
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    Exit Sub
    If BAS_PRD_VIW.Enabled = True Then
        Text3.Visible = False
        Text3 = ""
        V_PRD_LST
        BAS_PRD_VIW.Enabled = True
        BAS_PRD_VIW.SetFocus
        BAS_PRD_VIW.Col = 1
    Else
        Text3.Visible = False
        Text3 = ""
        V_PRD_LST
        FRE_PRD_VIW.Enabled = True
        FRE_PRD_VIW.SetFocus
        FRE_PRD_VIW.Col = 1
    End If
End If
End Sub

Public Sub PRD_LST_V()
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim STK As Boolean
Dim REC_SET As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0", CON, adOpenKeyset
If SSTab1.TabEnabled(1) = True Then
    PRD_LST.ListItems.Clear
ElseIf SSTab1.TabEnabled(2) = True Then
    PRD_VIW.ListItems.Clear
ElseIf SSTab1.TabEnabled(3) = True Then
    SKU_LST.ListItems.Clear
ElseIf SSTab1.TabEnabled(4) = True Then
    UOM_LST.ListItems.Clear
End If
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        If SSTab1.TabEnabled(1) = True Then
            Set ls = PRD_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ElseIf SSTab1.TabEnabled(2) = True Then
            Set ls = PRD_VIW.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ElseIf SSTab1.TabEnabled(3) = True Then
            Set ls = SKU_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ElseIf SSTab1.TabEnabled(4) = True Then
            Set ls = UOM_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        End If
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("PRD_MRP") & "")))
        ls.ListSubItems.Add , , QTY(Val(REC_SET("PRD_QTY") & ""))
        ls.ListSubItems.Add , , REC_SET("PRD_UOM") & ""
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , FAC(REC_SET("con_fac") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("PRD_idy") & "")
        REC_SET.MoveNext
    Next
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    End If
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(K As Integer)
If K = 46 Then K = 0
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Then
ElseIf K = 13 And Trim(BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 0)) <> "" Then
    If BAS_PRD_VIW.Col = 3 Then
    BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 3) = M_FLX_AMT
    M_FLX_AMT.Visible = False
'    If Val(BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 3)) > 0 Then
        If FRE_PRD_VIW.Row = FRE_PRD_VIW.Rows - 1 Then
        Else
            FRE_PRD_VIW.Row = FRE_PRD_VIW.Row + 1
        End If
        If SUB_CMB.ListIndex = 0 And Trim(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Row, 0)) <> "" Then
            'BAS_PRD_VIW.Enabled = False
            'FRE_PRD_VIW.Enabled = True
            'If BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Rows - 1, 0) <> "" Then BAS_PRD_VIW.AddItem ""
        ElseIf SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2 Then
            'BAS_PRD_VIW.Enabled = False
            'FRE_PRD_VIW.Enabled = True
            If PRO_CHK.Visible = True Then PRO_CHK.SetFocus
'            If Val(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Rows - 1, 5)) = 0 Then
'                FRE_PRD_VIW.SetFocus
'            Else
'                BAS_PRD_VIW.SetFocus
'            End If
        End If
        If FRE_PRD_VIW.Row = FRE_PRD_VIW.Rows - 1 Then
        Else
            FRE_PRD_VIW.Row = FRE_PRD_VIW.Row + 1
        End If
        If SUB_CMB.ListIndex = 0 Then
            FRE_PRD_VIW.Col = 1
            FRE_PRD_VIW.Enabled = True
            FRE_PRD_VIW.SetFocus
        ElseIf SUB_CMB.ListIndex = 1 And Val(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Rows - 1, 5)) = 0 Then
            PRO_CHK.SetFocus
            'FRE_PRD_VIW.Col = 5
            'FRE_PRD_VIW.SetFocus
        ElseIf SUB_CMB.ListIndex = 2 And Val(FRE_PRD_VIW.TextMatrix(FRE_PRD_VIW.Rows - 1, 6)) = 0 Then
            PRO_CHK.SetFocus
            'FRE_PRD_VIW.Col = 6
            'FRE_PRD_VIW.SetFocus
        End If
        Text3.Visible = False
        FLX_AMT.Visible = False
        'BAS_PRD_VIW.Enabled = False
        'FRE_PRD_VIW.Enabled = True

'    Else
'        MsgBox "Enter Quantity", vbExclamation, "Effmcg"
'        BAS_PRD_VIW.Col = 3
'        BAS_PRD_VIW.SetFocus
'        Exit Sub
'    End If
    End If
Else
    'BAS_PRD_VIW.Enabled = True
    BAS_PRD_VIW.Col = 1
    BAS_PRD_VIW.SetFocus
End If
End Sub

Public Sub P_DIS_PRD(prdid As String)
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
MSROW = BAS_PRD_VIW.Row
Dim REC_SET As New ADODB.Recordset
REC_SET.Open "select * from BRD,PRD where BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.prd_idy like '" & Trim(prdid) & "' AND BRD.PRD_QTY > 0", CON, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    PRD_LST.Visible = False
    BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 0) = REC_SET("PRD_IDY")
    BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 1) = Trim(REC_SET("PRD_NME") & "")
    BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 2) = convert(AMT(Val(REC_SET("PRD_MRP"))) & "")
    BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 3) = ""
    BAS_PRD_VIW.TextMatrix(BAS_PRD_VIW.Row, 4) = Trim(REC_SET("BAT_IDY")) & ""
    Text3.Visible = False
    SSTab1.Tab = 1
    BAS_PRD_VIW.Row = BAS_PRD_VIW.Row
    BAS_PRD_VIW.Col = 3
    'BAS_PRD_VIW.Enabled = True
    BAS_PRD_VIW.SetFocus
End If
End Sub

Public Sub PRODUCTLIST()
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If SSTab1.TabEnabled(2) = True Then
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(Text1) & "%'", CON, adOpenKeyset
ElseIf SSTab1.TabEnabled(3) = True Then
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(Text2) & "%'", CON, adOpenKeyset
ElseIf SSTab1.TabEnabled(4) = True Then
    REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(Text5) & "%'", CON, adOpenKeyset
End If
If SSTab1.TabEnabled(2) = True Then
    PRD_VIW.ListItems.Clear
ElseIf SSTab1.TabEnabled(3) = True Then
    SKU_LST.ListItems.Clear
ElseIf SSTab1.TabEnabled(4) = True Then
    UOM_LST.ListItems.Clear
End If
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        If SSTab1.TabEnabled(3) = True Then
            Set ls = SKU_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ElseIf SSTab1.TabEnabled(4) = True Then
            Set ls = UOM_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ElseIf SSTab1.TabEnabled(2) = True Then
            Set ls = PRD_VIW.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        End If
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("PRD_MRP") & "")))
        ls.ListSubItems.Add , , QTY(Val(REC_SET("PRD_QTY") & ""))
        ls.ListSubItems.Add , , REC_SET("PRD_UOM") & ""
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , FAC(REC_SET("con_fac") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("PRD_idy") & "")
        REC_SET.MoveNext
    Next
Else
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    Exit Sub
        If SSTab1.TabEnabled(3) = True Then
            Text2.Visible = False
            Text2 = ""
            PRODUCTLIST
            If MS_BAS.Enabled = True Then
                MS_BAS.SetFocus
                MS_BAS.Col = 1
            Else
                MS_FRE.Enabled = True
                MS_FRE.SetFocus
                MS_FRE.Col = 1
            End If
        ElseIf SSTab1.TabEnabled(4) = True Then
            Text5.Visible = False
            Text5 = ""
            PRODUCTLIST
            If BAS_MS.Enabled = True Then
                BAS_MS.SetFocus
                BAS_MS.Col = 1
            Else
                FRE_MS.SetFocus
                FRE_MS.Col = 1
            End If
        Else
            Text1.Visible = False
            Text1 = ""
            PRODUCTLIST
            MS_BILL.Enabled = True
            MS_BILL.SetFocus
            MS_BILL.Col = 1
        End If
End If
End Sub

Public Sub PRODUCTLIST_DIS()
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0", CON, adOpenKeyset
PRD_VIW.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = PRD_VIW.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("PRD_MRP") & "")))
        ls.ListSubItems.Add , , QTY(Val(REC_SET("PRD_QTY") & ""))
        ls.ListSubItems.Add , , REC_SET("PRD_UOM") & ""
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , FAC(REC_SET("con_fac") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("PRD_idy") & "")
        REC_SET.MoveNext
    Next
Else
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    Exit Sub
        Text1.Visible = False
        Text1 = ""
        PRODUCTLIST_DIS
        MS_BILL.SetFocus
        MS_BILL.Col = 1
End If
End Sub

Public Sub skulist()
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(Text3) & "%'", CON, adOpenKeyset
If SSTab1.TabEnabled(4) = True Then
    UOM_LST.ListItems.Clear
Else
    SKU_LST.ListItems.Clear
End If
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        If SSTab1.TabEnabled(4) = True Then
            Set ls = UOM_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        Else
            Set ls = SKU_LST.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        End If
        ls.ListSubItems.Add , , convert(AMT(Val(REC_SET("PRD_MRP") & "")))
        ls.ListSubItems.Add , , QTY(Val(REC_SET("PRD_QTY") & ""))
        ls.ListSubItems.Add , , REC_SET("PRD_UOM") & ""
        ls.ListSubItems.Add , , Trim(REC_SET("bat_idy") & "")
        ls.ListSubItems.Add , , FAC(REC_SET("con_fac") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("PRD_idy") & "")
        REC_SET.MoveNext
    Next
Else
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
    If SSTab1.TabEnabled(4) = True Then
    Text5.Visible = False
    Text5 = ""
    skulist
        If BAS_MS.Enabled = True Then
            BAS_MS.SetFocus
            BAS_MS.Col = 1
        Else
            FRE_MS.SetFocus
            FRE_MS.Col = 1
        End If
    Else
         Text2.Visible = False
         Text2 = ""
         skulist
        If MS_BAS.Enabled = True Then
            MS_BAS.SetFocus
            MS_BAS.Col = 1
        Else
            MS_FRE.Enabled = True
            MS_FRE.SetFocus
            MS_FRE.Col = 1
        End If
    End If
End If
End Sub

Public Sub DISPLAYALL()
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_QTY > 0 AND prd.PRD_NME LIKE '" & Trim(Text3) & "%'", CON, adOpenKeyset
If SSTab1.TabEnabled(3) = True Then
    For A = 1 To REC_SET.RecordCount
        If MS_BAS.Enabled = True And BAS_CHK.Value = 1 Then
            MS_BAS.TextMatrix(A, 0) = Trim(REC_SET("PRD_IDY"))
            MS_BAS.TextMatrix(A, 1) = Trim(REC_SET("PRD_NME"))
            MS_BAS.TextMatrix(A, 2) = convert(AMT(Val(REC_SET("PRD_MRP"))))
            MS_BAS.TextMatrix(A, 4) = Trim(REC_SET("BAT_IDY"))
            REC_SET.MoveNext
            MS_BAS.AddItem ""
        ElseIf FRE_CHK.Value = 1 Then
            MS_FRE.TextMatrix(A, 0) = Trim(REC_SET("PRD_IDY"))
            MS_FRE.TextMatrix(A, 1) = Trim(REC_SET("PRD_NME"))
            MS_FRE.TextMatrix(A, 2) = convert(AMT(Val(REC_SET("PRD_MRP"))))
            MS_FRE.TextMatrix(A, 4) = Trim(REC_SET("BAT_IDY"))
            REC_SET.MoveNext
            FRE_CHK.Enabled = True
            MS_FRE.AddItem ""
        End If
    Next
    If BAS_CHK.Value = 1 Then
        MS_BAS.Col = 3
        MS_BAS.Enabled = True
        MS_BAS.SetFocus
    Else
        MS_FRE.Col = 3
        MS_FRE.Enabled = True
        MS_FRE.SetFocus
    End If
Else
    For A = 1 To REC_SET.RecordCount
        If BAS_MS.Enabled = True And ALL_BAS.Value = 1 Then
            BAS_MS.TextMatrix(A, 0) = Trim(REC_SET("PRD_IDY"))
            BAS_MS.TextMatrix(A, 1) = Trim(REC_SET("PRD_NME"))
            BAS_MS.TextMatrix(A, 2) = convert(AMT(Val(REC_SET("PRD_MRP"))))
            BAS_MS.TextMatrix(A, 4) = Trim(REC_SET("BAT_IDY"))
            REC_SET.MoveNext
            BAS_MS.AddItem ""
        ElseIf ALL_FRE.Value = 1 Then
            FRE_MS.TextMatrix(A, 0) = Trim(REC_SET("PRD_IDY"))
            FRE_MS.TextMatrix(A, 1) = Trim(REC_SET("PRD_NME"))
            FRE_MS.TextMatrix(A, 2) = convert(AMT(Val(REC_SET("PRD_MRP"))))
            FRE_MS.TextMatrix(A, 4) = Trim(REC_SET("BAT_IDY"))
            REC_SET.MoveNext
            ALL_FRE.Enabled = True
            FRE_MS.AddItem ""
        End If
    Next
    If ALL_BAS.Value = 1 Then
        BAS_MS.Col = 3
        BAS_MS.Enabled = True
        BAS_MS.SetFocus
    Else
        FRE_MS.Col = 3
        FRE_MS.Enabled = True
        FRE_MS.SetFocus
    End If
End If
End Sub

Private Sub UOM_AMT_Change()
If Len(Trim(UOM_AMT)) = 7 Then SendKeys "{home}+{end}"
End Sub

Private Sub UOM_AMT_KeyPress(K As Integer)
If SUB_CMB.ListIndex = 0 And K = 46 Then K = 0
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Then

ElseIf K = 13 And BASPRD = "BASE" And Trim(BAS_MS.TextMatrix(BAS_MS.Row, 0)) <> "" Then
    BASPRD = ""
    If BAS_MS.Col = 3 Then
    If BAS_OR_OPT.Value = False Then BAS_MS.TextMatrix(BAS_MS.Row, 3) = Val(UOM_AMT)
    If BAS_OR_OPT.Value = True Then BAS_MS.TextMatrix(1, 3) = Val(UOM_AMT)
    UOM_AMT.Visible = False
        BAS_MS.Enabled = True
        If BAS_MS.Row = BAS_MS.Rows - 1 Then
        Else
            BAS_MS.Row = BAS_MS.Row + 1
        End If
        If Trim(BAS_MS.TextMatrix(BAS_MS.Row, 0)) <> "" Then
            BAS_MS.Col = 3
        Else
            BAS_MS.Col = 1
        End If
        BAS_MS.SetFocus
    End If
'ElseIf SUB_CMB.ListIndex = 0 Then
'    If FRE_MS.Col = 3 Or FRE_MS.Col = 7 Then
'    If FRE_MS.Col = 3 Then FRE_MS.TextMatrix(FRE_MS.Row, 3) = Val(UOM_AMT)
'    If FRE_MS.Col = 7 Then FRE_MS.TextMatrix(1, 7) = Val(UOM_AMT)
'   UOM_AMT.Visible = False
'    If Val(FRE_MS.TextMatrix(FRE_MS.Row, 3)) > 0 Then
'        FRE_MS.Enabled = True
'        If FRE_MS.Row = FRE_MS.Rows - 1 Then
'            FRE_MS.AddItem ""
'            ALL_FRE.Enabled = True
'            FRE_MS.Row = FRE_MS.Row + 1
'        Else
'            FRE_MS.Row = FRE_MS.Row + 1
'        End If
'        If Trim(FRE_MS.TextMatrix(FRE_MS.Row, 0)) <> "" Then
'            FRE_MS.Col = 1
'        Else
'            FRE_MS.Col = 1
'        End If
'        FRE_MS.SetFocus
'    ElseIf Val(FRE_MS.TextMatrix(FRE_MS.Row, 7)) > 0 Then
'        FRE_MS.Col = 7
'        FRE_MS.SetFocus
'    Else
'        If FRE_MS.Col = 3 Then
'            MsgBox "Enter Product Quantity", vbExclamation, "Effmcg"
'            FRE_MS.Col = 3
'            FRE_MS.SetFocus
'            Exit Sub
'        Else
'            BAS_MS.Enabled = True
'            BAS_MS.SetFocus
'        End If
'    End If
'
'    End If
ElseIf K = 13 And BASPRD = "FREE" And (SUB_CMB.ListIndex = 1 Or SUB_CMB.ListIndex = 2 Or SUB_CMB.ListIndex = 0) Then
    BASPRD = ""
    If FRE_MS.Col = 3 Or FRE_MS.Col = 5 Or FRE_MS.Col = 6 Or FRE_MS.Col = 7 Then
        If FRE_MS.Col = 3 And FRE_OR_OPT = False Then FRE_MS.TextMatrix(FRE_MS.Row, 3) = Val(UOM_AMT)
        If FRE_MS.Col = 3 And FRE_OR_OPT = True Then FRE_MS.TextMatrix(1, 3) = Val(UOM_AMT)
        If FRE_MS.Col = 7 And FRE_OR_OPT = False Then FRE_MS.TextMatrix(FRE_MS.Row, 7) = Val(UOM_AMT)
        If FRE_MS.Col = 7 And FRE_OR_OPT = True Then FRE_MS.TextMatrix(1, 7) = Val(UOM_AMT)
        If FRE_MS.Col = 5 Then FRE_MS.TextMatrix(FRE_MS.Row, 5) = convert(AMT(Val(UOM_AMT)))
        If FRE_MS.Col = 6 Then FRE_MS.TextMatrix(FRE_MS.Row, 6) = discount(UOM_AMT)
        UOM_AMT.Visible = False
        If SUB_CMB.ListIndex = 0 Then
            If Val(FRE_MS.TextMatrix(FRE_MS.Row, 3)) > 0 Then
                FRE_MS.Enabled = True
                ALL_FRE.Enabled = True
                If FRE_MS.Rows = 1 Then FRE_MS.AddItem ""
                If FRE_MS.Row = FRE_MS.Rows - 1 Then
                    FRE_MS.AddItem ""
                Else
                    FRE_MS.Row = FRE_MS.Row + 1
                End If
                FRE_MS.Col = 1
                FRE_MS.SetFocus
            End If
        ElseIf convert(AMT(Val(FRE_MS.TextMatrix(FRE_MS.Row, 5)))) > 0 Or discount(FRE_MS.TextMatrix(FRE_MS.Row, 6)) Then
            If FRE_MS.Row = FRE_MS.Rows - 1 Then
            Else
                FRE_MS.Row = FRE_MS.Row + 1
            End If
            FRE_MS.Col = 1
            FRE_MS.Enabled = True
            FRE_MS.SetFocus
        Else
            FRE_MS.Col = 1
            FRE_MS.Row = FRE_MS.Rows - 1
        End If
    End If
ElseIf K = 46 Then
Else
'    If BAS_MS.Enabled = True Then
'        BAS_MS.Col = 1
'        BAS_MS.SetFocus
'    Else
'        FRE_MS.Col = 1
'        FRE_MS.SetFocus
'    End If
End If
End Sub

Private Sub UOM_LST_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If UOM_LST.SortOrder = lvwAscending Then
UOM_LST.SortOrder = lvwDescending
ElseIf UOM_LST.SortOrder = lvwDescending Then
UOM_LST.SortOrder = lvwAscending
End If
UOM_LST.SortKey = CH.Index - 1
End Sub

Private Sub UOM_LST_KeyPress(KeyAscii As Integer)
Dim REC_SET As New ADODB.Recordset
If KeyAscii = 13 Then
    If UOM_LST.ListItems.Count = 0 Then
        Exit Sub
    End If
    If UOM_LST.ListItems.Count > 0 Then
        If BASPRD = "BASE" Then
            For A = 1 To BAS_MS.Rows - 1
                If Trim(UOM_LST.SelectedItem.ListSubItems.Item(4)) = Trim(BAS_MS.TextMatrix(A, 4)) Then
                    SSTab1.Tab = 4
                    MsgBox "Product Already Exists!", vbExclamation, "Effmcg"
                    Text5.Visible = False
                    If BAS_MS.Row = BAS_MS.Rows - 1 Then
                    Else
                        BAS_MS.Row = BAS_MS.Row + 1
                    End If
                    BAS_MS.SetFocus
                    BAS_MS.Col = 1
                    Exit Sub
                End If
            Next
            BAS_MS.AddItem ""
            BAS_MS.TextMatrix(BAS_MS.Row, 1) = UOM_LST.SelectedItem
            BAS_MS.TextMatrix(BAS_MS.Row, 2) = UOM_LST.SelectedItem.ListSubItems.Item(1)
            BAS_MS.TextMatrix(BAS_MS.Row, 4) = UOM_LST.SelectedItem.ListSubItems.Item(4)
            BAS_MS.TextMatrix(BAS_MS.Row, 0) = UOM_LST.SelectedItem.ListSubItems.Item(6)
            BAS_MS.Col = 3
            Text5.Visible = False
            PRD_LST_V
            BAS_MS.SetFocus
        Else
            If SUB_CMB.ListIndex = 0 Then
            For A = 1 To FRE_MS.Rows - 1
                If Trim(UOM_LST.SelectedItem.ListSubItems.Item(4)) = Trim(FRE_MS.TextMatrix(A, 4)) Then
                    SSTab1.Tab = 4
                    MsgBox "Product Already Exists!", vbExclamation, "Effmcg"
                    Text5.Visible = False
                    If FRE_MS.Row = FRE_MS.Rows - 1 Then
                    Else
                        FRE_MS.Row = FRE_MS.Row + 1
                    End If
                    FRE_MS.SetFocus
                    FRE_MS.Col = 1
                    Exit Sub
                End If
            Next

                FRE_MS.TextMatrix(FRE_MS.Row, 1) = UOM_LST.SelectedItem
                FRE_MS.TextMatrix(FRE_MS.Row, 2) = UOM_LST.SelectedItem.ListSubItems.Item(1)
                FRE_MS.TextMatrix(FRE_MS.Row, 4) = UOM_LST.SelectedItem.ListSubItems.Item(4)
                FRE_MS.TextMatrix(FRE_MS.Row, 0) = UOM_LST.SelectedItem.ListSubItems.Item(6)
            End If
            Text5.Visible = False
            PRD_LST_V
            If FRE_AND_OPT.Value = True Then
                FRE_MS.Col = 3
            Else
                ALL_FRE.Enabled = True
                FRE_MS.AddItem ""
                FRE_MS.Col = 1
                If FRE_MS.Row = FRE_MS.Rows - 1 Then
                Else
                    FRE_MS.Row = FRE_MS.Row + 1
                End If
            End If
            skulist
            FRE_MS.SetFocus
        End If
    End If
End If
End Sub

Public Sub dispsch()
If CHECKTB("sch.dbf,prd.dbf,brd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim REC_PRD As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
    If SCH_LST.ListItems.Count > 0 Then
        Dim lst As ListSubItem
        Set lst = SCH_LST.SelectedItem.ListSubItems.Item(1)
        If REC_SET.State = 1 Then REC_SET.Close
        If Len(Trim(SCH_IDY)) <> 10 Then
            REC_SET.Open "select * from sch where sch_idy like 'SCH" & Tincrement(SCH_IDY) & "' AND SCH_STS LIKE 'Y'", CON, adOpenStatic
        Else
            REC_SET.Open "select * from sch where sch_idy like '" & Trim(SCH_IDY) & "' AND SCH_STS LIKE 'Y'", CON, adOpenStatic
        End If
      If REC_SET.RecordCount > 0 Then
        If RS.State = 1 Then RS.Close
        RS.Open "select * from sch where sch_IDY like '" & Trim(REC_SET("SCH_IDY")) & "' AND TXN_TYP IN ('BAS')", CON, adOpenStatic
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select * from sch where sch_IDY like '" & Trim(REC_SET("SCH_IDY")) & "' AND TXN_TYP IN ('FRE','CAS','PCG')", CON, adOpenStatic
        MOD_REC = True
        SCH_IDY = Trim(REC_SET("SCH_IDY"))
        SCH_NME = Trim(REC_SET("sch_nme"))
        SCH_CMB = Trim(REC_SET("sch_typ"))
        SUB_CMB = Trim(REC_SET("F_SCH_TYP"))
        Dtpicker1 = REC_SET("sch_bdt")
        Dtpicker1.Enabled = False
        DTPicker2 = REC_SET("sch_edt")
        SCH_ACT = REC_SET("SCH_STS")
        SSTab1.Tab = 0
        If MOD_REC = True Then
            If Trim(SCH_CMB) = "SKU Level" Then
                SSTab1.TabEnabled(1) = True
                SSTab1.TabEnabled(2) = False
                SSTab1.TabEnabled(3) = False
                SSTab1.TabEnabled(4) = False
                SCH_IDY.Enabled = False
                PRD_LST.Enabled = False
                BAS_PRD_VIW.Enabled = False
                FRE_PRD_VIW.Enabled = False
                BAS_PRD_VIW.Rows = RS.RecordCount + 1
                FRE_PRD_VIW.Rows = RS1.RecordCount + 1
                For A = 1 To RS.RecordCount
                    If RS("txn_typ") = "BAS" Then
                        BAS_PRD_VIW.TextMatrix(A, 0) = RS("B_PRD_IDY") & ""
                        If REC_PRD.State = 1 Then REC_PRD.Close
                        REC_PRD.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.BAT_IDY LIKE '" & Trim(RS("BAT_IDY")) & "'", CON, adOpenDynamic, adLockOptimistic
                        BAS_PRD_VIW.TextMatrix(A, 1) = Trim(REC_PRD("PRD_NME") & "")
                        BAS_PRD_VIW.TextMatrix(A, 2) = convert(AMT(Val(REC_PRD("PRD_mrp") & "")))
                        BAS_PRD_VIW.TextMatrix(A, 3) = Val(RS("B_PRD_QTY") & "")
                        BAS_PRD_VIW.TextMatrix(A, 4) = RS("BAT_IDY")
                        RS.MoveNext
                    End If
                Next
                For b = 1 To RS1.RecordCount
                    If RS1("txn_typ") = "FRE" Or RS1("TXN_TYP") = "CAS" Or RS1("TXN_TYP") = "PCG" Then
                        If RS1("TXN_TYP") = "CAS" Then
                            FRE_PRD_VIW.ColWidth(1) = 0
                            FRE_PRD_VIW.ColWidth(2) = 0
                            FRE_PRD_VIW.ColWidth(3) = 0
                            FRE_PRD_VIW.ColWidth(5) = 900
                            FRE_PRD_VIW.ColWidth(6) = 0
                            LBL_QTY.Visible = False
                            LBL_PCG.Visible = False
                            LBL_AMT.Visible = True
                            If RS1("pro_rat") = "Y" Then
                            PRO_CHK.Value = 1
                            Else
                            PRO_CHK.Value = 0
                            End If
                            FRE_PRD_VIW.TextMatrix(b, 5) = convert(AMT(Val(RS1("B_DIS_AMT"))))
                            FRE_PRD_VIW.Enabled = False
                        ElseIf RS1("TXN_TYP") = "PCG" Then
                            FRE_PRD_VIW.ColWidth(1) = 0
                            FRE_PRD_VIW.ColWidth(2) = 0
                            FRE_PRD_VIW.ColWidth(3) = 0
                            FRE_PRD_VIW.ColWidth(5) = 0
                            FRE_PRD_VIW.ColWidth(6) = 1000
                            LBL_QTY.Visible = False
                            LBL_AMT.Visible = False
                            LBL_PCG.Visible = True
                            If RS1("pro_rat") = "Y" Then
                            PRO_CHK.Value = 1
                            Else
                            PRO_CHK.Value = 0
                            End If
                            FRE_PRD_VIW.TextMatrix(b, 6) = discount(RS1("B_DIS_PCG"))
                            FRE_PRD_VIW.Enabled = False
                        Else
                        FRE_PRD_VIW.TextMatrix(b, 0) = RS1("B_PRD_IDY") & ""
                        If REC_PRD.State = 1 Then REC_PRD.Close
                        REC_PRD.Open "SELECT * FROM PRD,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.BAT_IDY LIKE '" & Trim(RS1("BAT_IDY")) & "'", CON, adOpenDynamic, adLockOptimistic
                        FRE_PRD_VIW.TextMatrix(b, 1) = Trim(REC_PRD("PRD_NME") & "")
                        FRE_PRD_VIW.TextMatrix(b, 2) = convert(AMT(Val(REC_PRD("PRD_mrp") & "")))
                        FRE_PRD_VIW.TextMatrix(b, 3) = Val(RS1("B_PRD_QTY") & "")
                        FRE_PRD_VIW.TextMatrix(b, 4) = RS1("BAT_IDY")
                        If RS1("pro_rat") = "Y" Then
                            PRO_CHK.Value = 1
                        Else
                            PRO_CHK.Value = 0
                        End If
                            
                        End If
                        RS1.MoveNext
                    End If
                    Next
                    V_PRD_LST
             End If
             If Trim(SCH_CMB) = "Bill Level" Then
                SSTab1.TabEnabled(1) = False
                SSTab1.TabEnabled(2) = True
                SSTab1.TabEnabled(3) = False
                SSTab1.TabEnabled(4) = False
                PRD_VIW.Enabled = False
                GRS_AMT.Enabled = False
                SCH_IDY.Enabled = False
                MS_BILL.Enabled = False
                MS_BILL.Rows = RS1.RecordCount + 1
                For b = 1 To RS1.RecordCount
                If RS1("txn_typ") = "FRE" Or RS1("TXN_TYP") = "CAS" Or RS1("TXN_TYP") = "PCG" Then
                    If RS1("TXN_TYP") = "CAS" Then
                        MS_BILL.ColWidth(1) = 0
                        MS_BILL.ColWidth(2) = 0
                        MS_BILL.ColWidth(3) = 0
                        MS_BILL.ColWidth(5) = 900
                        MS_BILL.ColWidth(6) = 0
                        BILL_LBL.Caption = "Discount Amount"
                        GRS_AMT = convert(AMT(Val(RS1("b_bil_amt"))))
                        If RS1("pro_rat") = "Y" Then
                        PRO_CHK.Value = 1
                        Else
                        PRO_CHK.Value = 0
                        End If
                        MS_BILL.TextMatrix(b, 5) = convert(AMT(Val(RS1("B_DIS_AMT"))))
                        MS_BILL.Enabled = False
                    ElseIf RS1("TXN_TYP") = "PCG" Then
                        MS_BILL.ColWidth(1) = 0
                        MS_BILL.ColWidth(2) = 0
                        MS_BILL.ColWidth(3) = 0
                        MS_BILL.ColWidth(5) = 0
                        MS_BILL.ColWidth(6) = 1000
                        BILL_LBL.Caption = "Discount Percentage"
                        GRS_AMT = convert(AMT(Val(RS1("b_bil_amt"))))
                        If RS1("pro_rat") = "Y" Then
                        PRO_CHK.Value = 1
                        Else
                        PRO_CHK.Value = 0
                        End If
                        MS_BILL.TextMatrix(b, 6) = discount(RS1("B_DIS_PCG"))
                        MS_BILL.Enabled = False
                    Else
                    GRS_AMT = convert(AMT(Val(RS1("b_bil_amt"))))
                    MS_BILL.TextMatrix(b, 0) = RS1("B_PRD_IDY") & ""
                    If REC_PRD.State = 1 Then REC_PRD.Close
                    REC_PRD.Open "SELECT * FROM PRD,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.BAT_IDY LIKE '" & Trim(RS1("BAT_IDY")) & "'", CON, adOpenDynamic, adLockOptimistic
                    MS_BILL.TextMatrix(b, 1) = Trim(REC_PRD("PRD_NME") & "")
                    MS_BILL.TextMatrix(b, 2) = convert(AMT(Val(REC_PRD("PRD_mrp") & "")))
                    MS_BILL.TextMatrix(b, 3) = Val(RS1("B_PRD_QTY") & "")
                    MS_BILL.TextMatrix(b, 4) = RS1("BAT_IDY")
                    PRO_CHK.Visible = False
                    End If
                    RS1.MoveNext
                End If
                Next
            PRODUCTLIST
            End If
            If Trim(SCH_CMB) = "Multiple SKU Level" Then
                SSTab1.TabEnabled(1) = False
                SSTab1.TabEnabled(2) = False
                SSTab1.TabEnabled(3) = True
                SSTab1.TabEnabled(4) = False
                SCH_IDY.Enabled = False
                SKU_LST.Enabled = False
                BAS_CHK.Enabled = False
                FRE_CHK.Enabled = False
                Frame6.Enabled = False
                Frame7.Enabled = False
                MS_BAS.Enabled = False
                MS_FRE.Enabled = False
                MS_BAS.Rows = RS.RecordCount + 1
                MS_FRE.Rows = RS1.RecordCount + 1
                chk = False
                For A = 1 To RS.RecordCount
                    If RS("txn_typ") = "BAS" Then
                        If chk = False Then
                        If Trim(RS("b_con_app")) = "AND" Then BAS_AND.Value = True
                        If Trim(RS("b_con_app")) = "ORO" Then BAS_OR.Value = True
                        
                        If Trim(RS("b_con_app")) = "ORM" Then
                        BAS_OR.Value = True
                        Check1.Value = 1
                        Check1.Enabled = False
                        End If
                        chk = True
                        End If
                        MS_BAS.TextMatrix(A, 0) = RS("B_PRD_IDY") & ""
                        If REC_PRD.State = 1 Then REC_PRD.Close
                        REC_PRD.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.BAT_IDY LIKE '" & Trim(RS("BAT_IDY")) & "'", CON, adOpenDynamic, adLockOptimistic
                        MS_BAS.TextMatrix(A, 1) = Trim(REC_PRD("PRD_NME") & "")
                        MS_BAS.TextMatrix(A, 2) = convert(AMT(Val(REC_PRD("PRD_mrp") & "")))
                        MS_BAS.TextMatrix(A, 3) = Val(RS("B_PRD_QTY") & "")
                        MS_BAS.TextMatrix(A, 4) = RS("BAT_IDY")
                        RS.MoveNext
                    End If
                Next
                For b = 1 To RS1.RecordCount
                    If RS1("txn_typ") = "FRE" Or RS1("TXN_TYP") = "CAS" Or RS1("TXN_TYP") = "PCG" Then
                        If RS1("TXN_TYP") = "CAS" Then
                            MS_FRE.ColWidth(1) = 0
                            MS_FRE.ColWidth(2) = 0
                            MS_FRE.ColWidth(3) = 0
                            MS_FRE.ColWidth(5) = 900
                            MS_FRE.ColWidth(6) = 0
                            MS_FRE.ColWidth(7) = 0
                            Frame7.Visible = False
                            If RS1("pro_rat") = "Y" Then
                            PRO_CHK.Value = 1
                            Else
                            PRO_CHK.Value = 0
                            End If
                            MS_FRE.TextMatrix(b, 5) = convert(AMT(Val(RS1("B_DIS_AMT"))))
                            MS_FRE.Enabled = False
                        ElseIf RS1("TXN_TYP") = "PCG" Then
                            MS_FRE.ColWidth(1) = 0
                            MS_FRE.ColWidth(2) = 0
                            MS_FRE.ColWidth(3) = 0
                            MS_FRE.ColWidth(5) = 0
                            MS_FRE.ColWidth(6) = 1000
                            MS_FRE.ColWidth(7) = 0
                            Frame7.Visible = False
                            If RS1("pro_rat") = "Y" Then
                            PRO_CHK.Value = 1
                            Else
                            PRO_CHK.Value = 0
                            End If
                            MS_FRE.TextMatrix(b, 6) = discount(RS1("B_DIS_PCG"))
                            MS_FRE.Enabled = False
                        Else
                        MS_FRE.TextMatrix(b, 0) = RS1("B_PRD_IDY") & ""
                        If REC_PRD.State = 1 Then REC_PRD.Close
                        REC_PRD.Open "SELECT * FROM PRD,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.bat_IDY LIKE '" & Trim(RS1("BAT_IDY")) & "'", CON, adOpenDynamic, adLockOptimistic
                        MS_FRE.TextMatrix(b, 1) = Trim(REC_PRD("PRD_NME") & "")
                        MS_FRE.TextMatrix(b, 2) = convert(AMT(Val(REC_PRD("PRD_mrp") & "")))
                        MS_FRE.TextMatrix(b, 4) = RS1("BAT_IDY")
                        PRO_CHK.Visible = False
                        FRE_CHK.Visible = True
                        Frame6.Visible = True
                        If Trim(RS1("f_con_app")) = "ORO" Then
                            MS_FRE.ColWidth(1) = 2000
                            MS_FRE.ColWidth(2) = 1000
                            MS_FRE.ColWidth(5) = 0
                            MS_FRE.ColWidth(6) = 0
                            MS_FRE.ColWidth(3) = 0
                            MS_FRE.ColWidth(7) = 900
                            FRE_OR.Value = True
                            MS_FRE.TextMatrix(1, 7) = Val(RS1("B_PRD_QTY") & "")
                        Else
                            MS_FRE.ColWidth(1) = 2000
                            MS_FRE.ColWidth(2) = 1000
                            MS_FRE.ColWidth(5) = 0
                            MS_FRE.ColWidth(6) = 0
                            MS_FRE.ColWidth(3) = 900
                            MS_FRE.ColWidth(7) = 0
                            FRE_AND.Value = True
                            MS_FRE.TextMatrix(b, 3) = Val(RS1("B_PRD_QTY") & "")
                        End If
                        End If
                        RS1.MoveNext
                    End If
                    Next
                    PRODUCTLIST
            End If
            If Trim(SCH_CMB) = "UOM Product Group" Then
                SSTab1.TabEnabled(1) = False
                SSTab1.TabEnabled(2) = False
                SSTab1.TabEnabled(3) = False
                SSTab1.TabEnabled(4) = True
                SCH_IDY.Enabled = False
                UOM_LST.Enabled = False
                ALL_BAS.Enabled = False
                ALL_FRE.Enabled = False
                Frame8.Enabled = False
                Frame9.Enabled = False
                BAS_MS.Enabled = False
                FRE_MS.Enabled = False
                BAS_MS.Rows = RS.RecordCount + 1
                FRE_MS.Rows = RS1.RecordCount + 1
                For A = 1 To RS.RecordCount
                    If RS("txn_typ") = "BAS" Then
                        BAS_MS.TextMatrix(A, 0) = RS("B_PRD_IDY") & ""
                        If REC_PRD.State = 1 Then REC_PRD.Close
                        REC_PRD.Open "SELECT * FROM BRD,PRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_IDY LIKE '" & Trim(RS("BAT_IDY")) & "'", CON, adOpenDynamic, adLockOptimistic
                        BAS_MS.TextMatrix(A, 1) = Trim(REC_PRD("PRD_NME") & "")
                        BAS_MS.TextMatrix(A, 2) = convert(AMT(Val(REC_PRD("PRD_mrp") & "")))
                        BAS_MS.TextMatrix(A, 3) = Val(RS("B_PRD_QTY") & "")
                        BAS_MS.TextMatrix(A, 4) = RS("BAT_IDY")
                        If Trim(RS("b_con_app")) = "AND" Then BAS_AND_OPT.Value = True
                        If Trim(RS("b_con_app")) = "ORO" Then BAS_OR_OPT.Value = True
                        RS.MoveNext
                    End If
                Next
                For b = 1 To RS1.RecordCount
                    If RS1("txn_typ") = "FRE" Or RS1("TXN_TYP") = "CAS" Or RS1("TXN_TYP") = "PCG" Then
                        If RS1("TXN_TYP") = "CAS" Then
                            FRE_MS.ColWidth(1) = 0
                            FRE_MS.ColWidth(2) = 0
                            FRE_MS.ColWidth(3) = 0
                            FRE_MS.ColWidth(5) = 900
                            FRE_MS.ColWidth(6) = 0
                            FRE_MS.ColWidth(7) = 0
                            Frame8.Visible = False
                            If RS1("pro_rat") = "Y" Then
                            PRO_CHK.Value = 1
                            Else
                            PRO_CHK.Value = 0
                            End If
                            FRE_MS.TextMatrix(b, 5) = convert(AMT(Val(RS1("B_DIS_AMT"))))
                            FRE_MS.Enabled = False
                        ElseIf RS1("TXN_TYP") = "PCG" Then
                            FRE_MS.ColWidth(1) = 0
                            FRE_MS.ColWidth(2) = 0
                            FRE_MS.ColWidth(3) = 0
                            FRE_MS.ColWidth(5) = 0
                            FRE_MS.ColWidth(6) = 1000
                            FRE_MS.ColWidth(7) = 0
                            Frame8.Visible = False
                            If RS1("pro_rat") = "Y" Then
                            PRO_CHK.Value = 1
                            Else
                            PRO_CHK.Value = 0
                            End If
                            FRE_MS.TextMatrix(b, 6) = discount(RS1("B_DIS_PCG"))
                            FRE_MS.Enabled = False
                        Else
                        FRE_MS.TextMatrix(b, 0) = RS1("B_PRD_IDY") & ""
                        If REC_PRD.State = 1 Then REC_PRD.Close
                        REC_PRD.Open "SELECT * FROM PRD,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND BRD.PRD_IDY LIKE '" & Trim(RS1("BAT_IDY")) & "'", CON, adOpenDynamic, adLockOptimistic
                        FRE_MS.TextMatrix(b, 1) = Trim(REC_PRD("PRD_NME") & "")
                        FRE_MS.TextMatrix(b, 2) = convert(AMT(Val(REC_PRD("PRD_mrp") & "")))
                        FRE_MS.TextMatrix(b, 4) = RS1("BAT_IDY")
                        PRO_CHK.Visible = False
                        ALL_FRE.Visible = True
                        Frame8.Visible = True
                        If Trim(RS1("f_con_app")) = "ORO" Then
                            FRE_MS.ColWidth(1) = 2000
                            FRE_MS.ColWidth(2) = 1000
                            FRE_MS.ColWidth(5) = 0
                            FRE_MS.ColWidth(6) = 0
                            FRE_MS.ColWidth(3) = 0
                            FRE_MS.ColWidth(7) = 900
                            FRE_OR_OPT.Value = True
                            FRE_MS.TextMatrix(1, 7) = Val(RS1("B_PRD_QTY") & "")
                        Else
                            FRE_MS.ColWidth(1) = 2000
                            FRE_MS.ColWidth(2) = 1000
                            FRE_MS.ColWidth(5) = 0
                            FRE_MS.ColWidth(6) = 0
                            FRE_MS.ColWidth(3) = 900
                            FRE_MS.ColWidth(7) = 0
                            FRE_AND_OPT.Value = True
                            FRE_MS.TextMatrix(b, 3) = Val(RS1("B_PRD_QTY") & "")
                        End If
                        End If
                        RS1.MoveNext
                    End If
                    Next
                    PRODUCTLIST
            End If
        Else
            SCH_IDY.Enabled = False
            SCH_NME.SetFocus
            
        End If
            Schemelist
            SCH_ACT.SetFocus
            Exit Sub
      Else

          Schemelist
      End If
        If MOD_REC = True Then
            SCH_ACT.SetFocus
            SCH_NME.Locked = True
            SCH_CMB.Locked = True
            SUB_CMB.Locked = True
            Dtpicker1.Enabled = False
        End If
    Else
        MsgBox "No Schemes found!", vbExclamation
        SendKeys "{home}+{end}"
        SCH_IDY.SetFocus
    End If
    Toolbar1.Buttons(12).Enabled = True
End Sub
