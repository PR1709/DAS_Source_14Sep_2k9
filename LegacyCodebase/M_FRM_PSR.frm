VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_PCR 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Push Cart Return Stock Voucher"
   ClientHeight    =   5625
   ClientLeft      =   1425
   ClientTop       =   1845
   ClientWidth     =   8820
   ForeColor       =   &H00404040&
   Icon            =   "M_FRM_PSR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483637
      ForeColor       =   8388608
      TabCaption(0)   =   "&Invoice Details"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "M_SMN_NME"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label40"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "M_GRS_AMT"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "M_INV_DAT"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "M_INV_IDY"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "m_van_idy"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&Products"
      TabPicture(1)   =   "M_FRM_PSR.frx":0442
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Total"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "units"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "tax"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MS"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "List6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "a"
      TabPicture(2)   =   "M_FRM_PSR.frx":045E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "View"
      TabPicture(3)   =   "M_FRM_PSR.frx":047A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label49"
      Tab(3).Control(1)=   "list1"
      Tab(3).Control(2)=   "list5"
      Tab(3).Control(3)=   "list4"
      Tab(3).Control(4)=   "Frame1"
      Tab(3).ControlCount=   5
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
         ItemData        =   "M_FRM_PSR.frx":0496
         Left            =   -72000
         List            =   "M_FRM_PSR.frx":04A3
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72960
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74880
         TabIndex        =   3
         Top             =   4500
         Width           =   8295
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.TextBox m_van_idy 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   6
         Top             =   780
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ProductName"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Product ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Short Name"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sub Units"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Instituational Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Billing Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "M.R.P Rate"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   2640
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24444931
         CurrentDate     =   37099
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4395
         Left            =   -75000
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView list5 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   9
         Top             =   780
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Invoice ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Invoice Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Van Challan Number"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   10
         Top             =   780
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
            Text            =   "Salsesman Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Salesman ID"
            Object.Width           =   3528
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
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   195
         Left            =   1890
         TabIndex        =   22
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   2085
         TabIndex        =   21
         Top             =   1388
         Width           =   345
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Push Cart Challan ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1005
         TabIndex        =   20
         Top             =   2115
         Width           =   1470
      End
      Begin VB.Label tax 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -71550
         TabIndex        =   19
         Top             =   4380
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label units 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -73515
         TabIndex        =   18
         Top             =   4485
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label M_GRS_AMT 
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
         Left            =   2760
         TabIndex        =   17
         Top             =   2715
         Width           =   45
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
         Left            =   -74760
         TabIndex        =   16
         Top             =   420
         Width           =   8295
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   5040
         TabIndex        =   15
         Top             =   4350
         Width           =   60
      End
      Begin VB.Label M_SMN_NME 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4080
         TabIndex        =   14
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -66525
         TabIndex        =   13
         Top             =   4785
         Width           =   45
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Challan Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   668
         Width           =   1125
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
            Picture         =   "M_FRM_PSR.frx":04B0
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":09F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":0DBC
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":1110
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":1748
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":1A9C
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":1EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":21F0
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":25FC
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":2708
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":2A5C
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":2E70
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PSR.frx":33BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
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
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
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
            Object.ToolTipText     =   "Prefix Salesmen ID "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
Attribute VB_Name = "M_FRM_PCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, bslab, REC_SET1 As New ADODB.Recordset, RS4 As New ADODB.Recordset, REC_SET5 As New ADODB.Recordset, MSROW, INVIDY, INC, dlr_idy

Private Sub CREATE()
MOD_REC = False
M_INV_IDY = ""
saverecord
End Sub

Public Sub GRS_CAL()
M_GRS_AMT = convert(amt(Val(Total)))
M_NET_AMT = convert(amt(Round(Val(M_GRS_AMT) - Val(billdisamt), 2)))
M_NET_AMT = convert(amt(Round(Val(M_NET_AMT) - Val(addisamt), 2)))
M_NET_AMT = convert(amt(Round(Val(M_NET_AMT) + Val(taxamt), 2)))
M_NET_AMT = convert(amt(Val(M_NET_AMT) - Val(damageamt)))
End Sub

Private Sub saverecord()
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "Effmcg") = vbNo Then
    Cancelrecord
    Exit Sub
End If

If check("INl", "INV_idy", Trim(m_van_idy)) = False Then
    MsgBox "Van Challan Not Found !", vbExclamation, "Effmcg"
    m_van_idy.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If MOD_REC = True Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like 'PCR" & Tincrement(M_INV_IDY) & "' and txn_typ like 'PCR'", con, adOpenKeyset
    For A = 1 To REC_SET.RecordCount
        con.Execute "update prd where prd_idy like '" & REC_SET("prd_idy") & "' set cur_lev = " & CONQTY(Val(REC_SET("prd_qty")), REC_SET("prd_idy"))
        REC_SET.MoveNext
    Next
    con.Execute "delete from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'PCR'"
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'PCR'", con, adOpenKeyset, adLockOptimistic
    INC = Trim(M_INV_IDY)
Else
    REC_SET.Open "select * from inl", con, adOpenKeyset, adLockOptimistic
    REC_SET.AddNew
    REC_SET("inv_idy") = "PCR" & T7increment("idy_PCR")
    INC = REC_SET("inv_idy")
    INVIDY = INC
End If
REC_SET("inv_dat") = M_INV_DAT
REC_SET("act_idy") = Trim(credit & "")
REC_SET("ord_idy") = Trim(m_van_idy & "")
If REC_SET1.State = 1 Then REC_SET1.Close
REC_SET1.Open "select smn_idy,act_idy,cst_idy,NET_AMT from inl where inv_idy like '" & Trim(m_van_idy) & "' AND TXN_TYP LIKE 'PCS'", con, adOpenKeyset, adLockOptimistic
REC_SET("act_idy") = Trim(REC_SET1(1) & "")
REC_SET("cst_idy") = Trim(REC_SET1(2) & "")
REC_SET("smn_idy") = Trim(REC_SET1(0) & "")
REC_SET("txn_typ") = "PCR"
REC_SET("grs_amt") = Val(REC_SET1(3) & "")
REC_SET("net_amt") = Val(REC_SET1(3) & "")
REC_SET.Update
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inv", con, adOpenKeyset, adLockOptimistic
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
        REC_SET.AddNew
        REC_SET("inv_idy") = Trim(INC)
        REC_SET("inv_dat") = Trim(M_INV_DAT)
        REC_SET("prd_idy") = MS.TextMatrix(A, 0)
        REC_SET("lst_rat") = convert(amt(Val(MS.TextMatrix(A, 1))))
        REC_SET("lst_dis") = discount(MS.TextMatrix(A, 2))
        REC_SET("prd_qty") = QTY(Val(MS.TextMatrix(A, 3)))
        REC_SET("prd_rat") = convert(amt(Val(MS.TextMatrix(A, 4))))
        If Len(Trim(MS.TextMatrix(A, 0))) <> 0 Then
            REC_SET("rat_app") = MS.TextMatrix(A, 7) & ""
        Else
            REC_SET("rat_app") = "A"
        End If
        PRDNET = convert(amt(Val(MS.TextMatrix(A, 3)) * Val(MS.TextMatrix(A, 4))))
        ADDAMT = (PRDNET - TRDDIS) * Val(MS.TextMatrix(A, 7)) / 100
        TRDTAX = (PRDNET - (TRDDIS + ADDAMT)) * Val(MS.TextMatrix(A, 8)) / 100
        REC_SET("txn_typ") = "PCR"
        REC_SET("COL_AMT") = convert(amt(MS.TextMatrix(A, 9)))
        REC_SET.Update
        con.Execute "update prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "' set cur_lev = " & ADDQTY(Trim(MS.TextMatrix(A, 3)), Trim(MS.TextMatrix(A, 0)))
    End If
Next
If REC_SET1.State = 1 Then REC_SET1.Close
If MOD_REC = False Then MsgBox "New Van Challan ID is " & INC, vbExclamation, "Effmcg"
Label23 = labdisp("idy_PCR")
If check("PRT", "PRT_NME", "M_FRM_PCR") = True Then
    con.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
    con.Execute "update dir where dir_idy like 'REM' set dir_als = 'PCR'"
    Openreport App.Path & "\VOUCH.RP1", "Van Challan Ret", 2, "inv->inv_idy='" & Trim(INC) & "'.and.inv->txn_typ='PCR'", 2
    Cancelrecord
End If
Cancelrecord
End Sub

Private Sub Cancelrecord()
clrctr M_FRM_PCR
setval
Form_Load
MS.Rows = 1
MS.AddItem ""
SSTab1.Tab = 0
M_INV_IDY.Enabled = True
M_INV_IDY.SetFocus
MOD_REC = False
End Sub

Private Sub newrecord()
Cancelrecord
M_INV_IDY.Enabled = False
M_INV_DAT.SetFocus
SSTab1.Tab = 0
End Sub

Private Sub V_VAN_LST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
REC_SET.Open "select inl.inv_idy,inl.inv_dat,smn.smn_nme,smn.smn_ad4  from inl,smn where smn_nme like '" & Trim(m_van_idy) & "%' and inl.smn_idy=smn.smn_idy  and inl.txn_typ='PCS'  ", con, adOpenKeyset, adLockOptimistic
Else
REC_SET.Open "select inl.inv_idy,inl.inv_dat,smn.smn_nme,smn.smn_ad4  from inl,smn where smn_AD4 like '" & Trim(m_van_idy) & "%' and inl.smn_idy=smn.smn_idy  and inl.txn_typ='PCS'  ", con, adOpenKeyset, adLockOptimistic
End If
list5.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        If check("INL", "ORD_IDY", Trim(REC_SET(0)) & "' AND TXN_TYP LIKE 'PCR") = False Then
            Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
            ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
            ls.ListSubItems.Add , , Trim(REC_SET("smn_NME") & "")
            ls.ListSubItems.Add , , Trim(REC_SET("smn_AD4") & "")
        End If
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Van Challan"
    list5.ColumnHeaders.Item(1).Text = "Van Challan"
    list5.ColumnHeaders.Item(2).Text = "Date"
    list5.ColumnHeaders.Item(3).Text = "Salesman Name"
    list5.ColumnHeaders.Item(4).Text = "Place"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Van Challan Number Not Found...!", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        
        SendKeys "{HOME}+{END}"
        m_van_idy.SetFocus
    End If
End If
End Sub

Private Sub exitrecord()
Unload Me
End Sub

Private Sub List6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim REC_SET As New ADODB.Recordset
    List6.Visible = False
    MS.Text = List6
    List6.Visible = False
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from prd where prd_idy like '" & MS.TextMatrix(MS.Row, 0) & "'", con, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then
        If List6 = "A" Then
            MS.TextMatrix(MS.Row, 5) = amt(REC_SET("pur_rat") & "")
        ElseIf List6 = "B" Then
            MS.TextMatrix(MS.Row, 5) = amt(REC_SET("sal_rat") & "")
        ElseIf List6 = "C" Then
            MS.TextMatrix(MS.Row, 5) = amt(REC_SET("mrp_rat") & "")
        End If
    End If
    discalc
    MS.SetFocus
End If
End Sub

Private Sub M_INV_DAT_CloseUp()
M_INV_DAT.SetFocus
End Sub

Private Sub M_INV_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date Using Arrow Keys Or Press F4"
End Sub

Private Sub M_INV_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then m_van_idy.SetFocus
End Sub

Private Sub Form_Activate()
SSTab1.TabVisible(2) = False
checkdongle
SSTab1.Tab = 0
If M_INV_IDY.Enabled = True Then
    M_INV_IDY.SetFocus
Else
    M_INV_DAT.SetFocus
End If
Label23 = labdisp("idy_PCR")
SSTab1.TabEnabled(3) = False
setval
SETYEAR M_FRM_PCR
End Sub

Private Sub INVOICELIST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select inl.inv_idy,inl.inv_dat,smn.smn_nme,inl.ord_idy,smn.smn_nme from inl,smn where inl.txn_typ like 'PCR' AND INL.SMN_IDY LIKE SMN.SMN_IDY", con, adOpenKeyset, adLockOptimistic
Else
    REC_SET.Open "select inl.inv_idy,inl.inv_dat,SMN_NME,ord_idy,SMN_AD4 from SMN,inl where SMN_ad4 like '" & Trim(M_INV_IDY) & "%' and inl.smn_idy=smn.smn_idy and inl.txn_typ ='PCR'", con, adOpenKeyset, adLockOptimistic
End If
list5.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_nme") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ord_idy") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 3
    Label49 = "Select Van Return Challan"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Van Return Challan Number Not Found...!", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then Shell "c:\windows\calc.exe"
If list1.Visible = False And list4.Visible = False And list5.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then PRINTRECORD
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then exitrecord
    If KeyCode = 81 And Shift = 2 Then Cancelrecord
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
    
    
        If M_FRM_PCR.ActiveControl.name = "M_INV_IDY" Then
            INVOICELIST
        ElseIf M_FRM_PCR.ActiveControl.name = "m_van_idy" Then
            V_VAN_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_PCR.ActiveControl Is TextBox Then M_FRM_PCR.ActiveControl.Text = FUNKEY(KeyCode)
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
If list5.ListItems.Count > 0 Then
    LIST5_KeyPress 13
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        m_van_idy.SetFocus
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
            MS.SetFocus
        End If
    ElseIf SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        GRS_CAL
        discalc
        saverecord
    Else
        If M_FRM_PSR.ActiveControl.name = "M_INV_IDY" Then
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
MS.ColWidth(0) = 1100
MS.ColWidth(1) = 1200
MS.ColWidth(2) = 1200
MS.ColWidth(3) = 1050
MS.ColWidth(4) = 1000
MS.ColWidth(5) = 1200
MS.ColWidth(6) = 1500
MS.TextMatrix(0, 0) = "Product ID"
MS.TextMatrix(0, 1) = "Opening Stock"
MS.TextMatrix(0, 2) = "Return Stock"
MS.TextMatrix(0, 3) = "Billing Stock"
MS.TextMatrix(0, 4) = "Rate ID"
MS.TextMatrix(0, 5) = "Rate"
MS.TextMatrix(0, 6) = "Amount"
MS.Col = 0
MS.Col = 3
units.Left = (MS.CellWidth + MS.CellLeft) - units.Width
units.Top = 4850
MS.Col = 5
tax.Left = (MS.CellWidth + MS.CellLeft) - tax.Width
tax.Top = 4850
tax.Width = 0
MS.Col = 0
M_INV_DAT = Date
lrdate = Date
dcdate = Date
waydate = Date
Label17 = "0.00"
Label18 = "0"
Label19 = "0.00"
Label20 = "0.00"
Label23 = labdisp("idy_PCR")
lramount = "0.00"
deliverycharges = "0.00"
lrnumber = 0
frieghtamt = "0.00"
waybill = 0
M_NET_AMT = "0.00"
M_GRS_AMT = "0.00"
crdays = 0
waybill = 0
ad = "0.00"
cd = "0.00"
units = "0"
Total = "0.00"
tax = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Invoice Number"
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
If M_INV_IDY = "" And K = 27 Then Unload Me
End Sub

Private Sub LIST5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If list5.ListItems.Count > 0 Then
    If Label49 = "Select Van Challan" Then
        m_van_idy = list5.SelectedItem
        list5.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
        m_van_idy_keydown 13, 0
    Else
        M_INV_IDY = list5.SelectedItem
        dispdetails
    End If
End If
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
If Label49.Caption = "Select Van Voucher" Then
    V_VAN_LST
ElseIf Label49.Caption = "Select Invoice" Then
    INVOICELIST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If Label49.Caption = "Select Van Voucher" Then
    V_VAN_LST
ElseIf Label49.Caption = "Select Invoice" Then
    INVOICELIST
End If
'End If
End Sub


Private Sub MS_EnterCell()
If MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Type Product ID or Press <Enter> to get the List"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Type Units Required"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Rate of the Product"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Weightage"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Amount"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
End If
End Sub

Private Sub MS_KeyPress(K As Integer)
If MS.Col = 2 Then
    If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
        If K > 47 And K < 58 Or K = 46 Then
            text1 = ""
            text1.MaxLength = 7
            text1.Visible = True
            text1.Top = MS.Top + MS.CellTop
            text1.Left = MS.CellLeft + MS.Left
            text1.Width = MS.CellWidth
            text1.Height = MS.CellHeight
            text1.SetFocus
            If K <> 13 Then text1 = Chr(K)
            text1.SelStart = Len(text1)
        End If
    End If
ElseIf MS.Col = 4 Then
    If MS.TextMatrix(MS.Row, 0) <> "" Then
        If K <> 27 Then
            List6.Visible = True
            List6.Top = MS.Top + MS.CellTop
            List6.Left = MS.CellLeft + MS.Left
            List6.Width = MS.CellWidth
            List6.Height = MS.RowHeight(MS.Row)
            List6.Selected(0) = True
            List6 = Chr(K)
            List6.SetFocus
        End If
    End If
End If
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MS.Text = QTY(text1)
    text1.Visible = False
    MS.SetFocus
    discalc
End If
End Sub

Private Sub Text1_KeyPRESS(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 45 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
newrecord
ElseIf Button.KEY = "s" Then
saverecord
ElseIf Button.KEY = "e" Then
exitrecord
ElseIf Button.KEY = "p" Then
PRINTRECORD
ElseIf Button.KEY = "c" Then
Cancelrecord
ElseIf Button.KEY = "pre" Then
M_FRM_PRE.Caption = "Prefix Van Challan"
M_FRM_PRE.Show 1
ElseIf Button.KEY = "crt" Then
CREATE
ElseIf Button.KEY = "ed" Then
    If M_FRM_PSR.ActiveControl.name = "M_INV_IDY" Then
        Label49 = "Select Invoice"
        INVOICELIST
    ElseIf M_FRM_PSR.ActiveControl.name = "m_van_idy" Then
        Label49 = "Select Van Challan"
        V_VAN_LST
    End If
ElseIf Button.KEY = "DAY" Then
    If Button.Value = tbrPressed Then
        text1.Visible = True
        text1 = Format(Date, "dd/mm/yyyy") & " " & Format(Date, "dddd")
        text1.BackColor = vbBlue
        text1.ForeColor = vbWhite
    Else
        text1.Visible = False
    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
If check("prt", "prt_nme", "M_FRM_PSR") = False Then
CHECK3.Value = 0
Else
CHECK3.Value = 1
End If
CHECK3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Private Sub M_INV_IDY_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
If Len(Trim(M_INV_IDY)) = 0 Then
    INVOICELIST
Else
    inv = M_INV_IDY
    clrctr M_FRM_PSR
    Label7 = ""
    Label8 = ""
    Label9 = ""
    Label10 = ""
    Label11 = ""
    Label17 = "0.00"
    Label18 = "0"
    Label19 = "0.00"
    Label20 = "0.00"
    M_GRS_AMT = "0.00"
    M_NET_AMT = "0.00"
    M_SMN_NME = ""
    Label40 = ""
    M_INV_IDY = inv
    dispdetails
End If
End If
End Sub

Public Sub dispdetails()
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Len(Trim(M_INV_IDY)) <> 10 Then
REC_SET.Open "select * from inl where inv_idy like 'PCR" & Tincrement(Trim(M_INV_IDY)) & "' and txn_typ like 'PCR'", con, adOpenKeyset
    M_INV_IDY = "PCR" & Tincrement(Trim(M_INV_IDY))
Else
REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'PCR'", con, adOpenKeyset
    M_INV_IDY = Trim(M_INV_IDY)
End If
If REC_SET.RecordCount > 0 Then
    M_INV_IDY.Enabled = False
    MOD_REC = True
    M_INV_DAT = REC_SET("inv_dat")
    m_van_idy = Trim(REC_SET("ord_idy") & "")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from smn where smn_idy like '" & Trim(m_van_idy) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then M_SMN_NME = REC_SET("smn_nme")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ = 'PCR'", con, adOpenKeyset, adLockOptimistic
    MS.Rows = REC_SET.RecordCount + 1
    For A = 1 To MS.Rows - 1
        MS.TextMatrix(A, 0) = Trim(REC_SET("prd_idy"))
        MS.TextMatrix(A, 3) = QTY(Val(REC_SET("prd_qty")))
        MS.TextMatrix(A, 4) = convert(amt(REC_SET("prd_rat") & ""))
        MS.TextMatrix(A, 6) = Trim(REC_SET("rat_app") & "")
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT PRD_PRN,prd_nme,CON_FAC FROM PRD WHERE PRD_IDY LIKE '" & REC_SET("PRD_IDY") & "'", con, adOpenStatic
        If REC_SET2.RecordCount > 0 Then
            MS.TextMatrix(A, 10) = Trim(REC_SET2(0))
            MS.TextMatrix(A, 11) = Trim(REC_SET2(1))
            MS.TextMatrix(A, 2) = Val(REC_SET2("CON_FAC") & "")
        End If
        REC_SET.MoveNext
    Next
    discalc
    M_GRS_AMT = Total
    SSTab1.Tab = 0
    M_INV_DAT.SetFocus
Else
    INVOICELIST
End If
    If MS.Rows = 1 Then MS.Rows = 2
End Sub


Private Sub m_van_idy_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Challan"
End Sub
Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = ColumnHeader.Index - 1
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list1.ListItems.Count > 0 Then
        m_van_idy = list1.SelectedItem.ListSubItems.Item(1)
        M_SMN_NME = list1.SelectedItem
        SSTab1.Tab = 1
        MS.SetFocus
        
        list1.Visible = False
    End If
End If
End Sub

Private Sub List1_DblClick()
List1_KeyPress 13
End Sub
Private Sub m_van_idy_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K < 123 Then K = K - 32
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
    crdays.SetFocus
End If
If list1.Visible = True Then list1.Visible = False
If list4.Visible = True Then list4.Visible = False
If list5.Visible = True Then list5.Visible = False

If SSTab1.Tab <> 1 Then
    tax.Visible = False
    units.Visible = False
    Total.Visible = False
Else
    tax.Visible = True
    units.Visible = True
    Total.Visible = True
End If
If SSTab1.Tab = 3 Then
Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTab1.Tab = 1 Then MS.SetFocus
    If SSTab1.Tab = 2 Then crdays.SetFocus
    If SSTab1.Tab = 0 Then M_INV_DAT.SetFocus
End If
End Sub

Private Sub m_van_idy_keydown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    Dim REC_SET As New ADODB.Recordset
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select inv_idy from inl where inv_idy like '" & Trim(m_van_idy) & "' and txn_typ like 'PCS'", con, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then
        m_van_idy = Trim(m_van_idy)
        Loadproducts
        SSTab1.Tab = 1
        MS.SetFocus
        If K = 38 Then
            M_INV_DAT.SetFocus
        End If
    Else
        V_VAN_LST
    End If
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub
Public Sub calc()
M_GRS_AMT = 0
For A = 1 To MS.Rows - 1
M_GRS_AMT = Val(M_GRS_AMT) + Val(MS.TextMatrix(A, 5))
M_GRS_AMT = convert(amt(M_GRS_AMT & ""))
Next
M_NET_AMT = convert(amt((Val(M_GRS_AMT) + Val(deliverycharges)) & ""))
End Sub


Public Sub dispprd(prdid As String)
MSROW = MS.Row
Dim REC_SET As New ADODB.Recordset
For A = 1 To MS.Rows - 1
    If Trim(prdid) = MS.TextMatrix(A, 0) Then
        SSTab1.Tab = 1
        list4.Visible = False
        MsgBox "Product Already Exists In Invoice !", vbExclamation, "Effmcg"
        Text3.Visible = False
        MS.SetFocus
        MS.Col = 0
        MS.Row = A
        Exit Sub
    End If
Next
REC_SET.Open "select * from PRD where prd_idy like '" & Trim(prdid) & "'", con, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    list4.Visible = False
    MS.TextMatrix(MS.Row, 0) = REC_SET("PRD_IDY")
    'MS.TextMatrix(MS.Row, 1) = convert(amt(REC_SET("lst_rat") & ""))
    MS.TextMatrix(MS.Row, 2) = FAC(Val(REC_SET("CON_FAC") & ""))
    MS.TextMatrix(MS.Row, 3) = "0.000"
    If bslab = "A" Then
    MS.TextMatrix(MS.Row, 4) = convert(amt(REC_SET("SAL_rat") & ""))
    ElseIf bslab = "B" Then
    MS.TextMatrix(MS.Row, 4) = convert(amt(REC_SET("PUR_rat") & ""))
    ElseIf bslab = "C" Then
    MS.TextMatrix(MS.Row, 4) = convert(amt(REC_SET("mrp_rat") & ""))
    Else
    MS.TextMatrix(MS.Row, 4) = convert(amt(REC_SET("SAL_rat") & ""))
    End If
    'MS.TextMatrix(MS.Row, 6) = discount(REC_SET("ADd_dis") & "")
    'MS.TextMatrix(MS.Row, 7) = discount(REC_SET("CAS_dis") & "")
    'MS.TextMatrix(MS.Row, 8) = discount(REC_SET("sal_tax") & "")
    MS.TextMatrix(MS.Row, 10) = Trim(REC_SET("prd_prn"))
    MS.TextMatrix(MS.Row, 11) = Trim(REC_SET("prd_nme"))
    Text3.Visible = False
    calc
    SSTab1.Tab = 1
    MS.Row = MSROW
    MS.Col = 3
    MS.SetFocus
End If
End Sub

Public Sub discalc()
units = 0
tax = 0
Total = 0

If MOD_REC = True Then
    Dim REC_SET1 As New ADODB.Recordset
    If REC_SET1.State = 1 Then REC_SET1.Close
    REC_SET1.Open "select gRS_amt from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'PCR'", con, adOpenKeyset, adLockOptimistic
End If
For A = 1 To MS.Rows - 1
    Dim REC_SET As New ADODB.Recordset
    If REC_SET.State = 1 Then REC_SET.Close
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
    REC_SET.Open "select sub_unt from prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'", con, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then
    End If
    OPS_STK = P_SUB_UNT(MS.TextMatrix(A, 1), REC_SET(0))
    CLS_STK = P_SUB_UNT(MS.TextMatrix(A, 2), REC_SET(0))
    CUR_STK = Val(OPS_STK) - Val(CLS_STK)
    MS.TextMatrix(A, 3) = P_QTY_UNT(Trim(CUR_STK), REC_SET(0))
    MS.TextMatrix(A, 6) = RATCAL(MS.TextMatrix(A, 3), MS.TextMatrix(A, 5), REC_SET(0))
    tax = Val(tax) + Val(MS.TextMatrix(A, 5))
    units = QTY(Val(units) + Val(MS.TextMatrix(A, 3)))
    End If
Next
M_GRS_AMT = convert(amt(Total))
Total = convert(amt(Total))
M_NET_AMT = convert(amt(Val(M_GRS_AMT)))
GRS_CAL
End Sub

Public Sub PRINTRECORD()
If check("PRT", "PRT_NME", "M_FRM_PCR") = True Then
M_FRM_PRV.CHECK3.Value = 1
End If
M_FRM_PRV.Tag = "PCR"
M_FRM_PRV.Show 1
End Sub

Public Sub setval()
clrctr M_FRM_PCR
Label17 = "0.00"
Label18 = "0"
Label19 = "0.00"
Label20 = "0.00"
M_GRS_AMT = "0.00"
M_NET_AMT = "0.00"
label41 = "0.00"
M_SMN_NME = ""
Label40 = ""
ad = "0.00"
cd = "0.00"
units = "0.00"
tax = "0.00"
Total = "0.00"
M_INV_DAT = Date

End Sub

Public Sub Loadproducts()
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
With REC_SET
    .Open "SELECT * FROM INL WHERE INV_IDY LIKE '" & Trim(m_van_idy) & "' AND TXN_TYP LIKE 'PCS'", con, adOpenKeyset, adLockOptimistic
    If .RecordCount > 0 Then
        M_GRS_AMT = convert(amt(.Fields("NET_AMT") & ""))
        If .State = 1 Then .Close
        .Open "SELECT * FROM INV WHERE INV_IDY LIKE '" & Trim(m_van_idy) & "' AND TXN_TYP LIKE 'PCS'", con, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
        MS.Rows = .RecordCount + 1
        For A = 1 To .RecordCount
            If REC_SET1.State = 1 Then REC_SET1.Close
            REC_SET1.Open "select * from PRD where prd_idy like '" & Trim(.Fields("PRD_IDY")) & "'", con, adOpenKeyset
            If REC_SET1.RecordCount > 0 Then
                MS.TextMatrix(A, 0) = REC_SET("PRD_IDY")
                MS.TextMatrix(A, 1) = QTY(REC_SET("PRD_QTY"))
                MS.TextMatrix(A, 4) = bslab
                If bslab = "A" Then
                MS.TextMatrix(A, 5) = convert(amt(REC_SET1("SAL_rat") & ""))
                ElseIf bslab = "B" Then
                MS.TextMatrix(A, 5) = convert(amt(REC_SET1("PUR_rat") & ""))
                ElseIf bslab = "C" Then
                MS.TextMatrix(A, 5) = convert(amt(REC_SET1("mrp_rat") & ""))
                Else
                MS.TextMatrix(A, 5) = convert(amt(REC_SET1("SAL_rat") & ""))
                End If
                SSTab1.Tab = 1
                MS.Row = MS.Rows - 1
            End If
            .MoveNext
        Next
        End If
    End If
    discalc
End With
End Sub
