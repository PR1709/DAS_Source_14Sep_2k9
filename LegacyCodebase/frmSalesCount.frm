VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSalesCount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Cycle Count"
   ClientHeight    =   5625
   ClientLeft      =   1455
   ClientTop       =   330
   ClientWidth     =   8820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSalesCount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
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
            Picture         =   "frmSalesCount.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":2E02
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount.frx":334E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
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
         NumButtons      =   18
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Click to Check Post Empties After SAVE"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Voucher Details"
      TabPicture(0)   =   "frmSalesCount.frx":37A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "M_ACT_NME"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "M_SMN_NME"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label21(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line2(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "M_SUP_AD4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "M_SUP_AD3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "M_SUP_AD2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "M_SUP_AD1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "M_SUP_NME"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "D_ACT_NME"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label21(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "M_INV_DAT"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "M_SMN_IDY"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "M_SUP_IDY"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "M_ACT_IDY"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "M_INV_IDY"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "D_ACT_IDY"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "&Products"
      TabPicture(1)   =   "frmSalesCount.frx":37BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSF"
      Tab(1).Control(1)=   "Text5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "frmSalesCount.frx":37DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label49"
      Tab(2).Control(1)=   "list4"
      Tab(2).Control(2)=   "list2"
      Tab(2).Control(3)=   "list1"
      Tab(2).Control(4)=   "Check1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Check2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "list3"
      Tab(2).ControlCount=   7
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -72240
         MaxLength       =   10
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.ListView list3 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   17
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Party Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox D_ACT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Place"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69960
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -73200
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4680
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox M_ACT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox M_SUP_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox M_SMN_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   1800
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16449539
         CurrentDate     =   37099
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   4635
         Left            =   -74880
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   8176
         _Version        =   393216
         Cols            =   31
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         GridLines       =   2
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
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   21
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
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   22
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
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   23
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Name"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit (Stock Receipts)"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   31
         Top             =   4275
         Width           =   1605
      End
      Begin VB.Label D_ACT_NME 
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
         Left            =   4080
         TabIndex        =   30
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label M_SUP_NME 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   4200
         TabIndex        =   28
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label M_SUP_AD1 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   4200
         TabIndex        =   27
         Top             =   1650
         Width           =   3375
      End
      Begin VB.Label M_SUP_AD2 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   4200
         TabIndex        =   26
         Top             =   2010
         Width           =   3375
      End
      Begin VB.Label M_SUP_AD3 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   4200
         TabIndex        =   25
         Top             =   2370
         Width           =   3375
      End
      Begin VB.Label M_SUP_AD4 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   4200
         TabIndex        =   24
         Top             =   2640
         Width           =   3375
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
         TabIndex        =   18
         Top             =   480
         Width           =   8295
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   4080
         X2              =   7680
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit - Credit"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   750
         TabIndex        =   15
         Top             =   1155
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   675
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   5295
         TabIndex        =   13
         Top             =   675
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit (Stock Issues)"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   12
         Top             =   3788
         Width           =   1455
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
         Left            =   4080
         TabIndex        =   11
         Top             =   3240
         Visible         =   0   'False
         Width           =   3615
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
         Left            =   4080
         TabIndex        =   10
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   765
         TabIndex        =   9
         Top             =   3315
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1935
         Left            =   4080
         Top             =   1080
         Width           =   3615
      End
   End
   Begin VB.Label lblMinusGross 
      Caption         =   "#"
      Height          =   195
      Left            =   5520
      TabIndex        =   34
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label4 
      Caption         =   "#"
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblPlusGross 
      Caption         =   "#"
      Height          =   195
      Left            =   5520
      TabIndex        =   32
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label GST 
      Caption         =   "Label18"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label CST 
      Caption         =   "Label18"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmSalesCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, bslab, MSROW, INVIDY, INC, dlr_idy, ZIPCODE
Dim M_LST_RAT As Double
Dim M_LST_DIS As Double
Dim M_CPN_AMT As Double

Private Sub MSF_Click()
If (MSF.Col = 3 Or MSF.Col = 4) And MSF.Row > 0 And MSF.Row < MSF.Rows Then
    MSF_GotFocus
End If
End Sub

Private Sub MSF_GotFocus()
    MSF.HighLight = flexHighlightAlways
    Text5.Text = ""
    If MSF.Col = 3 Or MSF.Col = 4 Then
        Text5.Visible = True
        Text5.Width = MSF.CellWidth
        Text5.Top = MSF.Top + MSF.CellTop
        Text5.Height = MSF.RowHeight(MSF.Row)
        Text5.Left = MSF.Left + MSF.CellLeft
        If Trim(MSF.TextMatrix(MSF.Row, MSF.Col)) <> "" Then
            Text5.Text = QTY(Val(MSF.TextMatrix(MSF.Row, MSF.Col)))
        End If
        Text5.SelStart = 1
        Text5.Visible = True
        Text5.SetFocus
    End If
End Sub

Private Sub MSF_keypress(K As Integer)
    If K = 13 Then
        CalculateValues
        If MSF.Col = 3 Then
            MSF.Col = 4
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.RowHeight(MSF.Row)
            Text5.Left = MSF.Left + MSF.CellLeft
            If Trim(MSF.TextMatrix(MSF.Row, MSF.Col)) <> "" Then
                Text5.Text = QTY(Val(MSF.TextMatrix(MSF.Row, MSF.Col)))
            End If
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
        ElseIf MSF.Col = 4 Then
            MSF.Col = 3
            If MSF.Row > 0 And MSF.Row < MSF.Rows - 1 Then
                MSF.Row = MSF.Row + 1
            End If
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.CellHeight
            Text5.Left = MSF.Left + MSF.CellLeft
            If Trim(MSF.TextMatrix(MSF.Row, MSF.Col)) <> "" Then
                Text5.Text = QTY(Val(MSF.TextMatrix(MSF.Row, MSF.Col)))
            End If
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
        End If
    ElseIf K = 38 Then
            CalculateValues
            If MSF.Row > 1 Then
                MSF.Row = MSF.Row - 1
            End If
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.CellHeight
            Text5.Left = MSF.Left + MSF.CellLeft
            If Trim(MSF.TextMatrix(MSF.Row, MSF.Col)) <> "" Then
                Text5.Text = QTY(Val(MSF.TextMatrix(MSF.Row, MSF.Col)))
            End If
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
    ElseIf K = 40 Then
            CalculateValues
            If MSF.Row < MSF.Rows - 1 Then
                MSF.Row = MSF.Row + 1
            End If
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.RowHeight(MSF.Row)
            Text5.Left = MSF.Left + MSF.CellLeft
            If Trim(MSF.TextMatrix(MSF.Row, MSF.Col)) <> "" Then
                Text5.Text = QTY(Val(MSF.TextMatrix(MSF.Row, MSF.Col)))
            End If
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
    
    End If
End Sub

Private Sub MSF_LeaveCell()
If Text5.Visible = True Then Text5.Visible = False
If Text5.Text <> "" Then
    If MSF.Col = 3 Or MSF.Col = 4 And (MSF.Row > 0 And MSF.Row < MSF.Rows) Then
        MSF.TextMatrix(MSF.Row, MSF.Col) = QTY(Val(Text5.Text))
        Text5.Text = ""
    End If
    MSF_keypress 13
End If
End Sub

Private Sub MSF_Scroll()
If Text5.Visible = True Then
    Text5 = ""
    Text5.Visible = False
    MSF.SetFocus
End If
End Sub
Private Sub TEXT5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSF.Col = 3 Or MSF.Col = 4 Then
        MSF.TextMatrix(MSF.Row, MSF.Col) = QTY(Val(Text5.Text))
        Text5.Text = ""
    End If
    MSF_keypress 13
ElseIf KeyCode = 40 Then
    MSF.TextMatrix(MSF.Row, MSF.Col) = QTY(Val(Text5.Text))
    Text5.Text = ""
    MSF_keypress 40
ElseIf KeyCode = 38 Then
    MSF.TextMatrix(MSF.Row, MSF.Col) = QTY(Val(Text5.Text))
    Text5.Text = ""
    MSF_keypress 38
End If
End Sub

Private Sub CREATE()
MOD_REC = False
M_INV_IDY = ""
saverecord
End Sub

Public Sub GRS_CAL()
'On Error Resume Next
'G_PRD_VAL = convert(Val(AMT_TOT))
'T_PRD_VAL = convert(Val(AMT_TOT))
'
'G_PRD_DIS = convert(Val(DIS_TOT) + Val(ADD_TOT))
'T_PRD_DIS = convert(Val(AMT_TOT) - (Val(DIS_TOT) + Val(ADD_TOT)))
'
'G_PRD_TAX = convert(Val(TAX_TOT))
'T_PRD_TAX = convert(Val(TAX_TOT) + Val(T_PRD_DIS))
'
'G_GRS_AMT = T_PRD_TAX
'T_GRS_AMT = T_PRD_TAX
'
'G_CSD_PCG = convert(Val(M_CSD_AMT))
'T_CSD_AMT = convert(Val(T_GRS_AMT) - Val(G_CSD_PCG))
'
'G_SPD_PCG = convert(Val(M_SPD_AMT))
'T_SPD_AMT = convert(Val(T_CSD_AMT) - Val(G_SPD_PCG))
'
'G_TAX_PCG = convert(Val(M_TAX_AMT))
'T_TAX_AMT = convert(Val(M_TAX_AMT) + Val(T_SPD_AMT))
'
'G_SCT_PCG = convert(Val(M_EXD_AMT))
'T_SCT_PCG = convert(Val(M_EXD_AMT) + Val(T_TAX_AMT))
'
'G_PKG_CST = convert(Val(M_PKG_CST))
'T_PKG_CST = convert(Val(G_PKG_CST) + Val(T_SCT_PCG))
'
'G_FRT_AMT = convert(Val(M_FRT_AMT))
'T_FRT_AMT = convert(Val(T_PKG_CST) - Val(G_FRT_AMT))
'ROU_OFF
'T_ROU_OFF = convert(Val(T_FRT_AMT) + Val(M_ROU_OFF))
'T_NET_AMT = convert(T_ROU_OFF)
End Sub

Private Sub saverecord()
'On Error Resume Next
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
Dim REC_IRT As New ADODB.Recordset
Dim REC_PNL As New ADODB.Recordset
Dim INC As String

If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
Else
    CheckShutDownStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'")
    If Trim(CheckShutDownStatus) = "IMPROPER SHUTDOWN" Then
        MsgBox ("Please reset the Improper Shutdown flag in System-Shutdown Status"), vbCritical, "DAS Version 4.0"
        CANCELRECORD
        Exit Sub
    End If
End If
If CHECKTB("act.dbf,inv.dbf,prd.dbf,brd.dbf,inl.dbf,pnv.dbf,pnl.dbf,dir.dbf,mar.dbf,loc.dbf,smn.dbf,cat.dbf,rri.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
'If check("ACT", "(ACT_TYP LIKE 'SAL' OR ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'PUR') AND ACT_idy", Trim(M_SUP_IDY)) = False Then
If check("ACT", "ACT_IDY", Trim(M_SUP_IDY)) = False Then
    MsgBox "Party Not Found ,Define Party !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_SUP_IDY.SetFocus
    Exit Sub
End If

'If check("SMN", "SMN_IDY", Trim(M_SMN_IDY)) = False Then
'    MsgBox "Salesman Not Found ,Define Salesman !", vbExclamation, "DAS Version 4.0"
'    SSTab1.Tab = 0
'    M_SMN_IDY.SetFocus
'    Exit Sub
'End If

If check("act", "(grp_clf like 'Stock Issues ' and ACT_TYP LIKE '') AND act_idy", Trim(M_ACT_IDY)) = False Then
    MsgBox "Stock Issues Account Not Found, Define Stock Issues  Account...!", vbExclamation, "DAS Version 4.0"
    M_ACT_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If


' Day Closure Routine Added by PNR and PKB as Addon Control

Dim RS_DAYCLS As New ADODB.Recordset
If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
RS_DAYCLS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) like 'DAY' .and. alltrim(DIR_ALS) like '" & M_INV_DAT & "'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
If RS_DAYCLS.RecordCount > 0 Then
   If Trim(RS_DAYCLS!DIR_MSG) = "Y" Then
    MsgBox "Day Closure Effective ! Cannot Alter - Create Transaction FOR " & RS_DAYCLS("DIR_ALS"), vbCritical, "DAS Version 4.0"
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

'stock outward
If Trim(lblMinusGross.Caption) <> "" Then
    REC_SET.Open "select * from inl where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
    INC = UpdateRecordSetINL("STO", REC_SET)
    REC_PNL.Open "SELECT * FROM PNL", PNL, adOpenKeyset, adLockPessimistic
    UpdateRecordsetPNL INC, REC_PNL, "STO"
    REC_PNL.Close
    UpdateRecordsetINV "STO", INC
    UpdateRecordsetPNV "STO", INC
    MsgBox "New Stock Issues Voucher ID is " & INC, vbExclamation, "DAS Version 4.0"
End If
'stock Inward
If Trim(lblPlusGross.Caption) <> "" Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inl where inv_idy = ''", con, adOpenKeyset, adLockPessimistic
    INC = UpdateRecordSetINL("STI", REC_SET)
    If REC_PNL.State = 1 Then REC_PNL.Close
    REC_PNL.Open "SELECT * FROM PNL", PNL, adOpenKeyset, adLockPessimistic
    UpdateRecordsetPNL INC, REC_PNL, "STI"
    REC_PNL.Close
    UpdateRecordsetINV "STI", INC
    UpdateRecordsetPNV "STI", INC
    MsgBox "New Stock Receipt Voucher ID is " & INC, vbExclamation, "DAS Version 4.0"
End If
'PRINT DATABASE LOGIC END

'pnv update begin
AddShells Trim(INC), "STO", M_SUP_IDY, "Y", "", ""

If check("PRT", "PRT_NME", "STO_SHL") Then LOADSHELLS "STO", Trim(INC)
'pnv update close
CANCELRECORD
End Sub

Private Function UpdateRecordsetPNV(StockType As String, INC As String)
Dim Q As Double

Dim REC_PNL As New ADODB.Recordset

If REC_PNL.State = 1 Then REC_PNL.Close
REC_PNL.Open "select * from pnv where inv_idy = ''", PNL, adOpenKeyset, adLockPessimistic
For A = 1 To MSF.Rows - 1
    If Trim(MSF.TextMatrix(A, 0)) <> "" Then
        If StockType = "STO" Then
            Q = Val(MSF.TextMatrix(A, 4))
        Else
            Q = Val(MSF.TextMatrix(A, 3))
        End If
        If Q > 0 Then
        
            REC_PNL.AddNew
            REC_PNL("inv_idy") = Trim(INC)
            REC_PNL("inv_dat") = Trim(M_INV_DAT)
            REC_PNL("prd_idy") = MSF.TextMatrix(A, 0)
            
            
            REC_PNL("grp_idy") = Trim(MSF.TextMatrix(A, 27))
            REC_PNL("grp_cat") = Trim(MSF.TextMatrix(A, 28))
            REC_PNL("grp_cls") = Trim(MSF.TextMatrix(A, 29))
            
            If StockType = "STO" Then
                REC_PNL("prd_qty") = QTY(Val(MSF.TextMatrix(A, 4)))
                REC_PNL!PRD_UOM = QTY(Val(MSF.TextMatrix(A, 4))) * Val(MSF.TextMatrix(A, 11))
                REC_PNL!PRD_UNT = QTY(Val(MSF.TextMatrix(A, 4)))
            Else
                REC_PNL("prd_qty") = QTY(Val(MSF.TextMatrix(A, 3)))
                REC_PNL!PRD_UOM = QTY(Val(MSF.TextMatrix(A, 3))) * Val(MSF.TextMatrix(A, 11))
                REC_PNL!PRD_UNT = QTY(Val(MSF.TextMatrix(A, 3)))
            End If
            
            REC_PNL("prd_rat") = Val(MSF.TextMatrix(A, 8))
            REC_PNL("trd_dis") = Val(MSF.TextMatrix(A, 9))
            REC_PNL("trd_tax") = Val(MSF.TextMatrix(A, 10))
    
    
            REC_PNL!MRP_RAT = Val(MSF.TextMatrix(A, 30))
            
            
            
            REC_PNL("ADD_AMT") = 0
            
            
            If StockType = "STO" Then
                REC_PNL("DIS_AMT") = Val(MSF.TextMatrix(A, 17))
                REC_PNL("TAX_AMT") = Val(MSF.TextMatrix(A, 18))
            Else
                REC_PNL("DIS_AMT") = Val(MSF.TextMatrix(A, 15))
                REC_PNL("TAX_AMT") = Val(MSF.TextMatrix(A, 16))
            End If
            REC_PNL("txn_typ") = StockType
            
            
            
            REC_PNL!BAT_IDY = Trim(MSF.TextMatrix(A, 1))
            REC_PNL!SUB_UNT = Val(MSF.TextMatrix(A, 11))
            If StockType = "STO" Then
                REC_PNL!PRD_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 4)), Val(MSF.TextMatrix(A, 11))) Mod Val(MSF.TextMatrix(A, 11))  'P_SUB_UNT(QTY(msf.TextMatrix(A, 3)), Val(msf.TextMatrix(A, 12))) Mod Val(msf.TextMatrix(A, 12))
                REC_PNL!TOT_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 4)), Val(MSF.TextMatrix(A, 11)))
            Else
                REC_PNL!PRD_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 3)), Val(MSF.TextMatrix(A, 1))) Mod Val(MSF.TextMatrix(A, 11)) 'P_SUB_UNT(QTY(msf.TextMatrix(A, 3)), Val(msf.TextMatrix(A, 12))) Mod Val(msf.TextMatrix(A, 12))
                REC_PNL!TOT_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 3)), Val(MSF.TextMatrix(A, 11)))
            End If
            If StockType = "STO" Then
                REC_PNL("COL_AMT") = Val((MSF.TextMatrix(A, 20)))
            Else
                REC_PNL("COL_AMT") = Val((MSF.TextMatrix(A, 19)))
            End If
            REC_PNL("LST_RAT") = Val(MSF.TextMatrix(A, 24))
            REC_PNL("LST_DIS") = Val(MSF.TextMatrix(A, 25))
            REC_PNL!stk_upd = "Y"
            
            
            REC_PNL("prd_nme") = Trim(MSF.TextMatrix(A, 13))
            REC_PNL("STK_TYP") = Trim(MSF.TextMatrix(A, 21))
            REC_PNL("LST_RAT") = Val(MSF.TextMatrix(A, 22))
            REC_PNL("LST_DIS") = Val(MSF.TextMatrix(A, 23))
            
            REC_PNL("ret_qty") = 0
            REC_PNL("iss_qty") = 0
            REC_PNL("tot_qty") = 0
            REC_PNL("rat_app") = ""
            REC_PNL("cpn_rat") = 0
            REC_PNL("cpn_amt") = 0 'Round(Val(M_CPN_AMT), 2)
            REC_PNL("add_dis") = 0
            REC_PNL("fst_flg") = ""
            REC_PNL("opn_qty") = 0
            REC_PNL("sch_amt") = 0
            REC_PNL("ord_idy") = 0
            REC_PNL("bat_nme") = Trim(MSF.TextMatrix(A, 14))
            REC_PNL("sch_idy") = ""
            REC_PNL("INC_TTL") = 0
            REC_PNL.Update
        End If
    End If
Next

End Function
Private Function UpdateRecordsetINV(StockType As String, INC As String)
Dim Q As Double

Dim REC_SET As New ADODB.Recordset

If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from inv where inv_idy = ''", con, adOpenKeyset, adLockPessimistic

For A = 1 To MSF.Rows - 1
    If Trim(MSF.TextMatrix(A, 0)) <> "" Then
        If StockType = "STO" Then
            Q = Val(MSF.TextMatrix(A, 4))
        Else
            Q = Val(MSF.TextMatrix(A, 3))
        End If
        If Q > 0 Then
        
            REC_SET.AddNew
            REC_SET("inv_idy") = Trim(INC)
            REC_SET("inv_dat") = Trim(M_INV_DAT)
            REC_SET("prd_idy") = MSF.TextMatrix(A, 0)
            If StockType = "STO" Then
                REC_SET("prd_qty") = QTY(Val(MSF.TextMatrix(A, 4)))
                REC_SET!PRD_UOM = QTY(Val(MSF.TextMatrix(A, 4))) * Val(MSF.TextMatrix(A, 11))
                REC_SET!PRD_UNT = QTY(Val(MSF.TextMatrix(A, 4)))
            Else
                REC_SET("prd_qty") = QTY(Val(MSF.TextMatrix(A, 3)))
                REC_SET!PRD_UOM = QTY(Val(MSF.TextMatrix(A, 3))) * Val(MSF.TextMatrix(A, 11))
                REC_SET!PRD_UNT = QTY(Val(MSF.TextMatrix(A, 3)))
            End If
            REC_SET("prd_rat") = Val(MSF.TextMatrix(A, 8))
            REC_SET("trd_dis") = Val(MSF.TextMatrix(A, 9))
            REC_SET("trd_tax") = Val(MSF.TextMatrix(A, 10))
            REC_SET("ADD_AMT") = 0
            If StockType = "STO" Then
                REC_SET("DIS_AMT") = Val(MSF.TextMatrix(A, 17))
                REC_SET("TAX_AMT") = Val(MSF.TextMatrix(A, 18))
            Else
                REC_SET("DIS_AMT") = Val(MSF.TextMatrix(A, 15))
                REC_SET("TAX_AMT") = Val(MSF.TextMatrix(A, 16))
            End If
            REC_SET("txn_typ") = StockType
            REC_SET!BAT_IDY = Trim(MSF.TextMatrix(A, 1))
            REC_SET!SUB_UNT = Val(MSF.TextMatrix(A, 11))
            If StockType = "STO" Then
                REC_SET!PRD_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 4)), Val(MSF.TextMatrix(A, 11))) Mod Val(MSF.TextMatrix(A, 11))
                REC_SET!TOT_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 4)), Val(MSF.TextMatrix(A, 11)))
            Else
                REC_SET!PRD_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 3)), Val(MSF.TextMatrix(A, 1))) Mod Val(MSF.TextMatrix(A, 11))
                REC_SET!TOT_PCS = P_SUB_UNT(QTY(MSF.TextMatrix(A, 3)), Val(MSF.TextMatrix(A, 11)))
            End If
            If StockType = "STO" Then
                REC_SET("COL_AMT") = Val((MSF.TextMatrix(A, 20)))
            Else
                REC_SET("COL_AMT") = Val((MSF.TextMatrix(A, 19)))
            End If
            REC_SET("STK_TYP") = Trim(MSF.TextMatrix(A, 21))
            REC_SET("LST_RAT") = Val(MSF.TextMatrix(A, 22))
            REC_SET("LST_DIS") = Val(MSF.TextMatrix(A, 23))
            REC_SET("CPN_AMT") = 0 'Round(Val(M_CPN_AMT), 2)
            REC_SET!stk_upd = "Y"
            REC_SET("ORD_IDY") = ""
            REC_SET("INC_TTL") = 0
            REC_SET.Update
            If StockType = "STO" Then
                con.Execute "update PRD where prd_idy like '" & Trim(MSF.TextMatrix(A, 0)) & "' set CUR_LEV = " & CONQTY(Trim(Val(MSF.TextMatrix(A, 4))), Trim(MSF.TextMatrix(A, 0)))
                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MSF.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS - " & MTS(Val(MSF.TextMatrix(A, 4)), Val(MSF.TextMatrix(A, 11)))
                con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MSF.TextMatrix(A, 2)) & "' set prd_qty = " & CONBRDQTY(Trim(Val(MSF.TextMatrix(A, 4))), Trim(MSF.TextMatrix(A, 2)))
            Else
                con.Execute "update PRD where prd_idy like '" & Trim(MSF.TextMatrix(A, 0)) & "' set CUR_LEV = " & ADDQTY(Trim(Val(MSF.TextMatrix(A, 3))), Trim(MSF.TextMatrix(A, 0)))
                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(MSF.TextMatrix(A, 0)) & "' set PRD_PCS = PRD_PCS + " & MTS(Val(MSF.TextMatrix(A, 3)), Val(MSF.TextMatrix(A, 11)))
                con.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(MSF.TextMatrix(A, 2)) & "' set prd_qty = " & ADDBRDQTY(Trim(Val(MSF.TextMatrix(A, 3))), Trim(MSF.TextMatrix(A, 2)))
            End If
        End If
    End If
Next
End Function

Private Function UpdateRecordSetINL(StockType As String, REC_SET As ADODB.Recordset) As String
    REC_SET.AddNew
    REC_SET("inv_idy") = StockType & T7increment("IDY_" & StockType)
    INC = REC_SET("inv_idy")
    INVIDY = INC
    
    REC_SET("inv_dat") = M_INV_DAT
    REC_SET("ord_idy") = " "
    REC_SET("cst_idy") = Trim(M_SUP_IDY.Text & " ")
    REC_SET("smn_idy") = Trim(M_SMN_IDY.Text & " ")
    If StockType = "STO" Then
        REC_SET("act_idy") = Trim(M_ACT_IDY.Text & " ")
    Else
        REC_SET("act_idy") = Trim(D_ACT_IDY.Text & " ")
    End If
    REC_SET("dch_idy") = " "
    REC_SET("dch_dat") = M_INV_DAT
    REC_SET("csd_pcg") = 0
    REC_SET("csd_amt") = 0
    REC_SET("spd_pcg") = 0
    REC_SET("spd_amt") = 0
    REC_SET("tax_pcg") = 0
    REC_SET("tax_amt") = 0
    REC_SET("exd_pcg") = 0         'SURCHARGE PCG
    REC_SET("EXD_AMT") = 0             'SURCHARGE AMT
    REC_SET("pkg_cst") = 0             'AMOUNT (+)
    REC_SET("FRT_AMT") = 0        'AMOUNT (-)
    REC_SET("LRN_IDY") = 0
    REC_SET("LRN_AMT") = 0
    REC_SET("LRN_DAT") = M_INV_DAT
    REC_SET("WBL_IDY") = " "
    REC_SET("WBL_DAT") = M_INV_DAT
    REC_SET("TXN_RMK") = " "
    REC_SET("TTL_CAS") = 0
    REC_SET("txn_typ") = StockType
    
    If StockType = "STO" Then
        REC_SET("grs_amt") = lblMinusGross.Caption
        REC_SET("net_amt") = lblMinusGross.Caption
    Else
        REC_SET("grs_amt") = lblPlusGross.Caption
        REC_SET("net_amt") = lblPlusGross.Caption
    End If
    REC_SET("crt_day") = 0
    REC_SET("due_dat") = M_INV_DAT
    REC_SET("rou_off") = 0
    REC_SET("ROU_IDY") = ""
    REC_SET("TXN_TPT") = ""
    REC_SET("TXN_RM1") = ""
    REC_SET("TXN_RM2") = ""
    REC_SET!USR_IDY = Trim(EFF_LIB1.USERNAME)
    REC_SET!sys_dat = Date
    REC_SET!sys_tim = Format(Time, "HH:MM:SS")
    REC_SET("SCH_CHK") = "N"
    REC_SET!SRN_AMT = 0
    REC_SET!vch_idy = ""
    REC_SET!Sch_Idy = ""
    REC_SET.Update
    REC_SET.Close
    UpdateRecordSetINL = INC
End Function

Private Function UpdateRecordsetPNL(INC As String, REC_PNL As ADODB.Recordset, StockType As String)
    REC_PNL.AddNew
    REC_PNL!INV_IDY = INC
    REC_PNL("inv_dat") = M_INV_DAT
    REC_PNL("ord_idy") = ""
    REC_PNL("cst_idy") = Trim(M_SUP_IDY.Text & "")
    REC_PNL("cst_nme") = Trim(M_SUP_NME.Caption)
    REC_PNL("cst_ad1") = Trim(M_SUP_AD1.Caption)
    REC_PNL("cst_ad2") = Trim(M_SUP_AD2.Caption)
    REC_PNL("cst_ad3") = Trim(M_SUP_AD3.Caption)
    REC_PNL("cst_ad4") = Trim(M_SUP_AD4.Caption)
    REC_PNL("cst_zip") = ""
    If StockType = "STO" Then
        REC_PNL("act_idy") = Trim(M_ACT_IDY.Text & "")
    Else
        REC_PNL("act_idy") = Trim(D_ACT_IDY.Text & "")
    End If
    
    REC_PNL("dch_idy") = ""
    REC_PNL("dch_dat") = M_INV_DAT
    REC_PNL("csd_pcg") = 0
    REC_PNL("csd_amt") = 0
    REC_PNL("spd_pcg") = 0
    REC_PNL("spd_amt") = 0
    REC_PNL("tax_pcg") = 0
    REC_PNL("tax_amt") = 0
    REC_PNL("exd_pcg") = 0            'SURCHARGE PCG
    REC_PNL("EXD_AMT") = 0             'SURCHARGE AMT
    REC_PNL("pkg_cst") = 0             'AMOUNT (+)
    REC_PNL("FRT_AMT") = 0        'AMOUNT (-)
    REC_PNL("LRN_IDY") = ""
    REC_PNL("LRN_AMT") = 0
    REC_PNL("LRN_DAT") = M_INV_DAT
    REC_PNL("WBL_IDY") = ""
    REC_PNL("WBL_DAT") = M_INV_DAT
    REC_PNL("TXN_RMK") = ""
    REC_PNL("TTL_CAS") = 0
    REC_PNL("txn_typ") = StockType
    If StockType = "STO" Then
        REC_PNL("grs_amt") = lblMinusGross.Caption
        REC_PNL("net_amt") = lblMinusGross.Caption
    Else
        REC_PNL("grs_amt") = lblPlusGross.Caption
        REC_PNL("net_amt") = lblPlusGross.Caption
    End If
    REC_PNL("crt_day") = M_INV_DAT
    REC_PNL("due_dat") = Date
    REC_PNL("rou_off") = 0
    REC_PNL("ROU_IDY") = ""
    REC_PNL("TXN_TPT") = ""
    REC_PNL("TXN_RM1") = ""
    REC_PNL("TXN_RM2") = ""
    REC_PNL!usr_nme = Trim(USERNAME)
    REC_PNL!sys_dat = Date
    REC_PNL!sys_tim = Format(Time, "HH:MM:SS")
    REC_PNL("SCH_CHK") = "N"
    REC_PNL!SRN_AMT = 0
    REC_PNL!vch_idy = ""
    REC_PNL!Sch_Idy = ""
    REC_PNL!cst_gst = ""
    REC_PNL!cst_cst = ""
    REC_PNL!MAR_IDY = GetValue("mar_idy", "act where act_idy = '" & Trim(M_SUP_IDY.Text) & "'")
    REC_PNL!mar_nme = GetValue("mar_nme", "mar where mar_idy = '" & Trim(REC_PNL("mar_idy")) & "'")
    REC_PNL!SMN_NME = GetValue("smn_nme", "smn where smn_idy = '" & Trim(M_SMN_IDY.Text) & "'")
    REC_PNL!loc_IDY = GetValue("loc_idy", "act where act_idy = '" & Trim(M_SUP_IDY) & "'")
    REC_PNL!loc_nme = GetValue("loc_nme", "loc where loc_idy = '" & Trim(REC_PNL("loc_idy")) & "'")
    REC_PNL!loc_IDY = GetValue("loc_idy", "act where act_idy = '" & Trim(M_SUP_IDY) & "'")
    REC_PNL!cat_IDY = GetValue("cat_idy", "act where act_idy = '" & Trim(M_SUP_IDY) & "'")
    REC_PNL!cat_nme = GetValue("cat_nme", "cat where cat_idy = '" & Trim(REC_PNL("cat_idy")) & "'")
    REC_PNL!exd_cal = ""
    REC_PNL!SMN_IDY = M_SMN_IDY.Text
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
    REC_PNL!DIS_CHK = ""
    REC_PNL.Update

End Function

Private Sub CANCELRECORD()
clrctr frmSalesCount
setval
Form_Load
'MS.Rows = 1
'MS.AddItem ""
SSTab1.Tab = 0
D_ACT_NME.Caption = ""
M_ACT_NME.Caption = ""
MOD_REC = False
'M_INV_IDY.Enabled = True
M_INV_DAT.SetFocus
End Sub

Private Sub newrecord()
CANCELRECORD
M_INV_IDY.Enabled = False
M_INV_DAT.SetFocus
SSTab1.Tab = 0
End Sub

Private Sub V_SUP_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where  act_nme like '" & Trim(M_SUP_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
    '(ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL' OR ACT_TYP LIKE 'PUR') and  AND DBT_TYP='Y'
    ' ABOVE REC_SET QUERY CHANGED BY PNR AS PER REQUEST FROM PRINCE CHOPRA FOR BREAKAGE ACCOUNTING
ElseIf Check2.Value = 1 Then
    REC_SET.Open "select * from act where  act_ad4 like '" & Trim(M_SUP_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
    '(ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL' OR ACT_TYP LIKE 'PUR') and AND DBT_TYP='Y'
    ' ABOVE REC_SET QUERY CHANGED BY PNR AS PER REQUEST FROM PRINCE CHOPRA FOR BREAKAGE ACCOUNTING
End If
list3.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list3.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Party"
    list3.Visible = True
    list3.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Party Not Found, Define Party !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_SUP_IDY.SetFocus
    End If
End If
End Sub
Private Sub V_SMN_LST()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Label49 = "Select Salesman"
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
ElseIf Check2.Value = 1 Then
    REC_SET.Open "select * from SMN WHERE SMN_AD4 like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
End If
list1.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(REC_SET("smn_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 2
    list1.Visible = True
    list1.SetFocus
    
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Salesmen Not Found, Define Salesmen !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_SMN_IDY.SetFocus
    End If
End If
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
    REC_SET.Open "select * from act where grp_clf like 'Stock Issues ' and ACT_NME LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from act where grp_clf like 'Stock Issues ' and ACT_AD4 LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
End If
list2.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list2.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Account"
    list2.Visible = True
    list2.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Stock Issues  Account Not Found, Define Stock Issues  Account...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_ACT_IDY.SetFocus
End If
End If
End Sub

Private Sub D_ACT_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where grp_clf like 'Stock Receipts ' and ACT_NME LIKE '" & Trim(D_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "select * from act where grp_clf like 'Stock Receipts ' and ACT_AD4 LIKE '" & Trim(D_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockPessimistic
End If
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Account"
    list4.Visible = True
    list4.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Stock Receipts  Account Not Found, Define Stock Receipts  Account...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    D_ACT_IDY.SetFocus
End If
End If
End Sub

Private Sub EXT_REC()
Unload Me
End Sub



Private Sub M_SMN_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSSMN As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 38 Then
    If CHECKTB("smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    RSSMN.Open "select * from smn where smn_idy like '" & Sincrement(M_SMN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If RSSMN.RecordCount > 0 Then
        M_SMN_IDY = Sincrement(M_SMN_IDY)
        M_SMN_NME = Trim(RSSMN!SMN_NME)
        If KeyCode = 38 Then
            M_SUP_IDY.SetFocus
        Else
            M_ACT_IDY.SetFocus
        End If
    Else
        V_SMN_LST
    End If
End If
End Sub

Private Sub M_SMN_IDY_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
KeyAscii = CHGCASE(KeyAscii)
End Sub


Private Sub M_ACT_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_ACT_IDY)) & "' AND GRP_CLF LIKE 'Stock Issues ' AND ACT_TYP LIKE ''", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_ACT_IDY = Sincrement(M_ACT_IDY)
        If K = 38 Then
        M_SUP_IDY.SetFocus
        Else
        M_ACT_NME = Trim(REC_SET("act_nme"))
        SSTab1.Tab = 1
        D_ACT_IDY.SetFocus
        End If
    Else
        V_ACT_LST
    End If
End If
End Sub

Private Sub M_ACT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Stock Issues  Account ID or Press <Enter> to View Existing Accounts"
End Sub

Private Sub M_ACT_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub D_ACT_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(D_ACT_IDY)) & "' AND GRP_CLF LIKE 'Stock Issues ' AND ACT_TYP LIKE ''", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        D_ACT_IDY = Sincrement(D_ACT_IDY)
        If K = 38 Then
        M_SUP_IDY.SetFocus
        Else
        D_ACT_NME = Trim(REC_SET("act_nme"))
        SSTab1.Tab = 1
        MSF.SetFocus
        End If
    Else
        D_ACT_LST
    End If
End If
End Sub

Private Sub D_ACT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Stock Issues  Account ID or Press <Enter> to View Existing Accounts"
End Sub

Private Sub D_ACT_IDY_KeyPress(K As Integer)
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
M_SUP_IDY.SetFocus
End If
End Sub

Private Sub M_SUP_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Party ID or Press <Enter> to View Parties List"
End Sub

Private Sub M_SUP_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_SUP_IDY)) & "'", con, adOpenKeyset, adLockPessimistic
    ' AND (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL' OR ACT_TYP LIKE 'PUR' OR ACT_TYP LIKE '') AND DBT_TYP='Y'
    If REC_SET.RecordCount > 0 Then
        M_SUP_IDY = Sincrement(M_SUP_IDY)
        If Len(Trim(M_SUP_IDY)) = 0 Then Exit Sub
        M_SUP_NME.Caption = Trim(REC_SET("act_nme") & "")
        M_SUP_AD1.Caption = Trim(REC_SET("act_ad1") & "")
        M_SUP_AD2.Caption = Trim(REC_SET("act_ad2") & "")
        M_SUP_AD3.Caption = Trim(REC_SET("act_ad3") & "")
        M_SUP_AD4.Caption = Trim(REC_SET("act_ad4") & "")
        CST.Caption = Trim(REC_SET("act_CST") & "")
        GST.Caption = Trim(REC_SET("act_GST") & "")
        ZIPCODE = Trim(REC_SET("act_zip"))
        If Len(Trim(REC_SET("ROU_IDY"))) <> 0 Then M_ROU_IDY = Trim(REC_SET("ROU_IDY") & "")
        If MOD_REC = False Then M_CSD_PCG = discount(Val(REC_SET("crt_dis")))
        If REC_SET1.State = 1 Then REC_SET1.Close
        SSTab1.Tab = 0
        If K = 38 Then
            M_DCH_DAT.SetFocus
        Else
            M_ACT_IDY.SetFocus
        End If
    Else
        V_SUP_LST
    End If
End If
End Sub

Private Sub M_SUP_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Form_Activate()
checkdongle
If Me.Tag <> "Batch" Then
SSTab1.Tab = 0
List7 = "Y"
'If M_INV_IDY.Enabled = True Then
'    M_INV_IDY.SetFocus
'Else
    M_INV_DAT.SetFocus
'End If
Label23 = labdisp("IDY_STO")
SSTab1.TabEnabled(2) = False
setval
SETYEAR frmSalesCount
End If
End Sub

Private Sub V_INV_LST()
If CHECKTB("inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4 from act,inl where act_nme like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='STO'", con, adOpenKeyset, adLockPessimistic
Else
REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4 from act,inl where act_ad4 like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='STO'", con, adOpenKeyset, adLockPessimistic
End If
list5.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Voucher"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Voucher Number Not Found...!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If list1.Visible = False And list2.Visible = False And list3.Visible = False And list4.Visible = False Then 'And list5.Visible = False
        If KeyCode = 78 And Shift = 2 Then newrecord
        If KeyCode = 80 And Shift = 2 Then Printrecord
        If KeyCode = 83 And Shift = 2 Then saverecord
        If KeyCode = 88 And Shift = 2 Then EXT_REC
        If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
    
    
        If frmSalesCount.ActiveControl.name = "M_SUP_IDY" Then
            V_SUP_LST
        ElseIf frmSalesCount.ActiveControl.name = "M_SMN_IDY" Then
            V_SMN_LST
        ElseIf frmSalesCount.ActiveControl.name = "M_ACT_IDY" Then
            V_ACT_LST
        ElseIf frmSalesCount.ActiveControl.name = "D_ACT_IDY" Then
            D_ACT_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf frmSalesCount.ActiveControl Is TextBox Then frmSalesCount.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_SMN_IDY.SetFocus
    ElseIf list2.Visible = True Then
        list2.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_ACT_IDY.SetFocus
    ElseIf list3.Visible = True Then
        list3.Visible = False
        Check2.Caption = "Place"
        SSTab1.Tab = 0
        M_SUP_IDY.SetFocus
    ElseIf list4.Visible = True Then
        list4.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
        If Text3.Visible = True Then
            Text3.SetFocus
        Else
            Check2.Caption = "Place"
            SSTab1.Tab = 1
            MS.Row = MS.Rows - 1
            MS.SetFocus
        End If
    Else
        If frmSalesCount.ActiveControl.name = "M_INV_DAT" Then
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
FillGrid


M_INV_DAT = Date
M_LRN_DAT = Date
M_DCH_DAT = Date
waydate = Date

'list5.ColumnHeaders(1).Width = list2.Width / 4
'list5.ColumnHeaders(2).Width = list2.Width / 4
'list5.ColumnHeaders(3).Width = list2.Width / 4
'list5.ColumnHeaders(4).Width = list2.Width / 4

TAX_TOT = "0.00"
UNT_TOT = "0.000"
ADD_TOT = "0.00"
DIS_TOT = "0.00"
AMT_TOT = "0.00"
list1.ColumnHeaders(1).Width = 3800
list2.ColumnHeaders(1).Width = 3800
list3.ColumnHeaders(1).Width = 3800
Label23 = labdisp("IDY_STO")
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
M_PRD_TOT = "0.00"
Tax = "0.00"
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press Ctrl+N to Generate New Stock Issues Voucher ID by System, Press <Enter> to View Existing Vouchers"
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
If M_INV_IDY = "" And K = 27 Then Unload Me
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
        M_SMN_IDY = list1.SelectedItem.ListSubItems.Item(1)
        M_SMN_NME = list1.SelectedItem
        SSTab1.Tab = 0
        M_ACT_IDY.SetFocus
        list1.Visible = False
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
If list2.SortOrder = lvwAscending Then
list2.SortOrder = lvwDescending
ElseIf list2.SortOrder = lvwDescending Then
list2.SortOrder = lvwAscending
End If
list2.SortKey = ColumnHeader.Index - 1
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



Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
If Label49.Caption = "Select Salesman" Then
    V_SMN_LST
ElseIf Label49.Caption = "Select Party" Then
    V_SUP_LST
ElseIf Label49.Caption = "Select Voucher" Then
    V_INV_LST
ElseIf Label49.Caption = "Select Account" Then
    V_ACT_LST
ElseIf Label49.Caption = "Select Product" Then
    V_PRD_LST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If Label49.Caption = "Select Salesman" Then
    V_SMN_LST
ElseIf Label49.Caption = "Select Party" Then
    V_SUP_LST
ElseIf Label49.Caption = "Select Voucher" Then
    V_INV_LST
ElseIf Label49.Caption = "Select Account" Then
    V_ACT_LST
ElseIf Label49.Caption = "Select Product" Then
    If Check2.Value = 1 Then
    con.Execute "insert into prt values('p_prdnum')"
    Else
    con.Execute "delete from prt where prt_nme like 'p_prdnum'"
    End If
    V_PRD_LST
End If
'End If
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
M_FRM_PRE.Caption = "Prefix Purchase"
M_FRM_PRE.Show 1
ElseIf Button.KEY = "crt" Then
CREATE
ElseIf Button.KEY = "ed" Then
    If frmSalesCount.ActiveControl.name = "M_INV_IDY" Then
        Label49 = "Select Voucher"
        V_INV_LST
    ElseIf frmSalesCount.ActiveControl.name = "M_SUP_IDY" Then
        Label49 = "Select Party"
        
        V_SUP_LST
    ElseIf frmSalesCount.ActiveControl.name = "M_SMN_IDY" Then
        Label49 = "Select Salesman"
        V_SMN_LST
    ElseIf frmSalesCount.ActiveControl.name = "M_ACT_IDY" Then
        Label49 = "Select Account"
        V_ACT_LST
    ElseIf frmSalesCount.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        V_PRD_LST
    End If
ElseIf Button.KEY = "DAY" Then
'    If Button.Value = tbrPressed Then
'        Text1.Visible = True
'        Text1 = Format(Date, "dd/mm/yyyy") & " " & Format(Date, "dddd")
'        Text1.BackColor = vbBlue
'        Text1.ForeColor = vbWhite
'    Else
'        Text1.Visible = False
'    End If
ElseIf Button.KEY = "h" Then

M_HELP
ElseIf Button.KEY = "pre" Then
    If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    Picture1.Visible = True
    If check("prt", "prt_nme", "frmsalescount") = False Then
    Check3.Value = 0
    Else
    Check3.Value = 1
    End If
    Check3.SetFocus
    Toolbar1.Enabled = False
    SSTab1.Enabled = False
ElseIf Button.Index = 18 Then
    EMP_SHL "STO_SHL"
End If
End Sub

Private Sub M_INV_IDY_KeyDown(K As Integer, s As Integer)
If K = 13 Then

'If Len(Trim(M_INV_IDY)) = 0 Then
'    V_INV_LST
'Else
'    inV = M_INV_IDY
'    clrctr frmSalesCount
'    M_SUP_NME = ""
'    M_SUP_AD1 = ""
'    M_SUP_AD2 = ""
'    M_SUP_AD3 = ""
'    M_SUP_AD4 = ""
'
'    Label18 = "0"
'    Label19 = "0.00"
'    Label20 = "0.00"
'    M_GRS_AMT = "0.00"
'    M_NET_AMT = "0.00"
'    M_FFR_NME = ""
'    M_ACT_NME = ""
'    M_INV_IDY = inV
'    DISPDETAILS
'End If
End If
End Sub

'Public Sub DISPDETAILS()
'On Error GoTo LBL
'If CHECKTB("inl.dbf,smn.dbf,act.dbf,inv.dbf,brd.dbf,prd.dbf,pgr.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'    CANCELRECORD
'    Exit Sub
'End If
'Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
'If REC_SET.State = 1 Then REC_SET.Close
'If Len(Trim(M_INV_IDY)) <> 10 Then
'REC_SET.Open "select * from inl where inv_idy like 'STO" & Tincrement(Trim(M_INV_IDY)) & "' and txn_typ like 'STO'", con, adOpenKeyset
'Else
'REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'STO'", con, adOpenKeyset
'End If
'If REC_SET.RecordCount > 0 Then
'    If Len(Trim(M_INV_IDY)) <> 10 Then
'    M_INV_IDY = "STO" & Tincrement(Trim(M_INV_IDY))
'    End If
'    Check3.Enabled = False
'    M_INV_IDY.Enabled = False
'    MOD_REC = True
'    M_INV_DAT = REC_SET("inv_dat")
'    M_SUP_IDY = Trim(REC_SET("cst_idy") & "")
'    M_SMN_IDY = Trim(REC_SET("smn_idy") & "")
'    dlr_idy = M_SUP_IDY
'    If Len(Trim(M_SUP_IDY)) = 0 Then Exit Sub
'    M_PUR_IDY = Trim(REC_SET("ORD_IDY") & "")
'    M_ACT_IDY = Trim(REC_SET("act_idy") & "")
'    If REC_SET!SCH_CHK = "Y" Then
'        SCH_CHK.Value = 1
'    Else
'        SCH_CHK.Value = 0
'    End If
'    M_CRT_DAY.Locked = True
'    M_DCH_IDY = Trim(REC_SET("dch_idy") & "")
'    M_DCH_DAT = REC_SET("dch_dat")
'    M_EXD_AMT = convert(amt(REC_SET("EXD_AMT") & ""))
'    M_CRT_DAY = Val(REC_SET("crt_day"))
'    M_DUE_DAT = Format(REC_SET!due_dat, "dd/mm/yyyy")
'    M_ROU_OFF = Val(REC_SET("rou_off") & "")
'    If Val(REC_SET("rou_off") & "") = 0 Then M_ROU_OFF = "0.00"
'    M_SPD_PCG = discount(Val(REC_SET("spd_pcg") & ""))
'    M_SPD_AMT = convert(amt(Val(REC_SET("spd_amt") & "")))
'    M_CSD_PCG = discount(Val(REC_SET("csd_pcg") & ""))
'    M_CSD_AMT = convert(amt(Val(REC_SET("csd_amt") & "")))
'    M_TAX_PCG = discount(Val(REC_SET("tax_pcg") & ""))
'    M_TAX_AMT = convert(amt(Val(REC_SET("tax_amt") & "")))
'    M_WBL_IDY = Trim(REC_SET("WBL_IDY") & "")
'    M_FRT_TYP = Trim(REC_SET("FRT_typ") & "")
'    M_FRT_AMT = convert(amt(Val(REC_SET("frt_amt") & "")))
'    M_WBL_IDY = Trim(REC_SET("wbl_idy") & "")
'    M_WBL_DAT = REC_SET("wbl_dat")
'    M_PKG_CST = convert(amt(REC_SET("PKG_CSt") & ""))
'    M_TTL_CAS = Trim(REC_SET("ttl_cas") & "")
'    M_TXN_RMK = Trim(REC_SET("txn_rmk") & "")
'    M_TXN_TPT = Trim(REC_SET("txn_tpt") & "")
'    M_LRN_IDY = Trim(REC_SET("lrn_idy") & "")
'    M_LRN_AMT = convert(amt(REC_SET("lrn_amt") & ""))
'    M_LRN_DAT = Trim(REC_SET("lrn_dat") & "")
'    M_ROU_IDY = Trim(REC_SET("rou_idy") & "")
'    M_AMT_ADD = Trim(REC_SET!txn_rm1) & ""
'    M_AMT_LES = Trim(REC_SET!TXN_RM2) & ""
'    M_EXD_PCG = discount(Val(REC_SET!EXD_PCG & ""))
'    M_EXD_AMT = convert(amt(Val(REC_SET!EXD_AMT & "")))
'    If REC_SET.State = 1 Then REC_SET.Close
'    REC_SET.Open "select * from smn where smn_idy like '" & Trim(M_SMN_IDY) & "'", con, adOpenKeyset
'    If REC_SET.RecordCount > 0 Then M_SMN_NME = REC_SET("smn_nme")
'    If REC_SET.State = 1 Then REC_SET.Close
'    REC_SET.Open "select * from act where act_idy like '" & Trim(M_ACT_IDY) & "'", con, adOpenKeyset
'    If REC_SET.RecordCount > 0 Then M_ACT_NME = REC_SET("act_nme")
'    If REC_SET.State = 1 Then REC_SET.Close
'    REC_SET.Open "select * from ACT where ACT_idy like '" & Trim(M_SUP_IDY) & "'", con, adOpenKeyset
'    If REC_SET.RecordCount > 0 Then
'        M_SUP_NME = Trim(REC_SET("ACT_nme") & "")
'        M_SUP_AD1 = Trim(REC_SET("ACT_ad1") & "")
'        M_SUP_AD2 = Trim(REC_SET("ACT_ad2") & "")
'        M_SUP_AD3 = Trim(REC_SET("ACT_ad3") & "")
'        M_SUP_AD4 = Trim(REC_SET("ACT_ad4") & "")
'        ZIPCODE = Trim(REC_SET("act_zip"))
'        'Label17 = AMT(Trim(REC_SET("dbt_mcl")) & "")
'        Label18 = Trim(REC_SET("dbt_mcp") & "")
'    End If
'    If REC_SET.State = 1 Then REC_SET.Close
'    REC_SET.Open "select * from inv where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ = 'STO'", con, adOpenKeyset, adLockPessimistic
'    MS.Rows = REC_SET.RecordCount + 1
'    For A = 1 To MS.Rows - 1
'        MS.TextMatrix(A, 0) = Trim(REC_SET("prd_idy"))
'        MS.TextMatrix(A, 1) = QTY(Val(REC_SET("prd_qty")))
'        MS.TextMatrix(A, 2) = convert(amt(REC_SET("prd_rat") & ""))
'        MS.TextMatrix(A, 4) = convert(REC_SET("add_AMT") & "")
'        MS.TextMatrix(A, 5) = discount(REC_SET("trd_dis") & "")
'        MS.TextMatrix(A, 6) = discount(REC_SET("trd_tax") & "")
'        MS.TextMatrix(A, 13) = Trim(REC_SET!BAT_IDY)
'        MS.TextMatrix(A, 21) = Trim(REC_SET!stk_typ & "")
'        If Trim(REC_SET!stk_upd) = "N" Then
'            Check3.Value = 0
'        Else
'            Check3.Value = 1
'        End If
'        If REC_SET2.State = 1 Then REC_SET2.Close
'        REC_SET2.Open "SELECT * FROM BRD WHERE BAT_IDY LIKE '" & Trim(REC_SET!BAT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
'        If REC_SET2.RecordCount > 0 Then
'            MS.TextMatrix(A, 8) = convert(Val(REC_SET2!PRD_MRP))
'            MS.TextMatrix(A, 9) = convert(Val(REC_SET2!PRD_PDR))
'            MS.TextMatrix(A, 16) = Format(REC_SET2!bRD_dom, "dd/MM/yyyy")
'            MS.TextMatrix(A, 17) = Format(REC_SET2!brd_doe, "dd/MM/yyyy")
'            MS.TextMatrix(A, 22) = convert(Val(REC_SET2!P_INC_DIS))
'        End If
'        If REC_SET2.State = 1 Then REC_SET2.Close
'        REC_SET2.Open "SELECT prd_nme,CON_FAC,SUB_UNT,pgr.GRP_IDY,GRP_CAT,GRP_CLS, LST_RAT, LST_DIS FROM PRD,pgr WHERE prd.grp_idy Like pgr.grp_idy and PRD_IDY LIKE '" & REC_SET("PRD_IDY") & "'", con, adOpenStatic
'        If REC_SET2.RecordCount > 0 Then
'            MS.TextMatrix(A, 11) = Trim(REC_SET2(1))
'            MS.TextMatrix(A, 10) = Trim(REC_SET2(0))
'            MS.TextMatrix(A, 12) = Val(REC_SET2(2))
'            MS.TextMatrix(A, 18) = Trim(REC_SET2!GRP_IDY)
'            MS.TextMatrix(A, 19) = Trim(REC_SET2!GRP_CAT)
'            MS.TextMatrix(A, 20) = Trim(REC_SET2!GRP_CLS)
'            MS.TextMatrix(A, 24) = amt(Val(REC_SET2!LST_RAT))
'            MS.TextMatrix(A, 25) = amt(Val(REC_SET2!LST_DIS))
'        End If
'        REC_SET.MoveNext
'        If Val(MS.TextMatrix(A, 9)) <> 0 And Val(MS.TextMatrix(A, 2)) <> 0 Then MS.TextMatrix(A, 14) = discount(Round((Val(MS.TextMatrix(A, 9)) / Val(MS.TextMatrix(A, 2)) * 100), 2) - 100)
'        If Val(MS.TextMatrix(A, 6)) <> 0 And Val(MS.TextMatrix(A, 9)) <> 0 And Val(MS.TextMatrix(A, 2)) <> 0 Then
'            MS.TextMatrix(A, 15) = discount(Round((Val(MS.TextMatrix(A, 2)) * Val(MS.TextMatrix(A, 6)) / 100), 2))
'            MS.TextMatrix(A, 15) = discount(Round(((Val(MS.TextMatrix(A, 9)) / (Val(MS.TextMatrix(A, 15)) + Val(MS.TextMatrix(A, 2)))) * 100) - 100, 2))
'        ElseIf Val(MS.TextMatrix(A, 9)) <> 0 And Val(MS.TextMatrix(A, 2)) <> 0 Then
'            MS.TextMatrix(A, 15) = discount(Round(((Val(MS.TextMatrix(A, 9)) / (Val(MS.TextMatrix(A, 15)) + Val(MS.TextMatrix(A, 2)))) * 100) - 100, 2))
'        End If
'    Next
''    P_DIS_CAL
'    GRS_CAL
'    M_GRS_AMT = M_PRD_TOT
'    SSTab1.Tab = 0
'    M_PUR_IDY.SetFocus
'Else
'    V_INV_LST
'End If
'    If MS.Rows = 1 Then MS.Rows = 2
'Exit Sub
'LBL:
'If err.Number = 35773 Then
'    MsgBox "Invalid Cycle Count Voucher Selected for the Financial Year !", vbInformation, "DAS Version 4.0"
'    CANCELRECORD
'Else
'    MsgBox err.Description
'    CANCELRECORD
'End If
'End Sub



Private Sub list2_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list2.ListItems.count > 0 Then
        M_ACT_IDY = list2.SelectedItem.ListSubItems.Item(1)
        M_ACT_NME = list2.SelectedItem
        SSTab1.Tab = 0
        D_ACT_IDY.SetFocus
        list2.Visible = False
    End If
End If
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list3.ListItems.count > 0 Then
        M_SUP_IDY = list3.SelectedItem.ListSubItems(1)
        M_SUP_IDY_keydown 13, 0
        SSTab1.Tab = 0
        list3.Visible = False
        M_ACT_IDY.SetFocus
    End If
End If
End Sub

Private Sub list4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list4.ListItems.count > 0 Then
        D_ACT_IDY = list4.SelectedItem.ListSubItems.Item(1)
        D_ACT_NME = list4.SelectedItem
        SSTab1.Tab = 1
        MSF.SetFocus
        list4.Visible = False
    End If
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
'    M_PRD_TOT.Visible = True
'    If M_INV_IDY.Enabled = True Then
'        M_INV_IDY.SetFocus
'    Else
        M_INV_DAT.SetFocus
'    End If
ElseIf SSTab1.Tab = 1 Then
    MSF.Col = 0
    MSF.Row = 1
 '   M_PRD_TOT.Visible = True
    MSF.SetFocus
ElseIf SSTab1.Tab = 2 Then
  '  M_PRD_TOT.Visible = True
    'M_CSD_PCG.SetFocus
End If
If list1.Visible = True Then list1.Visible = False
If list2.Visible = True Then list2.Visible = False
If list3.Visible = True Then list3.Visible = False
If list4.Visible = True Then list4.Visible = False
'If list5.Visible = True Then list5.Visible = False

If SSTab1.Tab = 3 Then
    M_LRN_IDY.SetFocus
End If
If SSTab1.Tab = 2 Then
'Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTab1.Tab = 1 Then MSF.SetFocus
    If SSTab1.Tab = 0 Then M_INV_DAT.SetFocus
End If
End Sub

Public Sub P_DIS_PRD(PRDID As String, BATID As String)
If CHECKTB("prd.dbf,brd.dbf,pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
MSROW = MS.Row
Dim REC_SET As New ADODB.Recordset
For A = 1 To MS.Rows - 1
    If Trim(BATID) = MS.TextMatrix(A, 13) And "Salable" = MS.TextMatrix(A, 21) Then
        SSTab1.Tab = 1
        list4.Visible = False
        MsgBox "Product Already Exists In Voucher !", vbExclamation, "DAS Version 4.0"
        Text3.Visible = False
        MS.SetFocus
        MS.Col = 0
        MS.Row = A
        Exit Sub
    End If
Next
REC_SET.Open "select * from PRD,BRD where PRD.PRD_IDY LIKE BRD.PRD_IDY AND BRD.BAT_idy like '" & Trim(BATID) & "'", con, adOpenKeyset
If REC_SET.RecordCount > 0 Then
    list4.Visible = False
    MS.TextMatrix(MS.Row, 0) = REC_SET("PRD_IDY")
    MS.TextMatrix(MS.Row, 1) = "0.000"
    MS.TextMatrix(MS.Row, 2) = convert(REC_SET("PRD_PDR"))
    MS.TextMatrix(MS.Row, 3) = "0.000"
    MS.TextMatrix(MS.Row, 4) = "0.00"
    MS.TextMatrix(MS.Row, 5) = "0.00"
    MS.TextMatrix(MS.Row, 6) = "0.00"
    MS.TextMatrix(MS.Row, 7) = "0.00"
    MS.TextMatrix(MS.Row, 8) = convert(REC_SET("PRD_MRP"))
    MS.TextMatrix(MS.Row, 9) = convert(REC_SET("PRD_PDR"))
    MS.TextMatrix(MS.Row, 10) = Trim(REC_SET("PRD_NME"))
    MS.TextMatrix(MS.Row, 11) = Trim(REC_SET("CON_FAC"))
    MS.TextMatrix(MS.Row, 12) = Trim(REC_SET("SUB_UNT"))
    MS.TextMatrix(MS.Row, 13) = Trim(REC_SET("BAT_IDY"))
    MS.TextMatrix(MS.Row, 16) = Format(REC_SET("Brd_dom"), "dd/MM/yyyy")
    MS.TextMatrix(MS.Row, 17) = Format(REC_SET("Brd_doe"), "dd/MM/yyyy")
    MS.TextMatrix(MS.Row, 21) = "Salable"
    MS.TextMatrix(MS.Row, 22) = convert(REC_SET("P_INC_DIS"))
    MS.TextMatrix(MS.Row, 24) = convert(REC_SET("LST_RAT"))
    MS.TextMatrix(MS.Row, 25) = convert(REC_SET("LST_DIS"))
    If Val(MS.TextMatrix(MS.Row, 9)) > 0 And Val(MS.TextMatrix(MS.Row, 2)) > 0 Then MS.TextMatrix(MS.Row, 14) = discount(Round((Val(MS.TextMatrix(MS.Row, 9)) / Val(MS.TextMatrix(MS.Row, 2)) * 100), 2) - 100)
    If Val(MS.TextMatrix(MS.Row, 6)) <> 0 And Val(MS.TextMatrix(MS.Row, 9)) > 0 And Val(MS.TextMatrix(MS.Row, 2)) > 0 Then
        MS.TextMatrix(MS.Row, 15) = discount(Round((Val(MS.TextMatrix(MS.Row, 2)) * Val(MS.TextMatrix(MS.Row, 6)) / 100), 2))
        MS.TextMatrix(MS.Row, 15) = discount(Round(((Val(MS.TextMatrix(MS.Row, 9)) / (Val(MS.TextMatrix(MS.Row, 15)) + Val(MS.TextMatrix(MS.Row, 2)))) * 100) - 100, 2))
    ElseIf Val(MS.TextMatrix(MS.Row, 9)) <> 0 And Val(MS.TextMatrix(MS.Row, 2)) <> 0 Then
        MS.TextMatrix(MS.Row, 15) = discount(Round(((Val(MS.TextMatrix(MS.Row, 9)) / (Val(MS.TextMatrix(MS.Row, 15)) + Val(MS.TextMatrix(MS.Row, 2)))) * 100) - 100, 2))
    End If
    MS.TextMatrix(MS.Row, 18) = Trim(REC_SET!GRP_IDY)
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from pgr where grp_idy like '" & Trim(MS.TextMatrix(MS.Row, 18)) & "'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        MS.TextMatrix(MS.Row, 19) = Trim(REC_SET!GRP_CAT)
        MS.TextMatrix(MS.Row, 20) = Trim(REC_SET!GRP_CLS)
    End If
    Text3.Visible = False
    SSTab1.Tab = 1
    MS.Row = MSROW
    MS.Col = 1
    MS.SetFocus
End If
End Sub

Public Sub V_PRD_LST()
If CHECKTB("prt.dbf,prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("prt", "prt_nme", "p_prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "SELECT * FROM prd,BRD WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD_NME LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') and BRD.PRD_QTY>0 ORDER BY BRW_SEQ", con, adOpenKeyset
Else
    REC_SET.Open "SELECT * FROM brd,prd WHERE BRD.PRD_IDY LIKE PRD.PRD_IDY AND PRD_PRN LIKE '" & Trim(Text3) & "%' and prd.stk_typ in('Salable','Free','Promotional Item') and BRD.PRD_QTY>0 ORDER BY BRW_SEQ", con, adOpenKeyset
End If
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , convert(REC_SET!PRD_MRP)
        ls.ListSubItems.Add , , convert(REC_SET!PRD_PDR)
        If Val(REC_SET("PRD_QTY") & "") <> 0 Then
            ls.ListSubItems.Add , , Trim(QTY(REC_SET("PRD_QTY") & ""))
        Else
            ls.ListSubItems.Add , , "0.000"
        End If
        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("prd_idy") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_IDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BAT_NME") & "")
        ls.ListSubItems.Add , , convert(REC_SET("PRD_PDR") & "")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_dom"), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Format(REC_SET("Brd_doe"), "dd/MM/yyyy")
        REC_SET.MoveNext
    Next
    Label49 = "Select Product"
    SSTab1.Tab = 2
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product Not Found, Define Product !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 1
    If Text3.Visible = True Then
        Text3.SetFocus
    Else
        Check2.Caption = "Place"
        MS.SetFocus
    End If
    End If
End If
End Sub
Public Sub Printrecord()
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "frmsalescount") = True Then
    M_FRM_PRV.Check3.Value = 1
Else
    M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "STO"
M_FRM_PRV.Show 1
End Sub

Public Sub setval()
CST = ""
GST = ""
clrctr frmSalesCount
'M_ROU_IDY.Selected(0) = True
M_SUP_NME = ""
M_SUP_AD1 = ""
M_SUP_AD2 = ""
M_SUP_AD3 = ""
M_SUP_AD4 = ""
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
Tax = "0.00"
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
M_LST_RAT = "0.00"
M_LST_DIS = "0.00"
M_CRT_DAY = "0"
End Sub

Public Sub ROU_OFF()
If CHECKTB("inl.dbf,dir.dbf") = 1 Then
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
REC_SET.Open "select rou_off from inl where inv_idy like '" & Trim(M_INV_IDY) & "' AND TXN_TYP LIKE 'STO'", con, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then If (Val(REC_SET(0) & "") = 0 Or Val(M_ROU_OFF) = 0) And MOD_REC = True Then Exit Sub
dot = 0
rof = 0
dot = InStr(Trim(T_FRT_AMT), ".")
If dot > 0 Then
    rof = (Mid(Trim(T_FRT_AMT), dot + 1, 2))
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
        T_NET_AMT = Mid(Val(T_FRT_AMT), 1, dot - 1)
        T_NET_AMT = Val(M_NET_AMT) + 1 & ".00"
    Else
        M_ROU_OFF = "0.00"
    End If
Else
    M_ROU_OFF = "0.00"
    T_NET_AMT = convert(amt(Val(M_NET_AMT)))
End If
End Sub

Public Sub createbatch()
If MsgBox("Create New Price n Batch Record?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    Me.Tag = "Batch"
    If CHECKTB("dir.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If check("dir", "dir_msg", "Check Expiry") = True Then M_FRM_NBC.Width = 7530
    M_FRM_NBC.Tag = "STO"
    M_FRM_NBC.Show 1
    MS.SetFocus
Else
    MS.SetFocus
End If
End Sub

Private Sub InitializeGrid()
    With MSF
        .ColWidth(0) = 3000
        .ColWidth(1) = 1500
        .ColWidth(2) = 1000
        .ColWidth(3) = 900
        .ColWidth(4) = 900
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000

        .TextMatrix(0, 0) = "Product Name"
        .TextMatrix(0, 1) = "Batch Name"
        .TextMatrix(0, 2) = "Cur.Stk"
        .ColAlignment(2) = 7
        .TextMatrix(0, 3) = "+ Qty"
        SCl MSF, 3
        .ColAlignment(3) = 7
        .TextMatrix(0, 4) = "- Qty"
        SCl MSF, 4
        .ColAlignment(4) = 7
        .TextMatrix(0, 5) = "Balance"
        .ColAlignment(5) = 7
        .TextMatrix(0, 6) = "+  Amount"
        .ColAlignment(6) = 7
        .TextMatrix(0, 7) = "-  Amount"
        .ColAlignment(7) = 7
        .TextMatrix(0, 8) = "Rate"
        .ColAlignment(8) = 7
        .TextMatrix(0, 9) = "Dis %"
        .ColAlignment(9) = 7
        .TextMatrix(0, 10) = "Tax %"
        .ColAlignment(10) = 7
        .TextMatrix(0, 11) = "SubUnit"
        .ColAlignment(11) = 7
        .TextMatrix(0, 12) = "UOM"
        .ColAlignment(12) = 7
        .TextMatrix(0, 13) = "Product ID"
        .TextMatrix(0, 14) = "Batch"
        .TextMatrix(0, 15) = "+ Dis"
        .ColAlignment(15) = 7
        .TextMatrix(0, 16) = "+ Tax"
        .ColAlignment(16) = 7
        .TextMatrix(0, 17) = "- Dis"
        .ColAlignment(17) = 7
        .TextMatrix(0, 18) = "- Tax"
        .ColAlignment(18) = 7
        .TextMatrix(0, 19) = "+ Amount"
        .ColAlignment(19) = 7
        .TextMatrix(0, 20) = "- Amount"
        .ColAlignment(20) = 7
        .TextMatrix(0, 21) = "StockType"
        .TextMatrix(0, 22) = "List Rate"
        .TextMatrix(0, 23) = "List Dis"
        .TextMatrix(0, 24) = "List Tax"
        .ColWidth(21) = 0
        .ColWidth(22) = 0
        .ColWidth(23) = 0
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .ColWidth(27) = 0
        .ColWidth(28) = 0
        .ColWidth(29) = 0
        .ColWidth(30) = 0
        
    End With
    MSF.Rows = 2
End Sub

Private Sub FillGrid()
InitializeGrid
Dim rec As New ADODB.Recordset
rec.Open "select prd.prd_idy , PGR.GRP_IDY, PGR.GRP_CAT, PGR.GRP_CLS, brd.prd_pdr,prd.prd_uom, prd.pur_dis, prd.pur_tax, prd.prd_nme,prd.sub_unt,BRD.PRD_SDR,  prd.cur_lev, brd.bat_nme, bat_idy, prd.lst_rat,prd.lst_dis, prd.lst_cst,prd.stk_typ from prd, brd, PGR where TRIM(PGR.GRP_IDY) = TRIM(PRD.GRP_IDY) AND    trim(prd.prd_idy) = trim(brd.prd_idy)", con, adOpenKeyset, adLockPessimistic

While Not rec.EOF
    With MSF
        .TextMatrix(MSF.Rows - 1, 0) = Trim(rec!PRD_NME)
        .TextMatrix(MSF.Rows - 1, 1) = Trim(rec!Bat_nme)
        .TextMatrix(MSF.Rows - 1, 2) = Format(QTY(Val(Trim(rec!CUR_LEV))), "###############.000")
        .TextMatrix(MSF.Rows - 1, 8) = Trim(rec!PRD_PDR)
        .TextMatrix(MSF.Rows - 1, 9) = Format(Trim(rec!PUR_DIS), "###############.00")
        .TextMatrix(MSF.Rows - 1, 10) = Format(Trim(rec!PUR_TAX), "###############.00")
        .TextMatrix(MSF.Rows - 1, 11) = Trim(rec!SUB_UNT)
        .TextMatrix(MSF.Rows - 1, 12) = Trim(rec!PRD_UOM)
        .TextMatrix(MSF.Rows - 1, 13) = Trim(rec!PRD_IDY)
        .TextMatrix(MSF.Rows - 1, 14) = Trim(rec!BAT_IDY)
        .TextMatrix(MSF.Rows - 1, 21) = Trim(rec!stk_typ)
        .TextMatrix(MSF.Rows - 1, 22) = Format(IIf(IsNull(rec!LST_RAT) = True, "", Trim(rec!LST_RAT)), "###############.00")
        .TextMatrix(MSF.Rows - 1, 23) = Format(IIf(IsNull(rec!LST_DIS) = True, "", Trim(rec!LST_DIS)), "###############.00")
        .TextMatrix(MSF.Rows - 1, 24) = Format(IIf(IsNull(rec!lst_cst) = True, "", Trim(rec!lst_cst)), "###############.00")
        .TextMatrix(MSF.Rows - 1, 27) = Trim(rec!GRP_IDY)
        .TextMatrix(MSF.Rows - 1, 28) = Trim(rec!GRP_CAT)
        .TextMatrix(MSF.Rows - 1, 29) = Trim(rec!GRP_CLS)
        .TextMatrix(MSF.Rows - 1, 30) = Trim(rec!PRD_SDR)
        
    End With
    rec.MoveNext
    If Not rec.EOF Then MSF.Rows = MSF.Rows + 1
Wend
End Sub

Public Sub CalculateValues()
Dim discount As Double
Dim Tax As Double
Dim amount As Double
Dim GrossAmount As Double
lblMinusGross.Caption = "0"
lblPlusGross.Caption = "0"

For i = 1 To MSF.Rows - 1
    MSF.TextMatrix(i, 5) = QTY(Val(Val(MSF.TextMatrix(i, 2)) + Val(MSF.TextMatrix(i, 3)) - Val(MSF.TextMatrix(i, 4))))
    
    
    'plus amount
    discount = 0
    Tax = 0
    amount = 0
    amount = Val(MSF.TextMatrix(i, 8)) * Val(MSF.TextMatrix(i, 3)) '/ (Val(MSF.TextMatrix(i, 11)))
    dis = amount * Val(MSF.TextMatrix(i, 9)) / 100
    Tax = (amount - dis) * Val(MSF.TextMatrix(i, 10)) / 100
    GrossAmount = amount - dis + Tax
    'MSF.TextMatrix(i, 12) = Val(MSF.TextMatrix(i, 8)) * Val(MSF.TextMatrix(i, 11))
    MSF.TextMatrix(i, 6) = Format(GrossAmount, "###############.00")
    MSF.TextMatrix(i, 15) = Format(discount, "###############.00")
    MSF.TextMatrix(i, 16) = Format(Tax, "###############.00")
    MSF.TextMatrix(i, 19) = Format(amount, "###############.00")
    lblPlusGross.Caption = Format(Val(lblPlusGross.Caption) + GrossAmount, "###############.00")
    
    'minus amount
    discount = 0
    Tax = 0
    amount = 0
    amount = Val(MSF.TextMatrix(i, 8)) * Val(MSF.TextMatrix(i, 4)) '/ (Val(MSF.TextMatrix(i, 11)))
    dis = amount * Val(MSF.TextMatrix(i, 9)) / 100
    Tax = (amount - dis) * Val(MSF.TextMatrix(i, 10)) / 100
    GrossAmount = amount - dis + Tax
    MSF.TextMatrix(i, 7) = Format(GrossAmount, "###############.00")
    MSF.TextMatrix(i, 17) = Format(discount, "###############.00")
    MSF.TextMatrix(i, 18) = Format(Tax, "###############.00")
    MSF.TextMatrix(i, 20) = Format(amount, "###############.00")
    lblMinusGross.Caption = Format(Val(lblMinusGross.Caption) + GrossAmount, "###############.00")
Next
End Sub


