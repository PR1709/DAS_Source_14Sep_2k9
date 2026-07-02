VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_INC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salesman Incentive"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_INC1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "&Salesman Details"
      TabPicture(0)   =   "M_FRM_INC1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "smnname"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MS"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tdate"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fdate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtQuan"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "range"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "smnid"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txnid"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "All"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtIncAmt"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "inctyp"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtvar"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_INC1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lv"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   665
         Left            =   3670
         TabIndex        =   23
         Top             =   2280
         Visible         =   0   'False
         Width           =   1350
         Begin VB.TextBox TXTINCAMT1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   20
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   " "
            Top             =   330
            Visible         =   0   'False
            Width           =   1300
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   1360
            Y1              =   640
            Y2              =   640
         End
         Begin VB.Line Line4 
            X1              =   1330
            X2              =   1330
            Y1              =   0
            Y2              =   690
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   1360
            Y1              =   320
            Y2              =   320
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   690
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Incentive Amount"
            Height          =   170
            Left            =   30
            TabIndex        =   24
            Top             =   40
            Width           =   1245
         End
      End
      Begin VB.TextBox txtvar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox inctyp 
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
         ItemData        =   "M_FRM_INC1.frx":047A
         Left            =   3840
         List            =   "M_FRM_INC1.frx":0487
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtIncAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox All 
         Caption         =   "&All Products"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txnid 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox smnid 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox range 
         Caption         =   "&Range"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtQuan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker fdate 
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   37869
         MinDate         =   -106558
      End
      Begin MSComctlLib.ListView lv 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   10
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Category Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComCtl2.DTPicker tdate 
         Height          =   330
         Left            =   3240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   37869
         MaxDate         =   73050
         MinDate         =   37712
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2835
         Left            =   75
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2265
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   5001
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         GridLinesFixed  =   1
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
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   3650
         X2              =   3650
         Y1              =   2280
         Y2              =   2880
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive Type"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2640
         TabIndex        =   21
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Incentive Amount"
         Height          =   195
         Left            =   3720
         TabIndex        =   20
         Top             =   1965
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   195
         Left            =   1560
         TabIndex        =   19
         Top             =   1965
         Width           =   585
      End
      Begin VB.Label smnname 
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
         Left            =   2640
         TabIndex        =   18
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1140
         TabIndex        =   17
         Top             =   555
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   1515
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   960
         TabIndex        =   15
         Top             =   1515
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Salesman"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1028
         Width           =   1185
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Salesman"
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
         TabIndex        =   13
         Top             =   720
         Width           =   7900
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   600
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
            Picture         =   "M_FRM_INC1.frx":04A4
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":09E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":0DB0
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":1104
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":173C
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":1A90
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":1EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":21E4
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":25F0
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":26FC
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":2A50
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":2E64
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   22
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACT_LST"
                  Text            =   "&1 All Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_CAT"
                  Text            =   "&2 Retailer Category"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_LOC"
                  Text            =   "&3 Retailer Location"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_MAR"
                  Text            =   "&4 Retailer Market"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ROU"
                  Text            =   "&5 Retailer Beat"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_SMN"
                  Text            =   "&6 Retailer Salesman"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SUP_LST"
                  Text            =   "&7 Supplier List"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ID_LST"
                  Text            =   "&8 Retailer ID List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
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
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Account ID "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1200
      Top             =   360
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
            Picture         =   "M_FRM_INC1.frx":33B0
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":38F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":3CBC
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":4010
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":4648
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":499C
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":4DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":50F0
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":54FC
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":5608
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":595C
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_INC1.frx":5D70
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_INC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssmn As New ADODB.Recordset
Dim RSINC As New ADODB.Recordset
Dim RSPRD As New ADODB.Recordset
Dim MD As Boolean, FL As Boolean
Private Sub All_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If All.Value = 1 Then
        txtQuan.Visible = True
        txtQuan.Enabled = True
        txtQuan.SetFocus
        txtQuan = "0.000"
        txtIncAmt.Enabled = True
        txtIncAmt = "0.00"
        SendKeys "{home}+{end}"
    Else
        dispprd
        MS.SetFocus
        MS.Col = 1
    End If
End If
End Sub

Private Sub fdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    tdate.SetFocus
End If
End Sub

Private Sub Form_Activate()
txnid.SetFocus
inctyp.ListIndex = 0
MS.Rows = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 78 And Shift = 2 Then
    NewRecord
ElseIf KeyCode = 83 And Shift = 2 Then
    saverecord
ElseIf KeyCode = 81 And Shift = 2 Then
    cancel
ElseIf KeyCode = 69 And Shift = 2 Then
    If Me.ActiveControl.name = "txnid" Then
        inclist
    ElseIf Me.ActiveControl.name = "smnid" Then
        smnlist
    End If
ElseIf KeyCode = 73 And Shift = 2 Then
    TXTINCAMT1.SetFocus
    TXTINCAMT1 = "0.00"
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "txnid" And Len(Trim(txnid)) = 0 Then
        Unload Me
    
    ElseIf SSTab1.Tab = 1 Then
        If lv.Visible = True And UCase(Trim(Label20)) = UCase("Select Transaction") Then
            lv.Visible = False
            SSTab1.Tab = 0
            txnid.SetFocus
        ElseIf lv.Visible = True And UCase(Trim(Label20)) = UCase("Select Salesman") Then
            lv.Visible = False
            SSTab1.Tab = 0
            smnid.SetFocus
        End If
    Else
        saverecord
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAR Me
End Sub

Private Sub inctyp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(inctyp.Text) = "Quantity" Then
        txtQuan = "0.000"
        Label4.Caption = "Quantity"
    ElseIf Trim(inctyp.Text) = "Volume" Then
        Label4.Caption = "Volume"
        txtQuan = "0.000"
    Else
        Label4.Caption = "Value"
        txtQuan = "0"
    End If
    smnid.SetFocus
End If
End Sub

Private Sub lv_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If lv.SortOrder = lvwAscending Then
lv.SortOrder = lvwDescending
ElseIf lv.SortOrder = lvwDescending Then
lv.SortOrder = lvwAscending
End If
lv.SortKey = CH.index - 1
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 13 Then
    If lv.ListItems.Count > 0 Then
        If rs.State = 1 Then rs.Close
        If Trim(Label20.Caption) = "Select Salesman" Then
            Dim lst As ListSubItem
            Set lst = lv.SelectedItem.ListSubItems.Item(1)
            rs.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", CON, adOpenStatic
            smnid = Trim(rs("smn_idy"))
            smnname = Trim(rs("smn_nme"))
            lv.Visible = False
            SSTab1.Tab = 0
            fdate.SetFocus
            smnid.Enabled = False
        Else
            Dim ls As ListItem
            Set ls = lv.SelectedItem
            rs.Open "select * from inc where txn_idy like '" & Trim(ls.Text) & "'", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                MS.Visible = True
                txtvar.Visible = False
                txnid = rs("TXN_IDY")
                If rssmn.State = 1 Then rssmn.Close
                rssmn.Open "select smn_nme from smn where smN_idy ='" & Trim(rs("smn_idy")) & "'", CON, adOpenKeyset, adLockOptimistic
                If Not rssmn.EOF Then
                    smnid = rs("smn_idy")
                    smnname.Caption = rssmn(0)
                End If
                fdate.Value = rs("frm_dat")
                tdate.Value = rs("to_dat")
                If Trim(rs("inc_typ")) = "Quantity" Then
                    inctyp.ListIndex = 0
                ElseIf Trim(rs("inc_typ")) = "Volume" Then
                    inctyp.ListIndex = 1
                Else
                    inctyp.ListIndex = 2
                End If
                If rs("CHK_FLG") = "A" Then
                    All.Enabled = True
                    All.Value = 1
                    txtIncAmt = convert(rs("inc_AMT"))
                    If Trim(inctyp.Text) = "Quantity" Or Trim(inctyp) = "Volume" Then
                        txtQuan = QTY(Trim(rs("prd_qty")))
                    Else
                        txtQuan = convert(rs("prd_qty"))
                    End If
                    inctyp.Enabled = False
                    txnid.Enabled = False
                    'fdate.SetFocus
                    All.Enabled = False
                Else
                    All.Enabled = False
                    R = 1
                    MS.Rows = 2
                    MS.Cols = 5
                    MS.TextMatrix(0, 0) = "Product Name"
                    MS.TextMatrix(0, 1) = "Quantity"
                    MS.TextMatrix(0, 2) = "Incentive Amt"
                    MS.ColWidth(2) = 0
                    MS.ColWidth(3) = 0
                    MS.ColWidth(4) = 0
                    MS.ColWidth(0) = 2500
                    MS.ColWidth(1) = 1100
                    MS.ColAlignment(2) = 7
                    While Not rs.EOF
                        If RSPRD.State = 1 Then RSPRD.Close
                        RSPRD.Open "select prd_nme,SUB_UNT from prd wher prd_idy='" & rs("prd_idy") & "'", CON, adOpenKeyset, adLockOptimistic
                        MS.TextMatrix(R, 0) = RSPRD(0)
                        MS.TextMatrix(R, 4) = RSPRD(1)
                        MS.TextMatrix(R, 3) = rs("prd_idy")
                        If Trim(inctyp.Text) = "Quantity" Or Trim(inctyp) = "Volume" Then
                            MS.TextMatrix(R, 1) = QTY(rs("prd_qty"))
                        Else
                            MS.TextMatrix(R, 1) = convert(rs("prd_qty"))
                        End If
                        Frame1.Visible = True
                        TXTINCAMT1.Visible = True
                        TXTINCAMT1 = convert(Trim(rs("inc_amt")) & "")
                        R = R + 1
                        MS.Rows = MS.Rows + 1
                        rs.MoveNext
                    Wend
                    MS.Rows = MS.Rows - 1
                    inctyp.Enabled = False
                    'smnid.Enabled = False
                    MS.Visible = True
                    MS.Enabled = True
                    MS.SetFocus
                End If
                smnid.SetFocus
                SendKeys "{HOME}+{END}"
                lv.Visible = False
                SSTab1.Tab = 0
                MD = True
            End If
        End If
    End If
End If
End Sub
Private Sub MS_KeyPress(KeyAscii As Integer)
If MS.Col = 1 Then
If KeyAscii = 13 And MS.Row = MS.Rows - 1 Then
    TXTINCAMT1.SetFocus
    TXTINCAMT1 = "0.00"
    SendKeys "{HOME}+{END}"
Else
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Then
    txtvar.Visible = True
    txtvar.SetFocus
    txtvar.Left = MS.Left + MS.CellLeft
    txtvar.Top = MS.Top + MS.CellTop
    txtvar.Height = MS.CellHeight
    txtvar.Width = MS.CellWidth
    txtvar = Chr(KeyAscii)
    txtvar.SelStart = Len(Trim(txtvar))
End If
End If
End If
'If MS.Col = 2 Then
'    If MD = True Then
'        If Val(MS.TextMatrix(MS.Row, MS.Col)) > 0 Then
'            txtvar.Visible = True
'            txtvar.SetFocus
'            txtvar.Left = MS.Left + MS.CellLeft
'            txtvar.Top = MS.Top + MS.CellTop
'            txtvar.Height = MS.CellHeight
'            txtvar.Width = MS.CellWidth
'            txtvar = Chr(KeyAscii)
'            txtvar.SelStart = Len(Trim(txtvar))
'        End If
'    ElseIf MD = False Then
'        If FL = True Then
'            If Val(MS.TextMatrix(MS.Row, MS.Col)) > 0 Then
'                If KeyAscii >= 48 And KeyAscii <= 57 Then
'                txtvar.Visible = True
'                txtvar.SetFocus
'                txtvar.Left = MS.Left + MS.CellLeft
'                txtvar.Top = MS.Top + MS.CellTop
'                txtvar.Height = MS.CellHeight
'                txtvar.Width = MS.CellWidth
'                txtvar = Chr(KeyAscii)
'                txtvar.SelStart = Len(Trim(txtvar))
'                End If
'            End If
'        Else
'            If KeyAscii >= 48 And KeyAscii <= 57 Then
'                txtvar.Visible = True
'                txtvar.SetFocus
'                txtvar.Left = MS.Left + MS.CellLeft
'                txtvar.Top = MS.Top + MS.CellTop
'                txtvar.Height = MS.CellHeight
'                txtvar.Width = MS.CellWidth
'                txtvar = Chr(KeyAscii)
'                txtvar.SelStart = Len(Trim(txtvar))
'            End If
'    End If
'End If
'End If
End Sub

Private Sub smnid_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from smn where smn_idy like '" & Sincrement(smnid) & "'", CON, adOpenStatic
    If rs.RecordCount > 0 Then
        smnid = Sincrement(smnid)
        smnname = Trim(rs("smn_nme"))
        smnid.Enabled = False
        fdate.SetFocus
    Else
        smnlist
    End If
End If
End Sub

Private Sub tdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If tdate.Value < fdate.Value Then
        MsgBox "Invalid Date", vbExclamation
        tdate.SetFocus
        Exit Sub
    Else
        If All.Enabled = True Then
            All.SetFocus
        ElseIf All.Value = 1 Then
            txtQuan.Enabled = True
            txtQuan.SetFocus
            SendKeys "{HOME}+{END}"
        Else
            txtQuan.Enabled = False
            txtIncAmt.Enabled = False
            MS.SetFocus
            MS.Col = 1
        End If
    End If
End If
End Sub


Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from smn where smn_NME like '" & smnid & "%'", CON, adOpenStatic
Label20.Caption = "Select Salesman"
lv.ColumnHeaders.Clear
lv.ColumnHeaders.Add , , "Salesman Name", 2500
lv.ColumnHeaders.Add , , "ID", 1500
lv.ListItems.Clear
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    lv.Visible = True
    lv.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        smnid.Enabled = True
        smnid.SetFocus
    End If
End If
End Sub

Private Sub NewRecord()
MD = False
cancel
FL = False
inctyp.Enabled = True
'txtQuan.Visible = True
inctyp.SetFocus
txtQuan.Locked = False
MD = False
End Sub

Private Sub cancel()
clrctr Me
txnid = ""
smnid.Enabled = True
smnid = ""
smnname.Caption = ""
fdate.Value = Date
tdate.Value = Date
All.Enabled = True
All.Value = 0
txtQuan = "0.000"
txtIncAmt = ""
MS.Clear
txtvar = ""
txtvar.Visible = False
MS.Rows = 1
inctyp.Selected(0) = True
txnid.Enabled = True
If SSTab1.Tab = 1 Then
    lv.Visible = False
    SSTab1.Tab = 0
End If
Frame1.Visible = False
'Label8.Visible = False
'TXTINCAMT1.Visible = False
txnid.SetFocus
End Sub

Private Sub lv_DblClick()
lv_KeyDown 13, 0
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    NewRecord
ElseIf Button.KEY = "s" Then
    saverecord
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "p" Then

ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If Me.ActiveControl.name = "smnid" Then
        smnlist
    ElseIf Me.ActiveControl.name = "txnid" Then
        inclist
    End If
End If
End Sub

Private Sub txnid_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("inc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from Inc where txn_idy like '" & Sincrement(txnid) & "'", CON, adOpenStatic
    If rs.RecordCount > 0 Then
        'If Trim(rs("INC_TYP")) = "Quantity" Then
            txnid = Sincrement(txnid)
            If rssmn.State = 1 Then rssmn.Close
            rssmn.Open "select smn_nme from smn where smN_idy ='" & Trim(rs("smn_idy")) & "'", CON, adOpenKeyset, adLockOptimistic
            If Not rssmn.EOF Then
                smnid = rs("smn_idy")
                smnname.Caption = rssmn(0)
            End If
            fdate.Value = rs("frm_dat")
            tdate.Value = rs("to_dat")
            If Trim(rs("inc_typ")) = "Quantity" Then
                inctyp.ListIndex = 0
            ElseIf Trim(rs("inc_typ")) = "Volume" Then
                inctyp.ListIndex = 1
            Else
                inctyp.ListIndex = 2
            End If
            If rs("CHK_FLG") = "A" Then
                All.Enabled = True
                All.Value = 1
                txtIncAmt.Visible = True
                txtIncAmt = convert(Trim(rs("inc_amt")))
                If Trim(inctyp.Text) = "Quantity" Or Trim(inctyp) = "Volume" Then
                    txtQuan = QTY(Trim(rs("Prd_qty")))
                Else
                    txtQuan = convert(Trim(rs("prd_qty")))
                End If
                inctyp.Enabled = False
                txnid.Enabled = False
'                fdate.SetFocus
                All.Enabled = False
            Else
                All.Enabled = False
                R = 1
                MS.Rows = 1
                MS.Cols = 5
                R = 1
                MS.ColWidth(3) = 0
                MS.ColWidth(4) = 0
                MS.ColWidth(0) = 2500
                MS.ColWidth(1) = 1100
                MS.ColWidth(2) = 0
                MS.TextMatrix(0, 0) = "Product Name"
                MS.TextMatrix(0, 1) = "Quantity"
                MS.TextMatrix(0, 2) = "Incentive Amt"
                'MS.ColWidth(2) = 0
                'MS.ColAlignment(2) = 7
                While Not rs.EOF
                    MS.Rows = MS.Rows + 1
                    If RSPRD.State = 1 Then RSPRD.Close
                    RSPRD.Open "select prd_nme,SUB_UNT from prd wher prd_idy='" & rs("prd_idy") & "'", CON, adOpenKeyset, adLockOptimistic
                    MS.TextMatrix(R, 0) = RSPRD(0)
                    MS.TextMatrix(R, 4) = RSPRD(1)
                    MS.TextMatrix(R, 3) = rs("prd_idy")
                    If Trim(rs("INC_TYP")) = "Quantity" Then
                        inctyp.ListIndex = 0
                    ElseIf Trim(rs("inc_typ")) = "Volume" Then
                        inctyp.ListIndex = 1
                    Else
                        inctyp.ListIndex = 2
                    End If
                    If Trim(inctyp.Text) = "Quantity" Or Trim(inctyp) = "Volume" Then
                        MS.TextMatrix(R, 1) = QTY(rs("prd_qty"))
                    Else
                        MS.TextMatrix(R, 1) = convert(rs("prd_qty"))
                    End If
                    Frame1.Visible = True
                    TXTINCAMT1.Visible = True
                    TXTINCAMT1 = convert(Trim(rs("inc_amt")) & "")
                    R = R + 1
                    rs.MoveNext
                Wend
                'smnid.Enabled = False
                MS.SetFocus
            End If
            smnid.SetFocus
            SendKeys "{HOME}+{END}"
            MD = True
        'Else
            
        'End If
    Else
        inclist
    End If
End If
End Sub
Private Sub inclist()
If CHECKTB("inc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from inc where txn_idy like '" & Trim(txnid) & "%' group by txn_idy", CON, adOpenStatic
Label20.Caption = "Select Transaction"
lv.ColumnHeaders.Clear
lv.ColumnHeaders.Add , , "Transaction ID", 2000
lv.ColumnHeaders.Add , , "Salesman", 2500
'lv.ColumnHeaders.Add , , "Incentive Amt", 1500, 2
lv.ListItems.Clear
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rs("txn_idy")))
        If rssmn.State = 1 Then rssmn.Close
        rssmn.Open "select smn_nme from smn where smn_idy='" & rs("smn_idy") & "'", CON
        If Not rssmn.EOF Then
            ls.ListSubItems.Add , , Trim(rssmn("smn_nme"))
        End If
        'ls.ListSubItems.Add , , Trim(rs("inc_amt") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    lv.Visible = True
    lv.SetFocus
    MD = True
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Transaction ID Not Found, Define Transaction !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        txnid.Enabled = True
        txnid.SetFocus
        SendKeys "{HOME}+{END}"
        MD = False
    End If
End If
End Sub

Private Sub saverecord()
If CHECKTB("inc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If MsgBox("Save Record ?", vbYesNo + vbExclamation) = vbNo Then
    cancel
    Exit Sub
End If
If Trim(smnid) = "" Then
    MsgBox "Please select a Salesman ID", vbExclamation
    smnid.SetFocus
    Exit Sub
End If
If Trim(txtQuan) = "" Then
    MsgBox "Please enter Quantity", vbExclamation
    All.Value = 1
    txtQuan.SetFocus
    Exit Sub
End If
If MD = False Then
    If RSINC.State = 1 Then RSINC.Close
    RSINC.Open "select * from inc where txn_idy=''", CON, adOpenKeyset, adLockOptimistic
    R = 1
    txnid = increment("idy_txn")
    If All.Value = 1 Then
        If Val(txtIncAmt) <= 0 Then
            MsgBox "Enter Incentive Amount", vbExclamation
            txtIncAmt.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        RSINC.AddNew
        RSINC("txn_idy") = txnid
        RSINC("inc_typ") = Trim(inctyp.Text)
        RSINC("smn_idy") = Trim(smnid)
        RSINC("frm_dat") = fdate.Value
        RSINC("to_dat") = tdate.Value
        RSINC("chk_flg") = "A"
        RSINC("prd_qty") = Val(txtQuan)
        RSINC("inc_amt") = Val(txtIncAmt)
        RSINC.Update
        MsgBox "New Transaction ID is " & txnid, vbExclamation
        cancel
    Else
       If Val(TXTINCAMT1.Text) <= 0 Then
           MsgBox "Enter Incentive Amount", vbExclamation
           TXTINCAMT1.SetFocus
           SendKeys "{HOME}+{END}"
           Exit Sub
       End If
       For R = 1 To MS.Rows - 1
            If Val(MS.TextMatrix(R, 1)) > 0 Then
                RSINC.AddNew
                RSINC("txn_idy") = txnid
                RSINC("Inc_typ") = Trim(inctyp.Text)
                RSINC("smn_idy") = Trim(smnid)
                RSINC("frm_dat") = fdate.Value
                RSINC("to_dat") = tdate.Value
                RSINC("chk_flg") = "S"
                RSINC("prd_qty") = MS.TextMatrix(R, 1)
                RSINC("Inc_amt") = 0
                RSINC("prd_idy") = MS.TextMatrix(R, 3)
                RSINC.Update
            End If
        Next
        RSINC("inc_amt") = Val(TXTINCAMT1.Text)
        RSINC.Update
        MsgBox "New Transaction ID is " & txnid, vbExclamation
        cancel
    End If
Else
    If RSINC.State = 1 Then RSINC.Close
    RSINC.Open "SELECT * FROM Inc WHERE TXN_IDY='" & Trim(txnid) & "'", CON, adOpenKeyset, adLockOptimistic
    'If Trim(inctyp.Text) = "Quantity" Then
    If RSINC.RecordCount > 0 Then
        If RSINC("CHK_FLG") = "A" Then
            RSINC("txn_idy") = txnid
            RSINC("Inc_typ") = Trim(inctyp.Text)
            RSINC("smn_idy") = Trim(smnid)
            RSINC("frm_dat") = fdate.Value
            RSINC("to_dat") = tdate.Value
            RSINC("chk_flg") = "A"
            RSINC("PRD_QTY") = Val(txtQuan)
            RSINC("Inc_amt") = Val(txtIncAmt)
            RSINC.Update
            cancel
        Else
            R = 1
            While Not RSINC.EOF
                RSINC("Inc_typ") = Trim(inctyp.Text)
                RSINC("smn_idy") = Trim(smnid)
                RSINC("frm_dat") = fdate.Value
                RSINC("to_dat") = tdate.Value
                If All.Value = 1 Then
                    RSINC("chk_flg") = "A"
                Else
                    RSINC("chk_flg") = "S"
                End If
                RSINC("PRD_QTY") = MS.TextMatrix(R, 1)
                RSINC("Inc_amt") = Val(TXTINCAMT1)
                RSINC("prd_idy") = MS.TextMatrix(R, 3)
                RSINC.Update
                RSINC.MoveNext
                R = R + 1
            Wend
            cancel
        End If
    End If
End If
End Sub

Private Sub txtIncAmt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtIncAmt = convert(txtIncAmt.Text)
    saverecord
End If
End Sub

Private Sub txtIncAmt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Or KeyAscii = 46 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub TXTINCAMT1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    TXTINCAMT1 = convert(TXTINCAMT1.Text)
    saverecord
End If
End Sub

Private Sub txtQuan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(inctyp) = "Quantity" Or Trim(inctyp) = "Volume" Then
    Else
        txtQuan = convert(txtQuan)
    End If
    txtIncAmt.Enabled = True
    txtIncAmt.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtQuan_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    txtQuan.Locked = False
End If
If KeyAscii = 46 Then
    If Trim(inctyp.Text) = "Quantity" Then txtQuan.Locked = True
End If
If Trim(inctyp.Text) = "Quantity" Then
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 42 Then
Else
KeyAscii = 0
End If
Else
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 42 Or KeyAscii = 46 Then
Else
KeyAscii = 0
End If
End If
End Sub

Private Sub txtvar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MS.Col = 1 Then
        If Trim(inctyp.Text) = "Quantity" Then
            txtvar = QTY(Val(txtvar))
            If Val(CMP()) >= QTY(Val(MS.TextMatrix(MS.Row, 4))) Then
                MsgBox "Sub Units are " & Val(MS.TextMatrix(MS.Row, 4)), vbExclamation
                txtvar.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                MS.Text = QTY(Trim(Val(txtvar)))
                txtvar = ""
                txtvar.Visible = False
                If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                MS.SetFocus
            End If
        Else
            MS.Text = convert(Trim(Val(txtvar)))
            txtvar = ""
            txtvar.Visible = False
            If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
            MS.SetFocus
        End If
    Else
        MS.Text = convert(Trim(Val(txtvar)))
        txtvar = ""
        txtvar.Visible = False
        If MS.Row <> MS.Rows - 1 Then
            MS.Row = MS.Row + 1
            MS.Col = 1
        End If
        FL = True
        MS.SetFocus
    End If
End If
End Sub

Private Sub txtVar_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
Else
    KeyAscii = 0
End If
End Sub
Public Function CMP()
Dim chk As Integer, CHK1 As String
chk = InStr(Trim(STR(Val(txtvar))), ".")
If chk <> 0 Then
    CHK1 = Mid(Trim(STR(Val(txtvar))), chk + 1, Len(Trim(Val(txtvar))))
    If Len(Trim(CHK1)) = 1 Then
        CMP = Val(CHK1) * 100
    ElseIf Len(Trim(CHK1)) = 2 Then
        CMP = Val(CHK1) * 10
    ElseIf Len(Trim(CHK1)) = 3 Then
        CMP = Val(CHK1) * 1
    End If
End If
End Function

Private Sub dispprd()
If RSPRD.State = 1 Then RSPRD.Close
RSPRD.Open "select prd_idy,prd_nme,sub_unt from prd", CON, adOpenKeyset, adLockOptimistic
R = 1
MS.Rows = 1
MS.Cols = 5
R = 1
Frame1.Visible = True
Label8.Visible = True
TXTINCAMT1.Visible = True
MS.ColWidth(3) = 0
MS.ColWidth(4) = 0
MS.ColWidth(0) = 2500
MS.ColWidth(1) = 1100
MS.ColWidth(2) = 0
MS.TextMatrix(0, 0) = "Product Name"
MS.TextMatrix(0, 1) = Trim(inctyp.Text)
MS.TextMatrix(0, 2) = "Incentive Amt"
MS.ColAlignment(2) = 7
If Trim(inctyp.Text) = "Quantity" Then
    While Not RSPRD.EOF
        MS.Rows = MS.Rows + 1
        MS.TextMatrix(R, 0) = RSPRD(1)
        MS.TextMatrix(R, 3) = RSPRD(0)
        MS.TextMatrix(R, 4) = RSPRD(2)
        MS.TextMatrix(R, 1) = "0.000"
        R = R + 1
        RSPRD.MoveNext
    Wend
Else
    While Not RSPRD.EOF
        MS.Rows = MS.Rows + 1
        MS.TextMatrix(R, 0) = RSPRD(1)
        MS.TextMatrix(R, 3) = RSPRD(0)
        MS.TextMatrix(R, 4) = RSPRD(2)
        MS.TextMatrix(R, 1) = "0"
        R = R + 1
        RSPRD.MoveNext
    Wend
End If
End Sub

