VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_CRT 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipts"
   ClientHeight    =   5625
   ClientLeft      =   180
   ClientTop       =   720
   ClientWidth     =   8820
   Icon            =   "M_FRM_CRT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   24
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Cash Receipt Details"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "M_TXN_IDY"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "M_DBT_IDY"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "M_TXN_DAT"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Label7"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Other Details"
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MS"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "amount"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "date1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "View"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "list3"
      Tab(2).Control(2)=   "list1"
      Tab(2).Control(3)=   "list2"
      Tab(2).Control(4)=   "Label49"
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74880
         TabIndex        =   21
         Top             =   4680
         Width           =   4935
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
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
            Height          =   240
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
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
            Height          =   195
            Left            =   1800
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Value           =   1  'Checked
            Width           =   1695
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox M_TXN_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72915
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox M_DBT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72915
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker M_TXN_DAT 
         Height          =   330
         Left            =   -72915
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24576003
         CurrentDate     =   37097
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   14737632
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24576003
         CurrentDate     =   37099
         MaxDate         =   37256
         MinDate         =   36526
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         PromptInclude   =   0   'False
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4275
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   405
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   7541
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView list3 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5953
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
            Text            =   "Voucher No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Voucher Date"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5953
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
            Text            =   "Cash Account Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Account ID"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5953
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
            Text            =   "Account Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Account ID"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Address"
            Object.Width           =   3528
         EndProperty
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
         TabIndex        =   20
         Top             =   720
         Width           =   8415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7320
         TabIndex        =   16
         Top             =   4815
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6000
         TabIndex        =   15
         Top             =   4815
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher #"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -74040
         TabIndex        =   10
         Top             =   908
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   -73395
         TabIndex        =   9
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73635
         TabIndex        =   8
         Top             =   1508
         Width           =   345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73650
         TabIndex        =   7
         Top             =   1988
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Receipt"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -74250
         TabIndex        =   6
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   -70635
         TabIndex        =   5
         Top             =   1920
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72915
         TabIndex        =   4
         Top             =   2400
         Width           =   60
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
            Picture         =   "M_FRM_CRT.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":2E02
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRT.frx":334E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_CRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, IDNEW As Boolean, LS As ListItem

Private Sub amount_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Then

ElseIf K = 13 Then
    MS.Text = convert(amt(Val(amount)))
    amount.Visible = False
    MS.SetFocus
    MS.Col = MS.Col + 1
Else
K = 0
End If
End Sub

Private Sub SAVERECORD()
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset

If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "Effmcg") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If check("act", "act_idy", Trim(M_DBT_IDY)) = False Then
    MsgBox "Cash Account Not Found, Define Cash Account !", vbExclamation, "Effmcg"
    sstab1.Tab = 0
    M_DBT_IDY.SetFocus
    Exit Sub
End If
If MOD_REC = True Then
    con.Execute "delete from vch where txn_idy like 'CRT" & Tincrement(M_TXN_IDY) & "' and txn_typ='CRV'"
    INC = M_TXN_IDY
Else
    INC = "CRT" & T7increment("idy_crv")
End If
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from vch ", con, adOpenKeyset, adLockOptimistic
If MS.Rows >= 2 And MS.TextMatrix(1, 0) <> "" Then
    For A = 1 To MS.Rows - 1
        REC_SET.AddNew
        REC_SET("txn_idy") = INC
        If MS.TextMatrix(A, 0) <> "" Then
            REC_SET("txn_dat") = M_TXN_DAT
            REC_SET("crt_idy") = M_DBT_IDY
            REC_SET("dbt_idy") = MS.TextMatrix(A, 0)
            REC_SET("txn_typ") = "CRT"
            REC_SET("txn_amt") = Val(MS.TextMatrix(A, 2))
            REC_SET("clg_flg") = Trim(MS.TextMatrix(A, 3))
            If MS.TextMatrix(A, 4) <> "" Then
                REC_SET("clg_dat") = MS.TextMatrix(A, 4)
            Else
                REC_SET("clg_dat") = Date
            End If
            REC_SET("txn_rmk") = Trim(MS.TextMatrix(A, 5) & "")
            REC_SET.Update
        End If
    Next
 Else
    REC_SET.AddNew
    REC_SET("txn_idy") = INC
    REC_SET("txn_dat") = M_TXN_DAT
    REC_SET("crt_idy") = M_DBT_IDY
    REC_SET("txn_typ") = "CRT"
    REC_SET("clg_dat") = Date
    REC_SET.Update
End If
If MOD_REC = False Then MsgBox "New Cash Receipt Number is " & INC, vbExclamation, "Effmcg"
If check("PRT", "PRT_NME", "M_ACT_CRT") = True Then
    con.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
    con.Execute "update dir where dir_idy like 'REM' set dir_als = 'CRT'"
End If
CANCELRECORD
End Sub

Private Sub CANCELRECORD()
clrctr M_FRM_CRT
Label9 = "0"
Label7 = ""
Label6 = ""
M_TXN_IDY.Enabled = True
M_TXN_IDY.SetFocus
IDNEW = False
M_TXN_DAT = Date
date1 = Date
MOD_REC = False
MS.Rows = 1
MS.AddItem ""
Label2 = labdisp("idy_crv")
End Sub

Private Sub NEWRECORD()
CANCELRECORD
M_TXN_IDY.Enabled = False
M_TXN_DAT.SetFocus
End Sub

Private Sub ACCOUNTLIST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
REC_SET.Open "select * from act where act_idy not like '" & Sincrement(M_DBT_IDY) & "' and act_nme like '" & Trim(Text3) & "%'", con, adOpenStatic
Else
REC_SET.Open "select * from act where act_idy not like '" & Sincrement(M_DBT_IDY) & "' and act_ad4 like '" & Trim(Text3) & "%'", con, adOpenStatic
End If
list2.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set LS = list2.ListItems.Add(, , REC_SET("act_nme"))
        LS.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        LS.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        LS.ListSubItems.Add , , Trim(REC_SET("ACT_TYP") & "")
        REC_SET.MoveNext
    Next
    sstab1.Tab = 2
    Label49 = "Select Account"
    list2.Visible = True
    Frame1.Visible = True
    list2.SetFocus
Else
    If sstab1.Tab = 1 Then
    MsgBox "Accounts Not Found, Define Accounts !", vbExclamation, "Effmcg"
    MS.SetFocus
    End If
End If
End Sub

Private Sub EXITRECORD()
Unload Me
End Sub

Private Sub voucherlist()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select distinct(txn_idy),act_nme,act_ad4,txn_dat from vch,ACT where txn_typ like 'CRT' AND vch.crt_idy like act.act_idy ", con, adOpenStatic
If REC_SET.RecordCount > 0 Then
    List3.ListItems.Clear
    For A = 1 To REC_SET.RecordCount
    Set LS = List3.ListItems.Add(, , Trim(REC_SET(0)))
    LS.ListSubItems.Add , , Trim(REC_SET(3) & "")
    LS.ListSubItems.Add , , Trim(REC_SET(1) & "")
    LS.ListSubItems.Add , , Trim(REC_SET(2) & "")
    REC_SET.MoveNext
    Next
    Label49 = "Select Voucher"
    sstab1.Tab = 2
    List3.Visible = True
    Frame1.Visible = False
    List3.SetFocus
Else
    MsgBox "Transactions Not Found, Define Transactions !", vbExclamation, "Effmcg"
    M_TXN_IDY.Enabled = True
    SendKeys "{HOME}+{END}"
    M_TXN_IDY.SetFocus
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
If Label49.Caption = "Select Voucher" Then
    voucherlist
ElseIf Label49.Caption = "Select Account" Then
    ACCOUNTLIST
ElseIf Label49.Caption = "Select Cash Account" Then
    V_ACT_LST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If Label49.Caption = "Select Voucher" Then
    voucherlist
ElseIf Label49.Caption = "Select Account" Then
    ACCOUNTLIST
ElseIf Label49.Caption = "Select Cash Account" Then
    V_ACT_LST
End If
End Sub

Private Sub DATE1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
MS.Text = date1
MS.Text = Format(MS.Text, "dd/mm/yyyy")
date1.Visible = False
MS.SetFocus
MS.Col = MS.Col + 1
End If
End Sub
Private Sub DATE1_CloseUp()
date1.SetFocus
End Sub

Private Sub M_TXN_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date using UP or DOWN ARROW KEYS"
End Sub

Private Sub M_TXN_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
M_DBT_IDY.SetFocus
End If
End Sub

Private Sub M_TXN_DAT_CloseUp()
M_TXN_DAT.SetFocus
End Sub

Private Sub Form_Activate()
SETYEAR M_FRM_CRT
M_TXN_DAT = Date
date1 = Date
sstab1.Tab = 0
sstab1.TabEnabled(2) = False
If M_TXN_IDY.Enabled = True Then
M_TXN_IDY.SetFocus
Else
M_TXN_DAT.SetFocus
End If
MS.ColWidth(6) = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If M_FRM_CRT.ActiveControl.name = "M_TXN_IDY" And Trim(M_TXN_IDY.Text) = "" Then
    Unload Me
    Exit Sub
    End If
    If list1.Visible = True Then
        list1.Visible = False
        sstab1.Tab = 0
        M_DBT_IDY.SetFocus
    ElseIf list2.Visible = True Then
        list2.Visible = False
        sstab1.Tab = 1
        If Text3.Visible = True Then
            Text3.SetFocus
        Else
            MS.SetFocus
        End If
    ElseIf List3.Visible = True Then
        List3.Visible = False
        sstab1.Tab = 0
        If M_TXN_IDY.Enabled = True Then
            M_TXN_IDY.SetFocus
        Else
            M_DBT_IDY.SetFocus
        End If
    ElseIf sstab1.Tab = 1 Then
        Text3.Visible = False
        date1.Visible = False
        amount.Visible = False
        Label7 = "0.00"
        For A = 1 To MS.Rows - 1
        Label7 = convert(Label7 + Val(MS.TextMatrix(A, 2)))
        Next
        sstab1.Tab = 0
        SAVERECORD
    Else
        SAVERECORD
    End If
End If
End Sub

Private Sub Form_Load()
Dim REC_SET As New ADODB.Recordset
Label2 = labdisp("idy_crv")
MS.ColWidth(0) = 1250
MS.ColWidth(1) = 3000
MS.ColWidth(2) = 1100
MS.ColWidth(3) = 450
MS.ColWidth(4) = 1200
MS.ColWidth(5) = 3500

MS.TextMatrix(0, 0) = "Account"
MS.TextMatrix(0, 1) = "Account Name"
MS.TextMatrix(0, 2) = "      Amount"
MS.TextMatrix(0, 3) = "Clg"
MS.TextMatrix(0, 4) = "Clg Date"
MS.TextMatrix(0, 5) = "Remarks"

If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where grp_clf like 'Cash %' ", con, adOpenStatic
If REC_SET.RecordCount > 0 Then
    list1.ListItems.Clear
    For A = 1 To REC_SETRecordCount
        Set LS = list1.ListItems.Add(, , REC_SET("act_nme"))
        LS.ListSubItems.Add , , REC_SET("act_idy")
        LS.ListSubItems.Add , , REC_SET("act_ad4") & ""
        REC_SET.MoveNext
    Next
End If

If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act ", con, adOpenStatic
If REC_SET.RecordCount > 0 Then
    list2.ListItems.Clear
    For A = 1 To REC_SETRecordCount
        Set LS = list2.ListItems.Add(, , REC_SET("act_nme"))
        LS.ListSubItems.Add , , REC_SET("act_idy")
        LS.ListSubItems.Add , , REC_SET("act_ad4") & ""
        REC_SET.MoveNext
    Next
End If
MOD_REC = False
IDNEW = False

End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub LIST1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list2_DblClick()
List2_KeyDown 13, 0
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    M_DBT_IDY = list1.SelectedItem.ListSubItems.Item(1)
    Label6 = list1.SelectedItem
    list1.Visible = False
    sstab1.Tab = 1
    MS.SetFocus
    MS.Col = 0
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

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    list2.Visible = False
    Text3.Visible = False
    If list2.ListItems.Count > 0 Then
    MS.TextMatrix(MS.Row, 0) = list2.SelectedItem.ListSubItems(1)
    MS.TextMatrix(MS.Row, 1) = list2.SelectedItem
    Set LS = list2.SelectedItem
    MS.TextMatrix(MS.Row, 6) = LS.ListSubItems(3)
    End If
    sstab1.Tab = 1
    MS.SetFocus
    MS.Col = 2
    
End If
End Sub

Private Sub list2_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list2.SortOrder = lvwAscending Then
list2.SortOrder = lvwDescending
ElseIf list2.SortOrder = lvwDescending Then
list2.SortOrder = lvwAscending
End If
list2.SortKey = CH.Index - 1
End Sub

Private Sub list3_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If List3.SortOrder = lvwAscending Then
List3.SortOrder = lvwDescending
ElseIf List3.SortOrder = lvwDescending Then
List3.SortOrder = lvwAscending
End If
List3.SortKey = CH.Index - 1
End Sub

Private Sub list3_DblClick()
List3_KeyDown 13, 0
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    M_TXN_IDY = List3.SelectedItem
    List3.Visible = False
    sstab1.Tab = 0
    M_TXN_IDY_KeyDown 13, 0
End If
End Sub

Private Sub MS_EnterCell()
If MS.Col = 3 Then
    Label9 = "0.00"
    For A = 1 To MS.Rows - 1
        Label9 = convert(Val(Label9) + Val(MS.TextMatrix(A, 2)))
    Next
    M_MDI.STS_BAR.Panels(1).Text = "Enter Clearings 'Y' or 'N'"
ElseIf MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Account"
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Amount"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Select Clearing Date Using Arrow Keys or Press F4"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Remarks"
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_leaveCell()
date1.Visible = False
Text3.Visible = False
amount.Visible = False
MS.SetFocus
End Sub

Private Sub MS_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K <> 27 Then
    If MS.Col = 0 Then
    If K = 13 Then
        If Len(Trim(MS.TextMatrix(MS.Row, 0))) = 0 Then
            Text3 = ""
            ACCOUNTLIST
        Else
            MS.Col = 2
        End If
    Else
        Text3.Visible = True
        Text3.top = MS.top + MS.CellTop
        Text3.Left = MS.CellLeft + MS.Left
        Text3.Width = MS.CellWidth
        Text3.Height = MS.CellHeight
        Text3 = Chr(K)
        Text3.SetFocus
        Text3.SelStart = Len(Text3)
     End If
End If
If MS.Col = 5 Or MS.Col = 3 Then
    If K = 13 And MS.Col = 3 Then
        MS.Col = MS.Col + 1
    Else
        Text3.Visible = True
        If MS.Col = 3 Then Text3.MaxLength = 1
        Text3.top = MS.top + MS.CellTop
        Text3.Left = MS.CellLeft + MS.Left
        Text3.Width = MS.CellWidth
        Text3.Height = MS.CellHeight
        Text3 = Chr(K)
        Text3.SetFocus
        Text3.SelStart = Len(Text3)
    End If
ElseIf MS.Col = 2 Then
    If K = 13 Then
        MS.Col = 3
    ElseIf K > 47 And K < 58 Or K = 8 Or K = 46 Then
        amount.Text = MS.TextMatrix(MS.Row, 2)
        amount.Visible = True
        amount = ""
        amount.top = MS.top + MS.CellTop
        amount.Left = MS.CellLeft + MS.Left
        amount.Width = MS.CellWidth
        amount.Height = MS.CellHeight
        amount = Chr(K)
        amount.SetFocus
        amount.SelStart = Len(amount)
    Else
        K = 0
    End If
ElseIf MS.Col = 4 Then
    date1.Visible = True
    date1.top = MS.top + MS.CellTop
    date1.Left = MS.CellLeft + MS.Left
    date1.Width = MS.CellWidth
    date1.Height = MS.CellHeight
    date1.SetFocus
End If
End If
If MS.Col = 5 And MS.Row = MS.Rows - 1 Then
    If MS.TextMatrix(MS.Row, 0) <> "" And K = 13 Then
    MS.AddItem ""
    MS.Row = MS.Rows - 1
    MS.Col = 0
End If
End If
End Sub

Private Sub MS_LostFocus()
MS.HighLight = flexHighlightNever
End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)
If sstab1.Tab = 0 And M_TXN_IDY.Enabled = False Then
    M_TXN_DAT.SetFocus
ElseIf sstab1.Tab = 1 Then
    MS.SetFocus
ElseIf sstab1.Tab = 0 Then
M_TXN_IDY.SetFocus
End If
If sstab1.Tab = 2 Then
    Toolbar1.Enabled = False
Else
    Toolbar1.Enabled = True
End If
End Sub

Private Sub M_TXN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Voucher Number"
End Sub

Private Sub M_TXN_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If REC_SET.State = 1 Then REC_SET.Close
    If Len(Trim(M_TXN_IDY)) <> 0 Then
        If Len(Trim(M_TXN_IDY)) <> 10 Then
        REC_SET.Open "select * from vch where txn_idy like '" & Trim(M_TXN_IDY) & "' and txn_typ='CRV'", con, adOpenStatic
        Else
        REC_SET.Open "select * from vch where txn_idy like '" & Trim(M_TXN_IDY) & "' and txn_typ='CRV'", con, adOpenStatic
        End If
        If REC_SET.RecordCount > 0 Then
            Label7 = "0"
            Label9 = "0"
            MOD_REC = True
            M_TXN_DAT = REC_SET("txn_dat")
            M_DBT_IDY = REC_SET("crt_idy")
            If REC_SET1.State = 1 Then REC_SET1.Close
            REC_SET1.Open "select * from act where act_idy like '" & M_DBT_IDY & "'", con, adOpenStatic
            If REC_SET1.RecordCount > 0 Then Label6 = REC_SET1("act_nme")
            MS.Rows = REC_SET.RecordCount + 1
            For A = 1 To REC_SET.RecordCount
                 MS.TextMatrix(A, 0) = REC_SET("dbt_idy") & ""
                 If REC_SET1.State = 1 Then REC_SET1.Close
                 REC_SET1.Open "select * from act where act_idy like '" & Trim(REC_SET("dbt_idy")) & "'", con, adOpenStatic
                 If REC_SET1.RecordCount > 0 Then MS.TextMatrix(A, 1) = REC_SET1("act_nme")
                 MS.TextMatrix(A, 2) = convert(Val(REC_SET("txn_amt") & ""))
                 MS.TextMatrix(A, 3) = REC_SET("clg_flg") & ""
                 MS.TextMatrix(A, 4) = REC_SET("clg_dat") & ""
                 MS.TextMatrix(A, 5) = Trim(REC_SET("txn_rmk") & "")
                 If REC_SET1.RecordCount > 0 Then MS.TextMatrix(A, 6) = Trim(REC_SET1("ACT_TYP") & "")
                 Label7 = convert(amt(Val(Label7) + Val(REC_SET("txn_amt") & "")))
                 Label9 = convert(Val(Label9) + Val(REC_SET("txn_amt") & ""))
                 REC_SET.MoveNext
            Next
            M_TXN_IDY.Enabled = False
            M_TXN_IDY = Sincrement(M_TXN_IDY)
            M_TXN_DAT.SetFocus
        Else
            MsgBox "Voucher Number Not Found !", vbExclamation, "Effmcg"
            M_TXN_IDY.SetFocus
        End If
    Else
        voucherlist
    End If
End If
End Sub

Private Sub M_TXN_IDY_keypress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K = 27 And M_TXN_IDY = "" Then Unload Me
End Sub

Private Sub M_DBT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Cash Account"
End Sub

Private Sub M_DBT_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim REC_SET As New ADODB.Recordset

If KeyCode = 13 Or KeyCode = 40 Then
    If REC_SETState = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(M_DBT_IDY) & "' and grp_clf like 'Cash %'", con, adOpenStatic
    If REC_SET.RecordCount > 0 Then
        sstab1.Tab = 1
        MS.SetFocus
        MS.Col = 0
        Label6 = REC_SET("act_nme")
    Else
        V_ACT_LST
    End If
ElseIf KeyCode = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from act where act_idy like '" & M_DBT_IDY & "' and grp_clf like 'Cash %'", con, adOpenStatic
        If REC_SET.RecordCount > 0 Then
            M_TXN_DAT.SetFocus
            Label6 = REC_SET("act_nme")
        Else
            MsgBox "Cash Account Not Found, Define Accounts !", vbExclamation, "Effmcg"
            M_DBT_IDY.SetFocus
        End If
End If
End Sub

Private Sub M_DBT_IDY_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub

Private Sub Text3_KeyDown(K As Integer, S As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Then
    If MS.Col = 5 Then
        MS.Text = Trim(Text3)
        Text3.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 0 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Sincrement(Text3) & "' and act_idy not like '" & Sincrement(M_DBT_IDY) & "'", con, adOpenStatic
        If REC_SET.RecordCount > 0 Then
            MS.TextMatrix(MS.Row, 0) = REC_SET("act_IDY")
            MS.TextMatrix(MS.Row, 1) = REC_SET("act_nme")
            Text3.Visible = False
            MS.SetFocus
            MS.Col = 2
        Else
            ACCOUNTLIST
        End If
    End If
End If
End Sub


Private Sub Text3_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0

If MS.Col = 3 Then
    If K = 13 Then
        If UCase(Trim(Text3)) = "Y" Or UCase(Trim(Text3)) = "N" Then
            MS.Text = UCase(Text3)
            Text3.Visible = False
            Text3.MaxLength = 0
            MS.SetFocus
            MS.Col = MS.Col + 1
        Else
            K = 0
        End If
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then Shell "c:\windows\calc.exe"
If list1.Visible = False And list2.Visible = False And List3.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then NEWRECORD
    If KeyCode = 80 And Shift = 2 Then PRINTRECORD
    If KeyCode = 112 Then Shell "c:\windows\hh.exe " & App.Path & "\conthelp.chm", vbMaximizedFocus
    If KeyCode = 83 And Shift = 2 Then SAVERECORD
    If KeyCode = 88 And Shift = 2 Then EXITRECORD
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 69 And Shift = 2 Then
        If M_ACT_CRT.ActiveControl.name = "M_TXN_IDY" Then
            voucherlist
        ElseIf M_ACT_CRT.ActiveControl.name = "M_DBT_IDY" Then
            V_ACT_LST
        ElseIf M_ACT_CRT.ActiveControl.name = "Text3" Then
            ACCOUNTLIST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_ACT_CRT.ActiveControl Is TextBox Then M_ACT_CRT.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
NEWRECORD
ElseIf Button.KEY = "s" Then
SAVERECORD
ElseIf Button.KEY = "e" Then
EXITRECORD
ElseIf Button.KEY = "p" Then
PRINTRECORD
ElseIf Button.KEY = "pre" Then
M_FRM_PRE.Caption = "Prefix Cash Receipts"
M_FRM_PRE.Show 1
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "ed" Then
    If M_ACT_CRT.ActiveControl.name = "M_TXN_IDY" Then
        Check1.Value = 0
        Check2.Value = 0
        Label49 = "Select Voucher"
        voucherlist
    ElseIf M_ACT_CRT.ActiveControl.name = "MS" And MS.Col = 0 Then
        Check1.Value = 0
        Check2.Value = 0
        Label49 = "Select Account"
        ACCOUNTLIST
    ElseIf M_ACT_CRT.ActiveControl.name = "M_DBT_IDY" Then
        Check1.Value = 0
        Check2.Value = 0
        Label49 = "Select Cash Account"
        V_ACT_LST
    End If
ElseIf Button.KEY = "h" Then
Shell "c:\windows\hh.exe " & App.Path & "\conthelp.chm", vbMaximizedFocus
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
sstab1.Enabled = False
Toolbar1.Enabled = False
    If check("prt", "prt_nme", "M_ACT_CRT") = True Then
        Check3.Value = 1
    Else
        Check3.Value = 0
    End If
    Check3.SetFocus
End If
End Sub

Public Sub V_ACT_LST()
Dim REC_SET As New ADODB.Recordset
If REC_SETState = 1 Then REC_SET.Close
If Check1.Value = 1 Then
REC_SET.Open "select * from act where act_nme like '" & Trim(M_DBT_IDY) & "%' and grp_clf like 'Cash Book'", con, adOpenStatic
Else
REC_SET.Open "select * from act where act_ad4 like '" & Trim(M_DBT_IDY) & "%' and grp_clf like 'Cash Book'", con, adOpenStatic
End If
list1.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
    Set LS = list1.ListItems.Add(, , Trim(REC_SET("act_nme") & ""))
    LS.ListSubItems.Add , , Trim(REC_SET("act_idy"))
    LS.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
    REC_SET.MoveNext
    Next
    Label49 = "Select Cash Account"
    sstab1.Tab = 2
    list1.Visible = True
    Frame1.Visible = True
    list1.SetFocus
Else
    If sstab1.Tab = 0 Then
    MsgBox "Cash Account Not Found, Define Cash Account !", vbExclamation, "Effmcg"
    sstab1.Tab = 0
    M_DBT_IDY.SetFocus
    End If
End If
End Sub

Public Sub PRINTRECORD()
M_ACT_CRT.Tag = "CRT"
If check("PRT", "PRT_NME", "M_ACT_CRT") = True Then printvoucher.Check3.Value = 1
printvoucher.Show 1
End Sub
