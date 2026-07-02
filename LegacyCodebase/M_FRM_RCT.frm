VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_RCT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipts (Bank and Cash)"
   ClientHeight    =   5625
   ClientLeft      =   180
   ClientTop       =   720
   ClientWidth     =   8820
   Icon            =   "M_FRM_RCT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdXRef 
      Caption         =   "&Van - MER Reference"
      Height          =   375
      Left            =   7000
      TabIndex        =   28
      Top             =   50
      Width           =   1815
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
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Receipt Voucher"
      TabPicture(0)   =   "M_FRM_RCT.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblMERNum"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblVanRef"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DTPicker1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&Detail Entries"
      TabPicture(1)   =   "M_FRM_RCT.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List4"
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "date1"
      Tab(1).Control(3)=   "amount"
      Tab(1).Control(4)=   "MS"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label8"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "M_FRM_RCT.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "list3"
      Tab(2).Control(2)=   "list1"
      Tab(2).Control(3)=   "list2"
      Tab(2).Control(4)=   "Label49"
      Tab(2).ControlCount=   5
      Begin VB.ListBox List4 
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
         ItemData        =   "M_FRM_RCT.frx":0496
         Left            =   -70920
         List            =   "M_FRM_RCT.frx":04A3
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74760
         TabIndex        =   22
         Top             =   4560
         Width           =   4935
         Begin VB.CheckBox Check1 
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
            Height          =   240
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check2 
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
            Height          =   195
            Left            =   3120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -70800
         MaxLength       =   70
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2355
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1845
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74317827
         CurrentDate     =   37097
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   -68760
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   14737632
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   74317827
         CurrentDate     =   37099
         MaxDate         =   37256
         MinDate         =   36526
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   375
         Left            =   -69120
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
         Left            =   -74940
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   405
         Width           =   8720
         _ExtentX        =   15372
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
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
      Begin MSComctlLib.ListView list3 
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
            Text            =   "Place"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   20
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
      Begin VB.Label lblVanRef 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   27
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblMERNum 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         TabIndex        =   26
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   21
         Top             =   720
         Width           =   8415
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -70080
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
         Left            =   -72240
         TabIndex        =   15
         Top             =   4815
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   795
         TabIndex        =   10
         Top             =   1395
         Width           =   555
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2880
         TabIndex        =   9
         Top             =   1410
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1005
         TabIndex        =   8
         Top             =   1905
         Width           =   345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   540
         TabIndex        =   7
         Top             =   2430
         Width           =   810
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   810
         TabIndex        =   6
         Top             =   2955
         Width           =   540
      End
      Begin VB.Label Label6 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   2355
         Width           =   5925
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
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
            Picture         =   "M_FRM_RCT.frx":04B0
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":09F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":0DBC
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":1110
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":1748
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":1A9C
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":1EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":21F0
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":25FC
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":2708
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":2A5C
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RCT.frx":2E70
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Sets To Print the Record"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Cash Bills -> Receipts Conversion"
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
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_RCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, idnew As Boolean, ls As ListItem, CHECKBILL As Boolean
Private Sub amount_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 46 Or K = 8 Then
ElseIf K = 13 Then
    MS.Text = convert1(AMT1(Val(amount)))
    amount.Visible = False
    MS.SetFocus
    MS.Col = MS.Col + 1
Else
    K = 0
End If
End Sub
Private Sub saverecord()
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("act.dbf,vch.dbf,inl.dbf,prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("act", "act_idy", Trim(Text2)) = False Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text2.SetFocus
    Exit Sub
End If
If MS.TextMatrix(1, 0) <> "" And Val(Trim(MS.TextMatrix(1, 2))) >= "0.000" Then 'CHANGED BY PKB TO RESTRICT NUMBER INCREMENT
    If MOD1 = True Then
        con.Execute "delete from vch where txn_idy like '" & Trim(Text1) & "' and txn_typ='RCT'"
        INC = Trim(Text1)
    Else
        INC = "RCT" & T7increment("IDY_RCT")
    End If
End If
Dim Item As ListItem

If rs.State = 1 Then rs.Close
rs.Open "select * from vch where txn_idy = ''", con, adOpenKeyset, adLockPessimistic
If MS.Rows >= 2 Then
    If MS.TextMatrix(1, 0) <> "" Then
    For A = 1 To MS.Rows - 1
    If MS.TextMatrix(A, 0) <> "" And Val(Trim(MS.TextMatrix(A, 2))) >= "0.000" Then
        'If MS.TextMatrix(1, 0) <> "" And Val(Trim(MS.TextMatrix(A, 2))) <> "0.000" Then
        rs.AddNew
        rs("txn_idy") = INC
        If MS.TextMatrix(A, 0) <> "" Then
            rs("txn_dat") = DTPicker1
            rs("DBT_idy") = Text2
            rs("CRT_idy") = MS.TextMatrix(A, 0)
            rs("txn_typ") = "RCT"
            rs("txn_amt") = Val(MS.TextMatrix(A, 2))
            rs("clg_flg") = Trim(MS.TextMatrix(A, 3))
            If MS.TextMatrix(A, 4) <> "" Then
                rs("clg_dat") = MS.TextMatrix(A, 4)
            Else
                rs("clg_dat") = Date
            End If
            rs("txn_rmk") = Trim(MS.TextMatrix(A, 5) & "")
            rs("INV_IDY") = Trim(MS.TextMatrix(A, 7) & "")
            rs("INV_TYP") = Trim(MS.TextMatrix(A, 8) & "")
            rs!USR_IDY = Trim(USERNAME)
            rs!sys_dat = Date
            rs!sys_tim = Format(Time, "HH:MM:SS")
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
            If Trim(lblVanRef.Caption) <> "" Then
                rs("VAN_REF") = Trim(lblVanRef.Caption)
            End If
            If Trim(lblMERNum.Caption) <> "" Then
                rs("MER_NUM") = Trim(lblMERNum.Caption)
            End If

'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
            
            
            rs.Update
        End If
        End If
    Next
    Toolbar1.Buttons(12).Enabled = False
    If MOD1 = False Then MsgBox "New Receipt Number is " & INC, vbExclamation, "DAS Version 4.0"
    CANCELRECORD
 Else
  MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 1
    MS.SetFocus
'    rs.AddNew
'    rs("txn_idy") = INC
'    rs("txn_dat") = DTPicker1
'    rs("DBT_idy") = Text2
'    rs("txn_typ") = "RCT"
'    rs("clg_dat") = Date
'    rs("INV_IDY") = ""
'    rs.Update
End If

End If
If CHECKBILL = True Then
    con.Execute "Update inl where inv_dat = ctod('" & INVOICE_DATE & "') set sto_rmk='Y'"
    CHECKBILL = False
End If
'Toolbar1.Buttons(12).Enabled = False
'If MOD1 = False Then MsgBox "New Receipt Number is " & INC, vbExclamation, "DAS Version 4.0"
If check("PRT", "PRT_NME", "M_FRM_RCT") = True Then
con.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
con.Execute "update dir where dir_idy like 'REM' set dir_als = 'RCT'"
OpenReport CURDIR & "\vouch.RP1", "New Rct Voucher", 2, "vch->txn_idy='" & Trim(INC) & "'.and.vch->txn_typ='RCT'", 6
CANCELRECORD
End If
If Trim(MS.TextMatrix(1, 0)) <> "" Then
CANCELRECORD
End If
End Sub

Private Sub CANCELRECORD()
clrctr M_FRM_RCT
Label9 = "0"
Label7 = ""
Label6 = ""
SSTab1.Tab = 0
idnew = False
DTPicker1 = Date
DATE1 = Date
MOD1 = False
MS.Rows = 1
MS.AddItem ""
Toolbar1.Buttons(12).Enabled = False
Label2 = labdisp("IDY_RCT")
Text1.Enabled = True
Text1.SetFocus
lblMERNum.Caption = ""
lblVanRef.Caption = ""
End Sub

Private Sub newrecord()
CANCELRECORD
'Text1.Enabled = False
Text1.Enabled = True
Text1.SetFocus
DTPicker1.SetFocus
End Sub

Private Sub ACCOUNTLIST()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Check1.Value = 1 Then
rs.Open "select * from act where act_idy not like '" & Sincrement(Text2) & "' and act_nme like '" & Trim(Text3) & "%'", con, adOpenStatic
Else
rs.Open "select * from act where act_idy not like '" & Sincrement(Text2) & "' and act_ad4 like '" & Trim(Text3) & "%'", con, adOpenStatic
End If
List2.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = List2.ListItems.Add(, , rs("act_nme"))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        ls.ListSubItems.Add , , Trim(rs("ACT_TYP") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Account"
    List2.Visible = True
    Frame1.Visible = True
    List2.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Account Not Found, Define Account!", vbExclamation, "DAS Version 4.0"
    MS.SetFocus
    End If
End If
End Sub

Private Sub exitrecord()
Unload Me
End Sub

Private Sub voucherlist()
Dim rs As New ADODB.Recordset
Dim VDATE1, VDATE2 As String
If rs.State = 1 Then rs.Close
If CHECKTB("act.dbf,vch.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
rs.Open "select txn_idy,act_nme,act_ad4,txn_dat from vch,ACT where txn_typ = 'RCT' AND vch.DBT_idy = act.act_idy GROUP BY TXN_IDY", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list3.ListItems.CLEAR
    For A = 1 To rs.RecordCount
    Set ls = list3.ListItems.Add(, , Trim(rs(0)))
    VDATE1 = Format(rs(3), "MM/DD/YYYY")
    VDATE2 = Format(rs(3), "DD/MM/YYYY")
    ls.ListSubItems.Add , , Trim(VDATE2 & "")
    ls.ListSubItems.Add , , Trim(rs(1) & "")
    ls.ListSubItems.Add , , Trim(rs(2) & "")
    rs.MoveNext
    Next
    Label49 = "Select Receipt Voucher"
    SSTab1.Tab = 2
    list3.Visible = True
    Frame1.Visible = False
    list3.SetFocus
Else
    MsgBox "Receipt Not Found!", vbExclamation, "DAS Version 4.0"
    Text1.Enabled = True
    SendKeys "{HOME}+{END}"
    Text1.SetFocus
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
ElseIf Label49.Caption = "Select  Account" Then
    account
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
ElseIf Label49.Caption = "Select Account" Then
    account
End If
End Sub

Private Sub CHECK3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
If CHECKTB(PRT.dbf) = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Check3.Value = 0 Then
    con.Execute "delete from prt where prt_nme like 'M_FRM_RCT'"
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from prt", con, adOpenKeyset, adLockPessimistic
    rs.AddNew
        rs(0) = "M_FRM_RCT"
    rs.Update
End If
Picture1.Visible = False
SSTab1.Enabled = True
Toolbar1.Enabled = True
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub cmdXRef_Click()
'If Text1.Text = "" Then
'Text1.SetFocus
'Exit Sub
'End If
MerNum = lblMERNum.Caption
VanRef = lblVanRef.Caption
M_FRM_XRF.Show 1
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
MS.Text = DATE1
MS.Text = Format(MS.Text, "dd/mmm/yyyy")
DATE1.Visible = False
MS.SetFocus
MS.Col = MS.Col + 1
End If
End Sub
Private Sub date1_CloseUp()
DATE1.Visible = True
DATE1.SetFocus
End Sub

Private Sub DTPicker1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date using UP or DOWN ARROW KEYS"
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub DTPicker1_CloseUp()
DTPicker1.SetFocus
End Sub

Private Sub Form_Activate()
SETYEAR M_FRM_RCT
DTPicker1 = Date
DATE1 = Date
SSTab1.Tab = 0
SSTab1.TabEnabled(2) = False
If Text1.Enabled = True Then
Text1.SetFocus
Else
If DTPicker1.Enabled = True Then DTPicker1.SetFocus
End If
MS.ColWidth(6) = 0
lblVanRef.Caption = EFF_LIB1.VanRef
lblMERNum.Caption = EFF_LIB1.MerNum
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If M_FRM_RCT.ActiveControl.name = "Text1" And Trim(Text1.Text) = "" Then
    Unload Me
    Exit Sub
    End If
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        Text2.SetFocus
    ElseIf List2.Visible = True Then
        List2.Visible = False
        SSTab1.Tab = 1
        If Text3.Visible = True Then
            Text3.SetFocus
        Else
            MS.SetFocus
        End If
    ElseIf list3.Visible = True Then
        list3.Visible = False
        SSTab1.Tab = 0
        If Text1.Enabled = True Then
            Text1.SetFocus
        Else
            Text2.SetFocus
        End If
    ElseIf SSTab1.Tab = 1 Then
        Text3.Visible = False
        DATE1.Visible = False
        amount.Visible = False
        Label7 = "0.00"
        For A = 1 To MS.Rows - 1
        Label7 = convert1(Label7 + Val(MS.TextMatrix(A, 2)))
        Next
        SSTab1.Tab = 0
        saverecord
    Else
        saverecord
    End If
End If
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Label2 = labdisp("IDY_RCT")
MS.ColWidth(0) = 1250
MS.ColWidth(1) = 3000
MS.ColWidth(2) = 1100
MS.ColWidth(3) = 600
MS.ColWidth(4) = 1500
MS.ColWidth(5) = 6000 ' changed to 6000 by PNR as per prince chopra request
MS.ColWidth(7) = 1100
MS.ColWidth(8) = 0

MS.TextMatrix(0, 0) = "Account"
MS.ColAlignment(0) = 1
SCl MS, 0
MS.TextMatrix(0, 1) = "Account Name"
MS.TextMatrix(0, 2) = "      Amount"
SCl MS, 2
MS.TextMatrix(0, 3) = "Clg"
SCl MS, 3
MS.TextMatrix(0, 4) = "Dep Date"
MS.ColAlignment(4) = 1
SCl MS, 4
MS.TextMatrix(0, 5) = "Remarks"
SCl MS, 5
MS.TextMatrix(0, 7) = "Invoice ID"


If rs.State = 1 Then rs.Close
rs.Open "select * from act where (grp_clf like 'Cash %' or grp_clf like 'Bank %') ", con, adOpenStatic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , rs("act_nme"))
        ls.ListSubItems.Add , , rs("act_idy")
        ls.ListSubItems.Add , , rs("act_ad4") & ""
        rs.MoveNext
    Next
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from act ", con, adOpenStatic
If rs.RecordCount > 0 Then
    List2.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = List2.ListItems.Add(, , rs("act_nme"))
        ls.ListSubItems.Add , , rs("act_idy")
        ls.ListSubItems.Add , , rs("act_ad4") & ""
        rs.MoveNext
    Next
End If
MOD1 = False
idnew = False
End Sub

Private Sub Form_Unload(cancel As Integer)
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
EFF_LIB1.VanRef = ""
EFF_LIB1.MerNum = ""
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub



Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list2_DblClick()
List2_Keydown 13, 0
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text2 = list1.SelectedItem.ListSubItems.Item(1)
    Label6 = list1.SelectedItem
    If MOD1 = False Then Toolbar1.Buttons(12).Enabled = True
    list1.Visible = False
    SSTab1.Tab = 1
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

Private Sub List2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    List2.Visible = False
    Text3.Visible = False
    r = MS.Row
    MS.TextMatrix(MS.Row, 0) = List2.SelectedItem.ListSubItems(1)
    MS.TextMatrix(MS.Row, 1) = List2.SelectedItem
    MS.TextMatrix(MS.Row, 3) = "Y"
    MS.TextMatrix(MS.Row, 4) = Format(Date, "dd/mmm/yyyy")
    Set ls = List2.SelectedItem
    MS.TextMatrix(MS.Row, 6) = ls.ListSubItems(3)
    SSTab1.Tab = 1
    MS.SetFocus
    MS.Col = 2
    MS.Row = r
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

Private Sub list3_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list3.SortOrder = lvwAscending Then
list3.SortOrder = lvwDescending
ElseIf list3.SortOrder = lvwDescending Then
list3.SortOrder = lvwAscending
End If
list3.SortKey = CH.Index - 1
End Sub

Private Sub list3_DblClick()
List3_KeyDown 13, 0
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1 = list3.SelectedItem
    list3.Visible = False
    SSTab1.Tab = 0
    TEXT1_KeyDown 13, 0
End If
End Sub

Private Sub list4_KeyPress(K As Integer)
If K = 13 Then
    MS.Text = list4
    list4.Visible = False
    MS.SetFocus
End If
End Sub


Private Sub MS_EnterCell()
If MS.Col = 3 Then
    Label9 = "0.00"
    For A = 1 To MS.Rows - 1
        Label9 = convert1(Val(Label9) + Val(MS.TextMatrix(A, 2)))
    Next
    M_MDI.STS_BAR.Panels(1).Text = "Enter Clearings 'Y' or 'N'"
ElseIf MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Account"
ElseIf MS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Account Name"
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

Private Sub MS_LeaveCell()
DATE1.Visible = False
Text3.Visible = False
list4.Visible = False
amount.Visible = False
MS.Enabled = True
If MS.Visible = True Then MS.SetFocus
End Sub

Private Sub MS_KeyPress(K As Integer)
' TO AVOID ESCAPE GIVEN BELOW FUNCTION K=27 ALSO CHECKED BY PKB
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
            If K = vbKeySpace Then
            Else
            If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
            Else
            Text3.Visible = True
            Text3.Top = MS.Top + MS.CellTop
            Text3.Left = MS.CellLeft + MS.Left
            Text3.Width = MS.CellWidth
            Text3.Height = MS.CellHeight
            Text3 = Chr(K)
            Text3.Visible = True
            Text3.SetFocus
            Text3.SelStart = Len(Text3)
            End If
            End If
         End If
    End If
End If
If MS.Col = 7 And K <> 13 And K <> 27 Then
    If K = vbKeySpace Then
    Else
        If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
            Text3.Visible = True
            Text3.Top = MS.Top + MS.CellTop
            Text3.Left = MS.CellLeft + MS.Left
            Text3.Width = MS.CellWidth
            Text3.Height = MS.CellHeight
            Text3 = Chr(K)
            Text3.Visible = True
            Text3.SetFocus
            Text3.SelStart = Len(Text3)
        End If
    End If
End If
If MS.Col = 5 And K <> 27 Or MS.Col = 3 And K <> 27 Then
    If MS.Col = 3 Then
        If K <> 13 Then
            list4.Visible = True
            list4.Top = MS.Top + MS.CellTop
            list4.Left = MS.CellLeft + MS.Left
            list4.Width = MS.CellWidth
            list4.Height = MS.CellHeight
            list4 = Chr(K)
            list4.Visible = True
            list4.SetFocus
        Else
            MS.Col = MS.Col + 1
            MS.SetFocus
        End If
    Else
        If K <> 13 Then
            Text3.Visible = True
            Text3.Top = MS.Top + MS.CellTop
            Text3.Left = MS.CellLeft + MS.Left
            Text3.Width = MS.CellWidth
            Text3.Height = MS.CellHeight
            If K <> 8 Then Text3 = Chr(K)
            Text3.Visible = True
            Text3.SetFocus
            Text3.SelStart = Len(Text3)
        End If
    End If
ElseIf MS.Col = 2 Then
    If K = 13 Then
        MS.Col = 3
    ElseIf K > 47 And K < 58 Or K = 8 Or K = 46 Then
        amount.Text = MS.TextMatrix(MS.Row, 2)
        amount.Visible = True
        amount = ""
        amount.Top = MS.Top + MS.CellTop
        amount.Left = MS.CellLeft + MS.Left
        amount.Width = MS.CellWidth
        amount.Height = MS.CellHeight
        amount = Chr(K)
        amount.Visible = True
        amount.SetFocus
        amount.SelStart = Len(amount)
    Else
        K = 0
    End If
ElseIf MS.Col = 4 Then
    DATE1.Visible = True
    DATE1.Top = MS.Top + MS.CellTop
    DATE1.Left = MS.CellLeft + MS.Left
    DATE1.Width = MS.CellWidth
    DATE1.Height = MS.CellHeight
    DATE1.Visible = True
    DATE1.SetFocus
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


Private Sub Picture1_LostFocus()
Picture1.Visible = False
End Sub

Private Sub MS_Scroll()
If amount.Visible = True Then
    amount = ""
    amount.Visible = False
    MS.SetFocus
ElseIf list4.Visible = True Then
    list4.Visible = False
    MS.SetFocus
ElseIf Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    MS.SetFocus
ElseIf DATE1.Visible = True Then
    DATE1.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 And Text1.Enabled = False Then
    If DTPicker1.Enabled = True Then DTPicker1.SetFocus
ElseIf SSTab1.Tab = 1 Then
    If MS.Rows > 1 Then
        MS.SetFocus
        MS.Row = 1
        MS.Col = 0
    End If
ElseIf SSTab1.Tab = 0 Then
Text1.SetFocus
End If
If SSTab1.Tab = 2 Then
    'Toolbar1.Enabled = False
Else
    Toolbar1.Enabled = True
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Voucher Number"
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LBL
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If CHECKTB("vch.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If rs.State = 1 Then rs.Close
    If Len(Trim(Text1)) <> 0 Then
        If Len(Trim(Text1)) = 10 Then
            rs.Open "select * from vch where txn_idy like '" & (Text1) & "' and txn_typ='RCT'", con, adOpenStatic
        Else
            rs.Open "select * from vch where txn_idy like 'RCT" & Tincrement(Text1) & "' and txn_typ='RCT'", con, adOpenStatic
        End If
        If rs.RecordCount > 0 Then
            If Len(Trim(rs("VAN_REF"))) > 0 Or Len(Trim(rs("MER_NUM"))) > 0 Then
                MsgBox ("This Voucher has Van Ref " & Trim(rs("VAN_REF")) & " MER Reference " & Trim(rs("MER_NUM"))), vbInformation, "DAS Version 4.0"
            End If
            If Len(Trim(Text1)) <> 10 Then Text1 = "RCT" & Tincrement(Text1)
            Label7 = "0"
            Label9 = "0"
            MOD1 = True
            DTPicker1 = rs("txn_dat")
            Text2 = rs("DBT_idy")
            If RS1.State = 1 Then RS1.Close
            RS1.Open "select * from act where act_idy like '" & Text2 & "'", con, adOpenStatic
            If RS1.RecordCount > 0 Then Label6 = RS1("act_nme")
            MS.Rows = rs.RecordCount + 1
            For A = 1 To rs.RecordCount
                MS.TextMatrix(A, 0) = rs("CRT_idy") & ""
                If RS1.State = 1 Then RS1.Close
                RS1.Open "select * from act where act_idy like '" & Trim(rs("CRt_idy")) & "'", con, adOpenStatic
                If RS1.RecordCount > 0 Then MS.TextMatrix(A, 1) = RS1("act_nme")
                MS.TextMatrix(A, 2) = convert1(Val(rs("txn_amt") & ""))
                MS.TextMatrix(A, 3) = rs("clg_flg") & ""
                MS.TextMatrix(A, 4) = Format(rs("clg_dat") & "", "DD/MMM/YYYY")
                MS.TextMatrix(A, 5) = Trim(rs("txn_rmk") & "")
                
                If RS1.RecordCount > 0 Then MS.TextMatrix(A, 6) = Trim(RS1("ACT_TYP") & "")
                MS.TextMatrix(A, 7) = Trim(rs("INV_IDY") & "")
                MS.TextMatrix(A, 8) = Trim(rs("INV_TYP") & "")
                Label7 = convert1(AMT1(Val(Label7) + Val(rs("txn_amt") & "")))
                Label9 = convert1(Val(Label9) + Val(rs("txn_amt") & ""))

'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
        If IsNull(rs("VAN_REF")) = False Then lblVanRef.Caption = Trim(rs("VAN_REF"))
        If IsNull(rs("MER_NUM")) = False Then lblMERNum.Caption = Trim(rs("MER_NUM"))
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================


                rs.MoveNext
            Next
            Text1.Enabled = False
            Text1 = Sincrement(Text1)
            DTPicker1.SetFocus
        Else
            MsgBox "Receipt Not Found !", vbExclamation, "DAS Version 4.0"
            Text1.SetFocus
        End If
    Else
        voucherlist
    End If
End If
Exit Sub
LBL:
If err.Number = 35773 Then
    MsgBox "Invalid Voucher Selected for this Financial Year", vbExclamation
    CANCELRECORD
Else
    MsgBox err.Description
    CANCELRECORD
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Account"
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "select * from act where act_idy like '" & Sincrement(Text2) & "' and (grp_clf like 'Cash %' or grp_clf like 'Bank %')", con, adOpenStatic
    If rs.RecordCount > 0 Then
        Text2 = Sincrement(Text2)
        Label6 = rs("act_nme")
        SSTab1.Tab = 1
        MS.Col = 0
        MS.SetFocus
        If MOD1 = False Then Toolbar1.Buttons(12).Enabled = True
    Else
        account
    End If
ElseIf KeyCode = 38 Then
    If rs.State = 1 Then rs.Close
        rs.Open "select * from act where act_idy like '" & Sincrement(Text2) & "' and (grp_clf like 'Cash %' or grp_clf like 'Bank %')", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text2 = Sincrement(Text2)
            Label6 = rs("act_nme")
            DTPicker1.SetFocus
            If MOD1 = False Then Toolbar1.Buttons(12).Enabled = True
        Else
            MsgBox "Account Not Found, Define Account!", vbExclamation, "DAS Version 4.0"
            Text2.SetFocus
        End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub Text3_KeyDown(K As Integer, s As Integer)
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Then
    If MS.Col = 5 Then
        MS.Text = Trim(Text3)
        Text3.Visible = False
        Text3 = ""
        MS.SetFocus
    ElseIf MS.Col = 0 Then
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Sincrement(Text3) & "' and act_idy not like '" & Sincrement(Text2) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            r = MS.Row
            MS.TextMatrix(MS.Row, 0) = rs("act_IDY")
            MS.TextMatrix(MS.Row, 1) = rs("act_nme")
            MS.TextMatrix(MS.Row, 3) = "Y"
            MS.TextMatrix(MS.Row, 4) = Format(Date, "DD/MMM/YYYY")
            Text3.Visible = False
            Text3 = ""
            MS.SetFocus
            MS.Col = 2
            MS.Row = r
        Else
            ACCOUNTLIST
        End If
    ElseIf MS.Col = 7 Then
        If Trim(Text3) = "" Then
            MS.Text = ""
            Text3.Visible = False
            MS.SetFocus
        Else
            If rs.State = 1 Then rs.Close
            rs.Open "select * from inl where inv_idy = '" & Trim(Text3) & "' and txn_typ='SAL' and cst_idy = '" & Trim(MS.TextMatrix(MS.Row, 0)) & "'", con, adOpenKeyset, adLockPessimistic
            If rs.RecordCount = 0 Then
                MsgBox "Invoice Number not related to " & MS.TextMatrix(MS.Row, 1), vbExclamation
                Text3.SetFocus
            Else
                MS.Text = Trim(Text3)
                Text3.Visible = False
                MS.SetFocus
            End If
        End If
    End If
End If
End Sub


Private Sub text3_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0

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
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If list1.Visible = False And List2.Visible = False And list3.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then exitrecord
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_RCT.ActiveControl.name = "Text1" Then
            voucherlist
        ElseIf M_FRM_RCT.ActiveControl.name = "Text2" Then
            account
        ElseIf M_FRM_RCT.ActiveControl.name = "Text3" Then
            ACCOUNTLIST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_RCT.ActiveControl Is TextBox Then M_FRM_RCT.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
newrecord
ElseIf Button.KEY = "s" Then
saverecord
ElseIf Button.KEY = "e" Then
exitrecord
ElseIf Button.KEY = "p" Then
Printrecord
ElseIf Button.KEY = "crt" Then
SYS_PSW.Caption = "Password"
SYS_PSW.Show 1
convertbills
ElseIf Button.KEY = "pre" Then
Prefix.Caption = "Prefix Receipts"
Prefix.Show 1
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "ed" Then
    If M_FRM_RCT.ActiveControl.name = "Text1" Then
        Check1.Value = 0
        Check2.Value = 0
        Label49 = "Select Voucher"
        voucherlist
    ElseIf M_FRM_RCT.ActiveControl.name = "MS" And MS.Col = 0 Then
        Check1.Value = 0
        Check2.Value = 0
        Label49 = "Select Account"
        ACCOUNTLIST
    ElseIf M_FRM_RCT.ActiveControl.name = "Text2" Then
        Check1.Value = 0
        Check2.Value = 0
        Label49 = "Select  Account"
        account
    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
SSTab1.Enabled = False
'Toolbar1.Enabled = False
    If CHECKTB(PRT.dbf) = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    If check("prt", "prt_nme", "M_FRM_RCT") = True Then
        Check3.Value = 1
    Else
        Check3.Value = 0
    End If
    Check3.SetFocus
End If
End Sub

Public Sub account()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Check1.Value = 1 Then
rs.Open "select * from act where act_nme like '" & Trim(Text2) & "%' and (grp_clf like 'Cash %' or grp_clf like 'Bank %')", con, adOpenStatic
Else
rs.Open "select * from act where act_ad4 like '" & Trim(Text2) & "%' and (grp_clf like 'Cash %'or grp_clf like 'Bank %')", con, adOpenStatic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
    Set ls = list1.ListItems.Add(, , Trim(rs("act_nme") & ""))
    ls.ListSubItems.Add , , Trim(rs("act_idy"))
    ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
    rs.MoveNext
    Next
    Label49 = "Select Account"
    SSTab1.Tab = 2
    list1.Visible = True
    Frame1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Account Not Found, Define Account!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text2.SetFocus
    End If
End If
End Sub

Public Sub Printrecord()
M_FRM_PRV.Tag = "RCT"
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_RCT") = True Then M_FRM_PRV.Check3.Value = 1
M_FRM_PRV.Show 1
End Sub

Public Sub convertbills()
If CHECKTB("inl.dbf, act.dbf") = 1 Then
    MsgBox UCase(TNAME) & ".DBF  in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Me.Tag <> "Password" Then
Label9 = "0.00"
CHECKBILL = True
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select INL.CST_IDY,ACT.ACT_NME,INL.NET_AMT,INL.STO_RMK from inl,ACT where INL.inv_dat = ctod('" & Format(INVOICE_DATE, "mm/dd/yyyy") & "') AND INL.CST_IDY LIKE ACT.ACT_IDY AND INL.TXN_TYP IN ('LSL','SAL')", con, adOpenKeyset, adLockPessimistic
MS.Rows = 1
For A = 1 To rs.RecordCount
    If Trim(rs("STO_RMK") & "") <> "Y" Then
        MS.AddItem ""
        MS.TextMatrix(MS.Rows - 1, 0) = Trim(rs("CST_IDY") & "")
        MS.TextMatrix(MS.Rows - 1, 1) = Trim(rs("ACT_NME") & "")
        MS.TextMatrix(MS.Rows - 1, 2) = convert1(AMT1(Trim(rs("NET_AMT") & "")))
        MS.TextMatrix(MS.Rows - 1, 3) = "Y"
        MS.TextMatrix(MS.Rows - 1, 4) = Format(INVOICE_DATE, "DD/MM/YYYY")
        Label9 = convert1(Val(Label9) + Val(MS.TextMatrix(MS.Rows - 1, 2)))
    End If
    rs.MoveNext
Next
Label7 = Label9 & ""
If MS.Rows = 1 Then MS.Rows = 2
End If
End Sub
