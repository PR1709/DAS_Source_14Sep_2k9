VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Returnables - Transactions (Empties Additional Entries)"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_SIN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
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
            Picture         =   "M_FRM_SIN.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIN.frx":2E02
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
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      TabHeight       =   573
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
      TabCaption(0)   =   "&Voucher Details"
      TabPicture(0)   =   "M_FRM_SIN.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "MS"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPicker1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "shl_idy"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TEXT6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "m_flx_amt"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Shell &Details"
      TabPicture(1)   =   "M_FRM_SIN.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_ret"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Command3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "M_FRM_SIN.frx":3386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "LIST8"
      Tab(2).Control(2)=   "list1"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   975
         Left            =   6195
         TabIndex        =   24
         Top             =   810
         Width           =   1635
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   105
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   495
            Width           =   1215
         End
         Begin VB.CheckBox vanchk 
            Caption         =   "Van &Reference"
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
            Left            =   105
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   225
            Width           =   1380
         End
      End
      Begin VB.TextBox m_flx_amt 
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
         Height          =   285
         Left            =   3900
         MaxLength       =   6
         TabIndex        =   21
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TEXT6 
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
         Left            =   2370
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1351
         Width           =   1215
      End
      Begin VB.TextBox shl_idy 
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
         Left            =   2370
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   472
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Update Entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
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
         ItemData        =   "M_FRM_SIN.frx":33A2
         Left            =   2370
         List            =   "M_FRM_SIN.frx":33C4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   919
         Width           =   2655
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3855
         Left            =   -74520
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6800
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
            Text            =   "Market Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   6495
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   435
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   16580611
         CurrentDate     =   37258
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2835
         Left            =   1290
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2160
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   14
         FixedCols       =   3
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   1
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
      Begin MSComctlLib.ListView LIST8 
         Height          =   3855
         Left            =   -74520
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6800
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman "
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Route"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Reference"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   2370
         TabIndex        =   29
         Top             =   1800
         Width           =   5415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Txn Type"
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
         Left            =   1470
         TabIndex        =   23
         Top             =   979
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Party"
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
         Left            =   1305
         TabIndex        =   19
         Top             =   1419
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Updating entries please wait...."
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
         Left            =   -74880
         TabIndex        =   17
         Top             =   4800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   1995
         TabIndex        =   13
         Top             =   540
         Width           =   165
      End
      Begin VB.Label lbl_ret 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   12
         Top             =   4800
         Width           =   75
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   6000
         TabIndex        =   11
         Top             =   510
         Width           =   345
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   4080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Market"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Top             =   600
         Width           =   7785
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
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
               NumButtonMenus  =   10
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a"
                  Text            =   "&a Empties Stock in Hand"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b"
                  Text            =   "&b Shell Pack ROL COL Variance"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "h"
                  Text            =   "&c Party Pack ROL COL Variance"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "c"
                  Text            =   "&d Party ROL COL Summary"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "d"
                  Text            =   "Shell Transaction  Account "
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "e"
                  Text            =   "&e Deposits Account"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "f"
                  Text            =   "Shell Deposit Party"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "g"
                  Text            =   "Shells Breakage"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BRK_REP"
                  Text            =   "&f Breakages Report"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EMP_REG"
                  Text            =   "&g Empties Register"
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
            Key             =   "VLK"
            Object.ToolTipText     =   "Click Here to Freeze Van Challan for Transactions Entry"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
                  Key             =   "B"
                  Text            =   "Prefix Product ID"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
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
      Begin VB.CommandButton Command1 
         Caption         =   "EDA &Print Flag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7125
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Click Here to Print EDA (Empty Dispatch Advice) After Sorted Glass Entry"
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "M_FRM_SIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_REP_CHK As String
Dim rsp As New ADODB.Recordset

Private Sub Combo1_Click()
If Combo1.ListIndex = 3 Or Combo1.ListIndex = 0 Then
    Text6.Enabled = False
Else
    Text6.Enabled = True
End If
End Sub

Private Sub Combo1_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    If Combo1.ListIndex = 0 Then
        MsgBox "Select Shell Transaction Type, Press Down Arrow to Scroll!", vbExclamation, "DAS Version 4.0"
        Combo1.SetFocus
    Else
        If Text6.Enabled = True Then
            Text6.SetFocus
        ElseIf Combo1.ListIndex = 3 Then
            LoadProducts
            MS.SetFocus
            MS.Col = 3
        Else
            LoadProducts
            MS.SetFocus
            MS.Col = 3
        End If
    End If
ElseIf K = 38 And Shift = 1 Then
    DTPicker1.SetFocus
End If
End Sub

Private Sub Combo1_Validate(cancel As Boolean)
If Combo1.ListIndex = 0 Then
    cancel = True
    MsgBox "Select Empty Transaction!", vbInformation, "DAS Version 4.0"
    Combo1.SetFocus
End If
End Sub

Private Sub Command1_Click()
EMP_SHL_SOR "SIN_SHL"
End Sub

Private Sub Command3_Click()
Me.MousePointer = vbHourglass
Label11.Visible = True
For A = 1 To MS.Rows - 1
    MS.TextMatrix(A, 7) = Val(MS.TextMatrix(A, 3)) - (Val(MS.TextMatrix(A, 9)) + Val(MS.TextMatrix(A, 13)) + Val(MS.TextMatrix(A, 5)))
    MS.TextMatrix(A, 8) = Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 10)) + Val(MS.TextMatrix(A, 14)) + Val(MS.TextMatrix(A, 6)))
    MS.TextMatrix(A, 11) = (Val(MS.TextMatrix(A, 3)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 7)) + Val(MS.TextMatrix(A, 9)) + Val(MS.TextMatrix(A, 13))))
    MS.TextMatrix(A, 12) = (Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 8)) + Val(MS.TextMatrix(A, 10)) + Val(MS.TextMatrix(A, 14))))
Next
Label11.Visible = False
If MS.Rows > 1 Then Command3.Enabled = False
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Combo1.Enabled = False Then
        If Text6.Enabled = True Then
            Text6.SetFocus
        Else
            MS.SetFocus
        End If
    Else
        Combo1.SetFocus
    End If
End If
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabVisible(1) = False
SSTab1.TabEnabled(2) = False
MS.ColWidth(0) = 0
If shl_idy.Enabled = True Then
    SSTab1.Tab = 0
    shl_idy.SetFocus
Else
    SSTab1.Tab = 0
End If
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_SIN.ActiveControl Is TextBox Then M_FRM_SIN.ActiveControl.Text = Trim(FUNKEY(K))
If K = 83 And s = 2 Then
    save
ElseIf K = 78 And s = 2 Then
    newrec
ElseIf K = 81 And s = 2 Then
    cancel
ElseIf K = 69 And s = 2 Then
    V_SHL_LST
ElseIf K = 88 And s = 2 Then
    exitfrm
ElseIf K = 112 Then
    M_HELP
ElseIf K = 80 And s = 2 Then

End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        If Trim(Label20) = "Select Party" Then
            Text6.SetFocus
        Else
            SSTab1.Tab = 0
            If shl_idy.Enabled = True Then
                shl_idy.SetFocus
            Else
                DTPicker1.SetFocus
            End If
        End If
    ElseIf LIST8.Visible = True Then
        LIST8.Visible = False
        SSTab1.Tab = 0
        cancel
    ElseIf shl_idy.Enabled = False Then
        save
    Else
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
Label12.Caption = ""
SETYEAR M_FRM_SIN
Combo1.ListIndex = 0
MS.FixedCols = 3
MS.MergeCells = flexMergeFree
MS.Cols = 22
MS.ColWidth(0) = 1100
MS.ColWidth(1) = 1830
MS.ColWidth(2) = 2000
For A = 3 To 15
    MS.ColWidth(A) = 800
Next
MS.MergeCol(0) = True
MS.MergeCol(1) = True
MS.ColAlignment(1) = 1
MS.ColAlignment(2) = 1
MS.TextMatrix(0, 0) = "Voucher No"
MS.TextMatrix(0, 1) = "Shell Name"
MS.TextMatrix(0, 2) = "SKU Pack - UPC"
MS.TextMatrix(0, 3) = "Shells"
SCl MS, 3
MS.TextMatrix(0, 4) = "Glass"
SCl MS, 4
MS.TextMatrix(0, 5) = "Recd.Shl"
MS.TextMatrix(0, 6) = "Recd.Bot"
MS.TextMatrix(0, 7) = "Ret.Shl"
MS.TextMatrix(0, 8) = "Ret.Bot"
MS.TextMatrix(0, 9) = "Shls Brk"
SCl MS, 9
MS.TextMatrix(0, 10) = "Gls Brk"
SCl MS, 10
MS.TextMatrix(0, 11) = "Dif.Shl"
MS.TextMatrix(0, 12) = "Dif.Bot"
MS.TextMatrix(0, 13) = "Shls Dep"
SCl MS, 13
MS.TextMatrix(0, 14) = "Gls Dep"
SCl MS, 14
MS.TextMatrix(0, 15) = "Amount Rs"
SCl MS, 15
MS.ColWidth(3) = 1300
MS.ColWidth(4) = 1300
MS.ColWidth(13) = 1300
MS.ColWidth(14) = 1300
MS.ColWidth(15) = 1300
MS.ColWidth(5) = 0
MS.ColWidth(6) = 0
MS.ColWidth(7) = 0
MS.ColWidth(8) = 0
MS.ColWidth(9) = 0
MS.ColWidth(10) = 0
MS.ColWidth(11) = 0
MS.ColWidth(12) = 0

MS.ColWidth(16) = 0
MS.ColWidth(17) = 0
MS.ColWidth(18) = 0
MS.ColWidth(19) = 0
MS.ColWidth(20) = 0
MS.ColWidth(21) = 0
For A = 3 To 15
    MS.ColAlignment(A) = 7
Next
MS.TextMatrix(0, 16) = "prdidy"
MS.TextMatrix(0, 17) = "shlidy"
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub List1_keypress(KeyAscii As Integer)
If KeyAscii = 27 And shl_idy <> "" Then
Label12.Caption = ""
'cancel
'Text6.SetFocus
shl_idy.SetFocus
Else

End If
End Sub

Private Sub list1_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If list1.ListItems.count > 0 Then
        If Label20 = "Select Party" Then
            Label7 = list1.SelectedItem.ListSubItems.Item(1)
            Text6 = list1.SelectedItem.ListSubItems.Item(1)
            Label12 = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            If Combo1.ListIndex = 1 Or Combo1.ListIndex = 2 Or Combo1.ListIndex = 8 Or Combo1.ListIndex = 9 Then
                Frame1.Enabled = True
                vanchk.SetFocus
            Else
                LoadProducts
                MS.SetFocus
            End If
        ElseIf Label20 = "Select Shell" Then
            dispshl list1.SelectedItem
        Else
            list1.Visible = False
            shl_idy = list1.SelectedItem
            DISPDETAILS
        End If
    End If
End If
End Sub

Private Sub LIST8_KeyPress(K As Integer)
If K = 13 Then
    Text1 = LIST8.SelectedItem
    SSTab1.Tab = 0
    If Trim(shl_idy) = "" Then
    LoadProducts
    End If
    MS.SetFocus
'''ElseIf K = 27 Then
'''    LIST8.Visible = False
'''    SSTab1.Tab = 0
'''    Text1.Enabled = True
'''    Text1.SetFocus
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(K As Integer)
If K = 13 Then
    If MS.Col = 3 Or MS.Col = 13 Or MS.Col = 9 Then
        If MS.Col = 3 Or MS.Col = 13 Or MS.Col = 9 Then MS.Text = Int(Val(M_FLX_AMT))
        If MS.Col = 4 Or MS.Col = 14 Then
            If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
            If MS.Col = 4 Then MS.Col = 3
            If MS.Col = 14 Then MS.Col = 13
        ElseIf MS.Col = 3 Then
            MS.TextMatrix(MS.Row, 4) = QTY(Val(MS.Text)) 'MTS(Val(MS.Text), GetValue("SUB_UNT", "PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 16)) & "'"))
            MS.Col = MS.Col + 1
        ElseIf MS.Col = 13 Then
            MS.TextMatrix(MS.Row, 14) = QTY(Val(MS.Text)) 'MTS(Val(MS.Text), GetValue("SUB_UNT", "PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 16)) & "'"))
            MS.Col = MS.Col + 1
        ElseIf MS.Col = 9 Then
            MS.TextMatrix(MS.Row, 10) = QTY(Val(MS.Text)) 'MTS(Val(MS.Text), GetValue("SUB_UNT", "PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 16)) & "'"))
            MS.Col = MS.Col + 1
        End If
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 4 Or MS.Col = 14 Or MS.Col = 10 Then
        If CSUBUNT(Val(M_FLX_AMT), Val(MS.TextMatrix(MS.Row, 0))) = True Then
            MS.Text = QTY(Val(M_FLX_AMT))
            M_FLX_AMT.Visible = False
            MS.SetFocus
        End If
    Else
        MS.Text = convert(M_FLX_AMT)
        M_FLX_AMT.Visible = False
        MS.SetFocus
    End If
End If
End Sub

Private Sub MS_EnterCell()
If MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Voucher Number"
ElseIf MS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Name"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Shells"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Bottles"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Previous Received Shells"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Previous Received Bottles"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Current Receivable Shells"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Current Receivable Bottles"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Breakage Shells"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Breakage Bottles"
ElseIf MS.Col = 11 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Diffrence Shells"
ElseIf MS.Col = 12 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Diffrence Bottles"
ElseIf MS.Col = 13 Then
    M_MDI.STS_BAR.Panels(1).Text = "Deposit Shells"
ElseIf MS.Col = 14 Then
    M_MDI.STS_BAR.Panels(1).Text = "Deposit Bottles"
ElseIf MS.Col = 15 Then
    M_MDI.STS_BAR.Panels(1).Text = "Deposit Amount"
End If
End Sub

Private Sub MS_KeyPress(K As Integer)
If MS.Row > 0 Then If MS.Col <> 1 Then If Trim(MS.TextMatrix(1, 1)) = "" Then Exit Sub
If (MS.Col = 3 Or MS.Col = 4 Or MS.Col = 13 Or MS.Col = 14 Or MS.Col = 15 Or MS.Col = 9 Or MS.Col = 10) And MS.Row > 0 Then
    If MS.Col = 15 Then
        If Val(MS.TextMatrix(MS.Row, 13)) Or Val(MS.TextMatrix(MS.Row, 14)) > 0 Then
        Else
            Exit Sub
        End If
    End If
    If K > 47 And K < 58 Or K = 46 Then
        M_FLX_AMT = ""
        M_FLX_AMT.MaxLength = 7
        M_FLX_AMT.Visible = True
        M_FLX_AMT.Top = MS.Top + MS.CellTop
        M_FLX_AMT.Left = MS.CellLeft + MS.Left
        M_FLX_AMT.Width = MS.CellWidth
        M_FLX_AMT.Height = MS.CellHeight
        M_FLX_AMT.Visible = True
        M_FLX_AMT.SetFocus
        If K = vbKeySubtract Then
            K = 0
        Else
            M_FLX_AMT = Chr(K)
            M_FLX_AMT.SelStart = Len(M_FLX_AMT)
        End If
    ElseIf K = 13 Then
        If MS.Col = 4 Then
            If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
            MS.Col = 3
        ElseIf MS.Col = 3 Then
            MS.Col = 4
        End If
        If MS.Col = 14 Then
            If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
            MS.Col = 13
        ElseIf MS.Col = 13 Then
            MS.Col = 14
        End If

    End If
ElseIf MS.Col = 1 Then
    Text3 = ""
    If Trim(MS.Text) <> "" Then Exit Sub
    If K = 13 Then
        DISPLIST
    ElseIf K <> 27 Then
        If K = vbKeySpace Then
        Else
            If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
            Else
                Text3.Visible = True
                Text3.Top = MS.Top + MS.CellTop
                Text3.Left = MS.CellLeft + MS.Left
                Text3.Width = MS.CellWidth
                Text3.Height = MS.CellHeight
                Text3.Visible = True
                Text3.SetFocus
                If K <> 13 Then Text3 = Chr(K)
                Text3.SelStart = Len(Text3)
            End If
        End If
    ElseIf K = 27 Then
        save
    End If
End If
End Sub

Private Sub MS_LeaveCell()
Text3.Visible = False
M_FLX_AMT.Visible = False
End Sub

Private Sub MS_RowColChange()
lbl_ret = Trim(MS.TextMatrix(MS.Row, 20))
End Sub

Private Sub MS_Scroll()
If Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    MS.SetFocus
ElseIf M_FLX_AMT.Visible = True Then
    M_FLX_AMT = ""
    M_FLX_AMT.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub shl_idy_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("snv.dbf,prd.dbf,shl.dbf,act.dbf,inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
txntyp = ""
Dim ESHL_IDY As String
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
ESHL_IDY = "SHT" & Tincrement(shl_idy)
txntyp = GetValue("TXN_TYP", "SNV WHERE TXN_IDY LIKE '" & ESHL_IDY & "'")
'RS.Open "select * from snv where snv.txn_idy like '" & Sincrement(shl_idy) & "' ORDER BY INV_IDY", CON, adOpenKeyset, adlockpessimistic
rs.Open "select * from snv where snv.txn_idy like '" & Trim(shl_idy) & "'", con, adOpenKeyset, adLockPessimistic
MS.Rows = 1
' MsgBox "" & Sincrement(shl_idy)
If rs.RecordCount > 0 Then
'MsgBox "OK" & RS.RecordCount
    DTPicker1 = rs!txn_dat
    If Trim(rs!TXN_TYP) = "OPS" Then
         Combo1.ListIndex = 3
    ElseIf Trim(rs!TXN_TYP) = "AIN" Then
        Combo1.ListIndex = 2
    ElseIf Trim(rs!TXN_TYP) = "AOU" Then
        Combo1.ListIndex = 1
    ElseIf Trim(rs!TXN_TYP) = "RCT" Then
        Combo1.ListIndex = 4
    ElseIf Trim(rs!TXN_TYP) = "PMT" Then
        Combo1.ListIndex = 5
    ElseIf Trim(rs!TXN_TYP) = "DPR" Then
        Combo1.ListIndex = 6
    ElseIf Trim(rs!TXN_TYP) = "DPC" Then
        Combo1.ListIndex = 7
    ElseIf Trim(rs!TXN_TYP) = "DPI" Then
        Combo1.ListIndex = 8
    ElseIf Trim(rs!TXN_TYP) = "DPO" Then
        Combo1.ListIndex = 9
    End If
    If Combo1.ListIndex <= 3 And Combo1.ListIndex > 0 Then
        MS.ColWidth(3) = 1300
        MS.ColWidth(4) = 1300
        MS.ColWidth(9) = 1300
        MS.ColWidth(10) = 1300
        MS.ColWidth(13) = 0
        MS.ColWidth(14) = 0
        MS.ColWidth(15) = 0
    ElseIf Combo1.ListIndex = 4 Or Combo1.ListIndex = 5 Then
        MS.ColWidth(3) = 1300
        MS.ColWidth(4) = 1300
        MS.ColWidth(9) = 0
        MS.ColWidth(10) = 0
        MS.ColWidth(13) = 0
        MS.ColWidth(14) = 0
        MS.ColWidth(15) = 0
    Else
        MS.ColWidth(3) = 0
        MS.ColWidth(4) = 0
        MS.ColWidth(9) = 0
        MS.ColWidth(10) = 0
        MS.ColWidth(13) = 1300
        MS.ColWidth(14) = 1300
        MS.ColWidth(15) = 1300
    End If

'''    If Combo1.ListIndex <> 3 Then
'''        Text6.Enabled = True
'''        Label7 = Trim(rs!cst_idy)
'''        Label1.Caption = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(rs!cst_idy) & "'")
'''        Text6.Text = Trim(rs!cst_idy)
'''    End If
'''    Combo1.Enabled = False
End If
'''For A = 1 To rs.RecordCount
'''    MS.AddItem ""
'''    MS.TextMatrix(A, 0) = Val(rs!prd_upc)
'''    MS.TextMatrix(A, 1) = GetValue("SHL_NME", "SHL WHERE SHL_IDY LIKE '" & Trim(rs!shl_idy) & "'")
'''    MS.TextMatrix(A, 2) = Trim(rs!PRD_PCK) & " - " & Val(rs!prd_upc)
'''    If Combo1.ListIndex = 4 Then
'''        'MS.TextMatrix(A, 3) = Trim(RS!i_shl_nos)
'''        'MS.TextMatrix(A, 4) = Trim(RS!i_pck_nos)
'''        'MS.TextMatrix(A, 4) = QTY(STM(Val(Trim(RS!i_pck_nos)), RS!PRD_UPC))
'''        MS.TextMatrix(A, 3) = Trim(rs!d_shl_nos & "") 'Added by PNR
'''        MS.TextMatrix(A, 4) = QTY(STM(Val(Trim(rs!d_pck_nos & "")), rs!prd_upc)) 'Added by PNR
'''    Else
'''        MS.TextMatrix(A, 3) = Trim(rs!d_shl_nos & "")
'''        MS.TextMatrix(A, 4) = QTY(STM(Val(Trim(rs!d_pck_nos & "")), rs!prd_upc))
'''    End If
'''    MS.TextMatrix(A, 9) = Val(rs!BRK_SHL)
'''    MS.TextMatrix(A, 10) = QTY(STM(Val(rs!BRK_PCK), Val(MS.TextMatrix(A, 0))))
'''    MS.TextMatrix(A, 13) = Trim(rs!dp_shl_nos & "")
'''    MS.TextMatrix(A, 14) = QTY(STM(Val(Trim(rs!dp_pck_nos)), rs!prd_upc))
'''    MS.TextMatrix(A, 15) = convert(Val(rs!dep_amt) & "")
'''    MS.TextMatrix(A, 16) = Trim(rs!PRD_PCK)
'''    MS.TextMatrix(A, 17) = Trim(rs!shl_idy)
'''    MS.TextMatrix(A, 18) = Format(rs!INV_DAT, "DD/MMM/YYYY")
'''    MS.TextMatrix(A, 19) = Trim(rs!TXN_TYP)
'''    MS.TextMatrix(A, 20) = Trim(rs!prd_idy) 'Added by PNR - field entry required
'''    cst_idy = Trim(rs!cst_idy)
'''    Text1 = Trim(rs!inv_idy)
'''    If Len(Trim(Text1)) <> 0 Then vanchk.Value = 1
'''    rs.MoveNext
'''Next
'''CST_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(cst_idy) & "'")
'''If rs.State = 1 Then rs.Close
'''rs.Open "select * from inl,act where inl.cst_idy like act.act_idy and inv_idy in(select inv_idy from snv where txn_idy like '" & Trim(shl_idy) & "' group by inv_idy)", CON, adOpenKeyset, adlockpessimistic
'''For A = 1 To rs.RecordCount
'''    For b = 1 To MS.Rows - 1
'''        If Trim(rs!inv_idy) = Trim(MS.TextMatrix(b, 0)) Then
'''            MS.TextMatrix(b, 20) = Trim(rs!ACT_NME)
'''            MS.TextMatrix(b, 21) = Trim(rs!cst_idy)
'''        End If
'''    Next
'''    rs.MoveNext
'''Next
'''er:
'''shl_idy.Enabled = False
'''SSTab1.Tab = 0
'''DTPicker1.SetFocus
'''Command3.Enabled = False
End If
End Sub
Private Sub shl_idy_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii = 13 Then
    If check("snv", "txn_idy", "SHT" & Tincrement(shl_idy)) = True Then
        shl_idy = "SHT" & Tincrement(shl_idy)
        DISPDETAILS
    Else
        V_SHL_LST
    End If
End If
End Sub

Public Sub DISPLIST()
If CHECKTB("shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
list1.View = lvwReport
With list1
    .ColumnHeaders.Add 1, , "Shell ID", 1600
    .ColumnHeaders.Add 2, , "Shell Name", 3600
    .ColumnHeaders.Add 3, , "Shell Type", 1600
    .ColumnHeaders.Add 4, , "Unit Rate", 1600, 1
If RS1.State = 1 Then RS1.Close
    RS1.Open "SELECT DISTINCT(SHL_IDY),SHL_NME,SHL_TYP,UNT_RAT,PRD_IDY FROM SHL WHERE SHL_TYP='Shell' AND SHL_NME LIKE '" & Trim(Text3) & "%'", con, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount > 0 Then
        .Visible = True
        Label20 = "Select Shell"
        Label20.Visible = True
        .SetFocus
        SSTab1.Tab = 2
        .ListItems.CLEAR
        RS1.MoveFirst
        For i = 0 To RS1.RecordCount - 1
            Set ls = .ListItems.Add(, , RS1.Fields(0))
            ls.ListSubItems.Add , , Trim(RS1.Fields("shl_nme"))
            ls.ListSubItems.Add , , Trim(RS1.Fields("shl_typ"))
            ls.ListSubItems.Add , , convert(Val(RS1.Fields("unt_rat") & ""))
            RS1.MoveNext
        Next
    Else
        MsgBox "Shell Not Found ! Define Shell!", vbExclamation, "DAS Version 4.0"
        Text3.Visible = False
        MS.SetFocus
    End If
End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    If shl_idy.Enabled = True Then
        shl_idy.SetFocus
    Else
        DTPicker1.SetFocus
    End If
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If CHECKTB("van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii = 13 Then
    If Toolbar1.Buttons(10).Value = tbrPressed And Len(Trim(Text1)) = 10 Then
        MS.SetFocus
        MS.Col = 3
    Else
        If Len(Trim(Text1)) = 0 Then
            'MS.SetFocus
            'MS.Col = 3
            M_VAN_LST
            
        Else
            Dim rs As New ADODB.Recordset
            rs.Open "select * from van where inv_idy = 'RSI" & Tincrement(Text1) & "' AND VAN_CLS='S' AND VAN_CHK='V'", con, adOpenKeyset, adLockPessimistic
            If rs.RecordCount > 0 Then
                Text1 = "RSI" & Tincrement(Text1)
                LoadProducts
                MS.SetFocus
                MS.Col = 3
            Else
                M_VAN_LST
            End If
        End If
    End If
End If
End Sub
Public Sub M_VAN_LST()
If CHECKTB("van.dbf,smn.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim rsVan As New ADODB.Recordset
rsVan.Open "SELECT * FROM VAN,SMN WHERE VAN.SMN_IDY like SMN.SMN_IDY and VAN_CLS='S' AND VAN_CHK = 'V' GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
If rsVan.RecordCount > 0 Then
    LIST8.ListItems.CLEAR
    For A = 1 To rsVan.RecordCount
        Set ls = LIST8.ListItems.Add(, , Trim(rsVan!INV_IDY))
        ls.ListSubItems.Add , , Format(rsVan!INV_DAT, "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(rsVan("SMN_NME") & "")
        ls.ListSubItems.Add , , Trim(rsVan("ROU_IDY") & "")
        ls.ListSubItems.Add , , Trim(rsVan("MNL_REF") & "")
        If rsVan("VAN_CLS") = "N" Then
            ls.ListSubItems.Add , , "Open"
        ElseIf rsVan("VAN_CLS") = "S" Then
            ls.ListSubItems.Add , , "Returns Posted"
        Else
            ls.ListSubItems.Add , , "Closed"
        End If
        rsVan.MoveNext
    Next
    SSTab1.Tab = 2
    Label20 = "Select Van Reference"
    SSTab1.Tab = 2
    LIST8.Visible = True
    LIST8.ListItems(1).Selected = True
    LIST8.SetFocus
Else
    MsgBox "Van Challan Reference Not Found!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    SendKeys "{home}+{end}"
    M_VAN_IDY.SetFocus
End If
End Sub
Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
    If CHECKTB("shl.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
    If check("SHL", "SHL_IDY", Sincrement(Trim(Text3))) = True Then
        dispshl Sincrement(Trim(Text3))
    Else
        DISPLIST
    End If
End If
End Sub
Private Sub Text6_Change()
If Len(Trim(Text6)) = 0 Then
    Label7 = ""
End If
End Sub
Private Sub Text6_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim EACTIDY As String
EACTIDY = Sincrement(Text6)
If check("ACT", "ACT_IDY", Sincrement(Text6)) = True Then
    Label7 = Sincrement(Text6)
    Text6 = EACTIDY
    
    Label12.Caption = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(EACTIDY) & "'")
    If Trim(shl_idy) = "" Then
        If Combo1.ListIndex = 1 Or Combo1.ListIndex = 2 Or Combo1.ListIndex = 8 Or Combo1.ListIndex = 9 Then
            Frame1.Enabled = True
            vanchk.SetFocus
        Else
            LoadProducts
            MS.SetFocus
        End If
        
    Else
        MS.SetFocus
    End If
    Exit Sub
End If

Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select ACT_IDY,MAR_NME,ACT_NME from ACT left join MAR on ACT.MAR_IDY like MAR.MAR_IDY where ACT_NME like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT' or act_typ like 'PUR') and mar.mar_idy like act.mar_idy", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
list1.ColumnHeaders.Add , , "Party Name"
list1.ColumnHeaders.Add , , "ID"
list1.ColumnHeaders.Add , , "Market"
list1.ColumnHeaders.Item(1).Width = 3500
list1.ColumnHeaders.Item(3).Width = 2000
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy") & "")
        ls.ListSubItems.Add , , Trim(rs("mar_nme") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label20 = "Select Party"
    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Party Not Found, Define Party!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text6.Enabled = True
        Text6.SetFocus
    End If
End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "s" Then
    If shl_idy <> "" Then
        save
    Else
        cancel
    End If
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    exitfrm
ElseIf Button.KEY = "ed" Then
    V_SHL_LST
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "d" Then
    'delete
ElseIf Button.KEY = "pre" Then
    'M_FRM_PRE.Caption = "Salesmen ID Prefix"
    'M_FRM_PRE.Show 1
    'M_MDI.STS_BAR.Panels(1).Text = "Set Prefix for Salesmen ID"
ElseIf Button.KEY = "p" Then
    'Openreport CURDIR & "\lists.RP1", "Salesmen list", 0, "", 6
ElseIf Button.KEY = "pr" Then
ElseIf Button.KEY = "VLK" Then
    Button.Value = tbrPressed
ElseIf Button.KEY = "crt" Then
    'CREATE
End If
End Sub
Public Sub cancel()
Combo1.Enabled = True
Label12.Caption = ""
MS.Rows = 1
MS.Rows = 2
MS.FixedCols = 3
Label4 = ""
Label5 = ""
Label6 = ""
Label7 = ""
Label8 = ""
Label10 = ""
lbl_ret = ""
shl_idy = ""
shl_idy.Enabled = True
Text6 = ""
SSTab1.Tab = 0
Command3.Enabled = True
Combo1.ListIndex = 0
Combo1.Enabled = True
SSTab1.Tab = 0
M_FLX_AMT.Visible = False
Text3.Visible = False
Frame1.Enabled = False
Text1.Enabled = False
If Toolbar1.Buttons(10).Value = tbrPressed Then
Else
Text1 = ""
End If
vanchk.Value = 0
shl_idy.SetFocus
End Sub
Public Sub save()
If MsgBox("Confirm Transaction Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If CHECKTB("act.dbf,snv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If Combo1.ListIndex = 0 Then
    MsgBox "Cannot Save, No Empty Transaction Values!", vbInformation, "DAS Version 4.0"
    cancel
    Exit Sub
End If
If Combo1.ListIndex = 1 Or Combo1.ListIndex = 2 Or Combo1.ListIndex = 4 Or Combo1.ListIndex = 5 Then
    If check("ACT", "ACT_IDY", Trim(Label7)) = False Then
        MsgBox "Party - Account Not Found!", vbExclamation, "DAS Version 4.0"
        Text6.Enabled = True
        Text6.SetFocus
        Label7 = ""
        Exit Sub
    End If
End If
Me.MousePointer = vbHourglass
Me.Refresh
Dim rs As New ADODB.Recordset

If Len(Trim(shl_idy)) = 0 Then
    If MS.Rows > 1 Then
'''''        If Len(Trim(MS.TextMatrix(MS.Row, 3))) > 0 Then
            IDY = "SHT" & T7increment("idy_sht")
            MsgBox "New Empty-Shell Txn ID is " & IDY, vbExclamation, "DAS Version 4.0"
'''''        End If
    End If
Else
    TxnMoveBackup "SNV", "SBK", Trim(shl_idy), ""
    con.Execute "delete from snv where txn_idy like '" & Trim(shl_idy) & "'"
    If rs.State = 1 Then rs.Close
    IDY = Trim(shl_idy)
End If

If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM SNV WHERE INV_IDY LIKE ''", con, adOpenKeyset, adLockPessimistic
'''''If Len(Trim(MS.TextMatrix(MS.Row, 3))) > 0 Or Len(Trim(MS.TextMatrix(MS.Row, 4))) > 0 Then
    For A = 1 To MS.Rows - 1
            rs.AddNew
            rs!TXN_IDY = IDY
            rs!txn_dat = DTPicker1
            rs!I_shl_nos = 0
            rs!I_pck_nos = 0
            rs!INV_IDY = Trim(Text1)
            rs!BRK_SHL = Val(MS.TextMatrix(A, 9))
            rs!BRK_PCK = MTS(Val(MS.TextMatrix(A, 10)), MS.TextMatrix(A, 0))
            If Combo1.ListIndex = 4 Then
                'RS!i_shl_nos = Val(MS.TextMatrix(A, 3)) 'Code need to be Checked
                'RS!i_pck_nos = MTS(Val(MS.TextMatrix(A, 4)), MS.TextMatrix(A, 0)) 'Code need to be Checked
                rs!R_shl_nos = 0
                rs!R_pck_nos = 0
                rs!d_shl_nos = Val(MS.TextMatrix(A, 3)) 'Added by PNR
                rs!d_pck_nos = MTS(Val(MS.TextMatrix(A, 4)), MS.TextMatrix(A, 0)) 'Added by PNR
            Else
                rs!I_shl_nos = 0
                rs!I_pck_nos = 0
                rs!d_shl_nos = Val(MS.TextMatrix(A, 3))
                rs!d_pck_nos = MTS(Val(MS.TextMatrix(A, 4)), MS.TextMatrix(A, 0))
            End If
            rs!dp_shl_nos = 0
            rs!dp_pck_nos = 0
            If Val(MS.TextMatrix(A, 13)) > 0 Then
                rs!d_shl_nos = Val(MS.TextMatrix(A, 13))
                rs!d_pck_nos = MTS(Val(MS.TextMatrix(A, 14)), MS.TextMatrix(A, 0))
                rs!dp_shl_nos = Val(MS.TextMatrix(A, 13))
                rs!dp_pck_nos = MTS(Val(MS.TextMatrix(A, 14)), MS.TextMatrix(A, 0))
            End If
            If Trim(Text1) <> "" Then
            Dim RSIRS As New ADODB.Recordset
            If RSIRS.State = 1 Then RSIRS.Close
            RSIRS.Open "SELECT * FROM VAN WHERE INV_IDY LIKE '" & Trim(Text1) & " '", con, adOpenKeyset, adLockPessimistic
            If RSIRS.RecordCount > 0 Then
            RSIRS.MoveFirst
            rs!SMN_IDY = RSIRS("SMN_IDY")
            rs!rou_idy = RSIRS("rou_idy")
            End If
            Else
            Dim rdyrs As New ADODB.Recordset
            If rdyrs.State = 1 Then rdyrs.Close
            rdyrs.Open "SELECT * FROM RDY WHERE ACT_IDY LIKE '" & Trim(Label7) & " '", con, adOpenKeyset, adLockPessimistic
            If rdyrs.RecordCount > 0 Then
                rdyrs.MoveFirst
                rs!SMN_IDY = rdyrs("SMN_IDY")
                rs!rou_idy = rdyrs("rou_idy")
            End If
            End If
            rs!dep_amt = Val(MS.TextMatrix(A, 15))
            rs!shl_idy = Trim(MS.TextMatrix(A, 17))
            rs!prd_idy = MS.TextMatrix(A, 20) 'Added by PNR
            rs!INV_DAT = Date
            rs!TXN_TYP = Trim(MS.TextMatrix(A, 19))
            rs!CST_IDY = Trim(Label7)
            rs!PRD_PCK = Trim(MS.TextMatrix(A, 16))
            rs!prd_upc = Val(MS.TextMatrix(A, 0))
            rs!stk_upd = "Y"
            rs.Update
    Next
'''''Else
'''''    MsgBox "No Empty Transaction Values for Save!", vbInformation, "DAS Version 4.0"
'''''    cancel
'''''End If

'Empties Sorting Interface will be opened only when the party is HCCBPL, code modified by PNR on request of PRINCE CHOPRA
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Trim(Label7) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    rs.MoveFirst
    If Trim(rs!ACT_TYP) = "PUR" Then If Combo1.ListIndex = 8 Or Combo1.ListIndex = 1 Then If check("PRT", "PRT_NME", "SIN_SHL") = True Then LOADSHELLSSORTING Trim(MS.TextMatrix(A - 1, 19)), Trim(IDY)
End If
cancel
Me.MousePointer = vbDefault
End Sub

Public Sub exitfrm()
Unload Me
End Sub

Public Sub Modify(st As String)
If CHECKTB("snv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
RS1.Open "SELECT * FROM SNV", con, adOpenKeyset, adLockPessimistic
If Trim(STR1) = "SAL" Then
    If RS1.State = 1 Then RS1.Close
    RS1.Open "SELECT R_SHL_NOS,R_PCK_NOS,BRK_SHL,BRK_PCK,D_SHL_NOS,D_PCK_NOS,DP_SHL_NOS,DP_PCK_NOS,DEP_AMT,PRD_IDY,INV_IDY,SHL_IDY FROM SNV WHERE INV_IDY ", con, adOpenKeyset, adLockPessimistic
Else
    If RS1.State = 1 Then RS1.Close
    RS1.Open "SELECT SUM(R_SHL_NOS),SUM(R_PCK_NOS),SUM(BRK_SHL),SUM(BRK_PCK),SUM(D_SHL_NOS),SUM(D_PCK_NOS),SUM(DP_SHL_NOS),SUM(DP_PCK_NOS),SUM(DEP_AMT),PRD_IDY,INV_IDY,SHL_IDY FROM SNV WHERE INV_IDY IN(SELECT INV_IDY FROM INL WHERE INl.INV_DAT >=CTOD('" & FRM_DAT & "') AND INL.INV_DAT <=CTOD('" & TO_DAT & "') AND INL.TXN_TYP LIKE '" & STR1 & "')GROUP BY INV_IDY,PRD_IDY", con, adOpenKeyset, adLockPessimistic
End If
A = 1
MS.Rows = 1
While Not rs.EOF
    MS.AddItem ""
    MS.TextMatrix(A, 0) = Trim(rs!INV_IDY)
    MS.TextMatrix(A, 1) = Trim(rs!shl_nme)
    MS.TextMatrix(A, 2) = Trim(rs!PRD_NME)
    MS.TextMatrix(A, 3) = Trim(rs!PRD_UNT)
    MS.TextMatrix(A, 4) = Trim(rs!TOT_PCS)
    MS.TextMatrix(A, 9) = 0
    MS.TextMatrix(A, 10) = 0
    MS.TextMatrix(A, 11) = 0
    MS.TextMatrix(A, 12) = 0
    MS.TextMatrix(A, 13) = 0
    MS.TextMatrix(A, 14) = 0
    MS.TextMatrix(A, 15) = "0.00"
    MS.TextMatrix(A, 16) = Trim(rs!prd_idy)
    MS.TextMatrix(A, 17) = Trim(rs!shl_idy)
    MS.TextMatrix(A, 18) = Format(rs!INV_DAT, "DD/MMM/YYYY")
    MS.TextMatrix(A, 19) = Trim(rs!TXN_TYP)
    MS.TextMatrix(A, 20) = Trim(rs!prd_idy)
    rs.MoveNext
    A = A + 1
Wend
While Not RS1.EOF
    For A = 1 To MS.Rows - 1
        If Trim(RS1!INV_IDY) = Trim(MS.TextMatrix(A, 0)) And Trim(RS1!prd_idy) = Trim(MS.TextMatrix(A, 16)) Then
            MS.TextMatrix(A, 5) = Val(RS1(0) & "")
            MS.TextMatrix(A, 6) = Val(RS1(1) & "")
            MS.TextMatrix(A, 7) = 0
            MS.TextMatrix(A, 8) = 0
            MS.TextMatrix(A, 9) = Val(RS1(2) & "")
            MS.TextMatrix(A, 10) = Val(RS1(3) & "")
            MS.TextMatrix(A, 13) = Val(RS1(6) & "")
            MS.TextMatrix(A, 14) = Val(RS1(7) & "")
            MS.TextMatrix(A, 15) = convert(Val(RS1(8) & ""))
            MS.TextMatrix(A, 11) = (Val(MS.TextMatrix(A, 3)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 9)) + Val(MS.TextMatrix(A, 13))))
            MS.TextMatrix(A, 12) = (Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 10)) + Val(MS.TextMatrix(A, 14))))
            If (Val(MS.TextMatrix(A, 3)) = (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 9)) + Val(MS.TextMatrix(A, 13)))) And (Val(MS.TextMatrix(A, 4)) = (Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 10)) + Val(MS.TextMatrix(A, 14)))) Then
                If MS.Rows = 2 Then
                    MS.Rows = 1
                Else
                    MS.RemoveItem A
                    Exit For
                End If
            End If
        End If
    Next
    RS1.MoveNext
Wend
For A = 1 To MS.Rows - 1
    MS.TextMatrix(A, 11) = (Val(MS.TextMatrix(A, 3)) - (Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 9)) + Val(MS.TextMatrix(A, 13))))
    MS.TextMatrix(A, 12) = (Val(MS.TextMatrix(A, 4)) - (Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 10)) + Val(MS.TextMatrix(A, 14))))
Next
End Sub

Public Sub V_SHL_LST()
If CHECKTB("snv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
Dim Sls As ListItem
rs.Open "SELECT * FROM SNV WHERE TXN_TYP NOT IN('VAN','SAL','PUR','SRN','PRT','STI','STO') GROUP BY TXN_IDY", con, adOpenKeyset, adLockPessimistic
rsact.Open "select act_nme,act_idy,mar_nme from act LEFT JOIN MAR ON ACT.MAR_IDY =MAR.MAR_IDY ", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
list1.ColumnHeaders.CLEAR
list1.ColumnHeaders.Add , , "ID"
list1.ColumnHeaders.Add , , "Date"
list1.ColumnHeaders.Add , , "Type"
list1.ColumnHeaders.Add , , "Party"
list1.ColumnHeaders.Add , , "Van Challan"
list1.ColumnHeaders.Add , , "Market"
While Not rs.EOF
    Set ls = list1.ListItems.Add(, , Trim(rs("TXN_IDY")))
    ls.ListSubItems.Add , , Format((rs("TXN_DAT") & ""), "DD/MMM/YYYY")
    ls.ListSubItems.Add , , Trim(rs("TXN_TYP") & "")
    If rsact.RecordCount > 0 Then
        rsact.MoveFirst
        If Val(Trim(rs!CST_IDY)) > 0 Then
            rsact.Find "ACT_IDY LIKE '" & Trim(rs!CST_IDY) & "'"
            If rsact.EOF = False Then
                ls.ListSubItems.Add , , Trim(rsact("ACT_NME") & "")
                ls.ListSubItems.Add , , Trim(rs("INV_IDY") & "")
                ls.ListSubItems.Add , , Trim(rsact("MAR_NME") & "")
            End If
        End If
    End If
    rs.MoveNext
Wend
If list1.ListItems.count > 0 Then
    Label20 = "Select Empty Transaction"
    SSTab1.Tab = 2
    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    MsgBox "No Existing Shell Transactions for View!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    shl_idy.Enabled = True
    shl_idy.SetFocus
End If
End Sub
Public Sub DISPDETAILS()
If CHECKTB("snv.dbf,prd.dbf,shl.dbf,act.dbf,inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
txntyp = ""
Dim rs As New ADODB.Recordset
txntyp = GetValue("TXN_TYP", "SNV WHERE TXN_IDY LIKE '" & Trim(shl_idy) & "'")
'If TXNTYP = "AIN" Or TXNTYP = "AOU" Or TXNTYP = "OPS" Or TXNTYP = "RCT" Or TXNTYP = "PMT" Then
'    rs.Open "select * from snv where snv.txn_idy like '" & Trim(shl_idy) & "'", CON, adOpenKeyset, adlockpessimistic
    'MS.FixedCols = 1
'Else
    'RS.Open "select * from snv,prd,shl where snv.prd_idy like prd.prd_idy and shl.PRD_idy like snv.PRD_idy and snv.txn_idy like '" & Trim(shl_idy) & "' AND SHL_TYP= 'Shell' ORDER BY INV_IDY", CON, adOpenKeyset, adlockpessimistic
    rs.Open "select * from snv where snv.txn_idy like '" & Trim(shl_idy) & "' ORDER BY INV_IDY", con, adOpenKeyset, adLockPessimistic
'End If
MS.Rows = 1
If rs.RecordCount > 0 Then
    DTPicker1 = rs!txn_dat
    If Trim(rs!TXN_TYP) = "OPS" Then
        Combo1.ListIndex = 3
    ElseIf Trim(rs!TXN_TYP) = "AIN" Then
        Combo1.ListIndex = 2
    ElseIf Trim(rs!TXN_TYP) = "AOU" Then
        Combo1.ListIndex = 1
    ElseIf Trim(rs!TXN_TYP) = "RCT" Then
        Combo1.ListIndex = 4
    ElseIf Trim(rs!TXN_TYP) = "PMT" Then
        Combo1.ListIndex = 5
    ElseIf Trim(rs!TXN_TYP) = "DPR" Then
        Combo1.ListIndex = 6
    ElseIf Trim(rs!TXN_TYP) = "DPC" Then
        Combo1.ListIndex = 7
    ElseIf Trim(rs!TXN_TYP) = "DPI" Then
        Combo1.ListIndex = 8
    ElseIf Trim(rs!TXN_TYP) = "DPO" Then
        Combo1.ListIndex = 9
    End If
    If Combo1.ListIndex <= 3 And Combo1.ListIndex > 0 Then
        MS.ColWidth(3) = 1300
        MS.ColWidth(4) = 1300
        MS.ColWidth(9) = 1300
        MS.ColWidth(10) = 1300
        MS.ColWidth(13) = 0
        MS.ColWidth(14) = 0
        MS.ColWidth(15) = 0
    ElseIf Combo1.ListIndex = 4 Or Combo1.ListIndex = 5 Then
        MS.ColWidth(3) = 1300
        MS.ColWidth(4) = 1300
        MS.ColWidth(9) = 0
        MS.ColWidth(10) = 0
        MS.ColWidth(13) = 0
        MS.ColWidth(14) = 0
        MS.ColWidth(15) = 0
    Else
        MS.ColWidth(3) = 0
        MS.ColWidth(4) = 0
        MS.ColWidth(9) = 0
        MS.ColWidth(10) = 0
        MS.ColWidth(13) = 1300
        MS.ColWidth(14) = 1300
        MS.ColWidth(15) = 1300
    End If
    If Combo1.ListIndex <> 3 Then
        Text6.Enabled = True
        Label7 = Trim(rs!CST_IDY)
        Label12.Caption = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(rs!CST_IDY) & "'")
        Text6 = Trim(rs!CST_IDY)
    End If
    Combo1.Enabled = False
End If
For A = 1 To rs.RecordCount
    MS.AddItem ""
    MS.TextMatrix(A, 0) = Val(rs!prd_upc)
    MS.TextMatrix(A, 1) = GetValue("SHL_NME", "SHL WHERE SHL_IDY LIKE '" & Trim(rs!shl_idy) & "'")
    MS.TextMatrix(A, 2) = Trim(rs!PRD_PCK) & " - " & Val(rs!prd_upc)
    If Combo1.ListIndex = 4 Then
        'MS.TextMatrix(A, 3) = Trim(RS!i_shl_nos)
        'MS.TextMatrix(A, 4) = Trim(RS!i_pck_nos)
        'MS.TextMatrix(A, 4) = QTY(STM(Val(Trim(RS!i_pck_nos)), RS!PRD_UPC))
        MS.TextMatrix(A, 3) = Trim(rs!d_shl_nos & "") 'Added by PNR
        MS.TextMatrix(A, 4) = QTY(STM(Val(Trim(rs!d_pck_nos & "")), rs!prd_upc)) 'Added by PNR
    Else
        MS.TextMatrix(A, 3) = Trim(rs!d_shl_nos & "")
        MS.TextMatrix(A, 4) = QTY(STM(Val(Trim(rs!d_pck_nos & "")), rs!prd_upc))
    End If
    MS.TextMatrix(A, 9) = Val(rs!BRK_SHL)
    MS.TextMatrix(A, 10) = QTY(STM(Val(rs!BRK_PCK), Val(MS.TextMatrix(A, 0))))
    MS.TextMatrix(A, 13) = Trim(rs!dp_shl_nos & "")
    MS.TextMatrix(A, 14) = QTY(STM(Val(Trim(rs!dp_pck_nos)), rs!prd_upc))
    MS.TextMatrix(A, 15) = convert(Val(rs!dep_amt) & "")
    MS.TextMatrix(A, 16) = Trim(rs!PRD_PCK)
    MS.TextMatrix(A, 17) = Trim(rs!shl_idy)
    MS.TextMatrix(A, 18) = Format(rs!INV_DAT, "DD/MMM/YYYY")
    MS.TextMatrix(A, 19) = Trim(rs!TXN_TYP)
    MS.TextMatrix(A, 20) = Trim(rs!prd_idy) 'Added by PNR - field entry required
    CST_IDY = Trim(rs!CST_IDY)
    Text1 = Trim(rs!INV_IDY)
    If Len(Trim(Text1)) <> 0 Then vanchk.Value = 1
    rs.MoveNext
Next
CST_NME = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(CST_IDY) & "'")

If rs.State = 1 Then rs.Close
rs.Open "select * from inl,act where inl.cst_idy like act.act_idy and inv_idy in(select inv_idy from snv where txn_idy like '" & Trim(shl_idy) & "' group by inv_idy)", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If Trim(rs!INV_IDY) = Trim(MS.TextMatrix(b, 0)) Then
            MS.TextMatrix(b, 20) = Trim(rs!ACT_NME)
            MS.TextMatrix(b, 21) = Trim(rs!CST_IDY)
        End If
    Next
    rs.MoveNext
Next
er:
shl_idy.Enabled = False
SSTab1.Tab = 0
DTPicker1.SetFocus
Command3.Enabled = False
End Sub

Public Sub newrec()
cancel
shl_idy.Enabled = False
DTPicker1.SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.KEY = "a" Then
   OpenReport CURDIR & "\SNN.RP1", "Shells SOH", 2, "", 6
ElseIf ButtonMenu.KEY = "b" Then
    OpenReport CURDIR & "\SNN.RP1", "Shell ROL COL", 2, "(ACT->CST_TYP <> 'Van')", 6
ElseIf ButtonMenu.KEY = "h" Then
    OpenReport CURDIR & "\SNN.RP1", "Party ROL COL", 2, "(ACT->CST_TYP <> 'Van')", 6
ElseIf ButtonMenu.KEY = "c" Then
    OpenReport CURDIR & "\SNN.RP1", "Party ROL COL Account", 2, "(ACT->CST_TYP <> 'Van')", 6
ElseIf ButtonMenu.KEY = "d" Then
ElseIf ButtonMenu.KEY = "e" Then
    OpenReport CURDIR & "\SNN.RP1", "Deposits Account", 2, "(ACT->CST_TYP <> 'Van')", 6
ElseIf ButtonMenu.KEY = "f" Then
    M_REP_SHL.Tag = "f"
    M_REP_SHL.Show 1
ElseIf ButtonMenu.KEY = "g" Then
    M_REP_SHL.Tag = "g"
    M_REP_SHL.Show 1
ElseIf ButtonMenu.KEY = "BRK_REP" Then
    OpenReport CURDIR & "\SNN.RP1", "Breakage Summary", 2, "", 6
ElseIf ButtonMenu.KEY = "EMP_REG" Then
    M_FRM_DSN.Tag = "EMP_REG"
    M_FRM_DSN.Show 1
End If
End Sub

Public Sub stkinhand()
If CHECKTB("snv.dbf,cld.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
rs.Open "select sum(R_SHL_NOS),sum(R_PCK_NOS),shl_idy,txn_typ from SNV WHERE STK_UPD='Y' group by txn_typ,shl_idy", con, adOpenKeyset, adLockPessimistic
con.Execute "delete from cld"
RS1.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    If rs!TXN_TYP = "PUR" Or rs!TXN_TYP = "STI" Or rs!TXN_TYP = "SRN" Or rs!TXN_TYP = "AOU" Or rs!TXN_TYP = "DPO" Then
        con.Execute "UPDATE CLD SET OPN_STK= OPN_STK + " & Val(rs(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!shl_idy) & "'", d
        con.Execute "UPDATE CLD SET PUR_STK= PUR_STK + " & Val(rs(1) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!shl_idy) & "'", d
        If d = 0 Then
            RS1.AddNew
            RS1(0) = Trim(rs!shl_idy)
            RS1(1) = Val(rs(0) & "")
            RS1(2) = Val(rs(1) & "")
            RS1(3) = 0
            RS1(4) = 0
            RS1.Update
        End If
    ElseIf rs!TXN_TYP = "RCT" Or rs!TXN_TYP = "PMT" Or rs!TXN_TYP = "DPR" Or rs!TXN_TYP = "DPC" Then
    
    Else
        con.Execute "UPDATE CLD SET PRT_STK= PRT_STK + " & Val(rs(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!shl_idy) & "'", d
        con.Execute "UPDATE CLD SET SAL_STK= SAL_STK + " & Val(rs(1) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs!shl_idy) & "'", d
        If d = 0 Then
            RS1.AddNew
            RS1(0) = Trim(rs!shl_idy)
            RS1(1) = 0
            RS1(2) = 0
            RS1(3) = Val(rs(0) & "")
            RS1(4) = Val(rs(1) & "")
            RS1.Update
        End If
    End If
    rs.MoveNext
Wend
con.Execute "UPDATE CLD SET SRN_STK=PRT_STK-OPN_STK"
con.Execute "UPDATE CLD SET STI_STK=SAL_STK-PUR_STK"
End Sub

Public Sub dispshl(st As String)
If CHECKTB("shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
NME = GetValue("shl_nme", "shl where shl_idy like '" & Trim(st) & "'")
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 17)) = Trim(st) Then
        SSTab1.Tab = 1
        list1.Visible = False
        MsgBox "Shell Already Exists !", vbExclamation, "DAS Version 4.0"
        Text3.Visible = False
        MS.SetFocus
        MS.Col = 3
        MS.Row = A
        Exit Sub
    End If
Next
MS.TextMatrix(MS.Row, 17) = st
MS.TextMatrix(MS.Row, 18) = Date
If Combo1.ListIndex = 4 Then
    MS.TextMatrix(MS.Row, 19) = "AIN"
ElseIf Combo1.ListIndex = 3 Then
    MS.TextMatrix(MS.Row, 19) = "AOU"
ElseIf Combo1.ListIndex = 9 Then
    MS.TextMatrix(MS.Row, 19) = "OPS"
ElseIf Combo1.ListIndex = 10 Then
    MS.TextMatrix(MS.Row, 19) = "RCT"
ElseIf Combo1.ListIndex = 11 Then
    MS.TextMatrix(MS.Row, 19) = "PMT"
End If
MS.TextMatrix(MS.Row, 1) = Trim(NME)
MS.TextMatrix(MS.Row, 3) = 0
MS.TextMatrix(MS.Row, 4) = 0
MS.TextMatrix(MS.Row, 5) = 0
MS.TextMatrix(MS.Row, 6) = 0
MS.TextMatrix(MS.Row, 7) = 0
MS.TextMatrix(MS.Row, 8) = 0
MS.TextMatrix(MS.Row, 9) = 0
MS.TextMatrix(MS.Row, 10) = 0
MS.TextMatrix(MS.Row, 11) = 0
MS.TextMatrix(MS.Row, 12) = 0
MS.TextMatrix(MS.Row, 13) = 0
MS.TextMatrix(MS.Row, 14) = 0
MS.TextMatrix(MS.Row, 15) = "0.00"
MS.TextMatrix(MS.Row, 21) = Trim(Label7)
If Combo1.ListIndex > 8 Then
    MS.Col = 3
Else
    MS.Col = 7
End If
MS.SetFocus
Label20.Visible = False
list1.Visible = False
Text3.Visible = False
SSTab1.Tab = 1
End Sub

Public Sub LoadProducts()
If CHECKTB("prd.dbf,shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Me.MousePointer = vbHourglass
Dim rs As New ADODB.Recordset
rs.Open "select prd.prd_pck,shl.shl_nme,prd.sub_unt,shl.SHL_IDY,shl.prd_idy from prd,shl where SHL.SHL_TYP ='Shell' and shl.prd_idy like prd.prd_idy group by shl.shl_idy,prd.prd_pck,prd.sub_unt", con, adOpenKeyset, adLockPessimistic
MS.Rows = 1
For A = 1 To rs.RecordCount
    MS.AddItem ""
    MS.TextMatrix(A, 0) = Val(rs!SUB_UNT)
    MS.TextMatrix(A, 1) = Trim(rs!shl_nme & "")
    MS.TextMatrix(A, 2) = Trim(rs!PRD_PCK & "") & " - " & Val(rs!SUB_UNT)
    MS.TextMatrix(A, 16) = Trim(rs!PRD_PCK & "")
    MS.TextMatrix(A, 17) = Trim(rs!shl_idy & "")
    MS.TextMatrix(A, 20) = Trim(rs!prd_idy & "") 'Added by PNR
    If Combo1.ListIndex = 1 Then
        MS.TextMatrix(A, 19) = "AOU"
    ElseIf Combo1.ListIndex = 2 Then
        MS.TextMatrix(A, 19) = "AIN"
    ElseIf Combo1.ListIndex = 3 Then
        MS.TextMatrix(A, 19) = "OPS"
    ElseIf Combo1.ListIndex = 4 Then
        MS.TextMatrix(A, 19) = "RCT"
    ElseIf Combo1.ListIndex = 5 Then
        MS.TextMatrix(A, 19) = "PMT"
    ElseIf Combo1.ListIndex = 6 Then
        MS.TextMatrix(A, 19) = "DPR"
    ElseIf Combo1.ListIndex = 7 Then
        MS.TextMatrix(A, 19) = "DPC"
    ElseIf Combo1.ListIndex = 8 Then
        MS.TextMatrix(A, 19) = "DPI"
    ElseIf Combo1.ListIndex = 9 Then
        MS.TextMatrix(A, 19) = "DPO"
    End If
    rs.MoveNext
Next
If Combo1.ListIndex > 5 Then
    MS.ColWidth(3) = 0
    MS.ColWidth(4) = 0
    MS.ColWidth(9) = 0
    MS.ColWidth(10) = 0
    MS.ColWidth(13) = 1300
    MS.ColWidth(14) = 1300
    MS.ColWidth(15) = 1300
ElseIf Combo1.ListIndex <= 3 Then
    MS.ColWidth(3) = 1300
    MS.ColWidth(4) = 1300
    MS.ColWidth(9) = 1300
    MS.ColWidth(10) = 1300
    MS.ColWidth(13) = 0
    MS.ColWidth(14) = 0
    MS.ColWidth(15) = 0
Else
    MS.ColWidth(3) = 1300
    MS.ColWidth(4) = 1300
    MS.ColWidth(9) = 0
    MS.ColWidth(10) = 0
    MS.ColWidth(13) = 0
    MS.ColWidth(14) = 0
    MS.ColWidth(15) = 0
End If
Me.MousePointer = vbDefault
End Sub

Public Function CSUBUNT(STK As Double, SUBUNT As Integer) As Boolean
S_UNT = Round((Val(STK) - Int(STK)) * 1000)
If S_UNT < SUBUNT Then
    CSUBUNT = True
Else
    MsgBox "Product Subunits are " & SUBUNT, vbExclamation
    CSUBUNT = False
    M_FLX_AMT.SetFocus
    SendKeys "{home}+{end}"
End If
End Function

Private Sub vanchk_keypress(K As Integer)
If K = 13 Then
    If vanchk.Value = 1 Then
        Text1.Enabled = True
        Text1.SetFocus
    Else
        Text1.Enabled = False
        If Trim(shl_idy) = "" Then
        LoadProducts
        End If
        MS.SetFocus
    End If
End If
End Sub
