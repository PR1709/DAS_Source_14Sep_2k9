VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_COM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commissions - Salesmen"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_COM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
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
      TabPicture(0)   =   "M_FRM_COM.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "msr"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MS"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtvar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "comtyp"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "smnid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "List1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "All"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "range"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_COM.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lv"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox range 
         Caption         =   "&Range"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox All 
         Caption         =   "&All Products"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   810
         Left            =   3480
         TabIndex        =   20
         Top             =   840
         Width           =   5175
         Begin MSComCtl2.DTPicker fdate 
            Height          =   330
            Left            =   1020
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   285
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24510467
            CurrentDate     =   37869
         End
         Begin MSComCtl2.DTPicker tdate 
            Height          =   330
            Left            =   3300
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   285
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24510467
            CurrentDate     =   37869
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "From"
            Height          =   195
            Left            =   345
            TabIndex        =   24
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "To"
            Height          =   195
            Left            =   2865
            TabIndex        =   23
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   1920
         TabIndex        =   13
         Top             =   1680
         Width           =   7215
         Begin VB.TextBox txtIncAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            MaxLength       =   11
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtQuan 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2760
            TabIndex        =   15
            Text            =   "PER CASE"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox comper 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            MaxLength       =   5
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Commission Amount"
            Height          =   195
            Left            =   4080
            TabIndex        =   19
            Top             =   285
            Width           =   1410
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Quantity"
            Height          =   195
            Left            =   1800
            TabIndex        =   18
            Top             =   285
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Commission Amount"
            Height          =   195
            Left            =   2040
            TabIndex        =   17
            Top             =   285
            Visible         =   0   'False
            Width           =   1410
         End
      End
      Begin VB.ListBox List1 
         Height          =   255
         ItemData        =   "M_FRM_COM.frx":047A
         Left            =   8000
         List            =   "M_FRM_COM.frx":0484
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1935
         Width           =   735
      End
      Begin VB.TextBox smnid 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1635
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   495
         Width           =   1335
      End
      Begin VB.ListBox comtyp 
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
         ItemData        =   "M_FRM_COM.frx":048E
         Left            =   4875
         List            =   "M_FRM_COM.frx":0498
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   495
         Width           =   1335
      End
      Begin VB.TextBox txtvar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lv 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   4
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
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2835
         Left            =   75
         TabIndex        =   5
         Top             =   2265
         Width           =   8750
         _ExtentX        =   15425
         _ExtentY        =   5001
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
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
      Begin MSFlexGridLib.MSFlexGrid msr 
         Height          =   2835
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Active"
         Height          =   195
         Left            =   7440
         TabIndex        =   12
         Top             =   1965
         Width           =   450
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
         TabIndex        =   8
         Top             =   720
         Width           =   7900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Salesman"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commision Type"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3390
         TabIndex        =   6
         Top             =   555
         Width           =   1155
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
            Picture         =   "M_FRM_COM.frx":04AD
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":09F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":0DB9
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":110D
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":1745
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":1A99
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":1EF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":21ED
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":25F9
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":2705
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":2A59
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":2E6D
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACT_LST"
                  Text            =   "Commision Value Report"
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
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CVR"
                  Text            =   "Commision Value Report"
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
            Picture         =   "M_FRM_COM.frx":33B9
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":38FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":3CC5
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":4019
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":4651
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":49A5
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":4E05
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":50F9
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":5505
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":5611
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":5965
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COM.frx":5D79
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "m_frm_com"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSCOM As New ADODB.Recordset
Dim RSsmn As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim R As Integer, ro As Integer, MD As Boolean
Private Sub All_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If All.Value = 1 Then
        txtIncAmt.Enabled = True
        txtIncAmt.SetFocus
        SendKeys "{HOME}+{END}"
        MS.Enabled = False
    Else
        dispprd
        MS.Visible = True
        MS.Enabled = True
        MS.SetFocus
        MS.Col = 2
    End If
End If
End Sub

Private Sub comper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Val(comper) <= 99.99 Then
        comper = convert(comper)
        saverecord
    Else
        MsgBox "Invalid Percentage", vbExclamation
        comper.SetFocus
        SendKeys "{home}+{end}"
    End If
End If
End Sub

Private Sub comper_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Or KeyAscii = 46 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub comtyp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(comtyp.Text) = "Quantity" Then
        SMNID.SetFocus
        All.Visible = True
        range.Visible = False
        Label8.Visible = False
        comper.Visible = False
        chk (True)
        Label8.Caption = "Commission Amount"
        MS.TextMatrix(0, 0) = "Product Name"
        MS.TextMatrix(0, 1) = "Quantity"
        MS.TextMatrix(0, 2) = "Commission Amt"
        SCl MS, 2
    Else
        MS.Visible = False
        All.Visible = False
        msr.Visible = True
        SMNID.SetFocus
        range.Visible = True
        Label8.Visible = True
        comper.Visible = True
        chk (False)
        Label8 = "Commission %"
        DISPMSR
        msr.Rows = 1
        SMNID.SetFocus
End If
End If
End Sub

'Private Sub comtyp_KeyPress(KeyAscii As Integer)
'MS.TextMatrix(0, 0) = "Product Name"
'If Trim(comtyp.Text) = "Quantity" Then
'    MS.TextMatrix(0, 1) = "Quantity"
'    dispprd
'Else
'    MS.TextMatrix(0, 1) = "Value"
'End If
'End Sub

Private Sub fdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    TDATE.SetFocus
End If
End Sub

Private Sub Form_Activate()
SMNID.SetFocus
comtyp.ListIndex = 0
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
list1.Selected(0) = True
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
    'If Me.ActiveControl.name = "txnid" Then
    '    comlist
    'Else
    If Me.ActiveControl.name = "smnid" Then
        smnlist
    End If
ElseIf KeyCode = 88 And Shift = 2 Then
    Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "smnid" And Len(Trim(SMNID)) = 0 Then
       Unload Me
   
    ElseIf SSTab1.Tab = 1 Then
        If LV.Visible = True And UCase(Trim(Label20)) = UCase("Select Transaction") Then
            LV.Visible = False
            SSTab1.Tab = 0
            SMNID.SetFocus
        ElseIf LV.Visible = True And UCase(Trim(Label20)) = UCase("Select Salesman") Then
            LV.Visible = False
            SSTab1.Tab = 0
            SMNID.SetFocus
        End If
    Else
        saverecord
    End If
End If
End Sub

Private Sub Form_Load()
MS.Rows = 1
SETYEAR Me
MS.Rows = 2
MS.Cols = 5
R = 1
MS.ColWidth(3) = 0
MS.ColWidth(4) = 0
MS.ColWidth(0) = 2500
MS.ColWidth(1) = 1100
MS.ColWidth(2) = 1500
MS.TextMatrix(0, 0) = "Product Name"
MS.TextMatrix(0, 1) = "Quantity"
MS.TextMatrix(0, 2) = "Commission Amt"
SCl MS, 2
MS.ColAlignment(2) = 7
End Sub

Private Sub lv_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If LV.SortOrder = lvwAscending Then
LV.SortOrder = lvwDescending
ElseIf LV.SortOrder = lvwDescending Then
LV.SortOrder = lvwAscending
End If
LV.SortKey = CH.Index - 1
End Sub

Private Sub MS_KeyPress(KeyAscii As Integer)
ro = MS.Row
If keyasii = 13 Or KeyAscii = 27 Then

ElseIf KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Then
     If MS.Col = 2 Then
        txtvar.Visible = True
        txtvar.Left = MS.Left + MS.CellLeft
        txtvar.Top = MS.Top + MS.CellTop
        txtvar.Height = MS.CellHeight
        txtvar.Width = MS.CellWidth
        txtvar.Visible = True
        txtvar.SetFocus
        txtvar = Chr(KeyAscii)
        txtvar.SelStart = Len(Trim(txtvar))
    End If
End If
End Sub

Private Sub MS_LeaveCell()
txtvar.Visible = False
End Sub

Private Sub MS_Scroll()
If txtvar.Visible = True Then
    txtvar = ""
    txtvar.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub msr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And KeyAscii = 13 Then
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Then
    txtvar.Visible = True
    txtvar.Left = msr.Left + msr.CellLeft
    txtvar.Top = msr.Top + msr.CellTop
    txtvar.Height = msr.CellHeight
    txtvar.Width = msr.CellWidth
    txtvar.Visible = True
    txtvar.SetFocus
    txtvar = Chr(KeyAscii)
    txtvar.SelStart = Len(Trim(txtvar))
End If
End Sub

Private Sub msr_Scroll()
If txtvar.Visible = True Then
    txtvar = ""
    txtvar.Visible = False
    msr.SetFocus
End If
End Sub

Private Sub range_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If range.Value = 1 Then
        DISPMSR
        msr.Row = 1
        msr.Col = 0
    Else
        comper.Visible = True
        comper.Enabled = True
        comper = "0.00"
        comper.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End If
End Sub

Private Sub smnid_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If K = 13 Or K = 40 Then
If RS.State = 1 Then RS.Close
    RS.Open "select * from smn where smn_idy like '" & Sincrement(SMNID) & "'", CON, adOpenStatic
    If RS.RecordCount > 0 Then
        SMNID = Sincrement(SMNID)
        smnname = Trim(RS("smn_nme"))
        SMNID.Enabled = False
        If FDATE.Enabled = False Then
            If All.Visible = True And All.Value = 1 Then
                txtIncAmt.Enabled = True
                txtIncAmt.SetFocus
                SendKeys "{HOME}+{END}"
            End If
            If range.Visible = True And range.Value = 0 Then
                comper.Enabled = True
                comper.SetFocus
                SendKeys "{HOME}+{END}"
            End If
            If MS.Visible = True And MS.Rows > 1 Then
                MS.SetFocus
                MS.Col = 2
            End If
            If msr.Visible = True And msr.Rows > 1 Then
                msr.SetFocus
                msr.Col = 2
            End If
        Else
            FDATE.SetFocus
        End If
    Else
        smnlist
    End If
End If
End Sub
Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from smn where smn_nme like '" & SMNID & "%'", CON, adOpenStatic
Label20.Caption = "Select Salesman"
LV.ColumnHeaders.Clear
LV.ColumnHeaders.Add , , "Salesman Name", 2500
LV.ColumnHeaders.Add , , "ID", 1500
LV.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = LV.ListItems.Add(, , Trim(RS("smn_nme")))
        ls.ListSubItems.Add , , Trim(RS("smn_idy") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    LV.Visible = True
    LV.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        SMNID.Enabled = True
        SMNID.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End If
End Sub
Private Sub dispprd()
If rsprd.State = 1 Then rsprd.Close
rsprd.Open "select prd_idy,prd_nme,sub_unt from prd", CON, adOpenKeyset, adLockOptimistic
R = 1
MS.Rows = 2
MS.Cols = 5
R = 1
MS.ColWidth(3) = 0
MS.ColWidth(4) = 0
MS.ColWidth(0) = 2500
MS.ColWidth(1) = 1100
MS.ColWidth(2) = 1500
MS.TextMatrix(0, 0) = "Product Name"
MS.TextMatrix(0, 1) = "Quantity"
MS.TextMatrix(0, 2) = "Commission Amt"
SCl MS, 2
MS.ColAlignment(2) = 7
If Trim(comtyp.Text) = "Quantity" Then
    While Not rsprd.EOF
            MS.TextMatrix(R, 0) = rsprd(1)
            MS.TextMatrix(R, 3) = rsprd(0)
            MS.TextMatrix(R, 4) = rsprd(2)
            MS.TextMatrix(R, 1) = "PER CASE"
            R = R + 1
            MS.Rows = MS.Rows + 1
            rsprd.MoveNext
    Wend
Else
'    While Not rsprd.EOF
'        MS.TextMatrix(R, 0) = rsprd(1)
'        MS.TextMatrix(R, 3) = rsprd(0)
'        MS.TextMatrix(R, 4) = rsprd(2)
'        MS.TextMatrix(R, 1) = "0"
'        R = R + 1
'        MS.Rows = MS.Rows + 1
'        rsprd.MoveNext
'    Wend
End If
MS.Rows = MS.Rows - 1
End Sub
Private Sub NewRecord()
MD = False
cancel
comtyp.Enabled = True
comtyp.SetFocus
range.Enabled = True
txtIncAmt.Enabled = False
End Sub

Private Sub cancel()
clrctr Me
SMNID = ""
SMNID.Enabled = True
SMNID = ""
'smnname.Caption = ""
FDATE.Value = Date
TDATE.Value = Date
FDATE.Enabled = True
TDATE.Enabled = True
range.Enabled = True
All.Enabled = True
All.Value = 0
txtQuan = "PER CASE"
txtIncAmt = ""
MS.Clear
msr.Clear
txtvar = ""
txtvar.Visible = False
MS.Rows = 1
msr.Rows = 1
If comper.Visible = True Then
comper = ""
End If
comtyp.Selected(0) = True
SMNID.Enabled = True
If SSTab1.Tab = 1 Then
    LV.Visible = False
    SSTab1.Tab = 0
End If
SMNID.SetFocus
If range.Visible = True Then range.Value = 0
msr.Visible = False
MS.Visible = True
MS.TextMatrix(0, 0) = "Product Name"
MS.TextMatrix(0, 1) = "Quantity"
MS.TextMatrix(0, 2) = "Commission Amt"
SCl MS, 2
End Sub

Private Sub lv_DblClick()
lv_KeyDown 13, 0
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
If KeyCode = 13 Then
    If LV.ListItems.Count > 0 Then
        If RS.State = 1 Then RS.Close
        If Trim(Label20.Caption) = "Select Salesman" Then
            Dim lst As ListSubItem
            Set lst = LV.SelectedItem.ListSubItems.Item(1)
            RS.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", CON, adOpenStatic
            SMNID = Trim(RS("smn_idy"))
            smnname = Trim(RS("smn_nme"))
            LV.Visible = False
            SSTab1.Tab = 0
            If FDATE.Enabled = False Then
                If All.Visible = True And All.Value = 1 Then
                    txtIncAmt.Enabled = True
                    txtIncAmt.SetFocus
                    SendKeys "{HOME}+{END}"
                End If
                If range.Visible = True And range.Value = 0 Then
                    comper.Enabled = True
                    comper.SetFocus
                    SendKeys "{HOME}+{END}"
                End If
                    If MS.Visible = True And MS.Rows > 1 Then
                        MS.SetFocus
                        MS.Col = 2
                    End If
                    If msr.Visible = True And msr.Rows > 1 Then
                        msr.SetFocus
                        msr.Col = 2
                    End If
                Else
                    FDATE.SetFocus
            End If
        Else
            Dim ls As ListItem
            Set ls = LV.SelectedItem
            RS.Open "select * from cOM where txn_idy like '" & Trim(ls.Text) & "'", CON, adOpenKeyset, adLockOptimistic
            If RS.RecordCount > 0 Then
                FDATE.Enabled = False
                TDATE.Enabled = False
                range.Enabled = False
                If RS("chk_flg") = "A" Or RS("chk_flg") = "S" Then
                    MS.Visible = True
                    txtvar.Visible = False
                    chk (True)
                    All.Visible = True
                    range.Visible = False
                    Label8.Visible = False
                    comper.Visible = False
                    msr.Visible = False
                    'smnid = RS("TXN_IDY")
                    'txnid.Enabled = False
                    If RSsmn.State = 1 Then RSsmn.Close
                    RSsmn.Open "select smn_nme from smn where smN_idy ='" & Trim(RS("smn_idy")) & "'", CON, adOpenKeyset, adLockOptimistic
                    If Not RSsmn.EOF Then
                        SMNID = RS("smn_idy")
                        smnname.Caption = RSsmn(0)
                    End If
                    FDATE.Value = RS("frm_dat")
                    TDATE.Value = RS("to_dat")
                    If Trim(RS("cOM_typ")) = "Quantity" Then
                        comtyp.ListIndex = 0
                    Else
                        comtyp.ListIndex = 1
                    End If
                    comtyp.Enabled = False
                    If RS("CHK_FLG") = "A" Then
                        All.Enabled = True
                        All.Value = 1
                        txtIncAmt = convert(Trim(RS("COM_AMT")) & "")
                        comtyp.Enabled = False
                        'txnid.Enabled = False
                        All.Enabled = False
                        'CHANGED
                    Else
                        All.Enabled = False
                        R = 1
                        MS.Rows = 2
                        MS.TextMatrix(0, 0) = "Product Name"
                        MS.TextMatrix(0, 1) = "Quantity"
                        MS.TextMatrix(0, 2) = "Commission Amt"
                        SCl MS, 2
                        MS.ColAlignment(2) = 7
                        MS.ColWidth(2) = 1500
                        While Not RS.EOF
                            If rsprd.State = 1 Then rsprd.Close
                            rsprd.Open "select prd_nme,SUB_UNT from prd wher prd_idy='" & RS("prd_idy") & "'", CON, adOpenKeyset, adLockOptimistic
                            MS.TextMatrix(R, 0) = rsprd(0)
                            MS.TextMatrix(R, 4) = rsprd(1)
                            MS.TextMatrix(R, 3) = RS("prd_idy")
                            MS.TextMatrix(R, 1) = "PER CASE"
                            MS.TextMatrix(R, 2) = convert(RS("cOM_amt"))
                            R = R + 1
                            MS.Rows = MS.Rows + 1
                            RS.MoveNext
                        Wend
                        MS.Rows = MS.Rows - 1
                        comtyp.Enabled = False
                        'smnid.Enabled = False
                        MS.Visible = True
                        MS.Enabled = True
                        MS.SetFocus
                    End If
                    SMNID.SetFocus
                    SendKeys "{HOME}+{END}"
                    LV.Visible = False
                    SSTab1.Tab = 0
                    MD = True
                Else
                    If RS("CHK_FLG") = "R" Then
                        DISPMSR
                        'txnid = RS("txn_idy")
                        'txnid.Enabled = False
                        comtyp.ListIndex = 1
                        If RSsmn.State = 1 Then RSsmn.Close
                        RSsmn.Open "select SMN_NME from smn where smn_idy='" & RS("smn_idy") & "'", CON, adOpenKeyset, adLockOptimistic
                        If Not RSsmn.EOF Then
                            SMNID = RS("SMN_IDY")
                            smnname = RSsmn(0)
                        End If
                        FDATE.Value = RS("FRM_DAT")
                        TDATE.Value = RS("TO_DAT")
                        comtyp.Enabled = False
                        range.Value = 1
                        R = 1
                        msr.Rows = 1
                        msr.Cols = 3
                        While Not RS.EOF
                            msr.Visible = True
                            msr.Rows = msr.Rows + 1
                            MS.Visible = False
                            msr.TextMatrix(R, 0) = convert(Trim(RS("beg_rng")) & "")
                            msr.TextMatrix(R, 1) = convert(Trim(RS("end_rng")) & "")
                            msr.TextMatrix(R, 2) = convert(Trim(RS("com_per") & ""))
                            R = R + 1
                            RS.MoveNext
                        Wend
                        range.Visible = True
                        Label8.Visible = True
                        comper.Enabled = True
                        comper.Visible = True
                        All.Visible = False
                        chk (False)
                        'fdate.SetFocus
                        MD = True
                    Else
'                        txnid = RS("TXN_IDY")
                        comtyp.ListIndex = 1
                        comtyp.Enabled = False
                        If RSsmn.State = 1 Then RSsmn.Close
                        RSsmn.Open "SELECT SMN_NME FROM SMN WHERE SMN_IDY='" & RS("SMN_IDY") & "'", CON, adOpenKeyset, adLockOptimistic
                        If Not RSsmn.EOF Then
                            SMNID = RS("SMN_IDY")
                            smnname = RSsmn(0)
                        End If
                        FDATE.Value = RS("FRM_DAT")
                        TDATE.Value = RS("TO_DAT")
                        range.Value = 0
                        comper = convert(RS("COM_AMT"))
                        range.Visible = True
                        Label8.Visible = True
                        comper.Visible = True
                        All.Visible = False
                        chk (False)
                        SMNID.SetFocus
                        SendKeys "{HOME}+{END}"
                        'fdate.SetFocus
                        MD = True
                    End If
                    SMNID.SetFocus
                    SendKeys "{HOME}+{END}"
                    SSTab1.Tab = 0
                    LV.Visible = False
                  '  txnid.Enabled = True
                  '  txnid.SetFocus
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub tdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If FDATE.Value > TDATE.Value Then
        MsgBox "Invalid Date ", vbExclamation
        FDATE.SetFocus
        Exit Sub
    End If
    If All.Visible = True Then
        If All.Enabled = True Then
            All.SetFocus
        ElseIf All.Value = 0 Then
            MS.SetFocus
            MS.Col = 2
        Else
            txtIncAmt.Enabled = True
            txtIncAmt.SetFocus
            SendKeys "{HOME}+{END}"
        End If
    Else
        If range.Enabled = True Then
            range.SetFocus
        ElseIf range.Value = 1 Then
            msr.SetFocus
            msr.Col = 2
        Else
            comper.Enabled = True
            comper.SetFocus
            SendKeys "{HOME}+{END}"
        End If
    End If
End If
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
    'ElseIf Me.ActiveControl.name = "txnid" Then
     '   comlist
    End If
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If RS.State = 1 Then RS.Close
'dates
RS.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    RS("DIR_ALS") = Format(FDATE, "DD/MM/YYYY")
    RS("DIR_MSG") = Format(TDATE, "DD/MM/YYYY")
    RS.Update
End If
CON.Execute "delete from scl"
'If comtyp.ListIndex = 1 Then
'    If range.Value = 0 And range.Enabled = False Then
'        If RS.State = 1 Then RS.Close
'        RS.Open "SELECT SUM(NET_AMT) FROM INL WHERE SMN_IDY LIKE '" & Trim(SMNID) & "' AND INV_DAT >=CTOD('" & FDATE & "') AND INV_DAT<=CTOD('" & TDATE & "') and txn_typ like 'SAL'", CON, adOpenKeyset, adLockOptimistic
'        If RS.RecordCount > 0 Then
'            If RS1.State = 1 Then RS1.Close
'            RS1.Open "select * from scl", CON, adOpenKeyset, adLockOptimistic
'            RS1.AddNew
'            RS1(0) = Date
'            RS1(1) = SMNID
'            RS1(3) = Val(RS(0) & "")
'            RS1(4) = Val(comper)
'            If RS.State = 1 Then RS.Close
'            RS.Open "select com_amt from van where smn_idy like '" & SMNID & "' AND INV_DAT >=CTOD('" & FDATE & "') AND INV_DAT<=CTOD('" & TDATE & "') group by inv_idy", CON, adOpenKeyset, adLockOptimistic
'            For A = 1 To RS.RecordCount
'                RS1(5) = Val(RS1(5)) + Val(RS(0) & "")
'                RS.MoveNext
'            Next
'            RS1.Update
'        End If
'        Openreport CURDIR & "\rep_lib.rp1", "Commission Report 1", 0, "", 6
'    Else
'        If RS.State = 1 Then RS.Close
'        If RS.State = 1 Then RS.Close
'        RS.Open "SELECT SUM(NET_AMT) FROM INL WHERE SMN_IDY LIKE '" & Trim(SMNID) & "' AND INV_DAT >=CTOD('" & FDATE & "') AND INV_DAT<=CTOD('" & TDATE & "') and txn_typ like 'SAL'", CON, adOpenKeyset, adLockOptimistic
'        If RS.RecordCount > 0 Then
'            sa = RS(0) & ""
'            If RS1.State = 1 Then RS1.Close
'            RS1.Open "select * from scl", CON, adOpenKeyset, adLockOptimistic
'            For A = 1 To msr.Rows - 1
'                RS1.AddNew
'                RS1(0) = Date
'                RS1(1) = SMNID
'                RS1(3) = Val(msr.TextMatrix(A, 0))
'                RS1(4) = Val(msr.TextMatrix(A, 1))
'                RS1(6) = Val(msr.TextMatrix(A, 2))
'                If A = msr.Rows - 1 Then
'                    sa = Val(sa) - (Val(RS1(4)))
'                    RS1(5) = Val(sa)
'                Else
'                    sa = Val(sa) - (Val(RS1(4)) - Val(RS1(3)))
'                    If sa < 0 Then
'                        RS1(5) = 0
'                    Else
'                        RS1(5) = Val(msr.TextMatrix(A, 1)) - Val(msr.TextMatrix(A, 0))
'                    End If
'                End If
'                RS1(7) = Val(RS1(5)) * (Val(RS1(6)) / 100)
'                RS1.Update
'            Next
'                If RS.State = 1 Then RS.Close
'                sa = 0
'                RS.Open "select com_amt from van where smn_idy like '" & SMNID & "' AND INV_DAT >=CTOD('" & FDATE & "') AND INV_DAT<=CTOD('" & TDATE & "') group by inv_idy", CON, adOpenKeyset, adLockOptimistic
'                For A = 1 To RS.RecordCount
'                    sa = Val(sa) + Val(RS(0) & "")
'                    RS.MoveNext
'                Next
'                CON.Execute "update scl set sal_stk=" & Val(sa)
'        End If
'        Openreport CURDIR & "\rep_lib.rp1", "Commission Report 2", 0, "", 6
'    End If

    If RS.State = 1 Then RS.Close
    RS.Open "select sum(tot_pcs) as tot_pcs,sub_unt,prd_idy from inv where inv_idy in(select inv_idy from inl where inv_dat>=ctod('" & FDATE.Value & "') and inv_dat<=ctod('" & TDATE.Value & "') and smn_idy='" & Trim(SMNID) & "') AND stk_typ in('Salable','Free') group by prd_idy", CON, adOpenKeyset, adLockOptimistic
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from scl", CON, adOpenKeyset, adLockOptimistic
    While Not RS.EOF
        If All.Value = 1 Then
            RS1.AddNew
            RS1(0) = Date
            RS1(1) = Trim(SMNID)
            RS1(2) = Trim(RS(2))
            RS1(3) = STM(Val(RS(0) & ""), Val(RS(1) & ""))
            RS1(4) = Val(txtIncAmt)
            RS1.Update
        Else
            RS1.AddNew
            RS1(0) = Date
            RS1(1) = Trim(SMNID)
            RS1(2) = Trim(RS(2))
            RS1(3) = STM(Val(RS(0) & ""), Val(RS(1) & ""))
            For A = 1 To MS.Rows - 1
                If Trim(RS(2)) = Trim(MS.TextMatrix(A, 3)) Then
                    RS1(4) = Val(Trim(MS.TextMatrix(A, 2)))
                End If
            Next
            RS1.Update
        End If
        RS.MoveNext
    Wend
    If RS.State = 1 Then RS.Close
    sa = 0
    RS.Open "select com_amt from van where smn_idy like '" & SMNID & "' AND INV_DAT >=CTOD('" & FDATE & "') AND INV_DAT<=CTOD('" & TDATE & "') group by inv_idy", CON, adOpenKeyset, adLockOptimistic
    For A = 1 To RS.RecordCount
        sa = Val(sa) + Val(RS(0) & "")
        RS.MoveNext
    Next
    CON.Execute "update scl set iss_stk=" & Val(sa)
    Openreport CURDIR & "\rep_lib.rp1", "Commission Report 3", 0, "", 6
'End If
End Sub

''Private Sub txnid_KeyDown(K As Integer, Shift As Integer)
'If CHECKTB("com.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'    Exit Sub
'End If
'Dim RS As New ADODB.Recordset
'If K = 13 Or K = 40 Then
'    If RS.State = 1 Then RS.Close
'    RS.Open "select * from com where txn_idy like '" & Sincrement(txnid) & "'", CON, adOpenStatic
'    If RS.RecordCount > 0 Then
'        fdate.Enabled = False
'        tdate.Enabled = False
'        If RS("CHK_FLG") = "A" Or RS("CHK_FLG") = "S" Then
'            txnid.Enabled = False
'            txnid = Sincrement(txnid)
'            range.Enabled = False
'            'rs.Open "select * from inc where txn_idy like '" & Trim(txnid) & "'", CON, adOpenStatic
'                If RSsmn.State = 1 Then RSsmn.Close
'                RSsmn.Open "select smn_nme from smn where smN_idy ='" & Trim(RS("smn_idy")) & "'", CON, adOpenKeyset, adLockOptimistic
'                If Not RSsmn.EOF Then
'                    smnid = RS("smn_idy")
'                    smnname.Caption = RSsmn(0)
'                End If
'                fdate.Value = RS("frm_dat")
'                tdate.Value = RS("to_dat")
'                If Trim(RS("cOM_typ")) = "Quantity" Then
'                    comtyp.ListIndex = 0
'                Else
'                    comtyp.ListIndex = 1
'                End If
'                comtyp.Enabled = False
'                If RS("CHK_FLG") = "A" Then
'                    All.Enabled = True
'                    All.Value = 1
'                    txtIncAmt.Visible = True
'                    txtIncAmt = convert(RS("Com_AMT"))
'                    comtyp.Enabled = False
'                    txnid.Enabled = False
'                    All.Enabled = False
'                    smnid.SetFocus
'                    SendKeys "{HOME}+{END}"
'                Else
'                    All.Enabled = False
'                    All.Value = 0
'                    R = 1
'                    MS.Rows = 1
'                    MS.Cols = 5
'                    R = 1
'                    MS.ColWidth(3) = 0
'                    MS.ColWidth(4) = 0
'                    MS.ColWidth(0) = 2500
'                    MS.ColWidth(1) = 1100
'                    MS.ColWidth(2) = 1500
'                    MS.TextMatrix(0, 0) = "Product Name"
'                    MS.TextMatrix(0, 1) = "Quantity"
'                    MS.TextMatrix(0, 2) = "Commission Amt"
'                    SCl MS, 2
'                    MS.ColAlignment(2) = 7
'                    While Not RS.EOF
'                        MS.Rows = MS.Rows + 1
'                        If RSPRD.State = 1 Then RSPRD.Close
'                        RSPRD.Open "select prd_nme,SUB_UNT from prd wher prd_idy='" & RS("prd_idy") & "'", CON, adOpenKeyset, adLockOptimistic
'                        MS.TextMatrix(R, 0) = RSPRD(0)
'                        MS.TextMatrix(R, 4) = RSPRD(1)
'                        MS.TextMatrix(R, 3) = RS("prd_idy")
'                        MS.TextMatrix(R, 1) = "PER CASE"
'                        MS.TextMatrix(R, 2) = convert(RS("cOM_amt"))
'                        R = R + 1
'                        RS.MoveNext
'                    Wend
'                    'smnid.Enabled = False
'                    MS.SetFocus
'                    MS.Col = 2
'                End If
'                If MS.Rows > 1 Then MS.SetFocus
'                MD = True
'        Else
'            ' VALUE
'            Label8.Caption = "Commission %"
'            If RS("CHK_FLG") = "R" Then
'                txnid = RS("txn_idy")
'                txnid.Enabled = False
'                comtyp.ListIndex = 1
'                If RSsmn.State = 1 Then RSsmn.Close
'                RSsmn.Open "select * from smn where smn_idy='" & RS("smn_idy") & "'", CON, adOpenKeyset, adLockOptimistic
'                If Not RSsmn.EOF Then
'                    smnid = RS("SMN_IDY")
'                    smnname = RSsmn(1)
'                End If
'                fdate.Value = RS("FRM_DAT")
'                tdate.Value = RS("TO_DAT")
'                range.Value = 1
'                R = 1
'                DISPMSR
'                msr.Rows = 1
'                msr.Cols = 3
'                While Not RS.EOF
'                    msr.Visible = True
'                    msr.Rows = msr.Rows + 1
'                    MS.Visible = False
'                    msr.TextMatrix(R, 0) = convert(Trim(RS("beg_rng")) & "")
'                    msr.TextMatrix(R, 1) = convert(Trim(RS("end_rng")) & "")
'                    msr.TextMatrix(R, 2) = convert(Trim(RS("com_per") & ""))
'                    R = R + 1
'                    RS.MoveNext
'                Wend
'            Else
'                txnid = RS("TXN_IDY")
'                comtyp.ListIndex = 1
'                comtyp.Enabled = False
'                range.Enabled = False
'                If RSsmn.State = 1 Then RSsmn.Close
'                RSsmn.Open "SELECT SMN_NME FROM SMN WHERE SMN_IDY='" & RS("SMN_IDY") & "'", CON, adOpenKeyset, adLockOptimistic
'                If Not RSsmn.EOF Then
'                    smnid = RS("SMN_IDY")
'                    smnname = RSsmn(0)
'                End If
'                fdate.Value = RS("FRM_DAT")
'                tdate.Value = RS("TO_DAT")
'                range.Value = 0
'                comper = convert(RS("COM_AMT"))
'            End If
'            range.Enabled = False
'            comtyp.Enabled = False
'            range.Visible = True
'            Label8.Visible = True
'            comper.Enabled = True
'            comper.Visible = True
'            All.Visible = False
'            chk (False)
'            smnid.SetFocus
'            SendKeys "{HOME}+{END}"
'            'fdate.SetFocus
'            MD = True
'        End If
'    Else
'        comlist
'    End If
'    End If
'End Sub
'
'Private Sub txnid_KeyPress(K As Integer)
'If K = 39 Or K = 41 Or K = 37 Then K = 0
'If K > 47 And K < 58 Or K = 8 Or K = 13 Then
'ElseIf K > 96 And K < 123 Then
'K = K - 32
'ElseIf K > 64 And K < 91 Then
'Else
'K = 0
'End If
'End Sub

Private Sub txtIncAmt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtIncAmt = convert(txtIncAmt.Text)
    saverecord
End If
End Sub

Private Sub txtIncAmt_KeyPress(KeyAscii As Integer)
'If Len(Trim(txtIncAmt)) = 8 Then txtIncAmt.Locked = True
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 8 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtvar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If msr.Visible = True Then
    If msr.Row > 1 And msr.Col = 0 Then
        If Val(msr.TextMatrix(msr.Row - 1, 1)) >= Val(txtvar) Then
            MsgBox "Invalid Range", vbExclamation
            txtvar.Visible = True
            txtvar.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    End If
    If msr.Col = 0 Then
        If Val(msr.TextMatrix(msr.Row, 1)) > 0 Then
            If Val(txtvar) >= Val(msr.TextMatrix(msr.Row, 1)) Then
                MsgBox "Invalid Value", vbExclamation
                txtvar.Visible = True
                txtvar.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            Else 'If Val(txtvar) = 0 Then
                msr.Text = convert(txtvar)
            End If
        Else
            msr.Text = convert(txtvar)
        End If
    End If
    If msr.Col = 2 Then
        If Val(txtvar) >= 0 And Val(txtvar) <= 99.99 Then
            msr.Text = convert(Trim(txtvar))
        Else
            MsgBox "Invalid Percentage", vbExclamation
            txtvar.Visible = True
            txtvar.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    End If
    If msr.Col = 1 Then
        If Val(msr.TextMatrix(msr.Row, 0)) >= Val(txtvar) Then
            MsgBox "Invalid Range", vbExclamation
            txtvar.Visible = True
            txtvar.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        Else
            msr.Text = convert(Trim(txtvar))
            If msr.Row <> msr.Rows - 1 Then
                msr.TextMatrix(msr.Row + 1, 0) = convert(Val(txtvar) + 1)
            End If
        End If
    End If
    If msr.Col = 2 Then
        If msr.Row <> msr.Rows - 1 Then
            msr.Row = msr.Row + 1
            msr.Col = 1
        End If
    Else
        msr.Col = msr.Col + 1
    End If
    txtvar = ""
    txtvar.Visible = False
    msr.SetFocus
Else
    'MS.TextMatrix(ro, 2) = Val(txtvar)
    MS.Text = convert(Val(txtvar))
    txtvar = ""
    txtvar.Visible = False
    MS.SetFocus
    If MS.Row <> MS.Rows - 1 Then
        MS.Row = MS.Row + 1
    End If
End If
End If
End Sub

Private Sub txtVar_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
Else
KeyAscii = 0
End If
End Sub
Private Sub saverecord()
If MsgBox("Save Record ?", vbYesNo + vbExclamation) = vbNo Then
    cancel
    Exit Sub
End If
If Trim(SMNID) = "" Then
    MsgBox "Please select a Salesman ID", vbExclamation
    SMNID.SetFocus
    Exit Sub
End If
    
If FDATE.Value > TDATE.Value Then
    MsgBox "Invalid Date ", vbExclamation
    FDATE.SetFocus
    Exit Sub
End If

If Val(comper) > 99.99 Then
    MsgBox "Invalid Percentage", vbExclamation
    comper.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

'If msr.Visible = True Then
'    For R = 1 To msr.Rows - 1
'        If Val(msr.TextMatrix(R, 2)) > 0 Then
'            If Val(msr.TextMatrix(R, 1)) <= 0 Then
'                MsgBox "Enter End Range", vbExclamation
'                msr.SetFocus
'                msr.Row = R
'                msr.Col = 1
'                Exit Sub
'            End If
'        End If
'    Next
'End If
If Trim(comtyp.Text) = "Quantity" Then
'    If Trim(txtIncAmt.Text) = "" Then
'        MsgBox "Please enter Commission Amount", vbExclamation
'        All.Value = 1
'        txtIncAmt.Enabled = True
'        txtIncAmt.SetFocus
'        Exit Sub
'    End If
Else
    If msr.Visible = True And msr.Rows > 1 Then
    If Val(msr.TextMatrix(1, 2)) = 0 Then
    MsgBox "Enter Commission %", vbExclamation
        msr.SetFocus
        msr.Col = 2
        Exit Sub
    End If
    End If
End If

'If CHKROWS() = 1 Then Exit Sub
If MD = False Then
    If Trim(comtyp.Text) = "Quantity" Then
        If RSCOM.State = 1 Then RSCOM.Close
        RSCOM.Open "select * from com where txn_idy=''", CON, adOpenKeyset, adLockOptimistic
        R = 1
        If rschk.State = 1 Then rschk.Close
        rschk.Open "select * from com where smn_idy='" & Trim(SMNID) & "'", CON, adOpenKeyset, adLockOptimistic
        If Not rschk.EOF Then
            If FDATE.Value >= rschk("frm_dat") And FDATE.Value <= rschk("to_dat") Then
                MsgBox "Transaction Exists between " & Format(rschk("frm_dat"), "dd/mmm/yyyy") & " to " & Format(rschk("to_dat"), "dd/mmm/yyyy"), vbExclamation
                FDATE.SetFocus
                Exit Sub
            End If
        End If
'        txnid = increment("idy_tsn")
        If All.Value = 1 Then
            RSCOM.AddNew
            For I = 0 To RSCOM.Fields.Count - 1
                If RSCOM(I).Type = adVarChar Or RSCOM(I).Type = adChar Or RSCOM(I).Type = adBSTR Or RSCOM(I).Type = adLongVarChar Or RSCOM(I).Type = adLongVarWChar Then
                    RSCOM(I) = " "
                ElseIf RSCOM(I).Type = adInteger Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adSingle Or RSCOM(I).Type = adDouble Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adCurrency Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adSmallInt Or RSCOM(I).Type = adDecimal Then
                    RSCOM(I) = 0
                ElseIf RSCOM(I).Type = adDate Or RSCOM(I).Type = adDBDate Or RSCOM(I).Type = adDBTimeStamp Then
                    RSCOM(I) = Date
                End If
            Next
            'RSCOM("txn_idy") = txnid
            RSCOM("com_typ") = Trim(comtyp.Text)
            RSCOM("smn_idy") = Trim(SMNID)
            RSCOM("frm_dat") = FDATE.Value
            RSCOM("to_dat") = TDATE.Value
            If All.Value = 1 Then
                RSCOM("chk_flg") = "A"
            Else
                RSCOM("chk_flg") = "S"
            End If
            RSCOM("com_amt") = Val(txtIncAmt)
            RSCOM("act_inc") = Trim(list1.Text)
            RSCOM.Update
            'MsgBox "New Transaction ID is " & txnid, vbExclamation
            cancel
        Else
            For R = 1 To MS.Rows - 1
                If Val(MS.TextMatrix(R, 2)) > 0 Then
                    RSCOM.AddNew
                    For I = 0 To RSCOM.Fields.Count - 1
                        If RSCOM(I).Type = adVarChar Or RSCOM(I).Type = adChar Or RSCOM(I).Type = adBSTR Or RSCOM(I).Type = adLongVarChar Or RSCOM(I).Type = adLongVarWChar Then
                            RSCOM(I) = " "
                        ElseIf RSCOM(I).Type = adInteger Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adSingle Or RSCOM(I).Type = adDouble Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adCurrency Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adSmallInt Or RSCOM(I).Type = adDecimal Then
                            RSCOM(I) = 0
                        ElseIf RSCOM(I).Type = adDate Or RSCOM(I).Type = adDBDate Or RSCOM(I).Type = adDBTimeStamp Then
                            RSCOM(I) = Date
                        End If
                    Next
                    'RSCOM("txn_idy") = txnid
                    RSCOM("com_typ") = Trim(comtyp.Text)
                    RSCOM("smn_idy") = Trim(SMNID)
                    RSCOM("frm_dat") = FDATE.Value
                    RSCOM("to_dat") = TDATE.Value
                    If All.Value = 1 Then
                        RSCOM("chk_flg") = "A"
                    Else
                        RSCOM("chk_flg") = "S"
                    End If
                    RSCOM("com_amt") = Val(MS.TextMatrix(R, 2))
                    RSCOM("prd_idy") = Trim(MS.TextMatrix(R, 3))
                    RSCOM("act_inc") = Trim(list1.Text)
                    RSCOM.Update
                End If
            Next
            'MsgBox "New Transaction ID is " & txnid, vbExclamation
            cancel
        End If
    Else
        If RSCOM.State = 1 Then RSCOM.Close
        If RSCOM.State = 1 Then RSCOM.Close
        RSCOM.Open "select * from com where txn_idy=''", CON, adOpenKeyset, adLockOptimistic
        R = 1
        If rschk.State = 1 Then rschk.Close
        rschk.Open "select * from com where smn_idy='" & Trim(SMNID) & "'", CON, adOpenKeyset, adLockOptimistic
        If Not rschk.EOF Then
            If FDATE.Value >= rschk("frm_dat") And FDATE.Value <= rschk("to_dat") Then
                MsgBox "Transaction Exists between " & Format(rschk("frm_dat"), "dd/mmm/yyyy") & " to " & Format(rschk("to_dat"), "dd/mmm/yyyy"), vbExclamation
                FDATE.SetFocus
                Exit Sub
            End If
        End If
        'RSCOM.Open "SELECT * FROM COM WHERE TXN_IDY=''", CON, adOpenKeyset, adLockOptimistic
'        txnid = increment("idy_tsn")
            If range.Value = 1 Then
                For R = 1 To msr.Rows - 1
                    If Val(msr.TextMatrix(R, 2)) > 0 Then
                        RSCOM.AddNew
                        For I = 0 To RSCOM.Fields.Count - 1
                            If RSCOM(I).Type = adVarChar Or RSCOM(I).Type = adChar Or RSCOM(I).Type = adBSTR Or RSCOM(I).Type = adLongVarChar Or RSCOM(I).Type = adLongVarWChar Then
                                RSCOM(I) = " "
                            ElseIf RSCOM(I).Type = adInteger Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adSingle Or RSCOM(I).Type = adDouble Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adCurrency Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adSmallInt Or RSCOM(I).Type = adDecimal Then
                                RSCOM(I) = 0
                            ElseIf RSCOM(I).Type = adDate Or RSCOM(I).Type = adDBDate Or RSCOM(I).Type = adDBTimeStamp Then
                                RSCOM(I) = Date
                            End If
                        Next
                        'RSCOM("txn_idy") = txnid
                        RSCOM("com_typ") = Trim(comtyp.Text)
                        RSCOM("smn_idy") = Trim(SMNID)
                        RSCOM("frm_dat") = FDATE.Value
                        RSCOM("to_dat") = TDATE.Value
                        RSCOM("beg_rng") = Val(msr.TextMatrix(R, 0))
                        RSCOM("end_rng") = Val(msr.TextMatrix(R, 1))
                        RSCOM("com_per") = Val(msr.TextMatrix(R, 2))
                        RSCOM("chk_flg") = "R"
                        RSCOM("act_inc") = Trim(list1.Text)
                        RSCOM.Update
                    End If
                Next
            Else
                RSCOM.AddNew
                For I = 0 To RSCOM.Fields.Count - 1
                    If RSCOM(I).Type = adVarChar Or RSCOM(I).Type = adChar Or RSCOM(I).Type = adBSTR Or RSCOM(I).Type = adLongVarChar Or RSCOM(I).Type = adLongVarWChar Then
                        RSCOM(I) = " "
                    ElseIf RSCOM(I).Type = adInteger Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adSingle Or RSCOM(I).Type = adDouble Or RSCOM(I).Type = adBigInt Or RSCOM(I).Type = adCurrency Or RSCOM(I).Type = adNumeric Or RSCOM(I).Type = adSmallInt Or RSCOM(I).Type = adDecimal Then
                        RSCOM(I) = 0
                    ElseIf RSCOM(I).Type = adDate Or RSCOM(I).Type = adDBDate Or RSCOM(I).Type = adDBTimeStamp Then
                        RSCOM(I) = Date
                    End If
                Next
                'RSCOM("txn_idy") = txnid
                RSCOM("com_typ") = Trim(comtyp.Text)
                RSCOM("smn_idy") = Trim(SMNID)
                RSCOM("frm_dat") = FDATE.Value
                RSCOM("to_dat") = TDATE.Value
                RSCOM("chk_flg") = "P"
                RSCOM("com_amt") = Val(comper)
                RSCOM("act_inc") = Trim(list1.Text)
                RSCOM.Update
            End If
                'MsgBox "New Transaction ID is " & txnid, vbExclamation
            cancel
    End If
Else
    If RSCOM.State = 1 Then RSCOM.Close
    'RSCOM.Open "SELECT * FROM COM WHERE TXN_IDY='" & Trim(txnid) & "'", CON, adOpenKeyset, adLockOptimistic
    If Trim(comtyp.Text) = "Quantity" Then
        If All.Value = 1 Then
            'RSCOM("txn_idy") = txnid
            RSCOM("com_typ") = Trim(comtyp.Text)
            RSCOM("smn_idy") = Trim(SMNID)
            RSCOM("frm_dat") = FDATE.Value
            RSCOM("to_dat") = TDATE.Value
            If All.Value = 1 Then
                RSCOM("chk_flg") = "A"
            Else
                RSCOM("chk_flg") = "S"
            End If
            RSCOM("com_amt") = Val(txtIncAmt)
            RSCOM("act_inc") = Trim(list1.Text)
            RSCOM.Update
            cancel
        Else
            R = 1
            While Not RSCOM.EOF
                RSCOM("com_typ") = Trim(comtyp.Text)
                RSCOM("smn_idy") = Trim(SMNID)
                RSCOM("frm_dat") = FDATE.Value
                RSCOM("to_dat") = TDATE.Value
                If All.Value = 1 Then
                    RSCOM("chk_flg") = "A"
                Else
                    RSCOM("chk_flg") = "S"
                End If
                RSCOM("com_amt") = MS.TextMatrix(R, 2)
                RSCOM("prd_idy") = MS.TextMatrix(R, 3)
                RSCOM("act_inc") = Trim(list1.Text)
                RSCOM.Update
                RSCOM.MoveNext
                R = R + 1
            Wend
            cancel
        End If
    Else
        If RSCOM("CHK_FLG") = "R" Then
            RSCOM("com_typ") = Trim(comtyp.Text)
            RSCOM("smn_idy") = Trim(SMNID)
            RSCOM("frm_dat") = FDATE.Value
            RSCOM("to_dat") = TDATE.Value
            RSCOM("act_inc") = Trim(list1.Text)
            R = 1
            While Not RSCOM.EOF
                RSCOM("BEG_RNG") = Val(msr.TextMatrix(R, 0))
                RSCOM("END_RNG") = Val(msr.TextMatrix(R, 1))
                RSCOM("COM_PER") = Val(msr.TextMatrix(R, 2))
                R = R + 1
                RSCOM.Update
                RSCOM.MoveNext
            Wend
        Else
            RSCOM("com_typ") = Trim(comtyp.Text)
            RSCOM("smn_idy") = Trim(SMNID)
            RSCOM("frm_dat") = FDATE.Value
            RSCOM("to_dat") = TDATE.Value
            If range.Value = 1 Then
                RSCOM("chk_flg") = "R"
            Else
                RSCOM("chk_flg") = "P"
            End If
            RSCOM("com_amt") = Val(comper)
            RSCOM("act_inc") = Trim(list1.Text)
            RSCOM.Update
        End If
        cancel
    End If
End If
End Sub
Private Sub comlist()
If CHECKTB("com.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
'RS.Open "select * from com where txn_idy like '" & Trim(txnid) & "%' group by txn_idy", CON, adOpenStatic
Label20.Caption = "Select Transaction"
LV.ColumnHeaders.Clear
LV.ColumnHeaders.Add , , "Transaction ID", 2000
LV.ColumnHeaders.Add , , "Salesman", 2500
'lv.ColumnHeaders.Add , , "Commission Amt", 1500, 2
LV.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = LV.ListItems.Add(, , Trim(RS("txn_idy")))
        If RSsmn.State = 1 Then RSsmn.Close
        RSsmn.Open "select smn_nme from smn where smn_idy='" & RS("smn_idy") & "'", CON
        If Not RSsmn.EOF Then
            ls.ListSubItems.Add , , Trim(RSsmn("smn_nme"))
        End If
 '       ls.ListSubItems.Add , , Trim(rs("com_amt") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    LV.Visible = True
    LV.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Transaction ID Not Found, Define Transaction !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
'        txnid.Enabled = True
'        txnid.SetFocus
        MD = False
    End If
End If
End Sub

Private Sub chk(b As Boolean)
Label4.Visible = b
txtQuan.Visible = b
Label5.Visible = b
txtIncAmt.Visible = b
End Sub
Private Sub DISPMSR()
msr.Visible = True
msr.Cols = 3
msr.Rows = 6
msr.TextMatrix(0, 0) = "Start"
msr.ColWidth(0) = 1000
SCl msr, 0
msr.ColAlignment(0) = 7
msr.TextMatrix(0, 1) = "End"
SCl msr, 1
msr.ColWidth(1) = 1000
msr.ColAlignment(1) = 7
msr.TextMatrix(0, 2) = "Commission %"
msr.ColWidth(2) = 1200
SCl msr, 2
msr.ColAlignment(2) = 7
msr.TextMatrix(1, 0) = "0.00"
msr.SetFocus
End Sub
Private Function CHKROWS()
For R = 1 To msr.Rows - 1
    If Val(msr.TextMatrix(R, 0)) = 0 And Val(msr.TextMatrix(R, 1)) = 0 Then
    ElseIf Val(msr.TextMatrix(R, 0)) < Val(msr.TextMatrix(R, 1)) Then
    Else
        MsgBox "Invalid Range of Values", vbExclamation
        msr.SetFocus
        msr.Row = R
        msr.Col = 0
        CHKROWS = 1
        Exit Function
    End If
Next
End Function
