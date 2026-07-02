VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_HVO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HVO Discounts - Schemes"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "M_FRM_HVO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5300
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   405
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&HVO Discounts - Schemes"
      TabPicture(0)   =   "M_FRM_HVO.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CSR_END_DAT"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "MS"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "M_FLX_AMT"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "MER_REF_NUM"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "MER_REF_QTY"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "HVO_DIS_TYP"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "MER_REF_AMT"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CSR_SCH_NME"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CSR_BEG_DAT"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CSR_STATUS"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "M_CST_IDY"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CSR_SCH_IDY"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CSR_DIS_AMT"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_HVO.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "list4"
      Tab(1).Control(2)=   "V_LST_CAP"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command1 
         Caption         =   "&Fill "
         Height          =   375
         Left            =   7680
         TabIndex        =   35
         Top             =   2200
         Width           =   975
      End
      Begin VB.TextBox CSR_DIS_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   33
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox CSR_SCH_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3720
         TabIndex        =   32
         Top             =   487
         Width           =   1095
      End
      Begin VB.TextBox M_CST_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   487
         Width           =   975
      End
      Begin VB.ListBox CSR_STATUS 
         Height          =   255
         ItemData        =   "M_FRM_HVO.frx":047A
         Left            =   1440
         List            =   "M_FRM_HVO.frx":0484
         TabIndex        =   29
         Top             =   2280
         Width           =   495
      End
      Begin MSComCtl2.DTPicker CSR_BEG_DAT 
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   1350
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16515075
         CurrentDate     =   39783
         MinDate         =   39539
      End
      Begin VB.TextBox CSR_SCH_NME 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   885
         Width           =   3375
      End
      Begin VB.TextBox MER_REF_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   3840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2235
         Width           =   975
      End
      Begin VB.ComboBox HVO_DIS_TYP 
         Height          =   315
         ItemData        =   "M_FRM_HVO.frx":048E
         Left            =   6120
         List            =   "M_FRM_HVO.frx":0498
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1860
         Width           =   1335
      End
      Begin VB.TextBox MER_REF_QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   3840
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1830
         Width           =   975
      End
      Begin VB.TextBox MER_REF_NUM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3840
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox M_FLX_AMT 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   4560
         TabIndex        =   16
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74520
         TabIndex        =   9
         Top             =   4440
         Width           =   7815
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "HVO Details"
         Height          =   1335
         Left            =   4920
         TabIndex        =   6
         Top             =   487
         Width           =   3735
         Begin VB.Label M_CST_AD4 
            AutoSize        =   -1  'True
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   45
         End
         Begin VB.Label M_CST_AD1 
            AutoSize        =   -1  'True
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   45
         End
         Begin VB.Label M_CST_NME 
            AutoSize        =   -1  'True
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   45
         End
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3375
         Left            =   -74520
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   7815
         _ExtentX        =   13785
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Retail - Outlet Name"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Market"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Account Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Grade"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "National Key A/c"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2640
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
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
      Begin MSComCtl2.DTPicker CSR_END_DAT 
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   1815
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16515075
         CurrentDate     =   39783
         MinDate         =   39539
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Amt"
         Height          =   195
         Left            =   7680
         TabIndex        =   34
         Top             =   1920
         Width           =   270
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Scheme ID"
         Height          =   195
         Left            =   2880
         TabIndex        =   30
         Top             =   555
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Active"
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   2310
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         Height          =   195
         Left            =   600
         TabIndex        =   25
         Top             =   1905
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         Height          =   195
         Left            =   555
         TabIndex        =   24
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Scheme Name"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   975
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ref Amount"
         Height          =   195
         Left            =   2880
         TabIndex        =   21
         Top             =   2310
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ref Quantity"
         Height          =   195
         Left            =   2880
         TabIndex        =   20
         Top             =   1905
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MER Ref No"
         Height          =   195
         Left            =   2880
         TabIndex        =   19
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Discount Type"
         Height          =   195
         Left            =   4920
         TabIndex        =   18
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "HVO ID"
         Height          =   195
         Left            =   720
         TabIndex        =   17
         Top             =   555
         Width           =   555
      End
      Begin VB.Label V_LST_CAP 
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
         Left            =   -74520
         TabIndex        =   8
         Top             =   600
         Width           =   7815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   120
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
            Picture         =   "M_FRM_HVO.frx":04AE
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":09F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":0DBA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":110E
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":1746
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":1A9A
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":1EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":21EE
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":25FA
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":2706
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":2A5A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":2E6E
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HVO.frx":33BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
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
            Enabled         =   0   'False
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
            Enabled         =   0   'False
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
            Object.ToolTipText     =   "Prefix Salesmen ID "
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
            Enabled         =   0   'False
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
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_HVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
V_DLR_LST
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
V_DLR_LST
End Sub
Private Sub Command1_Click()
For A = 1 To MS.Rows - 1
    MS.TextMatrix(A, 2) = amt(Val(CSR_DIS_AMT) & "")
Next
If MS.Rows > 1 Then
    MS.SetFocus
    MS.Col = 0
    MS.Row = 1
End If
End Sub
Private Sub CSR_DIS_AMT_GotFocus()
SendKeys "{home}+{End}"
End Sub
Private Sub CSR_DIS_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SSTab1.Tab = 0
    MS.SetFocus
ElseIf KeyCode = 38 Then
    If HVO_DIS_TYP.Enabled = True Then
        HVO_DIS_TYP.SetFocus
    Else
        CSR_DIS_AMT.SetFocus
    End If
End If
End Sub
Private Sub CSR_END_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If DateDiff("D", CSR_BEG_DAT, CSR_END_DAT) < 0 Then
        SSTab1.Tab = 0
        MsgBox "End Date Should Be Greater Than Start Date", vbInformation, "DAS Version 4.0"
        CSR_END_DAT.SetFocus
        Exit Sub
    Else
        CSR_STATUS.Enabled = True
        CSR_STATUS.SetFocus
    End If
ElseIf KeyCode = 38 Then
    CSR_BEG_DAT.SetFocus
End If
End Sub
Private Sub CSR_SCH_NME_GotFocus()
'If (Trim(CSR_SCH_NME)) = "Enter HVO Discount Description Here" Then
'    CSR_SCH_NME = ""
'End If
End Sub
Private Sub CSR_SCH_NME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(CSR_SCH_NME)) > 0 Then
        CSR_BEG_DAT.SetFocus
    Else
        MsgBox "Scheme Name Required!", vbInformation, "DAS Version 4.0"
        CSR_SCH_NME.SetFocus
    End If
ElseIf KeyCode = 38 Then
    If Len(Trim(CSR_SCH_NME)) > 0 Then
        MsgBox "Scheme Name Required!", vbInformation, "DAS Version 4.0"
        CSR_SCH_NME.SetFocus
    Else
        M_CST_IDY.SetFocus
    End If
End If
End Sub
Private Sub CSR_BEG_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CSR_END_DAT.SetFocus
ElseIf KeyCode = 38 Then
    CSR_BEG_DAT.SetFocus
End If
End Sub
Private Sub CSR_STATUS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    MER_REF_NUM.SetFocus
ElseIf KeyCode = 38 Then
    CSR_END_DAT.SetFocus
End If
End Sub
Private Sub Form_Activate()
SETYEAROPEN M_FRM_HVO
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
M_CST_IDY.SetFocus
End Sub
Private Sub V_DLR_LST()
If CHECKTB("act.dbf,smn.dbf,mar.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
'    REC_SET.Open "select ACT.*,MAR.MAR_NME from ACT,MAR where (ACT.ACT_TYP = 'BOT' OR ACT.ACT_TYP = 'SAL') and ACT_NME like '" & Trim(M_CST_IDY) & "%' and ACT.MAR_IDY = MAR.MAR_IDY  AND (KEY_ACT='Y' or CST_TYP in ('Key A/c','Local Key A/c','Regnl Key A/c'))", CON, adOpenKeyset, adlockpessimistic
    REC_SET.Open "select ACT.*,MAR.MAR_NME from ACT,MAR where (ACT.ACT_TYP = 'BOT' OR ACT.ACT_TYP = 'SAL') and (ACT.CST_TYP = 'HVO-P' OR ACT.CST_TYP = 'HVO-NP' OR ACT.CST_TYP = 'HVO-P-WO' OR ACT.CST_TYP = 'HVO-NP-WO') and ACT_NME like '" & Trim(M_CST_IDY) & "%' and ACT.MAR_IDY = MAR.MAR_IDY", con, adOpenKeyset, adLockPessimistic
ElseIf Check2.Value = 1 Then
'    REC_SET.Open "select ACT.*,MAR.MAR_NME from ACT,MAR where (ACT.ACT_TYP = 'BOT' OR ACT.ACT_TYP = 'SAL') and ACT_AD4 like '" & Trim(M_CST_IDY) & "%' and ACT.MAR_IDY = MAR.MAR_IDY  AND (KEY_ACT='Y' or CST_TYP in ('Key A/c','Local Key A/c','Regnl Key A/c'))", CON, adOpenKeyset, adlockpessimistic
    REC_SET.Open "select ACT.*,MAR.MAR_NME from ACT,MAR where (ACT.ACT_TYP = 'BOT' OR ACT.ACT_TYP = 'SAL') and (ACT.CST_TYP = 'HVO-P' OR ACT.CST_TYP = 'HVO-NP' OR ACT.CST_TYP = 'HVO-P-WO' OR ACT.CST_TYP = 'HVO-NP-WO') and ACT_AD4 like '" & Trim(M_CST_IDY) & "%' and ACT.MAR_IDY = MAR.MAR_IDY", con, adOpenKeyset, adLockPessimistic
End If
list4.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("ACT_NME")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_IDY"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("MAR_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("CST_TYP") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("RET_TYP") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("BBN_ONE") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    V_LST_CAP = "Select Key Account - Key Retailer"
    list4.Visible = True
'    list4.ListItems(1).Selected = True 'Added by PNR
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Key Account - Key Retailer Not Found!", vbInformation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    End If
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 83 And Shift = 2 Then saverecord
If KeyCode = 88 And Shift = 2 Then Unload Me
If KeyCode = 81 And Shift = 2 Then
'    MS.Rows = 1
'    M_CST_IDY = ""
'    M_CST_NME = ""
'    M_CST_AD1 = ""
'    M_CST_AD4 = ""
'    M_CST_IDY.SetFocus
    cancel
End If
If KeyCode = 112 Then M_HELP
If KeyCode = 69 And Shift = 2 Then If M_FRM_HVO.ActiveControl.name = "M_CST_IDY" Then V_DLR_LST
End Sub
Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If SSTab1.Tab = 1 And V_LST_CAP.Caption = "Select Key Retailer" Then
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    ElseIf M_FRM_HVO.ActiveControl.name = "M_CST_IDY" And M_CST_IDY = "" Then
        Unload Me
    Else
        saverecord
    End If
End If
End Sub
Private Sub Form_Load()
CSR_BEG_DAT = Date
CSR_END_DAT = Date
DISPMS
End Sub
Private Sub HVO_DIS_TYP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SSTab1.Tab = 0
   MS.SetFocus
End If
End Sub
Private Sub list4_DblClick()
list4_KeyPress 13
End Sub
Private Sub list4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list4.ListItems.count > 0 Then
'        list4.ListItems(1).Selected = True 'Added by PNR
        M_CST_IDY = list4.SelectedItem.ListSubItems(1)
        M_CST_IDY_keydown 13, 0
        list4.Visible = False
        SSTab1.Tab = 0
        CSR_SCH_NME.SetFocus
    '    DISP_PRD
    '    MS.SetFocus
    End If
End If
End Sub
Private Sub M_CST_IDY_keydown(K As Integer, s As Integer)
If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If K = 13 Then
    If REC_SET.State = 1 Then REC_SET.Close
    'REC_SET.Open "SELECT * FROM ACT WHERE ACT_IDY = '" & Sincrement(Trim(M_CST_IDY)) & "' AND (ACT_TYP = 'BOT' OR ACT_TYP = 'SAL') and (KEY_ACT = 'Y' or CST_TYP in ('Key A/c','Local Key A/c','Regnl Key A/c'))", CON, adOpenKeyset, adlockpessimistic
    REC_SET.Open "SELECT * FROM ACT WHERE ACT_IDY = '" & Sincrement(Trim(M_CST_IDY)) & "' AND (ACT_TYP = 'BOT' OR ACT_TYP = 'SAL') AND (CST_TYP = 'HVO-P' OR CST_TYP = 'HVO-NP' OR ACT.CST_TYP = 'HVO-P-WO' OR ACT.CST_TYP = 'HVO-NP-WO')", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        M_CST_IDY = Sincrement(M_CST_IDY)
        If Len(Trim(M_CST_IDY)) = 0 Then Exit Sub
        M_CST_NME = Trim(REC_SET("act_nme") & "")
        M_CST_AD1 = Trim(REC_SET("act_ad1") & "")
        M_CST_AD2 = Trim(REC_SET("act_ad2") & "")
        M_CST_AD4 = Trim(REC_SET("act_ad4") & "")
        If Trim(REC_SET("CST_TYP")) = "HVO-NP" Or Trim(REC_SET("CST_TYP")) = "HVO-NP-WO" Then
            HVO_DIS_TYP.ListIndex = 0
            CSR_STATUS.ListIndex = 0
            HVO_DIS_TYP.Enabled = False
        Else
            HVO_DIS_TYP.ListIndex = 1
            CSR_STATUS.ListIndex = 1
            HVO_DIS_TYP.Enabled = False
        End If
        DISP_PRD
    Else
        V_DLR_LST
    End If
End If
End Sub
Public Sub DISP_PRD()
If CHECKTB("prd.dbf,brd.dbf,csr.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
End If
Me.Refresh
Dim REC_SET As New ADODB.Recordset
Dim REC_SET_P As New ADODB.Recordset
REC_SET_P.Open "SELECT * FROM PRD,brd where prd.prd_idy = brd.prd_idy and stk_typ in('Salable','Free')", con, adOpenKeyset, adLockPessimistic
MS.Rows = REC_SET_P.RecordCount + 1
For A = 1 To REC_SET_P.RecordCount
    'MS.AddItem ""
    MS.TextMatrix(A, 0) = Trim(REC_SET_P("PRD_IDY"))
    MS.TextMatrix(A, 1) = convert(REC_SET_P("PRD_SDR"))
    MS.TextMatrix(A, 2) = "0.00"   ' convert(REC_SET_P("prd_sdr"))
    MS.TextMatrix(A, 3) = Trim(REC_SET_P("PRD_NME"))
    MS.TextMatrix(A, 4) = convert(REC_SET_P("PRD_MRP"))
    MS.TextMatrix(A, 5) = Trim(REC_SET_P("BAT_IDY"))
    REC_SET_P.MoveNext
Next
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT ACT_IDY, B_PRD_IDY, B_PRD_RAT, B_BAT_IDY, CSR_IDY, MER_IDY, MER_QTY, MER_AMT, DIS_TYP, CSR_NME, CSR_DAT, BEG_DAT, END_DAT, CSR_STS, SCH_IDY, CSR_IDY, DIS_PCG, DIS_AMT FROM CSR WHERE ACT_IDY = '" & Trim(M_CST_IDY) & "'", con, adOpenKeyset, adLockPessimistic

If REC_SET.RecordCount > 0 Then
    REC_SET.MoveFirst
    If IsNull(Trim(REC_SET(5))) = False Then
        MER_REF_NUM = Trim(REC_SET(5))
    Else
        MER_REF_NUM = ""
    End If
    If IsNull(Trim(REC_SET(6))) = False Then
        MER_REF_QTY = QTY(Val(Trim(REC_SET(6))))
    Else
        MER_REF_QTY = "0.000"
    End If
    If IsNull(Trim(REC_SET(7))) = False Then
        MER_REF_AMT = amt(Val(Trim(REC_SET(7))))
    Else
        MER_REF_AMT = "0.00"
    End If
    If Trim(REC_SET(8)) = "OnSpot" Then
        HVO_DIS_TYP.ListIndex = 0
        HVO_DIS_TYP.Enabled = False
    ElseIf Trim(REC_SET(8)) = "Periodic" Then
        HVO_DIS_TYP.ListIndex = 1
        HVO_DIS_TYP.Enabled = False
    End If
    CSR_SCH_IDY.Enabled = True
    CSR_SCH_IDY = Trim(REC_SET(14))
    CSR_SCH_IDY.Enabled = False
    CSR_SCH_NME = Trim(REC_SET(9))
    CSR_BEG_DAT = Trim(REC_SET(11))
    CSR_END_DAT = Trim(REC_SET(12))
    If Trim(REC_SET(13)) = "Y" Then
        CSR_STATUS.Enabled = True
        CSR_STATUS.ListIndex = 0
    ElseIf Trim(REC_SET(13)) = "N" Then
        CSR_STATUS.ListIndex = 1
        CSR_STATUS.Enabled = True
    End If
End If
For A = 1 To MS.Rows - 1
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "SELECT ACT_IDY, B_PRD_IDY, B_PRD_RAT, B_BAT_IDY, CSR_IDY, MER_IDY, MER_QTY, MER_AMT, DIS_TYP, CSR_NME, CSR_DAT, BEG_DAT, END_DAT, CSR_STS, SCH_IDY, CSR_IDY, DIS_PCG, DIS_AMT FROM CSR WHERE B_BAT_IDY = '" & Trim(MS.TextMatrix(A, 5)) & "' AND ACT_IDY = '" & Trim(M_CST_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then MS.TextMatrix(A, 2) = amt(REC_SET("DIS_AMT") & "")
Next
If MS.Rows > 1 Then
    MS.SetFocus
    MS.Col = 0
    MS.Row = 1
End If
End Sub

Private Sub M_CST_IDY_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
'If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub

Private Sub M_FLX_AMT_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    MS.Text = amt(M_FLX_AMT)
    M_FLX_AMT.Visible = False
    MS.SetFocus
End If
End Sub
Private Sub M_FLX_AMT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 13 Then
Else
K = 0
End If
End Sub
Private Sub MER_REF_AMT_GotFocus()
SendKeys "{home}+{End}"
End Sub
Private Sub MER_REF_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If HVO_DIS_TYP.Enabled = True Then
        HVO_DIS_TYP.SetFocus
    Else
        CSR_DIS_AMT.SetFocus
    End If
ElseIf KeyCode = 38 Then
    MER_REF_QTY.SetFocus
End If
End Sub
Private Sub MER_REF_AMT_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub
Private Sub MER_REF_NUM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    MER_REF_QTY.SetFocus
ElseIf KeyCode = 38 Then
    CSR_STATUS.SetFocus
End If
End Sub
Private Sub MER_REF_NUM_KeyPress(KeyAscii As Integer)
'If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub
Private Sub MER_REF_QTY_GotFocus()
SendKeys "{home}+{End}"
End Sub
Private Sub MER_REF_QTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    MER_REF_AMT.SetFocus
ElseIf KeyCode = 38 Then
    MER_REF_QTY.SetFocus
End If
End Sub
Private Sub MER_REF_QTY_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
End Sub
Private Sub MS_GotFocus()
'ms.HighLight = flexHighlightAlways
End Sub
Private Sub MS_KeyPress(K As Integer)
If MS.Col = 2 Then
    If K > 47 And K < 58 Or K = 46 Then
        M_FLX_AMT = ""
        M_FLX_AMT.MaxLength = 9
        M_FLX_AMT.Visible = True
        M_FLX_AMT.Top = MS.Top + MS.CellTop
        M_FLX_AMT.Left = MS.CellLeft + MS.Left
        M_FLX_AMT.Width = MS.CellWidth
        M_FLX_AMT.Height = MS.CellHeight
        M_FLX_AMT.Visible = True
        M_FLX_AMT.SetFocus
        M_FLX_AMT = Chr(K)
        M_FLX_AMT.SelStart = Len(M_FLX_AMT)
    End If
End If
End Sub

Private Sub MS_LeaveCell()
If M_FLX_AMT.Visible = True Then M_FLX_AMT.Visible = False
End Sub

Private Sub MS_Scroll()
If M_FLX_AMT.Visible = True Then
    M_FLX_AMT = ""
    M_FLX_AMT.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
'newrecord
ElseIf Button.KEY = "s" Then
saverecord
ElseIf Button.KEY = "e" Then
Unload Me
ElseIf Button.KEY = "p" Then
    Printrecord
ElseIf Button.KEY = "c" Then
cancel
ElseIf Button.KEY = "crt" Then
'CREATE
ElseIf Button.KEY = "ed" Then
    If M_FRM_HVO.ActiveControl.name = "M_CST_IDY" Then
        V_LST_CAP = "Select Retailer"
        V_DLR_LST
    End If
ElseIf Button.KEY = "h" Then
M_HELP
End If
End Sub
Public Sub Printrecord()
    OpenReport CURDIR & "\lists.RP1", "Key Retailers", 0, SAL_QRY, 6
End Sub
Public Sub saverecord()
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbYes Then
    If CHECKTB("act.dbf,csr.dbf,sch.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
'
'    If check("act", "act_idy", Trim(M_CST_IDY) & "' AND KEY_ACT = 'Y") = False Then
'        MsgBox "Invalid Key Account - Key Retailer!", vbExclamation, "DAS Version 4.0"
'        M_CST_IDY.SetFocus
'        Exit Sub
'    End If
    
    Me.MousePointer = vbHourglass
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT ACT_IDY, B_PRD_IDY, B_PRD_RAT, B_BAT_IDY, CSR_IDY, MER_IDY, MER_QTY, MER_AMT, DIS_TYP, CSR_NME, CSR_DAT, BEG_DAT, END_DAT, CSR_STS, SCH_IDY, CSR_IDY, DIS_PCG, DIS_AMT, XCSR_IDY FROM CSR WHERE ACT_IDY = '" & Trim(M_CST_IDY & "") & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
'        If Trim(rs(13)) = "N" Then
'            MsgBox "Cannot Update HVO Details, HVO Scheme Is Not Active !", vbInformation, "DAS Version 4.0"
'            cancel
'            Exit Sub
'        End If
        If Len(Trim(MER_REF_NUM)) = 0 Then
            MsgBox "MER Ref Number Required!", vbInformation, "DAS Version 4.0"
            MER_REF_NUM.SetFocus
            Exit Sub
        End If
        If Len(Trim(CSR_SCH_NME)) = 0 Then
            MsgBox "HVO Discount Scheme Name Required!", vbInformation, "DAS Version 4.0"
            CSR_SCH_NME.SetFocus
            Exit Sub
        End If
        For A = 1 To MS.Rows - 1
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                rs.Find "B_BAT_IDY = '" & Trim(MS.TextMatrix(A, 5)) & "'"
            End If
            If rs.EOF = True Then
                rs.AddNew
                rs(0) = Trim(M_CST_IDY)
                rs(1) = Trim(MS.TextMatrix(A, 0))
                rs(16) = 0#
                rs(17) = Val(MS.TextMatrix(A, 2))   'Discount in Amount
                rs(3) = Trim(MS.TextMatrix(A, 5))
                rs(4) = ""
                rs(5) = Trim(MER_REF_NUM)
                rs(6) = Trim(MER_REF_QTY)
                rs(7) = Trim(MER_REF_AMT)
                HVO_DIS_TYP = GetValue("CST_TYP", "ACT WHERE ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'")
                If Trim(HVO_DIS_TYP) = "HVO-NP" Or Trim(HVO_DIS_TYP) = "HVO-NP-WO" Then
                    HVO_DIS_TYP = "OnSpot"
                Else
                    HVO_DIS_TYP = "Periodic"
                End If
                rs(8) = Trim(HVO_DIS_TYP)
                rs(9) = Trim(CSR_SCH_NME)
                rs(10) = Trim(CSR_BEG_DAT)
                rs(11) = Trim(CSR_BEG_DAT)
                rs(12) = Trim(CSR_END_DAT)
                If CSR_STATUS.ListIndex = 0 Then
                    rs(13) = Trim("Y")
                Else
                    rs(13) = Trim("N")
                End If
                rs(14) = Trim(CSR_SCH_IDY)
                rs(15) = Trim(M_CSR_IDY)
                rs(18) = ""
                rs.Update
            Else
                rs(16) = 0#
                rs(17) = Val(MS.TextMatrix(A, 2))   'Discount in Amount
                rs(5) = Trim(MER_REF_NUM)
                rs(6) = Trim(MER_REF_QTY)
                rs(7) = Trim(MER_REF_AMT)
                HVO_DIS_TYP = GetValue("CST_TYP", "ACT WHERE ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'")
                If Trim(HVO_DIS_TYP) = "HVO-NP" Or Trim(HVO_DIS_TYP) = "HVO-NP-WO" Then
                    HVO_DIS_TYP = "OnSpot"
                Else
                    HVO_DIS_TYP = "Periodic"
                End If
                rs(8) = Trim(HVO_DIS_TYP)
                rs(9) = Trim(CSR_SCH_NME)
                rs(10) = Trim(CSR_BEG_DAT)
                rs(11) = Trim(CSR_BEG_DAT)
                rs(12) = Trim(CSR_END_DAT)
                If CSR_STATUS.ListIndex = 0 Then
                    rs(13) = Trim("Y")
                Else
                    rs(13) = Trim("N")
                End If
                rs(18) = ""
                rs.Update
            End If
        Next
    Else
        M_CSR_IDY = "CSR" & T7increment("IDY_CSR")
        CSR_SCH_IDY = "SCH" & T7increment("IDY_SCH")
        For A = 1 To MS.Rows - 1
            rs.AddNew
            rs(0) = Trim(M_CST_IDY)
            rs(1) = Trim(MS.TextMatrix(A, 0))
            rs(16) = 0#
            rs(17) = Val(MS.TextMatrix(A, 2))   'Discount in Amt
            rs(3) = Trim(MS.TextMatrix(A, 5))
            rs(4) = Trim(M_CSR_IDY)
            rs(5) = Trim(MER_REF_NUM)
            rs(6) = Val(Trim(MER_REF_QTY))
            rs(7) = Val(Trim(MER_REF_AMT))
            HVO_DIS_TYP = GetValue("CST_TYP", "ACT WHERE ACT_IDY LIKE '" & Trim(M_CST_IDY) & "'")
            If Trim(HVO_DIS_TYP) = "HVO-NP" Or Trim(HVO_DIS_TYP) = "HVO-NP-WO" Then
                HVO_DIS_TYP = "OnSpot"
            Else
                HVO_DIS_TYP = "Periodic"
            End If
            rs(8) = Trim(HVO_DIS_TYP)
            rs(9) = Trim(CSR_SCH_NME)
            rs(10) = Trim(CSR_BEG_DAT)
            rs(11) = Trim(CSR_BEG_DAT)
            rs(12) = Trim(CSR_END_DAT)
            If CSR_STATUS.ListIndex = 0 Then
                rs(13) = Trim("Y")
            Else
                rs(13) = Trim("N")
            End If
            rs(14) = Trim(CSR_SCH_IDY)
            rs(18) = ""
            rs.Update
        Next
    End If
    rs.Close
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM SCH where SCH_IDY = '" & Trim(CSR_SCH_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 And Len(Trim(M_CSR_IDY)) > 0 And Len(Trim(M_CST_IDY)) > 0 Then
        rs("SCH_NME") = Trim(CSR_SCH_NME)
        rs("SCH_BDT") = Format(CSR_BEG_DAT, "DD/MM/YYYY")
        rs("SCH_EDT") = Format(CSR_END_DAT, "DD/MM/YYYY")
        If CSR_STATUS.ListIndex = 0 Then
            rs("SCH_STS") = Trim("Y")
        Else
            rs("SCH_STS") = Trim("N")
        End If
        rs("MER_NUM") = Trim(MER_REF_NUM)
        rs("MER_QTY") = Trim(MER_REF_QTY)
        rs("MER_AMT") = Trim(MER_REF_AMT)
        rs("TXN_TYP") = "CSR"
        rs("SYS_DAT") = Date
        rs("SYS_TIM") = Time
        rs.Update
    Else
        If Len(Trim(M_CSR_IDY)) > 0 And Len(Trim(M_CST_IDY)) > 0 Then
            rs.AddNew
            rs("SCH_IDY") = Trim(CSR_SCH_IDY)
            rs("SCH_NME") = Trim(CSR_SCH_NME)
            rs("SCH_BDT") = Format(CSR_BEG_DAT, "DD/MM/YYYY")
            rs("SCH_EDT") = Format(CSR_END_DAT, "DD/MM/YYYY")
            If CSR_STATUS.ListIndex = 0 Then
                rs("SCH_STS") = Trim("Y")
            Else
                rs("SCH_STS") = Trim("N")
            End If
            rs("MER_NUM") = Trim(MER_REF_NUM)
            rs("MER_QTY") = Val(Trim(MER_REF_QTY))
            rs("MER_AMT") = Val(Trim(MER_REF_AMT))
            rs("TXN_TYP") = "CSR"
            rs("SYS_DAT") = Date
            rs("SYS_TIM") = Format(Time, "HH:MM:SS")
            rs("BAT_IDY") = ""
            rs.Update
        End If
    End If
    MS.Rows = 1
    M_CST_IDY = ""
    M_CST_IDY.SetFocus
    M_CST_NME = ""
    M_CST_AD1 = ""
    M_CST_AD4 = ""
    M_CST_IDY.SetFocus
    Me.MousePointer = vbDefault
Else
    MS.Rows = 1
    M_CST_IDY = ""
    M_CST_IDY.SetFocus
    M_CST_NME = ""
    M_CST_AD1 = ""
    M_CST_AD4 = ""
    M_CST_IDY.SetFocus
    Me.MousePointer = vbDefault
End If
cancel
End Sub
Public Sub cancel()
    M_CST_IDY = ""
    M_CST_IDY.SetFocus
    M_CST_NME = ""
    M_CST_AD1 = ""
    M_CST_AD4 = ""
    CSR_SCH_IDY.Enabled = True
    CSR_SCH_IDY = ""
    CSR_SCH_IDY.Enabled = False
    CSR_SCH_NME = ""
    CSR_BEG_DAT = Date
    CSR_END_DAT = Date
    CSR_STATUS.ListIndex = 0
    MER_REF_NUM = ""
    MER_REF_QTY = "0.000"
    MER_REF_AMT = "0.00"
    CSR_DIS_AMT = "0.00"
    HVO_DIS_TYP.ListIndex = 0
    If SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        list4.Visible = False
    End If
    MS.Rows = 1
    M_FLX_AMT.Visible = False
    Me.MousePointer = vbDefault
    M_CST_IDY.SetFocus
End Sub
Private Sub DISPMS()
MS.TextMatrix(0, 0) = "Prd-SKU ID"
MS.ColAlignment(0) = 1
MS.TextMatrix(0, 1) = "Bill Rate - B"
MS.TextMatrix(0, 2) = " SKU Dis Amt"
SCl MS, 2
MS.TextMatrix(0, 3) = "Product - SKU"
MS.ColWidth(3) = 3500
MS.TextMatrix(0, 4) = "MRP"
MS.ColWidth(4) = 1150
MS.ColAlignment(1) = 7
MS.ColAlignment(2) = 7
MS.ColAlignment(4) = 7
MS.ColWidth(0) = 1200
MS.ColWidth(1) = 1200
MS.ColWidth(2) = 1200
MS.ColWidth(5) = 0
MER_REF_QTY = "0.000"
MER_REF_AMT = "0.00"
End Sub
