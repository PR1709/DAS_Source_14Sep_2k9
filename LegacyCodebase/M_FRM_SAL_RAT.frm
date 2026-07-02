VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SAL_RAT 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Sales (Rate Customisation)"
   ClientHeight    =   5625
   ClientLeft      =   1425
   ClientTop       =   1845
   ClientWidth     =   8820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "M_FRM_SAL_RAT.frx":0000
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
      TabIndex        =   45
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
         NumButtons      =   31
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
            Enabled         =   0   'False
            Key             =   "pre"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "crt"
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
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "smn1"
            Object.ToolTipText     =   "Freeze the Dealers for a Salesman"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAY"
            Object.ToolTipText     =   "Market Visit Day Filter for Today"
            ImageIndex      =   13
            Style           =   1
         EndProperty
      EndProperty
      Begin VB.ListBox text1 
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
         ItemData        =   "M_FRM_SAL_RAT.frx":0442
         Left            =   6960
         List            =   "M_FRM_SAL_RAT.frx":045B
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   50
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   340
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483637
      ForeColor       =   8388608
      TabCaption(0)   =   "&Invoice Details"
      TabPicture(0)   =   "M_FRM_SAL_RAT.frx":049F
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ImageList1"
      Tab(0).Control(1)=   "M_CST_IDY"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "M_SMN_IDY"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "M_ACT_IDY"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "M_DCH_IDY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "M_DCH_DAT"
      Tab(0).Control(6)=   "M_INV_IDY"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "M_INV_DAT"
      Tab(0).Control(8)=   "M_DBT_MCP"
      Tab(0).Control(9)=   "DBT_MCP"
      Tab(0).Control(10)=   "Shape2"
      Tab(0).Control(11)=   "Shape1"
      Tab(0).Control(12)=   "Line2(0)"
      Tab(0).Control(13)=   "Line1(0)"
      Tab(0).Control(14)=   "M_ACT_NME"
      Tab(0).Control(15)=   "M_SMN_NME"
      Tab(0).Control(16)=   "M_CST_AD4"
      Tab(0).Control(17)=   "M_CST_AD3"
      Tab(0).Control(18)=   "M_CST_AD2"
      Tab(0).Control(19)=   "M_CST_AD1"
      Tab(0).Control(20)=   "M_CST_NME"
      Tab(0).Control(21)=   "M_PRV_NUM"
      Tab(0).Control(22)=   "Label16"
      Tab(0).Control(23)=   "Label21"
      Tab(0).Control(24)=   "Label1"
      Tab(0).Control(25)=   "Label2"
      Tab(0).Control(26)=   "Label3"
      Tab(0).Control(27)=   "Label4"
      Tab(0).Control(28)=   "Label6"
      Tab(0).Control(29)=   "Label12"
      Tab(0).Control(30)=   "Label14"
      Tab(0).Control(31)=   "Label15"
      Tab(0).Control(32)=   "M_DBT_MCL"
      Tab(0).Control(33)=   "Label19"
      Tab(0).Control(34)=   "Label20"
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "&Products"
      TabPicture(1)   =   "M_FRM_SAL_RAT.frx":04BB
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "M_PRD_AMT"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "units"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "tax"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "M_Prd_Nme"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "MS"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "M_FLX_AMT"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "List6"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "&Taxes and  Levies"
      TabPicture(2)   =   "M_FRM_SAL_RAT.frx":04D7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "M_FRT_TYP"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "M_PKG_CST"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "M_FRT_AMT"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "M_TAX_DED"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "M_SPD_DED"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "M_CSD_DED"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "M_ROU_OFF"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "M_CSD_AMT"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "M_SPD_AMT"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "M_TAX_AMT"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "M_ROU_IDY"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "M_EXD_AMT"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "M_TAX_PCG"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "M_SPD_PCG"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "M_CSD_PCG"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "M_CRT_DAY"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label5"
      Tab(2).Control(17)=   "M_BEF_ROF"
      Tab(2).Control(18)=   "Label61"
      Tab(2).Control(19)=   "Label50"
      Tab(2).Control(20)=   "Label48"
      Tab(2).Control(21)=   "Label47"
      Tab(2).Control(22)=   "Label46"
      Tab(2).Control(23)=   "Label35"
      Tab(2).Control(24)=   "Label31"
      Tab(2).Control(25)=   "Label38"
      Tab(2).Control(26)=   "Label37"
      Tab(2).Control(27)=   "Label36"
      Tab(2).Control(28)=   "Label34"
      Tab(2).Control(29)=   "Label33"
      Tab(2).Control(30)=   "Label32"
      Tab(2).Control(31)=   "Label30"
      Tab(2).Control(32)=   "Label29"
      Tab(2).Control(33)=   "Label28"
      Tab(2).Control(34)=   "Label27"
      Tab(2).Control(35)=   "Label42"
      Tab(2).Control(36)=   "M_NET_AMT"
      Tab(2).Control(37)=   "Label44"
      Tab(2).Control(38)=   "Label24"
      Tab(2).Control(39)=   "M_GRS_AMT"
      Tab(2).Control(40)=   "Label22"
      Tab(2).Control(41)=   "Label26"
      Tab(2).Control(42)=   "Label25"
      Tab(2).ControlCount=   43
      TabCaption(3)   =   "&Shipment Details"
      TabPicture(3)   =   "M_FRM_SAL_RAT.frx":04F3
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label51"
      Tab(3).Control(1)=   "Label53"
      Tab(3).Control(2)=   "Label54"
      Tab(3).Control(3)=   "Label55"
      Tab(3).Control(4)=   "Label56"
      Tab(3).Control(5)=   "Label57"
      Tab(3).Control(6)=   "Label58"
      Tab(3).Control(7)=   "Label59"
      Tab(3).Control(8)=   "M_LRN_DAT"
      Tab(3).Control(9)=   "M_WBL_DAT"
      Tab(3).Control(10)=   "M_LRN_IDY"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "M_TTL_CAS"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "M_TXN_RMK"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "M_WBL_IDY"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "M_TXN_TPT"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "M_LRN_AMT"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "View"
      TabPicture(4)   =   "M_FRM_SAL_RAT.frx":050F
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "V_LST_CAP"
      Tab(4).Control(1)=   "list3"
      Tab(4).Control(2)=   "list1"
      Tab(4).Control(3)=   "list2"
      Tab(4).Control(4)=   "list5"
      Tab(4).Control(5)=   "list4"
      Tab(4).Control(6)=   "Frame1"
      Tab(4).ControlCount=   7
      Begin VB.TextBox M_FRT_TYP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox M_LRN_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1125
         Width           =   1095
      End
      Begin VB.TextBox M_TXN_TPT 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72600
         MaxLength       =   30
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox M_WBL_IDY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox M_TXN_RMK 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72600
         MaxLength       =   40
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   3000
         Width           =   3615
      End
      Begin VB.TextBox M_TTL_CAS 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72600
         MaxLength       =   5
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox M_LRN_IDY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox M_PKG_CST 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   9
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox M_FRT_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -70080
         MaxLength       =   9
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   3060
         Width           =   855
      End
      Begin VB.TextBox M_TAX_DED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68640
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox M_SPD_DED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68640
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox M_CSD_DED 
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
         Height          =   330
         Left            =   -68640
         MaxLength       =   9
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74760
         TabIndex        =   69
         Top             =   4560
         Width           =   8295
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Place"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.TextBox M_ROU_OFF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -70080
         MaxLength       =   5
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4320
         Width           =   615
      End
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
         ItemData        =   "M_FRM_SAL_RAT.frx":052B
         Left            =   2040
         List            =   "M_FRM_SAL_RAT.frx":0538
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox M_CSD_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox M_SPD_AMT 
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
         Height          =   330
         Left            =   -70080
         MaxLength       =   9
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox M_TAX_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -70080
         MaxLength       =   9
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2160
         Width           =   855
      End
      Begin VB.ListBox M_ROU_IDY 
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
         ItemData        =   "M_FRM_SAL_RAT.frx":0545
         Left            =   -70080
         List            =   "M_FRM_SAL_RAT.frx":05B5
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox M_EXD_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   9
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2610
         Width           =   855
      End
      Begin VB.TextBox M_TAX_PCG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox M_SPD_PCG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox M_CSD_PCG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox M_CRT_DAY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -75000
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
               Picture         =   "M_FRM_SAL_RAT.frx":0625
               Key             =   "p"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":0B69
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":0F31
               Key             =   "c"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":1285
               Key             =   "e"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":18BD
               Key             =   "s"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":1C11
               Key             =   "h"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":2071
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":2365
               Key             =   "pr"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":2771
               Key             =   "d"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":287D
               Key             =   "pre"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":2BD1
               Key             =   "brw"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":2FE5
               Key             =   "cpy"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_FRM_SAL_RAT.frx":3531
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox M_CST_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox M_SMN_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4012
         Width           =   1095
      End
      Begin VB.TextBox M_ACT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4402
         Width           =   1095
      End
      Begin VB.TextBox M_DCH_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker M_DCH_DAT 
         Height          =   330
         Left            =   -70080
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   37097
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   -70080
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   37099
      End
      Begin MSMask.MaskEdBox M_FLX_AMT 
         Height          =   375
         Left            =   4800
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
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
         Height          =   4395
         Left            =   80
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   360
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   13
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   72
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ProductName"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Product Type"
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
            Text            =   "Product ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Sub Units"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Institutional Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Billing Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "M.R.P Rate"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list5 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   74
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
            Text            =   "Dealer Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   75
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
            Text            =   "Account Name"
            Object.Width           =   2540
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
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   76
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
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Salesman  ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView list3 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   77
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dealer Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Dealer ID"
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
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Salesman Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker M_WBL_DAT 
         Height          =   330
         Left            =   -69480
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   37097
      End
      Begin MSComCtl2.DTPicker M_LRN_DAT 
         Height          =   330
         Left            =   -69480
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   37097
      End
      Begin VB.Label M_Prd_Nme 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4560
         TabIndex        =   109
         Top             =   4800
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   -72480
         TabIndex        =   108
         Top             =   3128
         Width           =   120
      End
      Begin VB.Label M_BEF_ROF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -71310
         TabIndex        =   104
         Top             =   4395
         Width           =   45
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -69630
         TabIndex        =   103
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LR Amount"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73665
         TabIndex        =   102
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LR Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -70260
         TabIndex        =   100
         Top             =   788
         Width           =   600
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WB Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -70320
         TabIndex        =   98
         Top             =   1628
         Width           =   660
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73545
         TabIndex        =   96
         Top             =   2588
         Width           =   675
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WB Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73740
         TabIndex        =   94
         Top             =   1628
         Width           =   870
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73500
         TabIndex        =   92
         Top             =   3068
         Width           =   630
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cases"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73305
         TabIndex        =   90
         Top             =   2108
         Width           =   435
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LR Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73680
         TabIndex        =   88
         Top             =   788
         Width           =   810
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing  Amount"
         Height          =   195
         Left            =   -73995
         TabIndex        =   86
         Top             =   3548
         Width           =   1215
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
         Height          =   255
         Left            =   -72480
         TabIndex        =   85
         Top             =   3555
         Width           =   255
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   -70845
         TabIndex        =   83
         Top             =   3128
         Width           =   540
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Surcharge on Tax"
         Height          =   195
         Left            =   -74055
         TabIndex        =   81
         Top             =   3128
         Width           =   1275
      End
      Begin VB.Label V_LST_CAP 
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
         Left            =   -74640
         TabIndex        =   73
         Top             =   480
         Width           =   8295
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
         Height          =   195
         Left            =   -70740
         TabIndex        =   62
         Top             =   2678
         Width           =   435
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
         Height          =   255
         Left            =   -72480
         TabIndex        =   61
         Top             =   2685
         Width           =   255
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less (-)"
         Height          =   195
         Left            =   -73290
         TabIndex        =   60
         Top             =   2685
         Width           =   510
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tax"
         Height          =   195
         Left            =   -73050
         TabIndex        =   59
         Top             =   2228
         Width           =   270
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Additonal Discount"
         Height          =   195
         Left            =   -74115
         TabIndex        =   58
         Top             =   1748
         Width           =   1335
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
         Height          =   255
         Left            =   -70560
         TabIndex        =   57
         Top             =   2235
         Width           =   255
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
         Height          =   255
         Left            =   -70560
         TabIndex        =   56
         Top             =   1755
         Width           =   255
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
         Height          =   255
         Left            =   -70560
         TabIndex        =   55
         Top             =   1275
         Width           =   255
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   -72480
         TabIndex        =   54
         Top             =   2205
         Width           =   120
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   -72480
         TabIndex        =   53
         Top             =   1725
         Width           =   120
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   -72480
         TabIndex        =   52
         Top             =   1245
         Width           =   120
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Bill Discount"
         Height          =   195
         Left            =   -73650
         TabIndex        =   51
         Top             =   1268
         Width           =   870
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3450
         TabIndex        =   44
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1485
         TabIndex        =   43
         Top             =   4485
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label M_DBT_MCP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -72360
         TabIndex        =   42
         Top             =   2715
         Width           =   45
      End
      Begin VB.Label DBT_MCP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit Period"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73800
         TabIndex        =   41
         Top             =   2715
         Width           =   1260
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1935
         Left            =   -70065
         Top             =   1935
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   -70080
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -70080
         X2              =   -66480
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -70080
         X2              =   -66480
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Label M_ACT_NME 
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
         Left            =   -69960
         TabIndex        =   40
         Top             =   4350
         Width           =   45
      End
      Begin VB.Label M_SMN_NME 
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
         Left            =   -69960
         TabIndex        =   39
         Top             =   3960
         Width           =   45
      End
      Begin VB.Label M_CST_AD4 
         BackColor       =   &H00800000&
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
         Height          =   330
         Left            =   -69960
         TabIndex        =   38
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label M_CST_AD3 
         BackColor       =   &H00800000&
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
         Height          =   330
         Left            =   -69960
         TabIndex        =   37
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label M_CST_AD2 
         BackColor       =   &H00800000&
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
         Height          =   330
         Left            =   -69960
         TabIndex        =   36
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label M_CST_AD1 
         BackColor       =   &H00800000&
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
         Height          =   330
         Left            =   -69960
         TabIndex        =   35
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label M_CST_NME 
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
         Height          =   330
         Left            =   -69960
         TabIndex        =   34
         Top             =   1950
         Width           =   3375
      End
      Begin VB.Label M_PRD_AMT 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   8475
         TabIndex        =   33
         Top             =   4785
         Width           =   45
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R/Off"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -70710
         TabIndex        =   29
         Top             =   4388
         Width           =   405
      End
      Begin VB.Label M_NET_AMT 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -67380
         TabIndex        =   28
         Top             =   4395
         Width           =   45
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -68835
         TabIndex        =   27
         Top             =   4388
         Width           =   840
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cr Days"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -73335
         TabIndex        =   26
         Top             =   548
         Width           =   555
      End
      Begin VB.Label M_GRS_AMT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Left            =   -67440
         TabIndex        =   25
         Top             =   840
         Width           =   75
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Before R/Off"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -73695
         TabIndex        =   24
         Top             =   4388
         Width           =   915
      End
      Begin VB.Label Label26 
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
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   -68640
         TabIndex        =   23
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due On"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   -69780
         TabIndex        =   22
         Top             =   480
         Width           =   555
      End
      Begin VB.Label M_PRV_NUM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   -72630
         TabIndex        =   21
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73440
         TabIndex        =   20
         Top             =   4080
         Width           =   900
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit (Sales)"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73470
         TabIndex        =   19
         Top             =   4470
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73665
         TabIndex        =   16
         Top             =   668
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73365
         TabIndex        =   15
         Top             =   1508
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -70680
         TabIndex        =   14
         Top             =   668
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -70680
         TabIndex        =   13
         Top             =   1508
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73110
         TabIndex        =   12
         Top             =   1988
         Width           =   570
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Credit Limit Amount"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73935
         TabIndex        =   11
         Top             =   2355
         Width           =   1395
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Outstanding "
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -74070
         TabIndex        =   10
         Top             =   3075
         Width           =   1530
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit Available"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73995
         TabIndex        =   9
         Top             =   3435
         Width           =   1455
      End
      Begin VB.Label M_DBT_MCL 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -72360
         TabIndex        =   8
         Top             =   2355
         Width           =   45
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -72360
         TabIndex        =   7
         Top             =   3075
         Width           =   45
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -72360
         TabIndex        =   6
         Top             =   3435
         Width           =   45
      End
   End
End
Attribute VB_Name = "M_FRM_SAL_RAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, idnew As Boolean, ls As ListItem, M_DBT_RAP, MSROW, INVIDY, INC, dlr_idy, taxamt As Double

Private Sub CREATE()
MOD_REC = False
M_INV_IDY = ""
saverecord
End Sub

Public Sub GRS_CAL()
On Error Resume Next
M_GRS_AMT = convert(amt(Val(M_PRD_AMT)))
M_CSD_AMT = convert(amt(Round(M_PRD_AMT * Val(M_CSD_PCG) / 100, 2)))
M_CSD_DED = convert(amt(Round((M_GRS_AMT - M_CSD_AMT), 2)))
M_SPD_AMT = convert(amt(Round(M_CSD_DED * Val(M_SPD_PCG) / 100, 2)))
M_SPD_DED = convert(amt(Round(Val(M_CSD_DED) - Val(M_SPD_AMT), 2)))
M_TAX_PCG = discount(M_TAX_PCG)
M_TAX_AMT = convert(amt(Round((Val(M_SPD_DED) * Val(M_TAX_PCG)) / 100, 2)))
M_TAX_DED = convert(amt(Val(M_SPD_DED) + Val(M_TAX_AMT)))
M_NET_AMT = convert(amt(Round(Val(M_TAX_DED) - Val(M_EXD_AMT), 2)))
M_FRT_AMT = convert(amt(Round((Val(taxamt) + Val(M_TAX_AMT)) * Val(M_FRT_TYP) / 100, 2)))
M_NET_AMT = Val(M_FRT_AMT) + Val(M_NET_AMT)
M_NET_AMT = convert(amt(Val(M_NET_AMT) + Val(M_PKG_CST)))
M_BEF_ROF = convert(amt(Val(M_NET_AMT) & ""))
ROU_OFF
End Sub

Private Sub saverecord()
On Error Resume Next
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "Effmcg") = vbNo Then
    Cancelrecord
    Exit Sub
End If

If check("ACT", "(ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') AND ACT_idy", Trim(M_CST_IDY)) = False Then
    MsgBox "Dealer Not Found ,Define Dealers !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    M_CST_IDY.SetFocus
    Exit Sub
End If

If check("smn", "smn_idy", Trim(M_SMN_IDY)) = False Then
    MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
    M_SMN_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If

If check("act", "(grp_clf like 'Sale%' and ACT_TYP LIKE '') AND act_idy", Trim(M_ACT_IDY)) = False Then
    MsgBox "Sales Account Not Found, Define Sales Account...!", vbExclamation, "Effmcg"
    M_ACT_IDY.SetFocus
    SSTab1.Tab = 0
    Exit Sub
End If
If Val(M_NET_AMT) < 0 Then
    MsgBox "Can't Save this bill, NetAmount is Invalid...!", vbExclamation, "Effmcg"
    SSTab1.Tab = 2
    M_EXD_AMT.SetFocus
    Exit Sub
End If
GRS_CAL
If MOD_REC = True Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Sincrement(M_INV_IDY) & "' and txn_typ like 'SAL'", con, adOpenKeyset
    For A = 1 To REC_SET.RecordCount
        con.Execute "update prd where prd_idy like '" & REC_SET("prd_idy") & "' set cur_lev = " & ADDQTY(Val(REC_SET("prd_qty")), REC_SET("prd_idy"))
        REC_SET.MoveNext
    Next
    con.Execute "delete from inv where inv_idy like '" & Sincrement(M_INV_IDY) & "' and txn_typ like 'SAL'"
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inl where inv_idy like '" & Sincrement(M_INV_IDY) & "' and txn_typ like 'SAL'", con, adOpenKeyset, adLockOptimistic
    INC = Sincrement(M_INV_IDY)
Else
    REC_SET.Open "select * from inl", con, adOpenKeyset, adLockOptimistic
    REC_SET.AddNew
    REC_SET("inv_idy") = "SAL" & T7increment("IDY_SAL")
    INC = REC_SET("inv_idy")
    INVIDY = INC
End If
REC_SET("inv_dat") = M_INV_DAT
REC_SET("ord_idy") = "1"
REC_SET("cst_idy") = Trim(M_CST_IDY & "")
REC_SET("act_idy") = Trim(M_ACT_IDY & "")
REC_SET("smn_idy") = Trim(M_SMN_IDY & "")
REC_SET("dch_idy") = M_DCH_IDY & ""
REC_SET("dch_dat") = M_DCH_DAT
REC_SET("csd_pcg") = discount(M_CSD_PCG)
REC_SET("csd_amt") = convert(amt(M_CSD_AMT))
REC_SET("spd_pcg") = discount(M_SPD_PCG)
REC_SET("spd_amt") = convert(amt(M_SPD_AMT))
REC_SET("tax_pcg") = discount(M_TAX_PCG)
REC_SET("tax_amt") = convert(amt(M_TAX_AMT))
REC_SET("pkg_cst") = convert(amt(M_PKG_CST))
REC_SET("exd_pcg") = discount(M_FRT_TYP)
REC_SET("FRT_AMT") = Val(M_FRT_AMT & "")
REC_SET("LRN_IDY") = Trim(M_LRN_IDY)
REC_SET("LRN_AMT") = Val(M_LRN_AMT & "")
REC_SET("LRN_DAT") = M_LRN_DAT
REC_SET("WBL_IDY") = Trim(M_WBL_IDY & "")
REC_SET("WBL_DAT") = M_WBL_DAT
REC_SET("TXN_RMK") = Trim(M_TXN_RMK)
REC_SET("TTL_CAS") = Val(M_TTL_CAS)
REC_SET("EXD_AMT") = Val(M_EXD_AMT)
REC_SET("txn_typ") = "SAL"
REC_SET("grs_amt") = convert(amt(Val(M_GRS_AMT)))
REC_SET("net_amt") = Val(M_NET_AMT)
REC_SET("crt_day") = Val(M_CRT_DAY)
REC_SET("due_dat") = Date
REC_SET("rou_off") = Val(M_ROU_OFF)
REC_SET("ROU_IDY") = Trim(M_ROU_IDY)
REC_SET("TXN_TPT") = Trim(M_TXN_TPT)
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
            REC_SET("rat_app") = MS.TextMatrix(A, 6) & ""
        Else
            REC_SET("rat_app") = "A"
        End If
        REC_SET("add_dis") = discount(MS.TextMatrix(A, 7))
        REC_SET("trd_dis") = discount(MS.TextMatrix(A, 8))
        REC_SET("trd_tax") = discount(MS.TextMatrix(A, 9))
        If REC_SET1.State = 1 Then REC_SET1.Close
        REC_SET1.Open "SELECT SUB_UNT FROM PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(A, 0)) & "'", con, adOpenKeyset, adLockOptimistic
        PRDNET = RATCAL(Val(MS.TextMatrix(A, 3) & ""), Val(MS.TextMatrix(A, 4) & ""), REC_SET1(0))
        REC_SET("ADD_AMT") = (PRDNET * Val(MS.TextMatrix(A, 7)) / 100)
        REC_SET("DIS_AMT") = ((PRDNET - REC_SET("ADD_AMT")) * Val(MS.TextMatrix(A, 8)) / 100)
        TRDTAX = (PRDNET - (REC_SET("ADD_AMT") + REC_SET("DIS_AMT"))) * Val(MS.TextMatrix(A, 9)) / 100
        REC_SET("TAX_AMT") = TRDTAX
        REC_SET("txn_typ") = "SAL"
        REC_SET("COL_AMT") = convert(amt(MS.TextMatrix(A, 10)))
        REC_SET.Update
        con.Execute "update prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "' set cur_lev = " & CONQTY(Trim(MS.TextMatrix(A, 3)), Trim(MS.TextMatrix(A, 0)))
    End If
Next
If MOD_REC = False Then MsgBox "New Credit Sales ID is " & INC, vbExclamation, "Effmcg"
M_PRV_NUM = labdisp("IDY_SAL")
If check("PRT", "PRT_NME", "M_FRM_SAL_RAT") = True Then
    con.Execute "update dir where dir_idy like 'INV' AND DIR_IDX LIKE 'NUM' set dir_als = '" & INC & "'"
    con.Execute "update dir where dir_idy like 'DOC' AND DIR_IDX LIKE 'TYP' set dir_als = 'SAL'"
    OpenDocument CURDIR & "\VOUCH.RP1", "Invoice", 2, Trim(INC), Trim(INC), 2, CURDIR & "\inv.cdx,c,inv_idy"
    Cancelrecord
End If
Cancelrecord
End Sub

Private Sub Cancelrecord()
TXT1 = text1
Txt2 = Text2
clrctr M_FRM_SAL_RAT
text1 = TXT1
Text2 = Txt2
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

Private Sub V_DLR_LST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close

If Check1.Value = 1 Then
    If Toolbar1.Buttons(31).Value = tbrPressed And Toolbar1.Buttons(30).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND MAR_VDA LIKE '" & UCase(text1) & "' AND SMN.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    ElseIf Toolbar1.Buttons(31).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_NME like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND MAR_VDA LIKE '" & UCase(text1) & "' AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    ElseIf Toolbar1.Buttons(30).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_NME like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND SMN.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    Else
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and act_nme like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    End If
ElseIf Check2.Value = 1 Then
    If Toolbar1.Buttons(31).Value = tbrPressed And Toolbar1.Buttons(30).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND MAR_VDA LIKE '" & UCase(text1) & "' AND SMN.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    ElseIf Toolbar1.Buttons(31).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND MAR.MAR_VDA LIKE '" & UCase(text1) & "' AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    ElseIf Toolbar1.Buttons(30).Value = 1 Then
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR act.ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND SMN.SMN_IDY LIKE '" & Trim(Text2) & "' AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    Else
        REC_SET.Open "select ACT.*,MAR.MAR_NME,act.ROU_IDY,smn.smn_nme from act,MAR,smn where (act.ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') and act_ad4 like '" & Trim(M_CST_IDY) & "%' AND ACT.MAR_IDY LIKE MAR.MAR_IDY AND ACT.SMN_IDY LIKE SMN.SMN_IDY AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    End If
End If
list3.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list3.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("MAR_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ROU_IDY") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("SMN_NME") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    V_LST_CAP = "Select Dealer"
    list3.Visible = True
    list3.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Dealer Not Found, Define Dealer !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        M_CST_IDY.SetFocus
    End If
End If
End Sub
Private Sub V_SMN_LST()
Dim REC_SET As New ADODB.Recordset
V_LST_CAP = "Select Salesman"
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    If Toolbar1.Buttons(30).Value = tbrPressed Then
        REC_SET.Open "select * from smn", con, adOpenKeyset, adLockOptimistic
    Else
        REC_SET.Open "select * from SMN WHERE SMN_nme like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockOptimistic
    End If
ElseIf Check2.Value = 1 Then
    If Toolbar1.Buttons(30).Value = tbrPressed Then
        REC_SET.Open "select * from smn", con, adOpenKeyset, adLockOptimistic
    Else
        REC_SET.Open "select * from SMN WHERE SMN_AD4 like '" & Trim(M_SMN_IDY) & "%'", con, adOpenKeyset, adLockOptimistic
    End If
End If
List1.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = List1.ListItems.Add(, , Trim(REC_SET("smn_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("smn_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    List1.Visible = True
    List1.SetFocus
    
Else
    If SSTab1.Tab = 0 Then
    MsgBox "Salesmen Not Found, Define Salesmen !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    M_SMN_IDY.SetFocus
    End If
End If
End Sub

Private Sub V_ACT_LST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
    REC_SET.Open "select * from act where grp_clf like 'Sales' and ACT_NME LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockOptimistic
Else
    REC_SET.Open "select * from act where grp_clf like 'Sales' and ACT_AD4 LIKE '" & Trim(M_ACT_IDY) & "%' and ACT_TYP LIKE '' ", con, adOpenKeyset, adLockOptimistic
End If
list2.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list2.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("act_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    V_LST_CAP = "Select Account"
    list2.Visible = True
    list2.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Sale Account Not Found, Define Sale Account...!", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    M_ACT_IDY.SetFocus
End If
End If
End Sub

Private Sub EXT_REC()
Unload Me
End Sub

Private Sub M_FRT_TYP_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Surcharge Percent"
End Sub

Private Sub M_LRN_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Lorry Receipt Date or Press F4 to get View"
End Sub


Private Sub M_ROU_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Route ID or Select using Arrow Keys"
End Sub

Private Sub M_SPD_AMT_Change()
If Len(Trim(M_SPD_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_SPD_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Additional Discount Amount"
End Sub

Private Sub M_SPD_PCG_Change()
If Len(Trim(M_SPD_PCG)) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_SPD_PCG_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Additional Discount Percentage"
End Sub

Private Sub M_CSD_AMT_Change()
If Len(Trim(M_CSD_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_CSD_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Bill Discount Amount"
End Sub

Private Sub M_CSD_PCG_Change()
If Len(Trim(M_CSD_PCG)) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_CSD_PCG_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Bill Discount Percentage"
End Sub

Private Sub M_CSD_PCG_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub

Private Sub M_CSD_PCG_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Or k = 38 Then
    M_CSD_PCG = discount(M_CSD_PCG)
    GRS_CAL
    If k = 38 Then
        M_CRT_DAY.SetFocus
    Else
        M_CSD_AMT.SetFocus
    End If
End If
End Sub

Private Sub M_CSD_AMT_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub

Private Sub M_CSD_AMT_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
    M_CSD_AMT = convert(amt(M_CSD_AMT))
    GRS_CAL
    M_SPD_PCG.SetFocus
ElseIf k = 38 Then
    M_CSD_AMT = convert(amt(M_CSD_AMT))
    GRS_CAL
    M_CSD_PCG.SetFocus
End If
End Sub
Private Sub M_SPD_PCG_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub
Private Sub M_SPD_PCG_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Or k = 38 Then
    M_SPD_PCG = discount(M_SPD_PCG)
    GRS_CAL
    If k = 13 Or k = 40 Then
        M_SPD_AMT.SetFocus
    Else
        M_CSD_AMT.SetFocus
    End If
End If
End Sub
Private Sub M_SPD_AMT_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub
Private Sub M_SPD_AMT_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
    M_SPD_AMT = convert(amt(M_SPD_AMT))
    GRS_CAL
    M_TAX_PCG.SetFocus
ElseIf k = 38 Then
    M_SPD_AMT = convert(amt(M_SPD_AMT))
    GRS_CAL
    M_SPD_PCG.SetFocus
End If
End Sub


Private Sub M_TTL_CAS_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Cases"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TTL_CAS_KeyDown(k As Integer, Shift As Integer)
If k = 13 Or k = 40 Then
    M_TXN_TPT.SetFocus
ElseIf k = 38 Then
    M_WBL_DAT.SetFocus
End If
End Sub

Private Sub M_TTL_CAS_KeYPRESS(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub

Private Sub M_EXD_AMT_Change()
If Len(Trim(M_EXD_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_EXD_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Damage Amount"
End Sub

Private Sub M_EXD_AMT_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub
Private Sub M_EXD_AMT_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Or k = 38 Then
    M_EXD_AMT = convert(amt(M_EXD_AMT))
    GRS_CAL
    If k = 38 Then
        M_TAX_AMT.SetFocus
    Else
        M_ROU_IDY.SetFocus
    End If
End If
End Sub

Private Sub M_FRT_AMT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Surcharge Amount"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_FRT_AMT_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
    M_FRT_AMT = convert(amt(M_FRT_AMT))
    P_DIS_CAL
    M_PKG_CST.SetFocus
ElseIf k = 38 Then
    M_FRT_AMT = convert(amt(M_FRT_AMT))
    P_DIS_CAL
    M_FRT_TYP.SetFocus
End If
End Sub

Private Sub M_FRT_AMT_KeyPress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub

Private Sub M_FRT_TYP_keydown(k As Integer, s As Integer)
If k = 13 Then
    GRS_CAL
    M_FRT_TYP = discount(M_FRT_TYP)
    M_FRT_AMT.SetFocus
ElseIf k = 38 Then
    GRS_CAL
    M_FRT_TYP = discount(M_FRT_TYP)
    M_ROU_IDY.SetFocus
End If
End Sub

Private Sub M_ROU_OFF_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter 0 to Retain Paise Values"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_ROU_OFF_KeyDown(k As Integer, Shift As Integer)
If k = 13 Or k = 40 Then
If Val(M_ROU_OFF) = 0 Then
    M_NET_AMT = M_BEF_ROF
    M_ROU_OFF = "0.00"
Else
    GRS_CAL
End If
    SSTab1.Tab = 3
    M_LRN_IDY.SetFocus
ElseIf k = 38 Then
    M_PKG_CST.SetFocus
End If
End Sub

Private Sub M_ROU_OFF_KeyPress(k As Integer)
If k > 47 And k < 58 Or k = 46 Or k = 8 Then
Else
k = 0
End If
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
            MS.TextMatrix(MS.Row, 4) = amt(REC_SET("pur_rat") & "")
        ElseIf List6 = "B" Then
            MS.TextMatrix(MS.Row, 4) = amt(REC_SET("sal_rat") & "")
        ElseIf List6 = "C" Then
            MS.TextMatrix(MS.Row, 4) = amt(REC_SET("mrp_rat") & "")
        End If
    End If
    P_DIS_CAL
    MS.Col = 10
    MS.SetFocus
End If
End Sub

Private Sub M_ROU_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    GRS_CAL
    M_FRT_TYP.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    M_EXD_AMT.SetFocus
End If
End Sub

Private Sub M_LRN_AMT_GotFocus()
SendKeys "{HOME}+{END}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Lorry Receipt Amount"
End Sub

Private Sub M_LRN_AMT_keypress(k As Integer)
If k > 47 And k < 58 Or k = 46 Or k = 8 Then
Else
k = 0
End If
End Sub
Private Sub M_LRN_AMT_Keydown(k As Integer, s As Integer)
If k = 38 And s = 2 Then
M_LRN_DAT.SetFocus
ElseIf k = 13 Then
M_WBL_IDY.SetFocus
End If
End Sub

Private Sub M_LRN_DAT_Keydown(k As Integer, s As Integer)
If k = 38 And s = 1 Then
M_LRN_IDY.SetFocus
ElseIf k = 13 Then
M_LRN_AMT.SetFocus
End If
End Sub

Private Sub M_LRN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Lorry Receipt Number"
SendKeys "{home}+{End}"
End Sub

Private Sub M_LRN_IDY_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
M_LRN_DAT.SetFocus
ElseIf KeyCode = 38 Then
SSTab1.Tab = 2
M_ROU_OFF.SetFocus
End If
End Sub

Private Sub M_LRN_IDY_KeyPress(k As Integer)
If k > 96 And k < 123 Then k = k - 32
End Sub

Private Sub M_PKG_CST_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Packing or Delivery Charges"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_PKG_CST_KeyDown(k As Integer, Shift As Integer)
If k = 13 Or k = 40 Then
    M_PKG_CST = convert(amt(M_PKG_CST))
    GRS_CAL
    M_ROU_OFF.SetFocus
ElseIf k = 38 Then
    M_PKG_CST = convert(amt(M_PKG_CST))
    GRS_CAL
    M_FRT_AMT.SetFocus
End If
End Sub

Private Sub M_PKG_CST_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub

Private Sub M_TXN_RMK_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Remarks"
SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TAX_AMT_Change()
If Len(Trim(M_TAX_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_TAX_AMT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Tax Amount"
End Sub

Private Sub M_TAX_PCG_Change()
If Len(Trim(M_TAX_PCG)) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_TAX_PCG_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Tax Percentage"
End Sub

Private Sub M_TAX_PCG_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub
Private Sub M_TAX_AMT_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
    M_TAX_AMT = convert(amt(M_TAX_AMT))
    GRS_CAL
    M_EXD_AMT.SetFocus
ElseIf k = 38 Then
    M_TAX_AMT = convert(amt(M_TAX_AMT))
    GRS_CAL
    M_TAX_PCG.SetFocus
End If
End Sub
Private Sub M_TAX_AMT_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub
Private Sub M_TAX_PCG_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Or k = 38 Then
    M_TAX_PCG = discount(M_TAX_PCG)
    GRS_CAL
    
    If k = 38 Then
        M_SPD_AMT.SetFocus
    Else
        M_TAX_AMT.SetFocus
    End If
End If
End Sub

Private Sub M_CRT_DAY_Change()
If Len(M_CRT_DAY) = 3 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_CRT_DAY_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Credit Days"
End Sub

Private Sub M_CRT_DAY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Val(M_CRT_DAY) > 365 Then
        MsgBox "Credit Days Of Bill Can't Go Beyond 365 Days !", vbExclamation, "Effmcg"
        M_CRT_DAY.SetFocus
    Else
        M_CSD_PCG.SetFocus
        Label26 = Format(DateAdd("d", Val(Trim(M_CRT_DAY)), M_INV_DAT.Value), "DD/MM/YYYY")
    End If
ElseIf KeyCode = 38 Then
    If Val(M_CRT_DAY) >= 365 Then
        MsgBox "Credit Days Of Bill Can't Go Beyond 365 Days !", vbExclamation, "Effmcg"
        M_CRT_DAY.SetFocus
    Else
        SSTab1.Tab = 0
        Label26 = Format(DateAdd("d", Val(Trim(M_CRT_DAY)), M_INV_DAT.Value), "DD/MM/YYYY")
        M_ACT_IDY.SetFocus
    End If
End If
End Sub

Private Sub M_CRT_DAY_Keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Then
Else
k = 0
End If
End Sub

Private Sub M_ACT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Account"
End Sub

Private Sub M_ACT_IDY_KeyPress(k As Integer)
If k = 39 Or k = 41 Then k = 0
If k > 96 And k < 123 Then k = k - 32
End Sub

Private Sub M_INV_DAT_CloseUp()
M_INV_DAT.SetFocus
End Sub

Private Sub M_INV_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date Using Arrow Keys Or Press F4"
End Sub

Private Sub M_INV_DAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then M_DCH_IDY.SetFocus
End Sub

Private Sub M_DCH_DAT_CloseUp()
M_DCH_DAT.SetFocus
End Sub

Private Sub M_DCH_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Dc Date Using Arrow Keys Or Press F4, To Move To Previous Fields Use Shift+UpArrow"
End Sub

Private Sub M_DCH_DAT_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
M_CST_IDY.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
M_DCH_IDY.SetFocus
End If
End Sub

Private Sub M_DCH_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Delivery Challan Number"
SendKeys "{home}+{End}"
End Sub

Private Sub M_DCH_IDY_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
M_DCH_DAT.SetFocus
ElseIf KeyCode = 38 Then
M_INV_DAT.SetFocus
End If
End Sub

Private Sub M_DCH_IDY_KeyPress(k As Integer)
If k > 96 And k < 123 Then k = k - 32
End Sub

Private Sub M_CST_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Dealer"
End Sub

Private Sub M_CST_IDY_keydown(k As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If k = 13 Or k = 40 Or k = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_CST_IDY)) & "' AND (ACT_TYP LIKE 'BOT' OR ACT_TYP LIKE 'SAL') AND DBT_TYP='Y'", con, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then
        M_CST_IDY = Sincrement(M_CST_IDY)
        If Len(Trim(M_CST_IDY)) = 0 Then Exit Sub
        M_CST_NME = Trim(REC_SET("act_nme") & "")
        M_CST_AD1 = Trim(REC_SET("act_ad1") & "")
        M_CST_AD2 = Trim(REC_SET("act_ad2") & "")
        M_CST_AD3 = Trim(REC_SET("act_ad3") & "")
        M_CST_AD4 = Trim(REC_SET("act_ad4") & "")
        M_DBT_MCL = amt(Trim(REC_SET("dbt_mcl")) & "")
        M_DBT_MCP = Trim(REC_SET("dbt_mcp") & "")
        If Len(Trim(REC_SET("ROU_IDY"))) <> 0 Then M_ROU_IDY = Trim(REC_SET("ROU_IDY") & "")
        If MOD_REC = False Then M_SMN_IDY = Trim(REC_SET("smn_idy") & "")
        If MOD_REC = False Then M_CSD_PCG = Val(Trim(REC_SET("dbt_dis")))
        If REC_SET1.State = 1 Then REC_SET1.Close
        REC_SET1.Open "select * from smn where smn_idy like '" & Trim(REC_SET("smn_idy")) & "'", con, adOpenKeyset, adLockOptimistic
        If REC_SET1.RecordCount > 0 Then M_SMN_NME = Trim(REC_SET1("smn_nme"))
        M_DBT_RAP = Trim(REC_SET("dbt_rap"))
        SSTab1.Tab = 0
        If k = 38 Then
            M_DCH_DAT.SetFocus
        Else
            M_SMN_IDY.SetFocus
        End If
    Else
        V_DLR_LST
    End If
End If
End Sub

Private Sub M_CST_IDY_KeyPress(k As Integer)
If k = 39 Or k = 41 Then k = 0
If k > 96 And k < 123 Then k = k - 32
End Sub
Private Sub Form_Activate()
checkdongle
SSTab1.Tab = 0
If M_INV_IDY.Enabled = True Then
    M_INV_IDY.SetFocus
Else
    M_INV_DAT.SetFocus
End If
M_PRV_NUM = labdisp("IDY_SAL")
SSTab1.TabEnabled(4) = False
setval
SETYEAR M_FRM_SAL_RAT
End Sub

Private Sub V_INV_LST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4 from act,inl where act_nme like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SAL' and ord_idy like '1'", con, adOpenKeyset, adLockOptimistic
Else
REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4 from act,inl where act_ad4 like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SAL' and ord_idy like '1'", con, adOpenKeyset, adLockOptimistic
End If
list5.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 4
    V_LST_CAP = "Select Invoice"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Invoice Number Not Found...!", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then Shell "c:\windows\calc.exe"
If List1.Visible = False And list2.Visible = False And list3.Visible = False And list4.Visible = False And list5.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then PRINTRECORD
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then EXT_REC
    If KeyCode = 81 And Shift = 2 Then Cancelrecord
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
    
    
        If M_FRM_SAL_RAT.ActiveControl.name = "M_INV_IDY" Then
            V_INV_LST
        ElseIf M_FRM_SAL_RAT.ActiveControl.name = "M_CST_IDY" Then
            V_DLR_LST
        ElseIf M_FRM_SAL_RAT.ActiveControl.name = "M_SMN_IDY" Then
            V_SMN_LST
        ElseIf M_FRM_SAL_RAT.ActiveControl.name = "M_ACT_IDY" Then
            V_ACT_LST
        ElseIf M_FRM_SAL_RAT.ActiveControl.name = "Text3" Then
            V_PRD_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_SAL_RAT.ActiveControl Is TextBox Then M_FRM_SAL_RAT.ActiveControl.Text = FUNKEY(KeyCode)
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
    list5.Visible = False
    
    M_INV_IDY = list5.SelectedItem
    dispdetails
End If
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If List1.Visible = True Then
        List1.Visible = False
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
        M_CST_IDY.SetFocus
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
            MS.Row = MS.Rows - 1
            MS.SetFocus
        End If
    ElseIf SSTab1.Tab = 1 Then
        SSTab1.Tab = 2
        Text3.Visible = False
        M_FLX_AMT.Visible = False
        GRS_CAL
        P_DIS_CAL
        M_CRT_DAY.SetFocus
    Else
        If M_FRM_SAL_RAT.ActiveControl.name = "M_INV_IDY" Then
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
MS.ColWidth(0) = 980
MS.ColWidth(1) = 0
MS.ColWidth(2) = 0
MS.ColWidth(3) = 1050
MS.ColWidth(4) = 1000
MS.ColWidth(5) = 900
MS.ColWidth(6) = 700
MS.ColWidth(7) = 700
MS.ColWidth(8) = 700
MS.ColWidth(9) = 700
MS.ColWidth(10) = 1500
MS.ColWidth(11) = 3000
MS.ColWidth(12) = 1600
MS.TextMatrix(0, 0) = "Product ID"
MS.TextMatrix(0, 1) = "        L Price"
MS.TextMatrix(0, 2) = "  LD%"
MS.TextMatrix(0, 3) = "            Units"
MS.TextMatrix(0, 4) = "            Rate"
MS.TextMatrix(0, 5) = "       UOM"
MS.TextMatrix(0, 6) = "Rate ID"
MS.TextMatrix(0, 7) = "  AD%"
MS.TextMatrix(0, 8) = "  CD%"
MS.TextMatrix(0, 9) = " Tax%"
MS.TextMatrix(0, 10) = "              Amount"
MS.TextMatrix(0, 11) = "Product Name"
MS.TextMatrix(0, 12) = "  Short Name"

MS.Col = 0
MS.Col = 3
units.Left = (MS.CellWidth + MS.CellLeft) - units.Width
units.Top = 4850
MS.Col = 5
tax.Left = (MS.CellWidth + MS.CellLeft) - tax.Width
tax.Top = 4850
tax.Width = 0
M_Prd_Nme.Top = units.Top
M_PRD_AMT.Top = units.Top
MS.Col = 0
M_INV_DAT = Date
M_LRN_DAT = Date
M_DCH_DAT = Date
waydate = Date

list5.ColumnHeaders(1).Width = list2.Width / 4
list5.ColumnHeaders(2).Width = list2.Width / 4
list5.ColumnHeaders(3).Width = list2.Width / 4
list5.ColumnHeaders(4).Width = list2.Width / 4

M_DBT_MCL = "0.00"
M_DBT_MCP = "0"
Label19 = "0.00"
Label20 = "0.00"
List1.ColumnHeaders(1).Width = 3800
list2.ColumnHeaders(1).Width = 3800
list3.ColumnHeaders(1).Width = 3800
M_PRV_NUM = labdisp("IDY_SAL")
M_LRN_AMT = "0.00"
deliverycharges = "0.00"
M_LRN_IDY = 0
M_FRT_TYP = "0.00"
waybill = 0
M_NET_AMT = "0.00"
M_GRS_AMT = "0.00"
M_CRT_DAY = 0
waybill = 0
M_SPD_DED = "0.00"
cd = "0.00"
units = "0"
M_PRD_AMT = "0.00"
tax = "0.00"
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Invoice Number"
End Sub

Private Sub M_INV_IDY_KeyPress(k As Integer)
If k = 39 Or k = 41 Then k = 0
If k > 96 And k < 123 Then k = k - 32
If M_INV_IDY = "" And k = 27 Then Unload Me
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If List1.SortOrder = lvwAscending Then
    List1.SortOrder = lvwDescending
ElseIf List1.SortOrder = lvwDescending Then
    List1.SortOrder = lvwAscending
End If
List1.SortKey = ColumnHeader.Index - 1
End Sub
Private Sub List1_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If List1.ListItems.Count > 0 Then
        If List1.Tag = "Salesman" Then
            Text2.Visible = True
            Text2 = List1.SelectedItem.ListSubItems.Item(1)
            List1.Visible = False
            SSTab1.Tab = 0
            List1.Tag = ""
        Else
            M_SMN_IDY = List1.SelectedItem.ListSubItems.Item(1)
            M_SMN_NME = List1.SelectedItem
            SSTab1.Tab = 0
            M_ACT_IDY.SetFocus
            List1.Visible = False
        End If
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
Private Sub list3_DblClick()
List3_KeyPress 13
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
List2_keypress 13
End Sub

Private Sub LIST5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If list5.ListItems.Count > 0 Then
    M_INV_IDY = list5.SelectedItem
    list5.Visible = False
    M_INV_IDY_KeyDown 13, 0
End If
End If
End Sub


Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
If V_LST_CAP.Caption = "Select Salesman" Then
    V_SMN_LST
ElseIf V_LST_CAP.Caption = "Select Dealer" Then
    V_DLR_LST
ElseIf V_LST_CAP.Caption = "Select Invoice" Then
    V_INV_LST
ElseIf V_LST_CAP.Caption = "Select Account" Then
    V_ACT_LST
ElseIf V_LST_CAP.Caption = "Select Product" Then
    V_PRD_LST
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If V_LST_CAP.Caption = "Select Salesman" Then
    V_SMN_LST
ElseIf V_LST_CAP.Caption = "Select Dealer" Then
    V_DLR_LST
ElseIf V_LST_CAP.Caption = "Select Invoice" Then
    V_INV_LST
ElseIf V_LST_CAP.Caption = "Select Account" Then
    V_ACT_LST
ElseIf V_LST_CAP.Caption = "Select Product" Then
    If Check2.Value = 1 Then
    con.Execute "insert into prt values('prdnum')"
    Else
    con.Execute "delete from prt where prt_nme like 'prdnum'"
    End If
    V_PRD_LST
End If
'End If
End Sub

Private Sub M_WBL_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Way Bill Date or Press F4 to get View"
End Sub

Private Sub MS_RowColChange()
M_Prd_Nme = MS.TextMatrix(MS.Row, 11)
End Sub
Private Sub MS_EnterCell()
If MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product ID View, Space Bar to Correct"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Product Units"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Rate"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Weight or Volume based on Conversion Factor"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Tax Percent"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Additional Discount Percent"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Cash Discount Percent"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Rate Slab (A-> Institutional Rate,B->Billing Rate,C-> M.R.P Rate)"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Amount"
ElseIf MS.Col = 11 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Name"
ElseIf MS.Col = 12 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
newrecord
ElseIf Button.KEY = "s" Then
saverecord
ElseIf Button.KEY = "e" Then
EXT_REC
ElseIf Button.KEY = "p" Then
PRINTRECORD
ElseIf Button.KEY = "c" Then
Cancelrecord
ElseIf Button.KEY = "pre" Then
M_FRM_PRE.Caption = "Prefix  Sales"
M_FRM_PRE.Show 1
ElseIf Button.KEY = "crt" Then
CREATE
ElseIf Button.KEY = "smn" Then
V_SMN_LST
ElseIf Button.KEY = "ed" Then
    If M_FRM_SAL_RAT.ActiveControl.name = "M_INV_IDY" Then
        V_LST_CAP = "Select Invoice"
        V_INV_LST
    ElseIf M_FRM_SAL_RAT.ActiveControl.name = "M_CST_IDY" Then
        V_LST_CAP = "Select Dealer"
        V_DLR_LST
    ElseIf M_FRM_SAL_RAT.ActiveControl.name = "M_SMN_IDY" Then
        V_LST_CAP = "Select Salesman"
        V_SMN_LST
    ElseIf M_FRM_SAL_RAT.ActiveControl.name = "M_ACT_IDY" Then
        V_LST_CAP = "Select Account"
        V_ACT_LST
    ElseIf M_FRM_SAL_RAT.ActiveControl.name = "Text3" Or (MS.Col = 0 And SSTab1.Tab = 1) Then
        V_PRD_LST
    End If
ElseIf Button.KEY = "DAY" Then
    If Button.Value = tbrPressed Then
        text1.Visible = True
        text1 = Format(Date, "dddd")
        text1.BackColor = vbBlue
        text1.ForeColor = vbWhite
    Else
        text1 = ""
        text1.Visible = False
    End If
ElseIf Button.KEY = "smn1" Then
    If Button.Value = tbrPressed Then
        Text2.Visible = True
        List1.Tag = "Salesman"
        V_SMN_LST
        Text2.BackColor = vbBlue
        Text2.ForeColor = vbWhite
    Else
        Text2 = ""
        Text2.Visible = False
    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
If check("prt", "prt_nme", "M_FRM_SAL_RAT") = False Then
Check3.Value = 0
Else
Check3.Value = 1
End If
Check3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Private Sub M_INV_IDY_KeyDown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
If Len(Trim(M_INV_IDY)) = 0 Then
    V_INV_LST
Else
    inv = M_INV_IDY
    TXT1 = text1
    Txt2 = Text2
    clrctr M_FRM_SAL_RAT
    text1 = TXT1
    Text2 = Txt2
    M_CST_NME = ""
    M_CST_AD1 = ""
    M_CST_AD2 = ""
    M_CST_AD3 = ""
    M_CST_AD4 = ""
    M_DBT_MCL = "0.00"
    M_DBT_MCP = "0"
    Label19 = "0.00"
    Label20 = "0.00"
    M_GRS_AMT = "0.00"
    M_NET_AMT = "0.00"
    M_SMN_NME = ""
    M_ACT_NME = ""
    M_INV_IDY = inv
    dispdetails
End If
End If
End Sub

Public Sub dispdetails()
On Error Resume Next
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Len(Trim(M_INV_IDY)) <> 10 Then
REC_SET.Open "select * from inl where inv_idy like 'SAL" & Tincrement(Trim(M_INV_IDY)) & "' and txn_typ like 'SAL' and ord_idy like '1'", con, adOpenKeyset
Else
REC_SET.Open "select * from inl where inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SAL' and ord_idy like '1'", con, adOpenKeyset
End If
If REC_SET.RecordCount > 0 Then
    If Len(Trim(M_INV_IDY)) <> 10 Then
    M_INV_IDY = "SAL" & Tincrement(Trim(M_INV_IDY))
    End If
    M_INV_IDY.Enabled = False
    MOD_REC = True
    M_INV_DAT = REC_SET("inv_dat")
    M_CST_IDY = Trim(REC_SET("cst_idy") & "")
    dlr_idy = M_CST_IDY
    If Len(Trim(M_CST_IDY)) = 0 Then Exit Sub
    M_ACT_IDY = Trim(REC_SET("act_idy") & "")
    M_SMN_IDY = Trim(REC_SET("smn_idy") & "")
    M_DCH_IDY = Trim(REC_SET("dch_idy") & "")
    M_DCH_DAT = REC_SET("dch_dat")
    M_FRT_TYP = discount(REC_SET("exd_pcg") & "")
    M_EXD_AMT = convert(amt(REC_SET("EXD_AMT")) & "")
    M_CRT_DAY = Val(REC_SET("crt_day"))
    M_ROU_OFF = Val(REC_SET("rou_off") & "")
    If Val(REC_SET("rou_off") & "") = 0 Then M_ROU_OFF = "0.00"
    M_SPD_PCG = discount(Val(REC_SET("spd_pcg") & ""))
    M_SPD_AMT = convert(amt(Val(REC_SET("spd_amt") & "")))
    M_CSD_PCG = discount(Val(REC_SET("csd_pcg") & ""))
    M_CSD_AMT = convert(amt(Val(REC_SET("csd_amt") & "")))
    M_TAX_PCG = discount(Val(REC_SET("tax_pcg") & ""))
    M_TAX_AMT = convert(amt(Val(REC_SET("tax_amt") & "")))
    M_WBL_IDY = Trim(REC_SET("WBL_IDY") & "")
    M_FRT_AMT = convert(amt(Val(REC_SET("frt_amt") & "")))
    M_WBL_IDY = Trim(REC_SET("wbl_idy") & "")
    M_WBL_DAT = REC_SET("wbl_dat")
    M_PKG_CST = convert(amt(REC_SET("PKG_CSt") & ""))
    M_TTL_CAS = Trim(REC_SET("ttl_cas") & "")
    M_TXN_RMK = Trim(REC_SET("txn_rmk") & "")
    M_TXN_TPT = Trim(REC_SET("txn_tpt") & "")
    M_LRN_IDY = Trim(REC_SET("lrn_idy") & "")
    M_LRN_AMT = convert(amt(REC_SET("lrn_amt") & ""))
    M_LRN_DAT = Trim(REC_SET("lrn_dat") & "")
    M_ROU_IDY = Trim(REC_SET("rou_idy") & "")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from smn where smn_idy like '" & Trim(M_SMN_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then M_SMN_NME = REC_SET("smn_nme")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Trim(M_ACT_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then M_ACT_NME = REC_SET("act_nme")
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from ACT where ACT_idy like '" & Trim(M_CST_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_CST_NME = Trim(REC_SET("ACT_nme") & "")
        M_CST_AD1 = Trim(REC_SET("ACT_ad1") & "")
        M_CST_AD2 = Trim(REC_SET("ACT_ad2") & "")
        M_CST_AD3 = Trim(REC_SET("ACT_ad3") & "")
        M_CST_AD4 = Trim(REC_SET("ACT_ad4") & "")
        M_DBT_MCL = amt(Trim(REC_SET("dbt_mcl")) & "")
        M_DBT_MCP = Trim(REC_SET("dbt_mcp") & "")
    End If
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from inv where inv_idy like '" & Sincrement(M_INV_IDY) & "' and txn_typ = 'SAL'", con, adOpenKeyset, adLockOptimistic
    MS.Rows = REC_SET.RecordCount + 1
    For A = 1 To MS.Rows - 1
        MS.TextMatrix(A, 0) = Trim(REC_SET("prd_idy"))
        MS.TextMatrix(A, 3) = QTY(Val(REC_SET("prd_qty")))
        MS.TextMatrix(A, 4) = convert(amt(REC_SET("prd_rat") & ""))
        MS.TextMatrix(A, 6) = Trim(REC_SET("rat_app") & "")
        MS.TextMatrix(A, 7) = discount(REC_SET("add_dis") & "")
        MS.TextMatrix(A, 8) = discount(REC_SET("trd_dis") & "")
        MS.TextMatrix(A, 9) = discount(REC_SET("trd_tax") & "")
        If REC_SET2.State = 1 Then REC_SET2.Close
        REC_SET2.Open "SELECT PRD_PRN,prd_nme,CON_FAC FROM PRD WHERE PRD_IDY LIKE '" & REC_SET("PRD_IDY") & "'", con, adOpenStatic
        If REC_SET2.RecordCount > 0 Then
            MS.TextMatrix(A, 11) = Trim(REC_SET2(1))
            MS.TextMatrix(A, 12) = Trim(REC_SET2(0))
            MS.TextMatrix(A, 2) = Val(REC_SET2("CON_FAC") & "")
        End If
        REC_SET.MoveNext
    Next
    P_DIS_CAL
    GRS_CAL
    M_GRS_AMT = M_PRD_AMT
    SSTab1.Tab = 0
    M_INV_DAT.SetFocus
Else
    V_INV_LST
End If
    If MS.Rows = 1 Then MS.Rows = 2
End Sub

Private Sub List2_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list2.ListItems.Count > 0 Then
        M_ACT_IDY = list2.SelectedItem.ListSubItems.Item(1)
        M_ACT_NME = list2.SelectedItem
        SSTab1.Tab = 1
        MS.SetFocus
        list2.Visible = False
    End If
End If
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list3.ListItems.Count > 0 Then
        M_CST_IDY = list3.SelectedItem.ListSubItems(1)
        M_CST_IDY_keydown 13, 0
        list3.Visible = False
    End If
End If
End Sub

Private Sub list4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list4.ListItems.Count > 0 Then
        list4.Visible = False
        If list4.ListItems.Count = 0 Then
            Option3.Visible = False
            Exit Sub
        End If
        Set ls = list4.SelectedItem
        P_DIS_PRD ls.SubItems(4)
    End If
End If
End Sub

Private Sub M_SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Salesmen"
End Sub

Private Sub M_SMN_IDY_KeyPress(k As Integer)
If k = 39 Or k = 41 Then k = 0
If k > 96 And k < 123 Then k = k - 32
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
    M_CRT_DAY.SetFocus
End If
If List1.Visible = True Then List1.Visible = False
If list2.Visible = True Then list2.Visible = False
If list3.Visible = True Then list3.Visible = False
If list4.Visible = True Then list4.Visible = False
If list5.Visible = True Then list5.Visible = False

If SSTab1.Tab <> 1 Then
    tax.Visible = False
    units.Visible = False
    M_PRD_AMT.Visible = False
Else
    tax.Visible = True
    units.Visible = True
    M_PRD_AMT.Visible = True
End If
Text3.Visible = False
M_FLX_AMT.Visible = False
If SSTab1.Tab = 3 Then
    M_LRN_IDY.SetFocus
End If
If SSTab1.Tab = 4 Then
Toolbar1.Enabled = False
Else
Toolbar1.Enabled = True
End If
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTab1.Tab = 1 Then MS.SetFocus
    If SSTab1.Tab = 2 Then M_CRT_DAY.SetFocus
    If SSTab1.Tab = 0 Then M_INV_DAT.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(k As Integer)
If k = 39 Or k = 41 Then k = 0
If k > 96 And k < 123 Then k = k - 32
End Sub

Private Sub M_SMN_IDY_keydown(k As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If k = 13 Or k = 40 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from smn where smn_idy like '" & Sincrement(Trim(M_SMN_IDY)) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_SMN_IDY = Sincrement(M_SMN_IDY)
        M_SMN_NME = Trim(REC_SET("smn_nme"))
        M_ACT_IDY.SetFocus
    Else
        V_SMN_LST
    End If
ElseIf k = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from smn where smn_idy like '" & Trim(M_SMN_IDY) & "'", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_CST_IDY.SetFocus
    Else
        V_SMN_LST
    End If
End If
End Sub
Private Sub M_ACT_IDY_keydown(k As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If k = 13 Or k = 40 Or k = 38 Then
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_ACT_IDY)) & "' AND GRP_CLF LIKE 'Sale%' AND ACT_TYP LIKE ''", con, adOpenKeyset
    If REC_SET.RecordCount > 0 Then
        M_ACT_IDY = Sincrement(M_ACT_IDY)
        If k = 38 Then
            M_SMN_IDY.SetFocus
        Else
        M_ACT_NME = Trim(REC_SET("act_nme"))
        SSTab1.Tab = 1
        MS.SetFocus
        End If
    Else
        V_ACT_LST
    End If
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_LeaveCell()
If Text3.Visible = True Then Text3.Visible = False
If M_FLX_AMT.Visible = True Then M_FLX_AMT.Visible = False
If List6.Visible = True Then List6.Visible = False
End Sub

Private Sub MS_KeyPress(k As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If k > 96 And k < 123 Then
k = k - 32
ElseIf k > 47 And k < 58 Or k = 46 Then

End If
If k = 39 Or k = 41 Then k = 0
If MS.Row <> 0 Then
    If MS.Col = 0 Then
        If k = 13 Then
            If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                MS.Col = 3
                MS.SetFocus
            Else
                If REC_SET.State = 1 Then REC_SET.Close
                REC_SET.Open "SELECT * FROM prd WHERE STK_TYP IN('Free','Salable')", con, adOpenKeyset
                If REC_SET.RecordCount > 0 Then
                    list4.ListItems.Clear
                    For A = 1 To REC_SET.RecordCount
                        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
                        ls.ListSubItems.Add , , Trim(REC_SET("STK_TYP") & "")
                        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
                        If Val(REC_SET("cur_lev")) <> 0 Then
                            ls.ListSubItems.Add , , Trim(QTY(REC_SET("cur_lev") & ""))
                        Else
                            ls.ListSubItems.Add , , "0.000"
                        End If
                        ls.ListSubItems.Add , , Trim(REC_SET("PRD_IDY") & "")
                        ls.ListSubItems.Add , , Val(REC_SET("sub_unt") & "")
                        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("pur_rat") & "")))
                        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("sal_rat") & "")))
                        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("mrp_rat") & "")))
                        REC_SET.MoveNext
                    Next
                    SSTab1.Tab = 4
                    If check("prt", "prt_nme", "prdnum") = True Then Check2.Value = 1
                    V_LST_CAP = "Select Product"
                    Check2.Caption = "Short Name"
                    list4.Visible = True
                    list4.SetFocus
                End If
            End If
        ElseIf k <> 27 Then
            Text3.Visible = True
            Text3.Top = MS.Top + MS.CellTop
            Text3.Left = MS.CellLeft + MS.Left
            Text3.Width = MS.CellWidth
            Text3.Height = MS.CellHeight
            Text3.SetFocus
            If k <> 13 Then Text3 = Chr(k)
            Text3.SelStart = Len(Text3)
        End If
    ElseIf MS.Col = 3 Then
        If MS.TextMatrix(MS.Row, 0) <> "" Then
           If k > 47 And k < 58 Or k = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.MaxLength = 9
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(k)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            ElseIf k = 13 Then
                If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                    MS.AddItem ""
                    MS.Col = 0
                    MS.Row = MS.Rows - 1
                End If
            Else
                k = 0
            End If
        End If
    ElseIf MS.Col = 5 Then
        If k > 47 And k < 58 Or k = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.MaxLength = 9
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(k)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
        ElseIf k = 13 Then
                If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                    MS.AddItem ""
                    MS.Col = 0
                    MS.Row = MS.Rows - 1
                End If
        Else
                k = 0
        End If
    ElseIf MS.Col = 7 Or MS.Col = 8 Or MS.Col = 9 Then
         If k > 47 And k < 58 Or k = 46 Then
                M_FLX_AMT = ""
                M_FLX_AMT.MaxLength = 9
                M_FLX_AMT.Visible = True
                M_FLX_AMT.Top = MS.Top + MS.CellTop
                M_FLX_AMT.Left = MS.CellLeft + MS.Left
                M_FLX_AMT.Width = MS.CellWidth
                M_FLX_AMT.Height = MS.CellHeight
                M_FLX_AMT.SetFocus
                M_FLX_AMT = Chr(k)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            End If
    ElseIf MS.Col = 10 Or MS.Col = 11 Then
        If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
            MS.AddItem ""
            MS.Col = 0
            MS.Row = MS.Rows - 1
        End If
    End If
End If
End Sub

Private Sub Text3_KeyDown(k As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If k = 39 Or k = 41 Then k = 0
If k = 13 Then
    If MS.Col = 0 Then
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM prd WHERE prd_IDY LIKE '" & Sincrement(Text3) & "' AND STK_TYP IN('Free','Salable')", con, adOpenKeyset
        If REC_SET.RecordCount > 0 Then
            P_DIS_PRD Sincrement(Text3)
        Else
            V_PRD_LST
       End If
       End If
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(k As Integer)
On Error Resume Next
Dim REC_SET As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET3 As New ADODB.Recordset
If k > 47 And k < 58 Or k = 46 Or k = 45 Or k = 8 Then

ElseIf k = 13 Then
    If MS.Col = 3 Then
        M_FLX_AMT = QTY(M_FLX_AMT)
        If REC_SET.State = 1 Then REC_SET.Close
        If REC_SET1.State = 1 Then REC_SET1.Close
        If REC_SET2.State = 1 Then REC_SET2.Close
        If REC_SET3.State = 1 Then REC_SET3.Close
        REC_SET.Open "select * from PRD where PRD_IDY like '" & Trim(MS.TextMatrix(MS.Row, 0)) & "'", con, adOpenKeyset, adLockOptimistic
        If Trim(M_INV_IDY) <> "" Then
            REC_SET2.Open "select prd_qty from inv where prd_idy like '" & Trim(MS.TextMatrix(MS.Row, 0)) & "' and inv_idy like '" & Trim(M_INV_IDY) & "' and txn_typ like 'SAL'", con, adOpenKeyset, adLockOptimistic
            If REC_SET2.RecordCount > 0 Then INVTOT = QTY(REC_SET2(0))
        End If
        B = InStr(M_FLX_AMT, ".")
        If B <> 0 Then M = Mid(M_FLX_AMT, B + 1, Len(M_FLX_AMT) - B + 1)
        If Len(M) = 1 Then
            M_FLX_AMT = M_FLX_AMT & "00"
        ElseIf Len(M) = 2 Then
            M_FLX_AMT = M_FLX_AMT & "0"
        ElseIf Len(M) = 0 Then
            M_FLX_AMT = M_FLX_AMT & ".000"
        End If
        If B > 0 And Val(REC_SET("sub_unt")) > 0 Then
            If InStr(B + 1, M_FLX_AMT, ".") > 0 Then
                MS.TextMatrix(MS.Row, 3) = "0.000"
                Exit Sub
            End If
            M = Mid(M_FLX_AMT, B + 1, 3)
            If M >= Val(REC_SET("sub_unt")) Then
                MsgBox "Product Sub Units Are " & REC_SET("sub_unt"), vbExclamation, "Effmcg"
                If M_FLX_AMT.Visible = True Then
                    M_FLX_AMT.SetFocus
                Else
                    MS.SetFocus
                End If
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                M_QTY = M_FLX_AMT
            End If
        Else
            M_QTY = M_FLX_AMT
        End If
        If REC_SET.RecordCount > 0 Then
            REMIND = ADDQTY(INVTOT, MS.TextMatrix(MS.Row, 0))
            If CONPRDQTY(REMIND, M_QTY, MS.TextMatrix(MS.Row, 0)) < 0 Then
                MsgBox "Stock Out Of Range, Check Stock !", vbExclamation, "Effmcg"
                If M_FLX_AMT.Visible = True Then
                    M_FLX_AMT.SetFocus
                Else
                    MS.SetFocus
                End If
                SendKeys "{home}+{end}"
                Exit Sub
             End If
            MS.Text = M_QTY
            M_FLX_AMT.Visible = False
            If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                MS.AddItem ""
                MS.Col = 0
                MS.Row = MS.Rows - 1
            End If
    End If
    MS.SetFocus
    P_DIS_CAL
ElseIf MS.Col = 4 Then
    MS.TextMatrix(MS.Row, 4) = convert(amt(Val(M_FLX_AMT)))
    P_DIS_CAL
    M_FLX_AMT.Visible = False
    MS.SetFocus
    MS.Col = 5
ElseIf MS.Col = 5 Then
    M_FLX_AMT = UNTCAL(M_FLX_AMT, MS.TextMatrix(MS.Row, 0))
    MS.Col = 3
    M_FLX_AMT_KeyPress 13
ElseIf MS.Col = 7 Or MS.Col = 8 Or MS.Col = 9 Then
    M_FLX_AMT.Visible = False
    MS.Text = discount(M_FLX_AMT)
    MS.Col = MS.Col + 1
    P_DIS_CAL
    MS.SetFocus
Else
    k = 0
End If
Else
k = 0
End If
End Sub

Public Sub P_DIS_PRD(prdid As String)
MSROW = MS.Row
Dim REC_SET1 As New ADODB.Recordset
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
    MS.TextMatrix(MS.Row, 2) = QTY(Val(REC_SET("CON_FAC") & ""))
    MS.TextMatrix(MS.Row, 3) = "0.000"
    If REC_SET1.State = 1 Then REC_SET1.Close
    REC_SET1.Open "select * from csr where act_idy like '" & Trim(M_CST_IDY) & "' and prd_idy like '" & prdid & "'", con, adOpenKeyset, adLockOptimistic
    If REC_SET1.RecordCount > 0 Then
        MS.TextMatrix(MS.Row, 4) = convert(amt(REC_SET1("prd_rat") & ""))
    Else
        MS.TextMatrix(MS.Row, 4) = convert(amt(REC_SET("SAL_rat") & ""))
    End If
    MS.TextMatrix(MS.Row, 6) = M_DBT_RAP
    MS.TextMatrix(MS.Row, 7) = discount(REC_SET("add_dis") & "")
    MS.TextMatrix(MS.Row, 8) = discount(REC_SET("cas_dis") & "")
    MS.TextMatrix(MS.Row, 9) = discount(REC_SET("EXC_ADD") & "")
    MS.TextMatrix(MS.Row, 12) = Trim(REC_SET("prd_prn"))
    MS.TextMatrix(MS.Row, 11) = Trim(REC_SET("prd_nme"))
    Text3.Visible = False
    SSTab1.Tab = 1
    MS.Row = MSROW
    MS.Col = 3
    MS.SetFocus
End If
End Sub

Public Sub P_DIS_CAL()
units = 0
taxamt = 0
tax = 0
M_PRD_AMT = 0
For A = 1 To MS.Rows - 1
    'Customer Discount calculating on units and rate
    Dim REC_SET As New ADODB.Recordset
    If REC_SET.State = 1 Then REC_SET.Close
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
    REC_SET.Open "select sub_unt from prd where prd_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'", con, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then
        If Val(REC_SET(0)) > 0 Then
            MS.TextMatrix(A, 10) = RATCAL(Val(MS.TextMatrix(A, 3) & ""), Val(MS.TextMatrix(A, 4) & ""), REC_SET(0))
        Else
            MS.TextMatrix(A, 10) = convert(amt(Round(Val(MS.TextMatrix(A, 3) & "") * Val(MS.TextMatrix(A, 4) & ""), 2)))
        End If
    End If
    MS.TextMatrix(A, 10) = Val(MS.TextMatrix(A, 10)) - Round(Val(Val(MS.TextMatrix(A, 10)) * Val(MS.TextMatrix(A, 7)) / 100), 2)
    MS.TextMatrix(A, 10) = Val(MS.TextMatrix(A, 10)) - Round(Val(Val(MS.TextMatrix(A, 10)) * Val(MS.TextMatrix(A, 8)) / 100), 2)
    taxamt = Val(taxamt) + Round(Val(Val(MS.TextMatrix(A, 10)) * Val(MS.TextMatrix(A, 9)) / 100), 2)
    MS.TextMatrix(A, 10) = convert(amt(Val(MS.TextMatrix(A, 10)) + Round(Val(Val(MS.TextMatrix(A, 10)) * Val(MS.TextMatrix(A, 9)) / 100), 2)))
    MS.TextMatrix(A, 5) = QTY(WGTCAL(MS.TextMatrix(A, 3), MS.TextMatrix(A, 2), REC_SET(0)))
    tax = QTY(Val(tax) + Val(MS.TextMatrix(A, 5)))
    units = QTY(Val(units) + Val(MS.TextMatrix(A, 3)))
    M_PRD_AMT = convert(amt(Val(M_PRD_AMT) + MS.TextMatrix(A, 10)))
    End If
Next
M_GRS_AMT = convert(amt(M_PRD_AMT))
M_PRD_AMT = convert(amt(M_PRD_AMT))
M_NET_AMT = convert(amt(Val(M_GRS_AMT)))
GRS_CAL
End Sub

Public Sub V_PRD_LST()
If check("prt", "prt_nme", "prdnum") = True Then Check2.Value = 1
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Check1.Value = 1 Then
REC_SET.Open "SELECT * FROM prd WHERE PRD_NME LIKE '" & Trim(Text3) & "%' AND STK_TYP IN('Free','Salable')", con, adOpenKeyset
Else
REC_SET.Open "SELECT * FROM prd WHERE PRD_PRN LIKE '" & Trim(Text3) & "%' AND STK_TYP IN('Free','Salable')", con, adOpenKeyset
End If
list4.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("prd_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("STK_TYP") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("prd_prn") & "")
        If Val(REC_SET("cur_lev")) <> 0 Then
            ls.ListSubItems.Add , , Trim(QTY(REC_SET("cur_lev") & ""))
        Else
            ls.ListSubItems.Add , , "0.000"
        End If
        ls.ListSubItems.Add , , Trim(REC_SET("PRD_IDY") & "")
        ls.ListSubItems.Add , , Val(REC_SET("sub_unt") & "")
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("pur_rat") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("sal_rat") & "")))
        ls.ListSubItems.Add , , convert(amt(Val(REC_SET("mrp_rat") & "")))
        REC_SET.MoveNext
    Next
    V_LST_CAP = "Select Product"
    SSTab1.Tab = 4
    Check2.Caption = "Short Name"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 1 Then
    MsgBox "Product Not Found, Define Product !", vbExclamation, "Effmcg"
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

Public Sub PRINTRECORD()
If check("PRT", "PRT_NME", "M_FRM_SAL_RAT") = True Then
    M_FRM_PRV.Check3.Value = 1
Else
    M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "SLR"
M_FRM_PRV.Show 1
End Sub

Public Sub setval()
TXT1 = text1
Txt2 = Text2
clrctr M_FRM_SAL_RAT
text1 = TXT1
Text2 = Txt2
M_ROU_IDY.Selected(0) = True
M_CST_NME = ""
M_CST_AD1 = ""
M_CST_AD2 = ""
M_CST_AD3 = ""
M_CST_AD4 = ""
M_DBT_MCL = "0.00"
M_DBT_MCP = "0"
Label19 = "0.00"
Label20 = "0.00"
M_GRS_AMT = "0.00"
M_NET_AMT = "0.00"
M_ROU_OFF = "0.00"
M_BEF_ROF = "0.00"
M_FRT_TYP = "0.00"
M_SMN_NME = ""
M_ACT_NME = ""
M_SPD_DED = "0.00"
cd = "0.00"
units = "0.00"
tax = "0.00"
M_PRD_AMT = "0.00"
M_INV_DAT = Date
M_DCH_DAT = Date
M_CSD_PCG = "0.00"
M_CSD_AMT = "0.00"
M_SPD_PCG = "0.00"
M_SPD_AMT = "0.00"
M_TAX_PCG = "0.00"
M_TAX_AMT = "0.00"
M_EXD_AMT = "0.00"
M_FRT_AMT = "0.00"
M_PKG_CST = "0.00"
M_CRT_DAY = "0"
End Sub

Public Sub ROU_OFF()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select rou_off from inl where inv_idy like '" & Trim(M_INV_IDY) & "' AND TXN_TYP LIKE 'SAL'", con, adOpenKeyset, adLockOptimistic
If REC_SET.RecordCount > 0 Then If (Val(REC_SET(0) & "") = 0 Or Val(M_ROU_OFF) = 0) And MOD_REC = True Then Exit Sub
dot = 0
rof = 0
dot = InStr(Val(M_NET_AMT), ".")
If dot > 0 Then
    rof = Mid(Val(M_NET_AMT), dot + 1, 2)
    If Len(rof) = 1 Then rof = rof & "0"
    If rof > 0 And rof < 50 Then
            M_ROU_OFF = "-0." & rof
        M_NET_AMT = Mid(Val(M_NET_AMT), 1, dot) & "00"
    ElseIf rof > 0 And rof >= 50 Then
        rof = 100 - rof
        If Len(rof) = 1 Then
            M_ROU_OFF = "+0.0" & rof
        ElseIf Len(rof) = 2 Then
            M_ROU_OFF = "+0." & rof
        End If
        M_NET_AMT = Mid(Val(M_NET_AMT), 1, dot - 1)
        M_NET_AMT = Val(M_NET_AMT) + 1 & ".00"
    End If
Else
    M_ROU_OFF = "0.00"
    M_NET_AMT = convert(amt(Val(M_NET_AMT)))
End If
End Sub

Private Sub M_TXN_TPT_GotFocus()
    M_MDI.STS_BAR.Panels(1).Text = "Enter Cases"
    SendKeys "{HOME}+{END}"
End Sub

Private Sub M_TXN_TPT_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    M_TXN_RMK.SetFocus
ElseIf KeyCode = 38 Then
    M_TTL_CAS.SetFocus
End If
End Sub

Private Sub M_TXN_TPT_KeyPress(k As Integer)
    If k > 96 And k < 123 Then k = k - 32
End Sub
Private Sub M_TXN_RMK_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    saverecord
ElseIf KeyCode = 38 Then
    M_TXN_TPT.SetFocus
End If
End Sub

Private Sub M_TXN_RMK_KeyPress(k As Integer)
If k > 96 And k < 123 Then k = k - 32
End Sub
Private Sub M_WBL_DAT_KeyDown(k As Integer, s As Integer)
If k = 38 And s = 1 Then
    M_WBL_IDY.SetFocus
ElseIf k = 13 Then
    M_TTL_CAS.SetFocus
End If
End Sub

Private Sub M_WBL_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Way Bill Number"
SendKeys "{home}+{End}"
End Sub

Private Sub M_WBL_IDY_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    M_WBL_DAT.SetFocus
ElseIf KeyCode = 38 Then
    M_LRN_AMT.SetFocus
End If
End Sub

